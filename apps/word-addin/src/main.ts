import { installAuthFetch, isWebViewHost, bridgeExecuteAction, bridgeGetStructuredContext, bridgeGetWordContext, bridgeSimpleInsert } from "./bridge.js";

// In production (bundled), API is on the same origin → empty string.
// In development, Vite proxies "/api" to the local-agent server.
// @ts-ignore __AGENT_BASE__ is injected by Vite at build time via `define`.
const agentBase: string = typeof __AGENT_BASE__ !== "undefined" ? __AGENT_BASE__ : "/api";

// ─── Types ───────────────────────────────────────────────────────────────────

type ChatMessage = {
  role: "system" | "user" | "assistant";
  content: string;
};

type DocumentParagraph = {
  index: number;
  text: string;
  style: string;
  headingLevel?: number;
  isTable: boolean;
  isList: boolean;
  font?: {
    name?: string;
    size?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
  };
};

type DocumentStructure = {
  title: string;
  totalParagraphs: number;
  totalCharacters: number;
  paragraphs: DocumentParagraph[];
  selection: {
    text: string;
    startParagraphIndex?: number;
    endParagraphIndex?: number;
  };
};

type WordAction = {
  action: string;
  params: Record<string, any>;
  description: string;
  toolCallId?: string;
};

type ActionPlan = {
  actions: WordAction[];
  explanation: string;
};

type Checkpoint = {
  id: string;
  time: number;
  label: string;
  messages: ChatMessage[];
  undoCount: number;
  undone?: boolean;
};

type Session = {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: number;
  updatedAt: number;
  titleGenerated?: boolean;
  checkpoints?: Checkpoint[];
};

type InsertMode = "chat_only" | "smart_action" | "replace_selection" | "append_end";

type KbFileEntry = {
  fileName: string;
  chunkCount: number;
};

type SimpleInsertPlan = {
  text: string;
  mode: "replace" | "append_end" | "insert_start";
};

// ─── State ───────────────────────────────────────────────────────────────────

const STORAGE_KEY = "word-agent-sessions";

const state: {
  sessions: Session[];
  activeSessionId: string | null;
  lastReply: string;
  pendingActionPlan: ActionPlan | null;
  pendingSimpleInsert: SimpleInsertPlan | null;
  pendingSessionId: string | null;
  pendingAttachment: { fileName: string; content: string } | null;
  isAgentRunning: boolean;
  autopilot: boolean;
} = {
  sessions: [],
  activeSessionId: null,
  lastReply: "",
  pendingActionPlan: null,
  pendingSimpleInsert: null,
  pendingSessionId: null,
  pendingAttachment: null,
  isAgentRunning: false,
  autopilot: false,
};

// Module-level AbortController shared by all in-flight fetch calls in the
// agent loop. Replaced before every request; aborted by the Stop button to
// immediately tear down SSE streams and any pending /v1/chat/agent-continue
// calls. See /v1/chat/abort for the server-side companion that deletes the
// session to prevent continued iteration.
let currentAbortController: AbortController | null = null;

// ─── DOM Helpers ─────────────────────────────────────────────────────────────

const $ = <T extends HTMLElement>(id: string): T => {
  const el = document.getElementById(id) as T | null;
  if (!el) {
    throw new Error(`Missing element: ${id}`);
  }
  return el;
};

const configStatus = $<HTMLParagraphElement>("configStatus");
const kbStatus = $<HTMLParagraphElement>("kbStatus");
const kbMeta = $<HTMLParagraphElement>("kbMeta");
const kbFileList = $<HTMLDivElement>("kbFileList");
const chatStatus = $<HTMLParagraphElement>("chatStatus");
const chatLog = $<HTMLDivElement>("chatLog");
const sessionSelect = $<HTMLSelectElement>("sessionSelect");
const userInput = $<HTMLTextAreaElement>("userInput");
const insertModeSelect = $<HTMLSelectElement>("insertMode");
const actionPlanPanel = $<HTMLDivElement>("actionPlanPanel");
const actionPlanContent = $<HTMLDivElement>("actionPlanContent");
const fileInput = $<HTMLInputElement>("fileInput");
const sendBtn = $<HTMLButtonElement>("sendMsg");
const stopBtn = $<HTMLButtonElement>("stopBtn");
const attachBtnEl = $<HTMLButtonElement>("attachBtn");
const contextMeter = $<HTMLDivElement>("contextMeter");
const contextRingFg = contextMeter.querySelector<SVGCircleElement>(".context-ring-fg")!;
const contextRingText = contextMeter.querySelector<HTMLElement>(".context-ring-text")!;

// ─── Context Usage Ring ─────────────────────────────────────────────────────

/**
 * Current session context usage, updated from the `usage` SSE events the
 * server emits after each LLM turn. `promptTokens` is the full context the
 * model saw on the latest turn (system prompt + history + tool results).
 */
const contextUsage: {
  promptTokens: number | null;
  completionTokens: number;
  estimated: boolean;
  contextWindow: number | null;
} = {
  promptTokens: null,
  completionTokens: 0,
  estimated: false,
  contextWindow: null,
};

const RING_CIRCUMFERENCE = 2 * Math.PI * 15.5; // r=15.5 → ≈97.39

function formatTokensShort(n: number): string {
  if (n >= 1_000_000) return `${Math.round(n / 100_000) / 10}M`;
  if (n >= 1_000) return `${Math.round(n / 100) / 10}k`;
  return String(n);
}

/**
 * Render the context-usage ring. Hidden until the first usage event arrives.
 * Colors: brand green < 70%, amber < 90%, red ≥ 90%.
 */
function renderContextMeter(): void {
  const { promptTokens, contextWindow } = contextUsage;
  if (promptTokens === null) {
    contextMeter.style.display = "none";
    return;
  }

  contextMeter.style.display = "";
  const ratio = contextWindow && contextWindow > 0 ? Math.min(promptTokens / contextWindow, 1) : 0;
  const offset = RING_CIRCUMFERENCE * (1 - ratio);
  contextRingFg.style.strokeDasharray = String(RING_CIRCUMFERENCE);
  contextRingFg.style.strokeDashoffset = String(offset);

  contextMeter.classList.toggle("warn", ratio >= 0.7 && ratio < 0.9);
  contextMeter.classList.toggle("danger", ratio >= 0.9);

  const pct = contextWindow && contextWindow > 0 ? `${Math.round(ratio * 100)}%` : formatTokensShort(promptTokens);
  contextRingText.textContent = contextWindow && contextWindow > 0 ? pct : formatTokensShort(promptTokens);

  const detail =
    `上下文：${formatTokensShort(promptTokens)} / ${contextWindow ? formatTokensShort(contextWindow) : "未配置"} tokens` +
    `（本轮输出 ${formatTokensShort(contextUsage.completionTokens)}）` +
    (contextUsage.estimated ? " · 估算值" : "");
  contextMeter.title = detail;
}

/** Reset the ring — called on session switch / clear / new session. */
function resetContextMeter(): void {
  contextUsage.promptTokens = null;
  contextUsage.completionTokens = 0;
  contextUsage.estimated = false;
  renderContextMeter();
}

/** Apply a `usage` SSE event to the ring state. */
function applyUsageEvent(ev: { promptTokens?: number; completionTokens?: number; estimated?: boolean; contextWindow?: number }): void {
  if (ev.promptTokens !== undefined && ev.promptTokens !== null) {
    contextUsage.promptTokens = ev.promptTokens;
  }
  if (ev.completionTokens !== undefined && ev.completionTokens !== null) {
    contextUsage.completionTokens = ev.completionTokens;
  }
  contextUsage.estimated = Boolean(ev.estimated);
  if (ev.contextWindow !== undefined && ev.contextWindow !== null && ev.contextWindow > 0) {
    contextUsage.contextWindow = ev.contextWindow;
  } else if (cachedProviderConfig?.contextWindow && cachedProviderConfig.contextWindow > 0) {
    contextUsage.contextWindow = cachedProviderConfig.contextWindow;
  }
  renderContextMeter();
}

/**
 * Toggle UI between "sending" and "idle" states.
 *
 * - `sendBtn` and `stopBtn` are mutually exclusive: the stop button is only
 *   visible while an agent loop is in flight.
 * - Input and attachment controls are locked while running so the user cannot
 *   trigger a second request mid-stream.
 * - Keyboard Enter-to-send is also blocked at the same time.
 */
function setRunningState(running: boolean): void {
  state.isAgentRunning = running;
  sendBtn.style.display = running ? "none" : "";
  stopBtn.style.display = running ? "" : "none";
  userInput.disabled = running;
  attachBtnEl.disabled = running;
  if (running) {
    sendBtn.classList.add("is-running");
  } else {
    sendBtn.classList.remove("is-running");
    currentAbortController = null;
  }
}

/** Grow the composer with its content, up to the CSS max-height (168px). */
function resizeComposerInput(): void {
  userInput.style.height = "auto";
  const next = Math.min(userInput.scrollHeight, 168);
  userInput.style.height = `${Math.max(next, 42)}px`;
}
const attachBtn = $<HTMLButtonElement>("attachBtn");
const dropZone = $<HTMLDivElement>("dropZone");
const dropOverlay = $<HTMLDivElement>("dropOverlay");
const attachmentPreview = $<HTMLDivElement>("attachmentPreview");
const attachmentName = $<HTMLSpanElement>("attachmentName");
const attachmentSize = $<HTMLSpanElement>("attachmentSize");
const removeAttachment = $<HTMLButtonElement>("removeAttachment");

// ─── Attachment Helpers ─────────────────────────────────────────────────────

const TEXT_EXTENSIONS = new Set([
  ".txt", ".md", ".json", ".csv", ".xml", ".yaml", ".yml",
  ".log", ".ini", ".cfg", ".conf", ".toml",
  ".py", ".js", ".ts", ".java", ".c", ".cpp", ".h", ".hpp",
  ".go", ".rs", ".rb", ".php", ".sh", ".bat", ".ps1",
  ".sql", ".html", ".css", ".svg", ".tsx", ".jsx", ".vue", ".svelte",
]);

function isPlainTextFile(fileName: string): boolean {
  const dot = fileName.lastIndexOf(".");
  if (dot < 0) return false;
  const ext = fileName.slice(dot).toLowerCase();
  return TEXT_EXTENSIONS.has(ext);
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

async function readTextFile(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsText(file, "utf-8");
  });
}

function showAttachmentPreview(fileName: string, size: number): void {
  attachmentName.textContent = `📄 ${fileName}`;
  attachmentSize.textContent = formatFileSize(size);
  attachmentPreview.style.display = "flex";
}

function hideAttachmentPreview(): void {
  state.pendingAttachment = null;
  attachmentPreview.style.display = "none";
  attachmentName.textContent = "";
  attachmentSize.textContent = "";
}

async function handleAttachmentFile(file: File): Promise<void> {
  if (!isPlainTextFile(file.name)) {
    setStatus(chatStatus, `不支持的文件类型，请选择纯文本文件`);
    return;
  }

  try {
    const content = await readTextFile(file);
    state.pendingAttachment = { fileName: file.name, content };
    showAttachmentPreview(file.name, file.size);
    setStatus(chatStatus, `已附加 ${file.name}`);
  } catch {
    setStatus(chatStatus, "文件读取失败");
  }
}

function setupAttachmentHandlers(): void {
  // Button click
  attachBtn.addEventListener("click", () => {
    fileInput.click();
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files && fileInput.files.length > 0) {
      void handleAttachmentFile(fileInput.files[0]);
      fileInput.value = "";
    }
  });

  // Remove attachment
  removeAttachment.addEventListener("click", () => {
    hideAttachmentPreview();
    setStatus(chatStatus, "已移除附件");
  });

  // Drag and drop on the drop zone
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropOverlay.style.display = "flex";
    userInput.classList.add("drag-over");
  });

  dropZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    e.stopPropagation();
    // Only hide if actually leaving the drop zone (not entering a child)
    const related = e.relatedTarget as Node | null;
    if (!related || !dropZone.contains(related)) {
      dropOverlay.style.display = "none";
      userInput.classList.remove("drag-over");
    }
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropOverlay.style.display = "none";
    userInput.classList.remove("drag-over");

    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      void handleAttachmentFile(files[0]);
    }
  });
}

// ─── Provider Config ────────────────────────────────────────────────────────

type ProviderConfigView = {
  baseUrl: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
  firstTokenTimeout?: number;
  overallTimeout?: number;
  hasApiKey: boolean;
  enableThinking?: boolean;
  includeReasoningContent?: boolean;
  thinkingEffort?: "medium" | "high";
  thinkingFormat?: "deepseek" | "openai";
  maxIterations?: number;
  contextWindow?: number;
};

/**
 * In-memory cache of the last server-persisted provider config. Used by
 * sendMessage so it can read `maxIterations` (and other server-side fields)
 * without re-querying. Refreshed on `loadProviderConfig()` and on
 * `saveProviderConfig()` success.
 */
let cachedProviderConfig: ProviderConfigView | null = null;

function setStatus(target: HTMLParagraphElement, text: string): void {
  target.textContent = text;
}

/**
 * Detect whether an error came from a fetch() that was aborted by the
 * Stop button. Treats both the standard DOMException("AbortError") and the
 * Node-style `name === "AbortError"` sentinel as user stops.
 */
function isAbortError(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  const name = (error as { name?: string }).name;
  if (name === "AbortError") return true;
  // fetch() in some browsers rejects with a TypeError carrying an internal
  // "aborted" flag — fall back to that.
  const message = (error as { message?: string }).message ?? "";
  return /aborted|abort\(\)/i.test(message);
}

// ─── Chat Log Constants ──────────────────────────────────────────────────────

const MAX_STORED_MESSAGES = 100; // keep last 100 messages in session
const LAZY_RENDER_BATCH = 20; // initially render last 20, load 20 more on demand
const LAZY_LOAD_THRESHOLD = 40; // px from top to trigger auto-load

// ─── Chat Log Rendering ──────────────────────────────────────────────────────

/**
 * Parse content for XML tags and render as rich HTML.
 * Supports: <thinking>/<reasoning> → collapsible block, Markdown → HTML
 */
function renderContentToHTML(content: string): string {
  let html = "";
  let lastEnd = 0;
  let tagIndex = 0;

  // Collect all <thinking>/<reasoning> block positions
  const blocks: Array<{ start: number; end: number; inner: string; tag: string }> = [];
  const thinkRe = /<thinking\b[^>]*>|<reasoning\b[^>]*>/gi;
  let m: RegExpExecArray | null;
  while ((m = thinkRe.exec(content)) !== null) {
    const tagName = m[0].startsWith("<thinking") ? "thinking" : "reasoning";
    const closeTag = `</${tagName}>`;
    const closeIdx = content.indexOf(closeTag, m.index + m[0].length);
    if (closeIdx === -1) continue;
    const inner = content.slice(m.index + m[0].length, closeIdx).trim();
    blocks.push({ start: m.index, end: closeIdx + closeTag.length, inner, tag: tagName });
  }
  // Sort and dedup overlapping blocks
  blocks.sort((a, b) => a.start - b.start);
  const merged: typeof blocks = [];
  for (const b of blocks) {
    if (merged.length > 0 && b.start < merged[merged.length - 1].end) continue;
    merged.push(b);
  }

  for (const block of merged) {
    // Render text before this block with Markdown
    if (block.start > lastEnd) {
      html += renderMarkdown(content.slice(lastEnd, block.start));
    }
    const label = block.tag === "reasoning" ? "推理" : "思考";
    const emoji = block.tag === "reasoning" ? "🧠" : "💭";
    const tagId = `think-${Date.now()}-${tagIndex++}`;
    html += `<details class="think-block" id="${tagId}">`;
    html += `<summary>${emoji} ${label}过程 (${block.inner.length} 字)</summary>`;
    html += `<div class="think-content">${renderMarkdown(block.inner)}</div>`;
    html += `</details>`;
    lastEnd = block.end;
  }

  // Render remaining text after last block
  if (lastEnd < content.length) {
    html += renderMarkdown(content.slice(lastEnd));
  }

  return html;
}

/**
 * Inline Markdown for a single line/paragraph. Input is already HTML-escaped.
 * Code spans are pulled out first so their contents are not styled as bold or
 * italic by the later passes.
 */
function renderInlineMarkdown(text: string): string {
  const codeSpans: string[] = [];
  let out = text.replace(/`([^`\n]+)`/g, (_m, code: string) => {
    const token = `\u0001${codeSpans.length}\u0001`;
    codeSpans.push(`<code>${code}</code>`);
    return token;
  });

  out = out.replace(/\*\*([^*\n]+)\*\*/g, "<strong>$1</strong>");
  out = out.replace(/__([^_\n]+)__/g, "<strong>$1</strong>");
  out = out.replace(/(^|[^*\w])\*([^*\n]+)\*/g, "$1<em>$2</em>");
  out = out.replace(/(^|[^_\w])_([^_\n]+)_/g, "$1<em>$2</em>");
  out = out.replace(/~~([^~\n]+)~~/g, "<del>$1</del>");

  return out.replace(/\u0001(\d+)\u0001/g, (_m, index: string) => codeSpans[Number(index)] ?? "");
}

/**
 * Render a subset of Markdown to block-level HTML: paragraphs, headings,
 * fenced code, ordered/unordered lists, blockquotes and inline emphasis.
 *
 * The previous implementation substituted `<br><br>` for blank lines, which
 * produced large gaps around headings/lists and left bare `<li>` elements
 * outside a list. Rendering line-by-line also behaves better while an answer
 * is still streaming (an unterminated fence is flushed as code at the end).
 */
function renderMarkdown(text: string): string {
  const lines = escapeHTML(text).replace(/\r\n?/g, "\n").split("\n");
  const out: string[] = [];
  let paragraph: string[] = [];
  let listTag: "ul" | "ol" | null = null;
  let inFence = false;
  let fenceLang = "";
  let fenceLines: string[] = [];

  const flushParagraph = (): void => {
    if (paragraph.length > 0) {
      out.push(`<p>${paragraph.join("<br>")}</p>`);
      paragraph = [];
    }
  };

  const closeList = (): void => {
    if (listTag) {
      out.push(`</${listTag}>`);
      listTag = null;
    }
  };

  for (const line of lines) {
    const fence = line.match(/^```([\w+-]*)\s*$/);
    if (fence) {
      if (!inFence) {
        flushParagraph();
        closeList();
        inFence = true;
        fenceLang = fence[1];
        fenceLines = [];
      } else {
        const langClass = fenceLang ? ` class="language-${fenceLang}"` : "";
        out.push(`<pre><code${langClass}>${fenceLines.join("\n")}</code></pre>`);
        inFence = false;
        fenceLang = "";
        fenceLines = [];
      }
      continue;
    }

    if (inFence) {
      fenceLines.push(line);
      continue;
    }

    if (line.trim() === "") {
      flushParagraph();
      closeList();
      continue;
    }

    const heading = line.match(/^(#{1,4})\s+(.+)$/);
    if (heading) {
      flushParagraph();
      closeList();
      const level = heading[1].length >= 3 ? 4 : heading[1].length + 1;
      out.push(`<h${level}>${renderInlineMarkdown(heading[2])}</h${level}>`);
      continue;
    }

    // ">" was escaped to "&gt;" above.
    const quote = line.match(/^&gt;\s?(.*)$/);
    if (quote) {
      flushParagraph();
      closeList();
      out.push(`<blockquote><p>${renderInlineMarkdown(quote[1])}</p></blockquote>`);
      continue;
    }

    const bullet = line.match(/^\s*[-*+]\s+(.+)$/);
    if (bullet) {
      flushParagraph();
      if (listTag !== "ul") {
        closeList();
        out.push("<ul>");
        listTag = "ul";
      }
      out.push(`<li>${renderInlineMarkdown(bullet[1])}</li>`);
      continue;
    }

    const numbered = line.match(/^\s*\d+[.)]\s+(.+)$/);
    if (numbered) {
      flushParagraph();
      if (listTag !== "ol") {
        closeList();
        out.push("<ol>");
        listTag = "ol";
      }
      out.push(`<li>${renderInlineMarkdown(numbered[1])}</li>`);
      continue;
    }

    paragraph.push(renderInlineMarkdown(line));
  }

  if (inFence) {
    out.push(`<pre><code>${fenceLines.join("\n")}</code></pre>`);
  }
  flushParagraph();
  closeList();
  return out.join("");
}

function escapeHTML(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Builds the inner markup shared by all chat bubbles: a small meta row
 * (avatar + role) and the rich body. Keeping this in one place means the
 * streaming update path and the history replay path render identically.
 */
function messageInner(role: "user" | "assistant" | "system", bodyHtml: string): string {
  const label = role === "user" ? "你" : role === "assistant" ? "助手" : "系统";
  const avatar = role === "user" ? "你" : "✦";
  return (
    `<div class="msg-meta"><span class="msg-avatar" aria-hidden="true">${avatar}</span>` +
    `<span class="msg-role">${label}</span></div>` +
    `<div class="msg-body">${bodyHtml}</div>`
  );
}

function appendMessage(role: "user" | "assistant" | "system", content: string): void {
  const div = document.createElement("div");
  div.className = `msg ${role}`;
  div.innerHTML = messageInner(role, renderContentToHTML(content));
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function appendToolCallCard(toolName: string, params: Record<string, any>): void {
  const div = document.createElement("div");
  div.className = "msg tool-call";
  const paramStr = JSON.stringify(params, null, 0);
  const shortName = toolName.replace(/_/g, " ");
  div.innerHTML =
    `<div class="tc-header">调用 · ${escapeHtml(shortName)}</div>` +
    `<div class="tc-params">${escapeHtml(paramStr.length > 300 ? paramStr.slice(0, 300) + "…" : paramStr)}</div>`;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function appendToolResultCard(toolName: string, success: boolean, result: string): void {
  const div = document.createElement("div");
  div.className = `msg tool-result ${success ? "tr-ok" : "tr-fail"}`;
  const shortName = toolName.replace(/_/g, " ");
  const body = result.length > 400 ? result.slice(0, 400) + "…" : result;
  div.innerHTML =
    `<div class="tr-header">${success ? "完成" : "失败"} · ${escapeHtml(shortName)}</div>` +
    `<div class="tr-body">${escapeHtml(body)}</div>`;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function createAssistantMessage(initial = ""): HTMLDivElement {
  const div = document.createElement("div");
  div.className = "msg assistant";
  const body = initial
    ? renderContentToHTML(initial)
    : `<div class="typing" aria-label="正在生成"><span></span><span></span><span></span></div>`;
  div.innerHTML = messageInner("assistant", body);
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
  return div;
}

function updateAssistantMessage(el: HTMLDivElement, content: string): void {
  const body = el.querySelector<HTMLDivElement>(".msg-body");
  if (body) {
    body.innerHTML = renderContentToHTML(content);
  } else {
    el.innerHTML = messageInner("assistant", renderContentToHTML(content));
  }
  // Don't auto-scroll during streaming to avoid jitter (user may be reading above)
}

/** Drop an assistant bubble that never received content (abort/transport error). */
function removeTypingBubbles(): void {
  chatLog.querySelectorAll<HTMLDivElement>(".msg.assistant").forEach((el) => {
    if (el.querySelector(".typing")) el.remove();
  });
}

function clearChatLog(): void {
  chatLog.innerHTML = "";
}

/**
 * Render messages lazily: show last LAZY_RENDER_BATCH, add load-more button for older.
 */
function renderSessionMessages(messages: ChatMessage[]): void {
  clearChatLog();
  const visible = messages.filter((m) => m.role !== "system");

  if (visible.length === 0) return;

  const total = visible.length;
  const startFrom = Math.max(0, total - LAZY_RENDER_BATCH);

  // Insert load-older button if there are hidden messages
  if (startFrom > 0) {
    const loadBtn = document.createElement("div");
    loadBtn.className = "load-older";
    loadBtn.innerHTML = `<button id="loadOlderBtn">显示更早的 ${startFrom} 条对话</button>`;
    chatLog.appendChild(loadBtn);

    const loadOlderBtn = loadBtn.querySelector("#loadOlderBtn") as HTMLButtonElement;
    if (loadOlderBtn) {
      loadOlderBtn.addEventListener("click", () => {
        loadOlderMessages(visible, startFrom);
      });
    }
  }

  // Render the last batch
  for (let i = startFrom; i < total; i++) {
    const msg = visible[i];
    appendMessage(msg.role, msg.content);
  }
}

function loadOlderMessages(visible: ChatMessage[], currentStart: number): void {
  const loadBtn = chatLog.querySelector(".load-older");
  if (loadBtn) loadBtn.remove();

  const newStart = Math.max(0, currentStart - LAZY_RENDER_BATCH);
  const fragment = document.createDocumentFragment();

  for (let i = newStart; i < currentStart; i++) {
    const msg = visible[i];
    const div = document.createElement("div");
    div.className = `msg ${msg.role}`;
    div.innerHTML = messageInner(msg.role, renderContentToHTML(msg.content));
    fragment.appendChild(div);
  }

  // Insert remaining load-more or prepend
  if (newStart > 0) {
    const moreBtn = document.createElement("div");
    moreBtn.className = "load-older";
    moreBtn.innerHTML = `<button id="loadOlderBtn">显示更早的 ${newStart} 条对话</button>`;
    fragment.appendChild(moreBtn);
  }

  // Insert before the first msg element
  const firstMsg = chatLog.querySelector(".msg");
  if (firstMsg) {
    firstMsg.before(fragment);
  } else {
    chatLog.appendChild(fragment);
  }

  // Rebind the new load-older button
  const newLoadBtn = chatLog.querySelector("#loadOlderBtn") as HTMLButtonElement;
  if (newLoadBtn) {
    newLoadBtn.addEventListener("click", () => {
      loadOlderMessages(visible, newStart);
    });
  }
}

/**
 * Trim session messages to stay within MAX_STORED_MESSAGES.
 */
function trimSessionMessages(session: Session): void {
  if (session.messages.length > MAX_STORED_MESSAGES) {
    const excess = session.messages.length - MAX_STORED_MESSAGES;
    session.messages.splice(0, excess);
    session.updatedAt = Date.now();
  }
}

// ─── Markdown Cleanup ───────────────────────────────────────────────────────

function cleanupMarkdownForWord(input: string): string {
  let text = input;

  // Convert literal \\n (LLM-typed backslash-n) to real newlines first
  text = text.replace(/\\n/g, "\n");
  text = text.replace(/\\r\\n/g, "\r\n");

  text = text.replace(/```[\s\S]*?```/g, (block) => block.replace(/```[a-zA-Z0-9_-]*\n?|```/g, ""));
  text = text.replace(/^\s{0,3}#{1,6}\s+/gm, "");
  text = text.replace(/\*\*(.*?)\*\*/g, "$1");
  text = text.replace(/__(.*?)__/g, "$1");
  text = text.replace(/(^|[^*])\*(?!\s)([^*]+?)\*(?!\*)/g, "$1$2");
  text = text.replace(/(^|[^_])_(?!\s)([^_]+?)_(?!_)/g, "$1$2");
  text = text.replace(/`([^`]+)`/g, "$1");
  text = text.replace(/^\s*[-*+]\s+/gm, "- ");
  text = text.replace(/^\s*\d+\.\s+/gm, "");
  text = text.replace(/^>\s?/gm, "");
  text = text.replace(/\[([^\]]+)\]\(([^)]+)\)/g, "$1 ($2)");
  text = text.replace(/\n{3,}/g, "\n\n");

  return text.trim();
}

// ─── Error Parsing ───────────────────────────────────────────────────────────

async function parseErrorMessage(response: Response): Promise<string> {
  const contentType = response.headers.get("content-type") || "";

  try {
    if (contentType.includes("application/json")) {
      const payload = (await response.json()) as { error?: string; message?: string };
      return payload.error || payload.message || JSON.stringify(payload);
    }
  } catch {
    return `HTTP ${response.status}`;
  }

  const text = await response.text();
  return text || `HTTP ${response.status}`;
}

// ─── Session Management ────────────────────────────────────────────────────

function generateSessionId(): string {
  return `session-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function loadSessionsFromStorage(): void {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      state.sessions = JSON.parse(raw) as Session[];
    }
  } catch {
    state.sessions = [];
  }
}

function saveSessionsToStorage(): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.sessions));
  } catch {
    // localStorage may be full or unavailable
  }
}

function getActiveSession(): Session | null {
  return state.sessions.find((s) => s.id === state.activeSessionId) ?? null;
}

function renderSessionSelect(): void {
  sessionSelect.innerHTML = "";
  for (const session of state.sessions) {
    const option = document.createElement("option");
    option.value = session.id;
    option.textContent = session.title || `会话 ${session.id.slice(8, 16)}`;
    if (session.id === state.activeSessionId) {
      option.selected = true;
    }
    sessionSelect.appendChild(option);
  }
}

function switchToSession(sessionId: string): void {
  // Save current session
  const current = getActiveSession();
  if (current) {
    current.updatedAt = Date.now();
    trimSessionMessages(current);
  }
  saveSessionsToStorage();

  state.activeSessionId = sessionId;
  const target = getActiveSession();
  clearChatLog();
  resetContextMeter();

  if (target) {
    renderSessionMessages(target.messages);
    state.lastReply = [...target.messages].reverse().find((m) => m.role === "assistant")?.content ?? "";
  } else {
    state.lastReply = "";
  }

  renderSessionSelect();
  saveSessionsToStorage();
}

function createNewSession(): void {
  const session: Session = {
    id: generateSessionId(),
    title: "",
    messages: [],
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
  state.sessions.unshift(session);
  switchToSession(session.id);
  setStatus(chatStatus, "已创建新会话");
}

function clearCurrentSession(): void {
  const session = getActiveSession();
  if (!session) {
    setStatus(chatStatus, "没有当前会话");
    return;
  }
  session.messages = [];
  session.title = "";
  session.updatedAt = Date.now();
  clearChatLog();
  state.lastReply = "";
  resetContextMeter();
  saveSessionsToStorage();
  setStatus(chatStatus, "已清空当前会话");
}

function deleteCurrentSession(): void {
  if (state.sessions.length === 0) {
    setStatus(chatStatus, "没有可删除的会话");
    return;
  }
  const currentId = state.activeSessionId;
  state.sessions = state.sessions.filter((s) => s.id !== currentId);

  if (state.sessions.length > 0) {
    switchToSession(state.sessions[0].id);
  } else {
    createNewSession();
  }
  setStatus(chatStatus, "已删除会话");
}

function autoTitleFromFirstMessage(session: Session): void {
  if (session.title) return;
  const firstUserMsg = session.messages.find((m) => m.role === "user");
  if (firstUserMsg) {
    session.title = firstUserMsg.content.slice(0, 30).replace(/\n/g, " ");
    renderSessionSelect();
  }
}

// ─── Knowledge Base ─────────────────────────────────────────────────────────

type KbStatsResponse = {
  ok: boolean;
  chunkCount: number;
  fileCount: number;
  files: string[];
};

async function loadKbStats(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/kb/stats`, { method: "GET" });
    if (!res.ok) {
      kbMeta.textContent = "知识库状态读取失败";
      return;
    }

    const data = (await res.json()) as KbStatsResponse;
    const latest = (Array.isArray(data.files) ? data.files : []).slice(-5).join("，");
    kbMeta.textContent = `知识库：${data.fileCount} 个文件 / ${data.chunkCount} 个分块${latest ? `。最近文件：${latest}` : ""}`;
  } catch {
    kbMeta.textContent = "知识库状态读取失败（本地 Agent 不可达）";
  }
}

/** Reflect the chosen knowledge-base file in the custom picker label. */
function syncKbFileLabel(): void {
  const input = $<HTMLInputElement>("kbFile");
  const picker = input.closest(".file-picker");
  const file = input.files?.[0];
  $("kbFileLabel").textContent = file ? file.name : "选择要索引的文档";
  picker?.classList.toggle("has-file", Boolean(file));
}

async function loadKbFileList(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/kb/files`, { method: "GET" });
    if (!res.ok) {
      kbFileList.innerHTML = '<div class="kb-file-item"><span>加载文件列表失败</span></div>';
      return;
    }

    const data = (await res.json()) as { ok: boolean; files: KbFileEntry[] };
    kbFileList.innerHTML = "";

    if (data.files.length === 0) {
      kbFileList.innerHTML = '<div class="kb-file-item"><span>暂无文件</span></div>';
      return;
    }

    for (const file of data.files) {
      const item = document.createElement("div");
      item.className = "kb-file-item";

      const nameSpan = document.createElement("span");
      nameSpan.className = "kb-file-name";
      nameSpan.textContent = file.fileName;
      nameSpan.title = file.fileName;

      const countSpan = document.createElement("span");
      countSpan.className = "kb-file-count";
      countSpan.textContent = `${file.chunkCount} 块`;

      const deleteBtn = document.createElement("button");
      deleteBtn.textContent = "删除";
      deleteBtn.addEventListener("click", () => {
        void deleteKbFile(file.fileName);
      });

      item.appendChild(nameSpan);
      item.appendChild(countSpan);
      item.appendChild(deleteBtn);
      kbFileList.appendChild(item);
    }
  } catch {
    kbFileList.innerHTML = '<div class="kb-file-item"><span>无法连接本地 Agent</span></div>';
  }
}

async function deleteKbFile(fileName: string): Promise<void> {
  try {
    setStatus(kbStatus, `正在删除 ${fileName}...`);
    const res = await fetch(`${agentBase}/v1/kb/files/${encodeURIComponent(fileName)}`, {
      method: "DELETE",
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `删除失败: ${err}`);
      return;
    }

    setStatus(kbStatus, `已删除 ${fileName}`);
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `删除失败: 无法连接本地 Agent (${message})`);
  }
}

async function clearKb(): Promise<void> {
  if (!confirm("确定要清空整个知识库吗？此操作不可撤销。")) {
    return;
  }

  try {
    setStatus(kbStatus, "正在清空知识库...");
    const res = await fetch(`${agentBase}/v1/kb/clear`, { method: "DELETE" });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `清空失败: ${err}`);
      return;
    }

    setStatus(kbStatus, "知识库已清空");
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `清空失败: 无法连接本地 Agent (${message})`);
  }
}

async function exportKb(): Promise<void> {
  try {
    setStatus(kbStatus, "正在导出...");
    const res = await fetch(`${agentBase}/v1/kb/export`, { method: "GET" });
    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `导出失败: ${err}`);
      return;
    }

    const data = (await res.json()) as { ok: boolean; chunks: Array<{ id: string; fileName: string; text: string }>; exportedAt: string };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `kb-archive-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    setStatus(kbStatus, `已导出 ${data.chunks.length} 个分块`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `导出失败: 无法连接本地 Agent (${message})`);
  }
}

async function importKb(): Promise<void> {
  const input = $<HTMLInputElement>("importKbFile");
  if (!input.files || input.files.length === 0) {
    setStatus(kbStatus, "请先选择存档文件");
    return;
  }

  const mode = $<HTMLSelectElement>("importMode").value as "merge" | "replace";
  const importStatus = $<HTMLParagraphElement>("kbImportStatus");

  try {
    setStatus(importStatus, "正在导入...");
    const file = input.files[0];
    const text = await file.text();

    let data: { chunks?: Array<{ id: string; fileName: string; text: string }> };
    try {
      data = JSON.parse(text);
    } catch {
      setStatus(importStatus, "导入失败: 文件不是有效的 JSON");
      return;
    }

    if (!Array.isArray(data.chunks)) {
      setStatus(importStatus, "导入失败: 文件格式不正确，缺少 chunks 字段");
      return;
    }

    const res = await fetch(`${agentBase}/v1/kb/import`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ chunks: data.chunks, mode }),
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(importStatus, `导入失败: ${err}`);
      return;
    }

    const result = (await res.json()) as { ok: boolean; importedChunks: number; totalChunks: number; mode: string };
    const modeLabel = mode === "replace" ? "替换" : "合并";
    setStatus(importStatus, `${modeLabel}导入成功：${result.importedChunks} 个分块，总计 ${result.totalChunks} 个分块`);
    input.value = "";
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(importStatus, `导入失败: 无法连接本地 Agent (${message})`);
  }
}

async function uploadKnowledgeFile(): Promise<void> {
  const input = $<HTMLInputElement>("kbFile");
  if (!input.files || input.files.length === 0) {
    setStatus(kbStatus, "请先选择文件");
    return;
  }

  try {
    setStatus(kbStatus, "上传中...");

    const form = new FormData();
    form.append("file", input.files[0]);

    const res = await fetch(`${agentBase}/v1/kb/upload`, {
      method: "POST",
      body: form,
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `上传失败: ${err}`);
      return;
    }

    const data = (await res.json()) as { chunkCount: number; fileName: string };
    setStatus(kbStatus, `已索引 ${data.fileName}，分块 ${data.chunkCount}`);
    input.value = "";
    syncKbFileLabel();
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `上传失败: 无法连接本地 Agent (${message})`);
  }
}

// ─── Provider Config ────────────────────────────────────────────────────────

async function loadProviderConfig(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/provider/config`, {
      method: "GET",
    });

    if (!res.ok) {
      if (res.status === 404) {
        setStatus(configStatus, "尚未保存模型配置");
        return;
      }

      const err = await parseErrorMessage(res);
      setStatus(configStatus, `读取配置失败: ${err}`);
      return;
    }

    const data = (await res.json()) as ProviderConfigView;
    cachedProviderConfig = data;
    $<HTMLInputElement>("baseUrl").value = data.baseUrl || "";
    $<HTMLInputElement>("model").value = data.model || "";
    $<HTMLInputElement>("temperature").value = String(data.temperature ?? 0.2);
    $<HTMLInputElement>("maxTokens").value = String(data.maxTokens ?? 900);
    $<HTMLInputElement>("firstTokenTimeout").value = String(data.firstTokenTimeout ?? 20);
    $<HTMLInputElement>("overallTimeout").value = String(data.overallTimeout ?? 240);
    $<HTMLInputElement>("maxIterations").value = String(data.maxIterations ?? 10);
    const cwInput = $<HTMLInputElement>("contextWindow");
    cwInput.value = data.contextWindow ? String(data.contextWindow) : "";
    cwInput.placeholder = data.contextWindow ? String(data.contextWindow) : "如 65536（用于会话上下文用量显示）";

    const apiKeyInput = $<HTMLInputElement>("apiKey");
    apiKeyInput.value = "";
    apiKeyInput.placeholder = data.hasApiKey ? "已保存（留空表示不修改）" : "sk-...";

    // Thinking / reasoning options
    const enableThinking = data.enableThinking ?? false;
    $<HTMLInputElement>("enableThinking").checked = enableThinking;
    $<HTMLInputElement>("includeReasoningContent").checked = data.includeReasoningContent ?? true;
    $<HTMLSelectElement>("thinkingEffort").value = data.thinkingEffort ?? "high";
    $<HTMLSelectElement>("thinkingFormat").value = data.thinkingFormat ?? "deepseek";
    updateThinkingOptionsVisibility(enableThinking);

    setStatus(configStatus, data.hasApiKey ? "已加载已保存配置" : "已加载配置（尚未保存 API Key）");

    // 加载配置后自动刷新模型列表
    void refreshModelList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `读取配置失败: 无法连接本地 Agent (${message})`);
  }
}

async function refreshModelList(): Promise<void> {
  const modelInput = $<HTMLInputElement>("model");
  const datalist = $<HTMLDataListElement>("modelList");
  const currentModel = modelInput.value.trim();

  setStatus(configStatus, "正在获取模型列表...");

  try {
    // Build query string from input fields so backend can fall back
    // to these values when no saved config exists.
    const params = new URLSearchParams();
    const inputBaseUrl = $<HTMLInputElement>("baseUrl").value.trim();
    const inputApiKey = $<HTMLInputElement>("apiKey").value.trim();
    if (inputBaseUrl) params.set("baseUrl", inputBaseUrl);
    if (inputApiKey) params.set("apiKey", inputApiKey);
    const qs = params.toString();
    const url = `${agentBase}/v1/provider/models${qs ? "?" + qs : ""}`;

    const res = await fetch(url, { method: "GET" });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `获取模型列表失败: ${err}`);
      return;
    }

    const data = (await res.json()) as { ok: boolean; models: string[]; currentModel?: string };
    const models = data.models ?? [];

    // 清空并填充 datalist
    datalist.innerHTML = "";
    for (const id of models) {
      const option = document.createElement("option");
      option.value = id;
      datalist.appendChild(option);
    }

    // 恢复当前选中的模型值
    if (currentModel) {
      modelInput.value = currentModel;
    } else if (data.currentModel) {
      modelInput.value = data.currentModel;
    }

    setStatus(configStatus, `已获取 ${models.length} 个模型${models.length > 0 ? "，可从列表选择或手动输入" : ""}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `获取模型列表失败: 无法连接本地 Agent (${message})`);
  }
}

function updateThinkingOptionsVisibility(visible: boolean): void {
  const div = $("thinkingOptions") as HTMLDivElement;
  div.style.display = visible ? "" : "none";
}

async function saveProviderConfig(): Promise<void> {
  try {
    setStatus(configStatus, "保存中...");

    const apiKeyRaw = $<HTMLInputElement>("apiKey").value.trim();
    const body: Record<string, any> = {
      baseUrl: $<HTMLInputElement>("baseUrl").value.trim(),
      model: $<HTMLInputElement>("model").value.trim(),
      temperature: Number($<HTMLInputElement>("temperature").value || "0.2"),
      maxTokens: Number($<HTMLInputElement>("maxTokens").value || "900"),
      firstTokenTimeout: Number($<HTMLInputElement>("firstTokenTimeout").value || "20"),
      overallTimeout: Number($<HTMLInputElement>("overallTimeout").value || "240"),
      enableThinking: $<HTMLInputElement>("enableThinking").checked,
      includeReasoningContent: $<HTMLInputElement>("includeReasoningContent").checked,
      thinkingEffort: $<HTMLSelectElement>("thinkingEffort").value,
      thinkingFormat: $<HTMLSelectElement>("thinkingFormat").value,
      maxIterations: Number($<HTMLInputElement>("maxIterations").value || "10"),
      contextWindow: Number($<HTMLInputElement>("contextWindow").value) || undefined,
    };

    // 只在用户实际输入了 apiKey 时才发送，否则由后端保留已保存的 key
    if (apiKeyRaw) {
      body.apiKey = apiKeyRaw;
    }

    const res = await fetch(`${agentBase}/v1/provider/config`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `保存失败: ${err}`);
      return;
    }

    // Update the in-memory cache so sendMessage sees the new value
    // without a roundtrip. We don't have the saved payload verbatim, so
    // merge the typed-in values over whatever was previously cached.
    cachedProviderConfig = {
      baseUrl: body.baseUrl,
      model: body.model,
      temperature: body.temperature,
      maxTokens: body.maxTokens,
      firstTokenTimeout: body.firstTokenTimeout,
      overallTimeout: body.overallTimeout,
      hasApiKey: cachedProviderConfig?.hasApiKey ?? false,
      enableThinking: body.enableThinking,
      includeReasoningContent: body.includeReasoningContent,
      thinkingEffort: body.thinkingEffort,
      thinkingFormat: body.thinkingFormat,
      maxIterations: body.maxIterations,
      contextWindow: body.contextWindow,
    };

    setStatus(configStatus, apiKeyRaw ? "保存成功" : "保存成功（API Key 已保留）");
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `保存失败: 无法连接本地 Agent (${message})`);
  }
}

// ─── Config Presets ─────────────────────────────────────────────────────────

type PresetView = {
  id: string;
  name: string;
  config: { baseUrl: string; model: string; apiKey: string; temperature?: number; maxTokens?: number; firstTokenTimeout?: number; overallTimeout?: number };
  createdAt: number;
  updatedAt: number;
};

async function loadPresets(): Promise<void> {
  const select = $<HTMLSelectElement>("presetSelect");
  try {
    const res = await fetch(`${agentBase}/v1/presets`);
    if (!res.ok) {
      setStatus(configStatus, "加载配置方案列表失败");
      return;
    }
    const data = (await res.json()) as { ok: boolean; presets: PresetView[] };
    const presets = data.presets ?? [];

    // Preserve current selection
    const current = select.value;
    select.innerHTML = '<option value="">-- 未保存的配置 --</option>';
    for (const p of presets) {
      const opt = document.createElement("option");
      opt.value = p.id;
      opt.textContent = p.name;
      select.appendChild(opt);
    }
    if (current && presets.some((p) => p.id === current)) {
      select.value = current;
    }
  } catch {
    setStatus(configStatus, "加载配置方案列表失败（本地 Agent 不可达）");
  }
}

async function loadPresetById(id: string): Promise<void> {
  if (!id) return;
  try {
    setStatus(configStatus, "正在加载方案...");
    const res = await fetch(`${agentBase}/v1/presets/${id}/activate`, { method: "POST" });
    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `加载方案失败: ${err}`);
      return;
    }
    await loadProviderConfig();
    setStatus(configStatus, "已加载方案");
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `加载方案失败: ${message}`);
  }
}

function showPresetNameInput(): Promise<string | null> {
  return new Promise((resolve) => {
    const row = $("presetNameRow") as HTMLDivElement;
    const input = $<HTMLInputElement>("presetNameInput");
    const confirmBtn = $("presetNameConfirm") as HTMLButtonElement;
    const cancelBtn = $("presetNameCancel") as HTMLButtonElement;

    input.value = "";
    row.style.display = "";
    input.focus();

    function cleanup() {
      row.style.display = "none";
      input.removeEventListener("keydown", onKeydown);
      confirmBtn.removeEventListener("click", onConfirm);
      cancelBtn.removeEventListener("click", onCancel);
    }

    function onConfirm() {
      const val = input.value.trim();
      if (!val) {
        input.style.borderColor = "red";
        input.placeholder = "名称不能为空，请输入方案名称";
        input.focus();
        return;
      }
      cleanup();
      resolve(val);
    }

    function onCancel() {
      cleanup();
      resolve(null);
    }

    function onKeydown(e: KeyboardEvent) {
      if (e.key === "Enter") {
        e.preventDefault();
        onConfirm();
      } else if (e.key === "Escape") {
        e.preventDefault();
        onCancel();
      }
      // Clear error styling on input
      input.style.borderColor = "";
      input.placeholder = "输入方案名称，如 DeepSeek";
    }

    input.addEventListener("keydown", onKeydown);
    confirmBtn.addEventListener("click", onConfirm);
    cancelBtn.addEventListener("click", onCancel);
  });
}

async function saveCurrentAsPreset(): Promise<void> {
  try {
    // Show inline name input (prompt() is unreliable in Office Add-in iframes)
    const name = await showPresetNameInput();
    if (!name) return;

    const apiKey = $<HTMLInputElement>("apiKey").value.trim();
    if (!apiKey) {
      setStatus(configStatus, "请先填写 API Key，方案需要包含完整的 API 凭据");
      return;
    }

    setStatus(configStatus, "正在保存方案...");
    const preset = {
      id: `preset_${Date.now()}`,
      name,
      config: {
        baseUrl: $<HTMLInputElement>("baseUrl").value.trim(),
        apiKey,
        model: $<HTMLInputElement>("model").value.trim(),
        temperature: Number($<HTMLInputElement>("temperature").value || "0.2"),
        maxTokens: Number($<HTMLInputElement>("maxTokens").value || "900"),
        firstTokenTimeout: Number($<HTMLInputElement>("firstTokenTimeout").value || "20"),
        overallTimeout: Number($<HTMLInputElement>("overallTimeout").value || "240"),
      },
    };

    const res = await fetch(`${agentBase}/v1/presets`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(preset),
    });
    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `保存方案失败: ${err}`);
      return;
    }
    await loadPresets();
    $<HTMLSelectElement>("presetSelect").value = preset.id;
    setStatus(configStatus, `方案「${name}」已保存`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `保存方案失败: ${message}`);
  }
}

async function deleteSelectedPreset(): Promise<void> {
  const select = $<HTMLSelectElement>("presetSelect");
  const id = select.value;
  if (!id) {
    setStatus(configStatus, "请先选择要删除的方案");
    return;
  }
  const name = select.options[select.selectedIndex]?.textContent || id;
  if (!confirm(`确定删除方案「${name}」？`)) return;

  try {
    const res = await fetch(`${agentBase}/v1/presets/${id}`, { method: "DELETE" });
    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `删除方案失败: ${err}`);
      return;
    }
    select.value = "";
    await loadPresets();
    setStatus(configStatus, `方案「${name}」已删除`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `删除方案失败: ${message}`);
  }
}

// ─── Word Context (Structured) ───────────────────────────────────────────────

async function getStructuredContext(): Promise<DocumentStructure> {
  if (isWebViewHost()) {
    try {
      return (await bridgeGetStructuredContext()) as DocumentStructure;
    } catch (error) {
      const message = error instanceof Error ? error.message : "未知错误";
      console.warn("getStructuredContext bridge failed:", message);
      return { title: "", totalParagraphs: 0, totalCharacters: 0, paragraphs: [], selection: { text: "" } };
    }
  }

  const emptyStructure: DocumentStructure = {
    title: "",
    totalParagraphs: 0,
    totalCharacters: 0,
    paragraphs: [],
    selection: { text: "" },
  };

  try {
    if (typeof Word === "undefined") {
      return emptyStructure;
    }

    return Word.run(async (context) => {
      const body = context.document.body;
      const selection = context.document.getSelection();

      // Load paragraphs
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      // Load tables in parallel so we can mark table-cell paragraphs as isTable=true.
      const tables = body.tables;
      tables.load("items");
      await context.sync();

      const paraList: DocumentParagraph[] = [];
      const MAX_PARAGRAPHS = 80;
      const MAX_TEXT_LENGTH = 200;

      // Detect which body-paragraph indices are inside tables. Word JS exposes
      // a table's content range via `table.getRange("Content")`; the paragraphs
      // inside that range are exactly the table cells (including empty ones).
      // We map the table's first cell text back to a body paragraph index,
      // then assume table paragraphs are contiguous from there.
      const tableParagraphIndices = new Set<number>();
      const paragraphTableIndex = new Map<number, number>();
      if (tables.items.length > 0) {
        for (const table of tables.items) {
          table.load(["rowCount", "columnCount"]);
        }
        await context.sync();

        for (let tIdx = 0; tIdx < tables.items.length; tIdx++) {
          const table = tables.items[tIdx];
          try {
            const tableRange = table.getRange("Content");
            const tableParas = tableRange.paragraphs;
            tableParas.load("items");
            await context.sync();

            if (tableParas.items.length === 0) continue;
            tableParas.items[0].load("text");
            await context.sync();

            const firstText = (tableParas.items[0].text || "").trim();
            let startIdx = -1;
            for (let j = 0; j < paragraphs.items.length; j++) {
              if ((paragraphs.items[j].text || "").trim() === firstText) {
                startIdx = j;
                break;
              }
            }
            if (startIdx < 0) continue;

            for (let j = 0; j < tableParas.items.length; j++) {
              const idx = startIdx + j;
              tableParagraphIndices.add(idx);
              paragraphTableIndex.set(idx, tIdx);
            }
          } catch (e) {
            console.warn(`[getStructuredContext] failed to map table ${tIdx}:`, e);
          }
        }
      }

      for (let i = 0; i < Math.min(paragraphs.items.length, MAX_PARAGRAPHS); i++) {
        const p = paragraphs.items[i];
        p.load(["text", "style", "isListItem"]);
        p.font.load(["name", "size", "color", "bold", "italic"]);
        (p as any).inlinePictures.load("items");
      }
      await context.sync();

      for (let i = 0; i < Math.min(paragraphs.items.length, MAX_PARAGRAPHS); i++) {
        const p = paragraphs.items[i];
        const text = (p.text || "").trim();
        const imgCount = (p as any).inlinePictures.items.length;
        if (!text && imgCount === 0) continue;

        const style = (p.style || "Normal").toString();
        const headingMatch = style.match(/Heading\s*(\d)/i);
        const fontName = p.font.name || undefined;
        const fontSize = p.font.size || undefined;
        const fontColor = p.font.color || undefined;
        const fontBold = p.font.bold === true ? true : undefined;
        const fontItalic = p.font.italic === true ? true : undefined;
        const hasFont = fontName || fontSize || fontColor || fontBold || fontItalic;
        const isInTable = tableParagraphIndices.has(i);
        const tableMarker = isInTable ? ` [表格${paragraphTableIndex.get(i)}]` : "";
        paraList.push({
          index: i,
          text: (imgCount > 0 && !text ? "（图片段落，无文字）" : text).slice(0, MAX_TEXT_LENGTH) + (imgCount > 0 ? ` [📷图片×${imgCount}]` : "") + tableMarker,
          style,
          headingLevel: headingMatch ? parseInt(headingMatch[1]) : undefined,
          isTable: isInTable,
          isList: p.isListItem,
          font: hasFont ? { name: fontName, size: fontSize, color: fontColor, bold: fontBold, italic: fontItalic } : undefined,
        });
      }

      // Load selection
      selection.load("text");
      await context.sync();

      let startParagraphIndex: number | undefined;
      let endParagraphIndex: number | undefined;

      if (selection.text) {
        const selParagraphs = selection.paragraphs;
        selParagraphs.load("items");
        await context.sync();

        if (selParagraphs.items.length > 0) {
          const firstText = (selParagraphs.items[0].text || "").trim();
          const lastText = (selParagraphs.items[selParagraphs.items.length - 1].text || "").trim();

          for (const p of paraList) {
            if (startParagraphIndex === undefined) {
              if (p.text === firstText || p.text.startsWith(firstText) || firstText.startsWith(p.text)) {
                startParagraphIndex = p.index;
              }
            }
            if (endParagraphIndex === undefined) {
              if (p.text === lastText || p.text.startsWith(lastText) || lastText.startsWith(p.text)) {
                endParagraphIndex = p.index;
              }
            }
          }
        }
      }

      // Get total text length (approximate)
      body.load("text");
      await context.sync();

      return {
        title: "",
        totalParagraphs: paragraphs.items.length,
        totalCharacters: (body.text || "").length,
        paragraphs: paraList,
        selection: {
          text: (selection.text || "").slice(0, 2000),
          startParagraphIndex,
          endParagraphIndex,
        },
      };
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    console.warn("getStructuredContext failed:", message);
    setStatus(chatStatus, `读取文档结构失败: ${message}`);
    return emptyStructure;
  }
}

// Legacy function for backward compatibility
async function getWordContext(): Promise<{ documentContext: string; selection: string }> {
  if (isWebViewHost()) {
    try {
      return await bridgeGetWordContext();
    } catch (error) {
      const message = error instanceof Error ? error.message : "未知错误";
      console.warn("getWordContext bridge failed:", message);
      return { documentContext: "", selection: "" };
    }
  }

  try {
    if (typeof Word === "undefined") {
      return { documentContext: "", selection: "" };
    }

    return Word.run(async (context) => {
      const body = context.document.body;
      const selection = context.document.getSelection();
      body.load("text");
      selection.load("text");
      await context.sync();

      return {
        documentContext: (body.text || "").slice(0, 3000),
        selection: selection.text || "",
      };
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    console.warn("getWordContext failed:", message);
    setStatus(chatStatus, `读取文档上下文失败: ${message}`);
    return { documentContext: "", selection: "" };
  }
}

// ─── Perception Tools (for Agent Loop) ───────────────────────────────────────

async function readDocument(params: Record<string, any>): Promise<string> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  return Word.run(async (context) => {
    const body = context.document.body;

    switch (params.mode) {
      case "paragraph_range": {
        const start = params.paragraph_index ?? 0;
        const count = params.count ?? 5;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        const result: string[] = [];
        for (let i = start; i < Math.min(start + count, paragraphs.items.length); i++) {
          const p = paragraphs.items[i];
          p.load(["text", "style", "isListItem"]);
          p.font.load(["name", "size", "color", "bold", "italic"]);
          (p as any).inlinePictures.load("items");
        }
        await context.sync();

        for (let i = start; i < Math.min(start + count, paragraphs.items.length); i++) {
          const p = paragraphs.items[i];
          const style = (p.style || "Normal").toString();
          const headingMatch = style.match(/Heading\s*(\d)/i);
          const headingLevel = headingMatch ? parseInt(headingMatch[1]) : undefined;
          const fontParts: string[] = [];
          if (p.font.name) fontParts.push(p.font.name);
          if (p.font.size) fontParts.push(`${p.font.size}pt`);
          if (p.font.bold) fontParts.push("加粗");
          if (p.font.italic) fontParts.push("斜体");
          if (p.font.color && !["#000000", "#000000ff", "#000"].includes((p.font.color || "").toLowerCase())) fontParts.push(p.font.color);
          const fontStr = fontParts.length > 0 ? ` [${fontParts.join(" ")}]` : "";
          const imgCount = (p as any).inlinePictures.items.length;
          const imgStr = imgCount > 0 ? ` [📷图片×${imgCount}]` : "";
          const displayText = p.text || (imgCount > 0 ? "（图片段落，无文字）" : "");
          result.push(`[段落${i}] ${headingLevel ? `标题${headingLevel} ` : ""}${displayText}${imgStr}${fontStr}`);
        }

        return result.length
          ? `读取段落 ${start} 起共 ${result.length} 段（总计 ${paragraphs.items.length} 段）：\n${result.join("\n")}`
          : "未读取到段落内容";
      }

      case "heading_context": {
        if (!params.heading_text) {
          throw new Error("heading_context 模式需要提供 heading_text");
        }
        const matches = body.search(params.heading_text, { matchCase: false, matchWholeWord: false });
        matches.load("items");
        await context.sync();

        if (matches.items.length === 0) {
          return `未找到标题"${params.heading_text}"`;
        }

        // Load all matched paragraphs to find the actual heading (not TOC entries)
        const allParagraphs = body.paragraphs;
        allParagraphs.load("items");
        await context.sync();

        // Find the paragraph index for each match
        const matchInfos: Array<{ paraIndex: number; text: string; style: string; isHeading: boolean; level: number }> = [];
        for (const match of matches.items) {
          const firstPara = match.paragraphs.getFirst();
          firstPara.load(["text", "style"]);
          await context.sync();
          const styleStr = (firstPara.style || "").toString();
          const levelMatch = styleStr.match(/Heading\s*(\d)/i);
          const level = levelMatch ? parseInt(levelMatch[1]) : 0;
          const matchText = (firstPara.text || "").trim();

          // Find the paragraph index in the body
          for (let i = 0; i < allParagraphs.items.length; i++) {
            if ((allParagraphs.items[i].text || "").trim() === matchText) {
              matchInfos.push({ paraIndex: i, text: matchText, style: styleStr, isHeading: level > 0, level });
              break;
            }
          }
        }

        if (matchInfos.length === 0) {
          return `未找到标题"${params.heading_text}"`;
        }

        // Prefer the actual heading (Heading style) over TOC entries
        const headingMatch = matchInfos.find((m) => m.isHeading);
        const bestMatch = headingMatch || matchInfos[0];

        // If there are multiple matches, report them all so the LLM knows
        let multiMatchNote = "";
        if (matchInfos.length > 1) {
          const matchList = matchInfos.map((m) => `  [段落${m.paraIndex}] ${m.isHeading ? `标题${m.level}` : "普通文本"}: ${m.text}`).join("\n");
          multiMatchNote = `\n⚠️ 找到 ${matchInfos.length} 处匹配，已自动选择${headingMatch ? "实际标题" : "第一个匹配"}（段落${bestMatch.paraIndex}）。所有匹配：\n${matchList}\n如需查看其他匹配项，请使用 read_document 的 paragraph_range 模式指定具体段落序号。\n`;
        }

        const startLevel = bestMatch.level;
        let found = false;
        const result: string[] = [];
        for (let i = 0; i < allParagraphs.items.length; i++) {
          const p = allParagraphs.items[i];
          p.load(["text", "style"]);
          (p as any).inlinePictures.load("items");
          await context.sync();

          const pLevelMatch = (p.style || "").toString().match(/Heading\s*(\d)/i);
          const pLevel = pLevelMatch ? parseInt(pLevelMatch[1]) : 0;

          if (!found) {
            if (i === bestMatch.paraIndex) {
              found = true;
              const imgCount = (p as any).inlinePictures.items.length;
              const imgStr = imgCount > 0 ? ` [📷图片×${imgCount}]` : "";
              result.push(`[段落${i}] 标题${pLevel} ${p.text}${imgStr}`);
            }
            continue;
          }

          if (pLevel > 0 && pLevel <= startLevel) {
            break;
          }
          const imgCount = (p as any).inlinePictures.items.length;
          const imgStr = imgCount > 0 ? ` [📷图片×${imgCount}]` : "";
          const displayText = p.text || (imgCount > 0 ? "（图片段落，无文字）" : "");
          result.push(`[段落${i}] ${displayText}${imgStr}`);
        }

        return `标题"${params.heading_text}"及其子内容：${multiMatchNote}\n${result.join("\n")}`;
      }

      case "selection": {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        const text = (selection.text || "").slice(0, 2000);
        return text ? `当前选区内容：${text}` : "当前无选中文本";
      }

      case "cursor_surrounding": {
        const chars = params.surrounding_chars ?? 500;
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        const selText = selection.text || "";
        const start = Math.max(0, Math.floor(selText.length / 2) - chars);
        const end = Math.min(selText.length, Math.floor(selText.length / 2) + chars);
        const surrounding = selText.slice(start, end);

        return `光标周围内容：${surrounding || "（无法获取光标周围内容）"}`;
      }

      default:
        return `未知的读取模式: ${params.mode}`;
    }
  });
}

async function getSelectionInfo(): Promise<string> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const text = selection.text || "";
    const isCursorOnly = text.length === 0;

    let startParagraphIndex: number | undefined;
    let endParagraphIndex: number | undefined;

    const selParagraphs = selection.paragraphs;
    selParagraphs.load("items");
    await context.sync();

    if (selParagraphs.items.length > 0) {
      const bodyParagraphs = context.document.body.paragraphs;
      bodyParagraphs.load("items");
      await context.sync();

      const firstSelText = (selParagraphs.items[0].text || "").trim();
      const lastSelText = (selParagraphs.items[selParagraphs.items.length - 1].text || "").trim();

      for (let i = 0; i < bodyParagraphs.items.length; i++) {
        const p = bodyParagraphs.items[i];
        p.load("text");
        await context.sync();
        const pText = (p.text || "").trim();
        if (startParagraphIndex === undefined && (pText === firstSelText || pText.startsWith(firstSelText))) {
          startParagraphIndex = i;
        }
        if (endParagraphIndex === undefined && (pText === lastSelText || pText.startsWith(lastSelText))) {
          endParagraphIndex = i;
        }
      }
    }

    const info = {
      isCursorOnly,
      textLength: text.length,
      text: text.slice(0, 200),
      startParagraphIndex,
      endParagraphIndex,
      paragraphCount: selParagraphs.items.length,
    };

    return JSON.stringify(info, null, 2);
  });
}

async function getDocumentStats(): Promise<string> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    const headings: Array<{ level: number; text: string; paragraphIndex: number }> = [];
    let listCount = 0;

    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      p.load(["text", "style", "isListItem"]);
    }
    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      const style = (p.style || "Normal").toString();
      const headingMatch = style.match(/Heading\s*(\d)/i);
      if (headingMatch) {
        headings.push({ level: parseInt(headingMatch[1]), text: (p.text || "").trim().slice(0, 100), paragraphIndex: i });
      }
      if (p.isListItem) {
        listCount++;
      }
    }

    body.load("text");
    await context.sync();

    const stats = {
      totalParagraphs: paragraphs.items.length,
      totalCharacters: (body.text || "").length,
      headingCount: headings.length,
      listParagraphCount: listCount,
      headings,
    };

    return JSON.stringify(stats, null, 2);
  });
}

// ─── Table Perception Tools (Word Tables) ───────────────────────────────────

/**
 * List all top-level Word tables in the document body with their dimensions,
 * style, and approximate paragraph-index span. Does NOT read cell contents
 * — use read_table for that. The paragraph-index span helps the LLM locate
 * each table when planning edits (e.g. "insert before the table").
 *
 * Why this exists: read_document returns every table cell as a separate
 * paragraph, so the LLM cannot tell where one table ends and the next
 * begins. This tool returns a clean structural overview instead.
 */
async function getDocumentTables(): Promise<string> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  return Word.run(async (context) => {
    const body = context.document.body;

    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    const tables = body.tables;
    tables.load("items");
    await context.sync();

    if (tables.items.length === 0) {
      return "文档中没有 Word 表格。";
    }

    // Load each table's dimensions and style.  columnCount may not be a
    // loadable scalar in older Office.js hosts — fall back to values[0].length.
    for (const table of tables.items) {
      table.load(["rowCount", "columnCount", "styleBuiltIn", "values"]);
    }
    await context.sync();

    // Build a paragraph-text index to map tables → paragraph spans
    for (let i = 0; i < paragraphs.items.length; i++) {
      paragraphs.items[i].load("text");
    }
    await context.sync();
    const paragraphTexts: string[] = [];
    for (let i = 0; i < paragraphs.items.length; i++) {
      paragraphTexts.push((paragraphs.items[i].text || "").trim());
    }

    const tableInfos: Array<{
      tableIndex: number;
      rowCount: number;
      columnCount: number;
      style: string;
      firstParagraphIndex?: number;
      lastParagraphIndex?: number;
      firstCellPreview?: string;
    }> = [];

    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];

      let firstParagraphIndex: number | undefined;
      let lastParagraphIndex: number | undefined;
      let firstCellPreview: string | undefined;

      try {
        // Pull paragraphs that belong to this table via its content range
        const tableRange = table.getRange("Content");
        const tableParas = tableRange.paragraphs;
        tableParas.load("items");
        await context.sync();

        if (tableParas.items.length > 0) {
          tableParas.items[0].load("text");
          const lastIdx = tableParas.items.length - 1;
          if (lastIdx > 0) {
            tableParas.items[lastIdx].load("text");
          }
          await context.sync();

          const firstText = (tableParas.items[0].text || "").trim();
          firstCellPreview = firstText.slice(0, 60);

          // Map first paragraph → body paragraph index by text matching.
          // The table's first cell text usually appears uniquely in the
          // body paragraph sequence (cell texts are aggregated and pushed
          // into body.paragraphs by Word JS). For ambiguous matches we
          // pick the first occurrence — table paragraphs are contiguous.
          for (let j = 0; j < paragraphTexts.length; j++) {
            if (paragraphTexts[j] === firstText) {
              firstParagraphIndex = j;
              break;
            }
          }

          if (lastIdx > 0) {
            const lastText = (tableParas.items[lastIdx].text || "").trim();
            // Walk backwards to find the last occurrence (most likely the
            // table's last cell) — tables are contiguous in paragraph
            // ordering so this is the cell's position.
            for (let j = paragraphTexts.length - 1; j >= 0; j--) {
              if (paragraphTexts[j] === lastText) {
                lastParagraphIndex = j;
                break;
              }
            }
          }
        }
      } catch (e) {
        console.warn(`[getDocumentTables] failed to map paragraph indices for table ${i}:`, e);
      }

      tableInfos.push({
        tableIndex: i,
        rowCount: table.rowCount,
        columnCount: table.columnCount || (table.values?.[0]?.length ?? 0),
        style: table.styleBuiltIn || "",
        firstParagraphIndex,
        lastParagraphIndex,
        firstCellPreview,
      });
    }

    const summary = tableInfos
      .map((t) => {
        const span =
          t.firstParagraphIndex !== undefined && t.lastParagraphIndex !== undefined
            ? `（位于段落 ${t.firstParagraphIndex} ~ ${t.lastParagraphIndex}）`
            : "";
        const preview = t.firstCellPreview ? `，首单元格："${t.firstCellPreview}"` : "";
        return `  表格 ${t.tableIndex}：${t.rowCount} 行 × ${t.columnCount} 列，样式=${t.style || "(默认)"}${span}${preview}`;
      })
      .join("\n");

    return `文档中共有 ${tables.items.length} 张表格：\n${summary}\n\n--- 结构化数据（便于精确引用） ---\n${JSON.stringify(tableInfos, null, 2)}`;
  });
}

/**
 * Read the full 2D content of a specific table. Returns both a human-readable
 * row-by-row view and a JSON dump so the LLM can pick the format that fits
 * the question being asked (e.g. "summarize the table" vs. "find cell at row 3").
 */
async function readTable(params: Record<string, any>): Promise<string> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  const tableIndex = typeof params.table_index === "number" ? params.table_index : 0;

  return Word.run(async (context) => {
    const body = context.document.body;
    const tables = body.tables;
    tables.load("items");
    await context.sync();

    if (tableIndex < 0 || tableIndex >= tables.items.length) {
      throw new Error(`表格序号 ${tableIndex} 越界（文档共有 ${tables.items.length} 张表格，序号范围 0 ~ ${tables.items.length - 1}）。请先调用 get_document_tables 确认有效索引。`);
    }

    const table = tables.items[tableIndex];
    table.load(["values", "rowCount", "columnCount", "styleBuiltIn"]);
    await context.sync();

    const values: string[][] = Array.isArray(table.values) ? table.values : [];
    const rowCount = table.rowCount || values.length;
    const columnCount = table.columnCount || (values[0]?.length ?? 0);
    const style = table.styleBuiltIn || "";

    if (values.length === 0) {
      return `表格 ${tableIndex}（${rowCount} 行 × ${columnCount} 列，样式: ${style || "(默认)"}）为空，没有可读取的单元格内容。`;
    }

    const lines: string[] = [];
    lines.push(`表格 ${tableIndex}（${rowCount} 行 × ${columnCount} 列，样式: ${style || "(默认)"}）：`);

    for (let r = 0; r < values.length; r++) {
      const row = values[r] || [];
      const cells = row.map((c, cIdx) => `[列${cIdx + 1}] ${c == null ? "" : String(c)}`);
      lines.push(`[行${r + 1}] ${cells.join("  ")}`);
    }

    // Append JSON for unambiguous data extraction (e.g. when LLM needs
    // to do regex or arithmetic on cell contents).
    lines.push("");
    lines.push("--- JSON 格式（便于精确解析）---");
    lines.push(JSON.stringify({ tableIndex, rowCount, columnCount, style, values }, null, 2));

    return lines.join("\n");
  });
}

// ─── Run-Level Format Parsing ───────────────────────────────────────────────

type RunInfo = {
  text: string;
  font?: { name?: string; size?: string; bold?: boolean; italic?: boolean; color?: string };
};

function parseRunsFromHtml(html: string): RunInfo[] {
  const runs: RunInfo[] = [];

  // Match <span style="...">text</span> patterns (Word getHtml produces these)
  const spanRegex = /<span[^>]*style="([^"]*)"[^>]*>([\s\S]*?)<\/span>/gi;
  let match: RegExpExecArray | null;

  while ((match = spanRegex.exec(html)) !== null) {
    const style = match[1];
    // Strip any inner HTML tags from the text content
    const text = match[2].replace(/<[^>]+>/g, "").trim();
    if (!text) continue;

    const font: RunInfo["font"] = {};
    const nameMatch = style.match(/font-family:\s*([^;]+)/i);
    if (nameMatch) font!.name = nameMatch[1].replace(/['"]/g, "").trim();

    const sizeMatch = style.match(/font-size:\s*([^;]+)/i);
    if (sizeMatch) font!.size = sizeMatch[1].trim();

    if (/font-weight:\s*(bold|700)/i.test(style)) font!.bold = true;
    if (/font-style:\s*italic/i.test(style)) font!.italic = true;

    const colorMatch = style.match(/(?:^|[^-])color:\s*([^;]+)/i);
    if (colorMatch) font!.color = colorMatch[1].trim();

    const hasFont = font!.name || font!.size || font!.bold || font!.italic || font!.color;
    runs.push({ text, font: hasFont ? font : undefined });
  }

  // Fallback: if no <span> found, extract plain text from <p> or body
  if (runs.length === 0) {
    const plainText = html.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").trim();
    if (plainText) {
      // Split by meaningful chunks (preserve line breaks as separate runs)
      const lines = plainText.split(/\n/).filter((l) => l.trim());
      for (const line of lines) {
        runs.push({ text: line.trim() });
      }
    }
  }

  return runs;
}

async function getParagraphFormat(params: Record<string, any>): Promise<string> {
  const { paragraph_index } = params;
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    if (paragraph_index < 0 || paragraph_index >= paragraphs.items.length) {
      throw new Error(`段落序号 ${paragraph_index} 超出范围（共 ${paragraphs.items.length} 段）`);
    }

    const para = paragraphs.items[paragraph_index];
    para.load(["text", "style", "isListItem"]);
    para.font.load(["name", "size", "color", "bold", "italic"]);

    const range = para.getRange("Whole");
    const htmlResult = range.getHtml();
    await context.sync();

    const style = (para.style || "Normal").toString();
    const headingMatch = style.match(/Heading\s*(\d)/i);

    const paragraphFont: Record<string, any> = {};
    if (para.font.name) paragraphFont.name = para.font.name;
    if (para.font.size) paragraphFont.size = para.font.size;
    if (para.font.color) paragraphFont.color = para.font.color;
    if (para.font.bold === true) paragraphFont.bold = true;
    if (para.font.italic === true) paragraphFont.italic = true;

    const runs = parseRunsFromHtml(htmlResult.value);
    const hasMixedFormatting = runs.length > 1 && runs.some((r) => r.font && Object.keys(r.font).length > 0);

    const result: Record<string, any> = {
      paragraphIndex: paragraph_index,
      text: (para.text || "").slice(0, 1000),
      style,
      headingLevel: headingMatch ? parseInt(headingMatch[1]) : undefined,
      isListItem: para.isListItem,
      paragraphFont: Object.keys(paragraphFont).length > 0 ? paragraphFont : undefined,
      hasMixedFormatting,
    };

    // Only include runs detail when there's mixed formatting or multiple runs
    if (runs.length > 0) {
      result.runs = runs.map((r) => {
        const runInfo: Record<string, any> = { text: r.text.length > 200 ? r.text.slice(0, 200) + "..." : r.text };
        if (r.font) runInfo.font = r.font;
        return runInfo;
      });
    }

    return JSON.stringify(result, null, 2);
  });
}

// ─── Word Action Execution ───────────────────────────────────────────────────

async function applyFormat(
  paragraph: Word.Paragraph,
  format: string,
  context: Word.RequestContext
): Promise<void> {
  if (!format) return;

  // Word JS API style names — try English first, then Chinese fallback
  // Chinese Word uses localized style names like "标题 1"
  const styleMap: Record<string, string[]> = {
    normal: ["Normal", "正文"],
    heading1: ["Heading 1", "标题 1", "标题1"],
    heading2: ["Heading 2", "标题 2", "标题2"],
    heading3: ["Heading 3", "标题 3", "标题3"],
    bullet_list: ["List Paragraph", "列表段落"],
    numbered_list: ["List Paragraph", "列表段落"],
  };

  const candidates = styleMap[format];
  if (!candidates) return;

  for (const styleName of candidates) {
    try {
      paragraph.style = styleName;
      await context.sync();
      return; // Success — exit
    } catch {
      // This style name didn't work, try the next one
      continue;
    }
  }

  console.warn(`无法设置样式 "${format}"，所有候选名称均失败，将使用默认格式`);
}

/**
 * Normalize content: convert literal two-char "\\n" sequences to real newlines,
 * then split into lines. Also collapse excessive blank lines.
 */
function normalizeContentLines(content: string): string[] {
  // Step 1: literal "\\n" (backslash + n, as typed by LLM) → real newline
  let text = content.replace(/\\n/g, "\n");
  // Step 2: also handle literal "\\r\\n"
  text = text.replace(/\\r\\n/g, "\n");
  // Step 3: collapse 3+ consecutive newlines into 2
  text = text.replace(/\n{3,}/g, "\n\n");
  // Split and filter out trailing empty string (but keep intentional blank lines)
  const lines = text.split("\n");
  // Remove trailing empty line if any (common LLM artifact)
  if (lines.length > 1 && lines[lines.length - 1].trim() === "") {
    lines.pop();
  }
  return lines;
}

/**
 * Insert multi-line content as multiple paragraphs after a reference paragraph.
 * Returns the last inserted paragraph.
 */
async function insertMultiParagraphAfter(
  anchor: Word.Paragraph,
  content: string,
  format: string,
  context: Word.RequestContext
): Promise<Word.Paragraph> {
  const lines = normalizeContentLines(content);
  let lastPara: Word.Paragraph = anchor;
  for (let i = 0; i < lines.length; i++) {
    const newPara = lastPara.insertParagraph(lines[i] || " ", "After");
    await applyFormat(newPara, format, context);
    lastPara = newPara;
  }
  return lastPara;
}

/**
 * Insert multi-line content as multiple paragraphs at the end of the body.
 * Returns the last inserted paragraph.
 */
async function insertMultiParagraphAtEnd(
  body: Word.Body,
  content: string,
  format: string,
  context: Word.RequestContext
): Promise<Word.Paragraph> {
  const lines = normalizeContentLines(content);
  let lastPara: Word.Paragraph | null = null;
  for (const line of lines) {
    try {
      const anchor: Word.Paragraph = lastPara ?? body.paragraphs.getLast();
      const newPara: Word.Paragraph = anchor.insertParagraph(line || " ", "After");
      await applyFormat(newPara, format, context);
      lastPara = newPara;
    } catch {
      const newPara: Word.Paragraph = body.insertParagraph(line || " ", "End");
      await applyFormat(newPara, format, context);
      lastPara = newPara;
    }
  }
  return lastPara!;
}

/**
 * Insert multi-line content as multiple paragraphs at the start of the body.
 * Returns the first inserted paragraph.
 */
async function insertMultiParagraphAtStart(
  body: Word.Body,
  content: string,
  format: string,
  context: Word.RequestContext
): Promise<Word.Paragraph> {
  const lines = normalizeContentLines(content);
  // Insert lines in reverse order before the first paragraph so they end up in correct order
  const firstPara = body.paragraphs.getFirst();
  let anchor = firstPara;
  // We need to insert from last to first so order is preserved
  let firstInserted: Word.Paragraph | null = null;
  for (let i = lines.length - 1; i >= 0; i--) {
    const newPara = anchor.insertParagraph(lines[i] || " ", "Before");
    await applyFormat(newPara, format, context);
    anchor = newPara;
    firstInserted = newPara;
  }
  return firstInserted!;
}

// ─── Mermaid Rendering ──────────────────────────────────────────────────────

let mermaidLib: typeof import("mermaid").default | null = null;
let mermaidRenderSeq = 0;

async function getMermaid(): Promise<typeof import("mermaid").default> {
  if (mermaidLib) return mermaidLib;
  const mod = await import("mermaid");
  mermaidLib = mod.default;
  mermaidLib.initialize({ startOnLoad: false, securityLevel: "strict", htmlLabels: false });
  return mermaidLib;
}

/**
 * Render mermaid source → SVG → Canvas raster → PNG base64.
 * htmlLabels:false avoids foreignObject, which would taint Canvas.
 */
async function renderMermaidImage(code: string): Promise<string> {
  const mermaid = await getMermaid();
  const id = `mermaid-render-${++mermaidRenderSeq}`;
  const { svg } = await mermaid.render(id, code);

  // Parse viewBox to get the real rendering dimensions, then rasterize at
  // a target width of 800 px (DPR-aware).  The default SVG from mermaid
  // often has no explicit width/height, so img.width returns ~0.
  const vb = svg.match(/viewBox="([^"]+)"/)?.[1];
  const parts = vb?.split(/\s+/).map(Number);
  const vbW = parts && parts.length >= 3 ? parts[2] : 0;
  const vbH = parts && parts.length >= 4 ? parts[3] : 0;
  const naturalW = vbW > 0 ? vbW : 800;
  const naturalH = vbH > 0 ? vbH : 600;
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const targetW = Math.round(800 * dpr);
  const targetH = Math.round((naturalH / naturalW) * 800 * dpr);

  const blob = new Blob([svg], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new Error("mermaid SVG 栅格化失败"));
      img.src = url;
    });
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, targetW);
    canvas.height = Math.max(1, targetH);
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas 2d context 不可用");
    ctx.drawImage(img, 0, 0, targetW, targetH);
    return canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, "");
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function executeAction(action: WordAction): Promise<string> {
  if (isWebViewHost()) {
    // Desktop path: execute in the COM host via the WebView2 bridge (WP3/D3).
    if (action.action === "insert_mermaid_image") {
      const dataUrl = await renderMermaidImage(String(action.params.mermaid_code ?? ""));
      action = { ...action, params: { ...action.params, image_data_url: dataUrl } };
    }
    return bridgeExecuteAction(action);
  }

  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  return Word.run(async (context) => {
    const body = context.document.body;

    switch (action.action) {
      case "insert_after_heading": {
        const { heading_text, content, format = "normal", content_format = "text" } = action.params;
        if (!content || String(content).trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符");
        }
        const isHtml = content_format === "html";
        const matches = body.search(heading_text, { matchCase: false, matchWholeWord: false });
        matches.load("items");
        await context.sync();

        if (matches.items.length > 0) {
          const range = matches.items[0];
          const paragraph = range.paragraphs.getFirst();
          paragraph.load("text,style,isListItem");
          if (isHtml) {
            paragraph.getRange("End").insertHtml(String(content), "After");
            await context.sync();
            return `已在"${heading_text}"后插入 HTML 格式内容`;
          }
          const lines = normalizeContentLines(String(content));
          if (lines.length <= 1) {
            const newPara = paragraph.insertParagraph(String(content), "After");
            await context.sync();
            await applyFormat(newPara, format, context);
          } else {
            await insertMultiParagraphAfter(paragraph, String(content), format, context);
          }
          return `已在"${heading_text}"后插入内容（${lines.length} 段）`;
        }
        const lastPara = body.paragraphs.getLast();
        if (isHtml) {
          lastPara.getRange("End").insertHtml(String(content), "After");
          await context.sync();
          return `未找到标题"${heading_text}"，已插入 HTML 格式内容到文档末尾`;
        }
        const lines = normalizeContentLines(String(content));
        if (lines.length <= 1) {
          const newPara = lastPara.insertParagraph(String(content), "After");
          await context.sync();
          await applyFormat(newPara, format, context);
        } else {
          await insertMultiParagraphAfter(lastPara, String(content), format, context);
        }
        return `未找到标题"${heading_text}"，已插入到文档末尾（${lines.length} 段）`;
      }

      case "replace_selection": {
        const { content, format = "normal", content_format = "text" } = action.params;
        if (!content || String(content).trim().length === 0) {
          throw new Error("替换内容为空或仅包含空白字符");
        }
        const selection = context.document.getSelection();
        if (content_format === "html") {
          selection.insertHtml(String(content), "Replace");
          await context.sync();
          return "已替换选区内容（HTML 格式）";
        }
        const lines = normalizeContentLines(String(content));
        if (lines.length <= 1) {
          const para = selection.insertParagraph(String(content), "Replace");
          await context.sync();
          await applyFormat(para, format, context);
        } else {
          // Insert first line as replacement, then add remaining lines after
          const firstPara = selection.insertParagraph(lines[0] || " ", "Replace");
          await applyFormat(firstPara, format, context);
          await insertMultiParagraphAfter(firstPara, lines.slice(1).join("\n"), format, context);
        }
        return "已替换选区内容";
      }

      case "insert_at_end": {
        const { content, format = "normal", content_format = "text" } = action.params;
        if (!content || String(content).trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符");
        }
        if (content_format === "html") {
          try {
            body.paragraphs.getLast().getRange("End").insertHtml(String(content), "After");
          } catch {
            body.insertHtml(String(content), "End");
          }
          await context.sync();
          return "已追加 HTML 格式内容到文档末尾";
        }
        const lines = normalizeContentLines(String(content));
        if (lines.length <= 1) {
          let newPara: Word.Paragraph;
          try {
            const lastPara = body.paragraphs.getLast();
            newPara = lastPara.insertParagraph(String(content), "After");
          } catch {
            newPara = body.insertParagraph(String(content), "End");
          }
          await context.sync();
          await applyFormat(newPara, format, context);
        } else {
          await insertMultiParagraphAtEnd(body, String(content), format, context);
        }
        return `已追加到文档末尾（${lines.length} 段）`;
      }

      case "insert_at_start": {
        const { content, format = "normal", content_format = "text" } = action.params;
        if (!content || String(content).trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符");
        }
        if (content_format === "html") {
          try {
            body.paragraphs.getFirst().getRange("Start").insertHtml(String(content), "Before");
          } catch {
            body.insertHtml(String(content), "Start");
          }
          await context.sync();
          return "已插入 HTML 格式内容到文档开头";
        }
        const lines = normalizeContentLines(String(content));
        if (lines.length <= 1) {
          let newPara: Word.Paragraph;
          try {
            const firstPara = body.paragraphs.getFirst();
            firstPara.load("text");
            await context.sync();
            newPara = firstPara.insertParagraph(String(content), "Before");
          } catch (e) {
            newPara = body.insertParagraph(String(content), "Start");
          }
          await context.sync();
          await applyFormat(newPara, format, context);
        } else {
          await insertMultiParagraphAtStart(body, String(content), format, context);
        }
        return `已插入到文档开头（${lines.length} 段）`;
      }

      case "insert_after_paragraph": {
        const { paragraph_index, content, format = "normal", content_format = "text" } = action.params;
        if (!content || String(content).trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符");
        }
        const isHtml = content_format === "html";
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraph_index >= 0 && paragraph_index < paragraphs.items.length) {
          if (isHtml) {
            paragraphs.items[paragraph_index].getRange("End").insertHtml(String(content), "After");
            await context.sync();
            return `已在段落 ${paragraph_index} 后插入 HTML 格式内容`;
          }
          const lines = normalizeContentLines(String(content));
          if (lines.length <= 1) {
            const newPara = paragraphs.items[paragraph_index].insertParagraph(String(content), "After");
            await context.sync();
            await applyFormat(newPara, format, context);
          } else {
            await insertMultiParagraphAfter(paragraphs.items[paragraph_index], String(content), format, context);
          }
          return `已在段落 ${paragraph_index} 后插入内容（${lines.length} 段）`;
        }
        const lastPara = body.paragraphs.getLast();
        if (isHtml) {
          lastPara.getRange("End").insertHtml(String(content), "After");
          await context.sync();
          return `段落序号 ${paragraph_index} 超出范围，已插入 HTML 格式内容到末尾`;
        }
        const lines = normalizeContentLines(String(content));
        if (lines.length <= 1) {
          const newPara = lastPara.insertParagraph(String(content), "After");
          await context.sync();
          await applyFormat(newPara, format, context);
        } else {
          await insertMultiParagraphAfter(lastPara, String(content), format, context);
        }
        throw new Error(`insert_after_paragraph 失败：段落序号 ${paragraph_index} 超出范围（文档共 ${paragraphs.items.length} 段）。请先调用 read_document 确认段落序号。`);
      }

      case "delete_paragraph": {
        const { paragraph_index } = action.params;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraph_index >= 0 && paragraph_index < paragraphs.items.length) {
          try {
            paragraphs.items[paragraph_index].delete();
            await context.sync();
          } catch (e: any) {
            // Office.js throws GeneralException when the document is
            // read-only, protected, or the user lacks edit permissions.
            const isGeneral = e?.code === "GeneralException" || e?.message === "GeneralException";
            if (isGeneral) {
              throw new Error(
                `delete_paragraph 失败：无法删除段落 ${paragraph_index}。` +
                `可能原因：文档处于只读模式、受保护视图、或当前没有编辑权限。` +
                `请通知用户检查文档是否可编辑（如点击"启用编辑"按钮），然后重试。`
              );
            }
            throw e;
          }
          return `已删除段落 ${paragraph_index}`;
        }
        throw new Error(`delete_paragraph 失败：段落序号 ${paragraph_index} 超出范围（文档共 ${paragraphs.items.length} 段）。请先调用 read_document 确认段落序号。`);
      }

      case "find_and_replace": {
        const { find_text, replace_text, match_case = false } = action.params;
        const ranges = body.search(find_text, { matchCase: match_case, matchWholeWord: false });
        ranges.load("items");
        await context.sync();

        if (ranges.items.length > 0) {
          // Replace first occurrence
          ranges.items[0].insertText(replace_text, "Replace");
          await context.sync();
          return `已替换 ${ranges.items.length} 处匹配中的第 1 处`;
        }
        throw new Error(`find_and_replace 失败：未找到"${find_text}"。请检查查找文本是否正确，或尝试使用 find_and_replace_v2 配合 match_case=false。`);
      }

      case "read_document": {
        const result = await readDocument(action.params);
        return result;
      }

      case "get_selection_info": {
        const result = await getSelectionInfo();
        return result;
      }

      case "get_document_stats": {
        const result = await getDocumentStats();
        return result;
      }

      case "get_paragraph_format": {
        const result = await getParagraphFormat(action.params);
        return result;
      }

      case "get_document_tables": {
        const result = await getDocumentTables();
        return result;
      }

      case "read_table": {
        const result = await readTable(action.params);
        return result;
      }

      case "insert_at_cursor": {
        const { content, format = "normal", content_format = "text" } = action.params;
        if (!content || String(content).trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符");
        }
        const selection = context.document.getSelection();
        if (content_format === "html") {
          selection.insertHtml(String(content), "Replace");
          await context.sync();
          return "已在光标处插入 HTML 格式内容";
        }
        const lines = normalizeContentLines(String(content));
        if (lines.length <= 1) {
          const para = selection.insertParagraph(String(content), "Replace");
          await context.sync();
          await applyFormat(para, format, context);
        } else {
          // Insert first line at cursor, then add remaining lines after
          const firstPara = selection.insertParagraph(lines[0] || " ", "Replace");
          await applyFormat(firstPara, format, context);
          await insertMultiParagraphAfter(firstPara, lines.slice(1).join("\n"), format, context);
        }
        return `已在光标处插入内容（${lines.length} 段）`;
      }

      case "find_and_replace_v2": {
        const { find_text, replace_text, replace_all = false, match_case = false, match_whole_word = false } = action.params;
        const ranges = body.search(find_text, { matchCase: match_case, matchWholeWord: match_whole_word });
        ranges.load("items");
        await context.sync();

        let replaced = 0;
        for (const range of ranges.items) {
          range.insertText(replace_text, "Replace");
          replaced++;
          if (!replace_all) break;
        }
        await context.sync();
        return `已替换 ${replaced} 处匹配（共找到 ${ranges.items.length} 处）`;
      }

      case "replace_paragraph": {
        const { paragraph_index, content, content_format = "text" } = action.params;
        if (!content || String(content).trim().length === 0) {
          throw new Error("替换内容为空或仅包含空白字符");
        }
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraph_index < 0 || paragraph_index >= paragraphs.items.length) {
          throw new Error(`段落序号 ${paragraph_index} 超出范围`);
        }
        const para = paragraphs.items[paragraph_index];
        if (content_format === "html") {
          para.getRange("Whole").insertHtml(String(content), "Replace");
          await context.sync();
          return `已替换段落 ${paragraph_index} 的内容（HTML 格式）`;
        }
        // If content has newlines, replace current paragraph and insert remaining as new paragraphs after
        const lines = normalizeContentLines(String(content));
        if (lines.length <= 1) {
          para.getRange("Whole").insertText(String(content), "Replace");
          await context.sync();
        } else {
          para.getRange("Whole").insertText(lines[0] || " ", "Replace");
          await insertMultiParagraphAfter(para, lines.slice(1).join("\n"), "normal", context);
        }
        return `已替换段落 ${paragraph_index} 的内容（${lines.length} 段）`;
      }

      case "set_paragraph_style": {
        const { paragraph_index, format } = action.params;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraph_index < 0 || paragraph_index >= paragraphs.items.length) {
          throw new Error(`段落序号 ${paragraph_index} 超出范围`);
        }
        await applyFormat(paragraphs.items[paragraph_index], format, context);
        return `已将段落 ${paragraph_index} 设置为 ${format} 格式`;
      }

      case "merge_paragraphs": {
        const { first_paragraph_index, second_paragraph_index, separator = " " } = action.params;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (first_paragraph_index < 0 || first_paragraph_index >= paragraphs.items.length) {
          throw new Error(`第一个段落序号 ${first_paragraph_index} 超出范围`);
        }
        if (second_paragraph_index < 0 || second_paragraph_index >= paragraphs.items.length) {
          throw new Error(`第二个段落序号 ${second_paragraph_index} 超出范围`);
        }
        if (first_paragraph_index === second_paragraph_index) {
          throw new Error("不能合并同一个段落");
        }

        const firstPara = paragraphs.items[first_paragraph_index];
        const secondPara = paragraphs.items[second_paragraph_index];
        firstPara.load("text");
        secondPara.load("text");
        await context.sync();

        const mergedText = (firstPara.text || "") + separator + (secondPara.text || "");
        firstPara.getRange("Whole").insertText(mergedText, "Replace");
        secondPara.delete();
        await context.sync();
        return `已合并段落 ${first_paragraph_index} 和段落 ${second_paragraph_index}`;
      }

      case "apply_rich_format": {
        const { target_mode, paragraph_index, font, text_to_format } = action.params;
        let targetRange: Word.Range;

        if (target_mode === "selection") {
          targetRange = context.document.getSelection();
        } else if (target_mode === "paragraph_index") {
          const paragraphs = body.paragraphs;
          paragraphs.load("items");
          await context.sync();
          if (paragraph_index < 0 || paragraph_index >= paragraphs.items.length) {
            throw new Error(`段落序号 ${paragraph_index} 超出范围`);
          }
          targetRange = paragraphs.items[paragraph_index].getRange("Whole");
        } else if (target_mode === "last_inserted") {
          throw new Error("last_inserted 模式暂未实现，请使用 paragraph_index 或 selection");
        } else {
          throw new Error(`未知的目标模式: ${target_mode}`);
        }

        if (text_to_format) {
          const found = targetRange.search(text_to_format, { matchCase: true });
          found.load("items");
          await context.sync();
          if (found.items.length > 0) {
            targetRange = found.items[0];
          } else {
            throw new Error(`在目标范围内未找到文本"${text_to_format}"，请确认文本完全匹配（区分大小写）`);
          }
        }

        targetRange.load("font");
        await context.sync();

        if (font) {
          if (font.name) targetRange.font.name = font.name;
          if (font.size) targetRange.font.size = font.size;
          if (font.color) targetRange.font.color = font.color;
          if (font.bold !== undefined) targetRange.font.bold = font.bold;
          if (font.italic !== undefined) targetRange.font.italic = font.italic;
        }

        await context.sync();
        return text_to_format
          ? `已对文本"${text_to_format}"应用富文本格式`
          : "已应用富文本格式";
      }

      case "reply_only":
        // No document operation needed
        return "仅回复文本，无需文档操作";

      case "task_complete":
        // No document operation needed — this signals agent loop completion
        return action.params.summary || "任务已完成";

      case "undo_last_action": {
        const count = action.params.count || 1;
        const doc = context.document;
        // Word JS API 1.3+ supports document.undo(count), but local type defs may not include it
        (doc as any).undo(count);
        await context.sync();
        return `已撤销最近 ${count} 步操作`;
      }

      case "delete_table": {
        const tableIndex = typeof action.params.table_index === "number" ? action.params.table_index : 0;
        const tables = body.tables;
        tables.load("items");
        await context.sync();
        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(`表格序号 ${tableIndex} 越界（文档共有 ${tables.items.length} 张表格，序号范围 0 ~ ${tables.items.length - 1}）。请先调用 get_document_tables 确认有效索引。`);
        }
        const targetTable = tables.items[tableIndex];
        // WordApi 1.3+: table.delete() directly deletes the entire table.
        // Fallback: delete via the table's body-level range.
        try {
          targetTable.delete();
        } catch {
          const range = targetTable.getRange("Whole");
          range.delete();
        }
        await context.sync();
        return `已删除表格 ${tableIndex}`;
      }

      case "insert_table": {
        const {
          location,
          heading_text,
          paragraph_index,
          headers = [],
          rows,
          style = "TableGrid",
        } = action.params;

        if (!Array.isArray(rows) || rows.length === 0) {
          throw new Error("rows 必须是非空数组");
        }
        if (!location) {
          throw new Error("缺少 location 参数");
        }

        // Normalize rows: ensure every row is an array of strings, and pad
        // short rows with empty strings to match the longest row (or the
        // headers length, whichever is greater). Auto-align is intentional
        // — see Phase 3 design decision.
        const stringRows: string[][] = rows.map((r: unknown) => {
          if (!Array.isArray(r)) {
            throw new Error("rows 元素必须是字符串数组");
          }
          return r.map((c) => (c == null ? "" : String(c)));
        });
        const allRows: string[][] = headers.length > 0
          ? [headers.map((h: unknown) => (h == null ? "" : String(h))), ...stringRows]
          : stringRows;
        const colCount = Math.max(...allRows.map((r) => r.length));
        if (colCount === 0) {
          throw new Error("表格至少需要 1 列");
        }
        const normalized = allRows.map((r) => {
          const padded = [...r];
          while (padded.length < colCount) padded.push("");
          return padded;
        });
        const rowCount = normalized.length;
        // Built-in style names like "TableGrid" must be set via Table.styleBuiltIn
        // (enum), not Table.style (which is for custom/localized names). Setting
        // an invalid string on `style` causes Word to throw a GeneralException
        // because the named style does not exist in the document's style table.
        const tableStyleBuiltIn = (style || "TableGrid") as string;

        // Office.js requires the proxy object to be `.load`-ed and `context.sync()`
        // round-tripped before setting `styleBuiltIn`. The set is deferred to
        // a follow-up sync to keep the change batch isolated from insertion.
        const applyTableStyle = async (table: Word.Table): Promise<void> => {
          try {
            (table as any).styleBuiltIn = tableStyleBuiltIn;
            await context.sync();
          } catch (e) {
            console.warn(`[insert_table] failed to apply style ${tableStyleBuiltIn}:`, e);
          }
        };

        if (location === "after_heading") {
          if (!heading_text) {
            throw new Error("location=after_heading 时必须提供 heading_text");
          }
          const matches = body.search(heading_text, { matchCase: false, matchWholeWord: false });
          matches.load("items");
          await context.sync();
          if (matches.items.length === 0) {
            throw new Error(`未找到标题 "${heading_text}"`);
          }
          const para = matches.items[0].paragraphs.getFirst();
          para.load("text");
          await context.sync();
          // Insert by first obtaining the paragraph's range — `insertTable` is a
          // method on Range, not on Paragraph, in the real Word API.
          const range: Word.Range = para.getRange("Content");
          const table: Word.Table = range.insertTable(rowCount, colCount, "After", normalized);
          await context.sync();
          await applyTableStyle(table);
          return `已在标题 "${heading_text}" 后插入 ${rowCount} × ${colCount} 表格`;
        }

        if (location === "after_paragraph") {
          if (typeof paragraph_index !== "number") {
            throw new Error("location=after_paragraph 时必须提供 paragraph_index");
          }
          const paragraphs = body.paragraphs;
          paragraphs.load("items");
          await context.sync();
          if (paragraph_index < 0 || paragraph_index >= paragraphs.items.length) {
            throw new Error(`paragraph_index ${paragraph_index} 越界（共 ${paragraphs.items.length} 段）`);
          }
          const target = paragraphs.items[paragraph_index];
          target.load("text");
          await context.sync();
          // Insert via the paragraph's own range so Word creates the table as
          // a sibling of the paragraph (After the paragraph's content range),
          // rather than trying to splice into the paragraph's body.
          const range: Word.Range = target.getRange("Content");
          const table: Word.Table = range.insertTable(rowCount, colCount, "After", normalized);
          await context.sync();
          await applyTableStyle(table);
          return `已在段落 ${paragraph_index} 后插入 ${rowCount} × ${colCount} 表格`;
        }

        if (location === "at_end") {
          const lastPara = body.paragraphs.getLast();
          lastPara.load("text");
          await context.sync();
          const range: Word.Range = lastPara.getRange("Content");
          const table: Word.Table = range.insertTable(rowCount, colCount, "After", normalized);
          await context.sync();
          await applyTableStyle(table);
          return `已追加 ${rowCount} × ${colCount} 表格到文档末尾`;
        }

        // at_cursor (default fallback): insert at the end of the current
        // selection so the table lands as a sibling, not replacing the
        // selection's content.
        const selection = context.document.getSelection();
        const table: Word.Table = selection.insertTable(rowCount, colCount, "After", normalized);
        await context.sync();
        await applyTableStyle(table);
        return `已在光标处插入 ${rowCount} × ${colCount} 表格`;
      }

      case "insert_mermaid_image": {
        const {
          mermaid_code,
          location = "at_cursor",
          heading_text,
          paragraph_index,
        } = action.params as {
          mermaid_code: string;
          location?: "at_cursor" | "at_end" | "after_heading" | "after_paragraph";
          heading_text?: string;
          paragraph_index?: number;
        };
        if (!mermaid_code || String(mermaid_code).trim().length === 0) {
          throw new Error("mermaid_code 不能为空");
        }

        // Render to PNG data URL (mermaid only touches the browser DOM).
        const dataUrl = await renderMermaidImage(String(mermaid_code));

        let target: Word.Range;
        let locationDesc: string;
        if (location === "after_heading") {
          if (!heading_text) throw new Error("location=after_heading 时必须提供 heading_text");
          const matches = body.search(heading_text, { matchCase: false, matchWholeWord: false });
          matches.load("items");
          await context.sync();
          if (matches.items.length === 0) throw new Error(`未找到标题 "${heading_text}"`);
          target = matches.items[0].paragraphs.getFirst().getRange("End");
          locationDesc = `标题 "${heading_text}" 后`;
        } else if (location === "after_paragraph") {
          if (typeof paragraph_index !== "number") {
            throw new Error("location=after_paragraph 时必须提供 paragraph_index");
          }
          const paragraphs = body.paragraphs;
          paragraphs.load("items");
          await context.sync();
          if (paragraph_index < 0 || paragraph_index >= paragraphs.items.length) {
            throw new Error(`paragraph_index ${paragraph_index} 越界（共 ${paragraphs.items.length} 段）`);
          }
          target = paragraphs.items[paragraph_index].getRange("End");
          locationDesc = `段落 ${paragraph_index} 后`;
        } else if (location === "at_end") {
          target = body.paragraphs.getLast().getRange("End");
          locationDesc = "文档末尾";
        } else {
          target = context.document.getSelection().getRange("End");
          locationDesc = "光标处";
        }

        target.insertInlinePictureFromBase64(dataUrl, "After");
        await context.sync();
        return `已在${locationDesc}插入流程图`;
      }

      default:
        return `未知操作: ${action.action}`;
    }
  });
}

function stringifyOfficeError(error: unknown): { message: string; details: string } {
  let message = "未知错误";
  let details = "";
  if (error instanceof Error) {
    message = error.message;
    details = `name: ${error.name}, message: ${error.message}`;
    // Office.js errors may have extra fields
    const ext = error as any;
    if (ext.code) {
      details += `, code: ${String(ext.code)}`;
    }
    if (ext.traceMessages && Array.isArray(ext.traceMessages)) {
      details += `, traceMessages: ${JSON.stringify(ext.traceMessages)}`;
    }
    if (ext.innerError) {
      details += `, innerError: ${JSON.stringify(ext.innerError)}`;
    }
    if (ext.debugInfo) {
      details += `, debugInfo: ${JSON.stringify(ext.debugInfo)}`;
    }
  } else {
    details = String(error);
  }
  console.error("[OfficeError]", details, error);
  return { message, details };
}

async function executeActionPlan(plan: ActionPlan): Promise<Array<{ toolCallId: string; toolName: string; result: string; success: boolean }>> {
  const results: Array<{ toolCallId: string; toolName: string; result: string; success: boolean }> = [];

  for (let i = 0; i < plan.actions.length; i++) {
    const action = plan.actions[i];
    try {
      const result = await executeAction(action);
      appendToolResultCard(action.action, true, result);
      results.push({
        toolCallId: action.toolCallId ?? `action-${i}`,
        toolName: action.action,
        result,
        success: true,
      });
    } catch (error) {
      const { message, details } = stringifyOfficeError(error);
      appendToolResultCard(action.action, false, `错误: ${details}`);
      results.push({
        toolCallId: action.toolCallId ?? `action-${i}`,
        toolName: action.action,
        result: `错误: ${details}`,
        success: false,
      });
      console.error(`[executeActionPlan] action ${i + 1} (${action.action}) failed:`, details);

      if (plan.actions.length > 1 && !state.autopilot) {
        const shouldContinue = confirm(`操作 ${i + 1} 失败: ${message}\n\n是否继续执行后续操作？`);
        if (!shouldContinue) break;
      }
    }
  }

  const statusLines = results.map((r) => `${r.success ? "✅" : "❌"} ${r.toolName}: ${r.result.slice(0, 100)}`);
  setStatus(chatStatus, statusLines.join("\n"));
  return results;
}

// ─── Action Plan Preview ────────────────────────────────────────────────────

function showActionPlanPreview(plan: ActionPlan): void {
  // Autopilot never pauses for confirmation — run any plan that reaches the
  // preview immediately (safety net over every call path).
  if (state.autopilot) {
    state.pendingActionPlan = plan;
    void confirmActionPlan();
    return;
  }
  state.pendingActionPlan = plan;

  // Show tool call cards in chat log for visibility
  for (const action of plan.actions) {
    appendToolCallCard(action.action, action.params);
  }

  // Build preview content
  let html = `<div class="action-plan-explanation">${escapeHtml(plan.explanation)}</div>`;

  for (const action of plan.actions) {
    html += `<div class="action-plan-item">`;
    html += `<div class="action-plan-action">${escapeHtml(action.description)}</div>`;
    html += `<div class="action-plan-params">`;

    for (const [key, value] of Object.entries(action.params)) {
      if (key === "content") {
        const truncated = String(value).length > 100 ? String(value).slice(0, 100) + "..." : String(value);
        html += `<span class="action-param">${escapeHtml(key)}: "${escapeHtml(truncated)}"</span>`;
      } else {
        html += `<span class="action-param">${escapeHtml(key)}: ${escapeHtml(String(value))}</span>`;
      }
    }

    html += `</div></div>`;
  }

  actionPlanContent.innerHTML = html;
  actionPlanPanel.style.display = "";
  setStatus(chatStatus, `预览：${plan.actions.length} 个操作 — 请确认或取消`);
}

function hideActionPlanPreview(): void {
  actionPlanPanel.style.display = "none";
  state.pendingActionPlan = null;
}

async function confirmActionPlan(): Promise<void> {
  if (state.isAgentRunning) {
    setStatus(chatStatus, "正在生成中，请先点 ⏹ 停止");
    return;
  }
  if (!state.pendingActionPlan) {
    setStatus(chatStatus, "没有待执行的操作计划");
    return;
  }

  const plan = state.pendingActionPlan;
  const sessionId = state.pendingSessionId;
  console.log("[confirmActionPlan] plan:", JSON.stringify(plan.actions.map(a => ({ action: a.action, toolCallId: a.toolCallId }))));
  console.log("[confirmActionPlan] sessionId:", sessionId);
  hideActionPlanPreview();

  setRunningState(true);

  try {
    const results = await executeActionPlan(plan);
    console.log("[confirmActionPlan] executeActionPlan results:", JSON.stringify(results.map(r => ({ toolCallId: r.toolCallId, toolName: r.toolName, success: r.success }))));

    // If this came from the agent loop (has sessionId), continue the loop
    if (sessionId) {
      console.log("[confirmActionPlan] Continuing agent loop with sessionId:", sessionId);
      setStatus(chatStatus, "操作已执行，等待模型继续...");

      const assistantEl = createAssistantMessage("");

      const data = await continueAgentLoop(
        sessionId,
        results,
        (partial) => updateAssistantMessage(assistantEl, partial)
      );

      // Save assistant reply to session
      const activeSession = getActiveSession();
      if (activeSession && data.reply) {
        activeSession.messages.push({ role: "assistant", content: data.reply });
        activeSession.updatedAt = Date.now();
        trimSessionMessages(activeSession);
        saveSessionsToStorage();
      }

      const citationText = data.citations?.length
        ? `\n\n来源: ${data.citations.map((c) => c.fileName).join(", ")}（命中 ${data.retrievalCount ?? data.citations.length} 段）`
        : "";

      if (data.actionPlan && data.actionPlan.actions.length > 0) {
        updateAssistantMessage(assistantEl, `${data.reply}${citationText}\n\n📋 操作计划：${data.actionPlan.explanation}`);
        showActionPlanPreview(data.actionPlan);
        state.pendingSessionId = data.sessionId;
      } else {
        updateAssistantMessage(assistantEl, `${data.reply}${citationText}`);
        setStatus(chatStatus, "已完成");
        state.pendingSessionId = null;
      }
    } else {
      // No sessionId — cannot continue agent loop
      console.log("[confirmActionPlan] Not continuing agent loop. sessionId:", sessionId);
      setStatus(chatStatus, "操作已执行（无会话ID，无法继续迭代）");
    }
  } catch (error) {
    if (isAbortError(error)) {
      setStatus(chatStatus, "已停止生成");
      state.pendingSessionId = null;
    } else {
      const message = error instanceof Error ? error.message : "未知错误";
      console.error("[confirmActionPlan] Error:", message);
      setStatus(chatStatus, `执行操作失败: ${message}`);
    }
  } finally {
    setRunningState(false);
  }
}

function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// ─── Legacy Insert Functions (kept for backward compatibility) ──────────────

async function applyTextToWordSelection(text: string): Promise<void> {
  if (isWebViewHost()) {
    await bridgeSimpleInsert("replace", text);
    return;
  }
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text, "Replace");
    await context.sync();
  });
}

async function insertTextAtCursor(text: string): Promise<void> {
  const normalized = cleanupMarkdownForWord(text);
  if (!normalized) {
    throw new Error("没有可插入的正文内容");
  }

  if (isWebViewHost()) {
    await bridgeSimpleInsert("at_cursor", normalized);
    return;
  }

  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(normalized, "End");
    await context.sync();
  });
}

// ─── Preview (legacy, for simple text insert) ────────────────────────────────

const previewPanel = $<HTMLDivElement>("previewPanel");
const previewContent = $<HTMLDivElement>("previewContent");

function showSimplePreview(text: string, mode: "replace" | "append_end" | "insert_start"): void {
  // Autopilot never pauses on the insert preview either.
  if (state.autopilot) {
    state.pendingSimpleInsert = { text, mode };
    void confirmSimplePreview();
    return;
  }
  const modeLabels: Record<string, string> = {
    replace: "替换当前选区",
    append_end: "追加到文末",
    insert_start: "插入到文档开头",
  };

  previewContent.textContent = text;
  state.pendingSimpleInsert = { text, mode };
  previewPanel.style.display = "";
  setStatus(chatStatus, `预览：${modeLabels[mode] ?? "插入"} — 请确认或取消`);
}

function hideSimplePreview(): void {
  previewPanel.style.display = "none";
  state.pendingSimpleInsert = null;
}

async function confirmSimplePreview(): Promise<void> {
  if (!state.pendingSimpleInsert) {
    setStatus(chatStatus, "没有待插入的内容");
    return;
  }

  const { text, mode } = state.pendingSimpleInsert;
  hideSimplePreview();

  if (isWebViewHost()) {
    try {
      await bridgeSimpleInsert(mode, text);
      setStatus(chatStatus, mode === "replace" ? "已替换当前选区（已去除 Markdown）" : mode === "append_end" ? "已追加到文档末尾" : "已插入到文档开头");
    } catch (error) {
      const message = error instanceof Error ? error.message : "未知错误";
      setStatus(chatStatus, `插入失败: ${message}`);
    }
    return;
  }

  try {
    if (mode === "replace") {
      await applyTextToWordSelection(text);
      setStatus(chatStatus, "已替换当前选区（已去除 Markdown）");
    } else if (mode === "append_end") {
      await Word.run(async (context) => {
        const lastPara = context.document.body.paragraphs.getLast();
        lastPara.insertParagraph(text, "After");
        await context.sync();
      });
      setStatus(chatStatus, "已追加到文档末尾");
    } else if (mode === "insert_start") {
      await Word.run(async (context) => {
        const body = context.document.body;
        console.log("[DEBUG insert_start simple] text length:", text?.length, "text preview:", JSON.stringify(text?.slice(0, 50)));

        if (!text || text.trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符");
        }

        let newPara: Word.Paragraph;
        try {
          const firstPara = body.paragraphs.getFirst();
          firstPara.load("text");
          await context.sync();
          console.log("[DEBUG insert_start simple] firstPara.text:", JSON.stringify(firstPara.text));
          newPara = firstPara.insertParagraph(text, "Before");
        } catch (e) {
          const errMsg = e instanceof Error ? e.message : String(e);
          console.log("[DEBUG insert_start simple] getFirst failed, fallback to body.insertParagraph(Start):", errMsg);
          newPara = body.insertParagraph(text, "Start");
        }
        await context.sync();
        console.log("[DEBUG insert_start simple] inserted successfully");
      });
      setStatus(chatStatus, "已插入到文档开头");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    console.error("[ERROR confirmSimplePreview] mode:", mode, "error:", message, error);
    setStatus(chatStatus, `插入失败: ${message}`);
  }
}

// ─── Chat Stream ────────────────────────────────────────────────────────────

type StreamEvent = {
  type: string;
  delta?: string;
  reply?: string;
  actionPlan?: ActionPlan | null;
  retrievalCount?: number;
  citations?: Array<{ fileName: string }>;
  error?: string;
  promptTokens?: number;
  completionTokens?: number;
  estimated?: boolean;
  contextWindow?: number;
};

type AgentStreamEvent =
  | { type: "start"; ts: number }
  | { type: "delta"; delta: string }
  | { type: "tool_call"; tools: Array<{ id: string; tool: string; params: Record<string, any> }> }
  | { type: "iterable_tool_call"; tools: Array<{ id: string; tool: string; params: Record<string, any> }> }
  | { type: "action_plan"; plan: ActionPlan }
  | { type: "task_complete"; summary: string }
  | { type: "session"; sessionId: string }
  | { type: "await_tool_result"; iteration: number }
  | { type: "fallback"; reason: string }
  | { type: "usage"; promptTokens?: number; completionTokens?: number; estimated?: boolean; contextWindow?: number }
  | { type: "done"; reply: string; actionPlan?: ActionPlan | null; sessionId?: string; retrievalCount?: number; citations?: Array<{ fileName: string }> }
  | { type: "error"; error: string };

async function parseSSEResponse(
  res: Response,
  onDelta?: (text: string) => void
): Promise<AgentStreamEvent[]> {
  if (!res.ok) {
    const err = await parseErrorMessage(res);
    throw new Error(err);
  }

  if (!res.body) {
    throw new Error("流式响应不可用");
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  const events: AgentStreamEvent[] = [];
  let accumulatedDelta = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const chunks = buffer.split("\n\n");
    buffer = chunks.pop() ?? "";

    for (const chunk of chunks) {
      const line = chunk.split("\n").find((item) => item.trim().startsWith("data:"));
      if (!line) continue;

      const raw = line.slice(5).trim();
      if (!raw || raw === "[DONE]") continue;

      let data: AgentStreamEvent;
      try {
        data = JSON.parse(raw) as AgentStreamEvent;
      } catch {
        continue;
      }

      events.push(data);

      if (data.type === "delta" && data.delta) {
        accumulatedDelta += data.delta;
        onDelta?.(accumulatedDelta);
      }

      if (data.type === "error") {
        throw new Error(data.error || "流式响应出错");
      }
    }
  }

  return events;
}

async function executeSingleTool(
  tool: { id: string; tool: string; params: Record<string, any> }
): Promise<{ toolCallId: string; toolName: string; result: string; success: boolean }> {
  const action: WordAction = {
    action: tool.tool,
    params: tool.params,
    description: `Agent Loop: ${tool.tool}`,
  };

  const startedMs = Date.now();
  try {
    const result = await executeAction(action);
    reportToolEvent(tool.tool, true, Date.now() - startedMs);
    trackCheckpointUndo(tool.tool, tool.params, true);
    return { toolCallId: tool.id, toolName: tool.tool, result, success: true };
  } catch (error) {
    const { details } = stringifyOfficeError(error);
    reportToolEvent(tool.tool, false, Date.now() - startedMs);
    return { toolCallId: tool.id, toolName: tool.tool, result: `错误: ${details}`, success: false };
  }
}

/** Tools that change paragraph indices (like iterator invalidation) */
const STRUCTURE_MUTATING_TOOLS = new Set([
  "insert_after_paragraph", "insert_after_heading", "insert_at_start", "insert_at_end",
  "insert_table", "delete_table",
  "insert_mermaid_image",
  "delete_paragraph", "merge_paragraphs", "undo_last_action",
]);

/**
 * After a structure-mutating tool executes, automatically call read_document
 * to refresh paragraph indices — prevents stale-index bugs (LLM "iterator invalidation").
 */
async function autoReadAfterMutation(
  executedTools: Array<{ tool: string; params: Record<string, any> }>,
  results: Array<{ toolCallId: string; toolName: string; result: string; success: boolean }>
): Promise<Array<{ toolCallId: string; toolName: string; result: string; success: boolean }>> {
  const lastMutation = [...results].reverse().find(r => r.success && STRUCTURE_MUTATING_TOOLS.has(r.toolName));
  if (!lastMutation) return results;

  // Find the params of the last mutating tool to determine read range
  const lastExecuted = [...executedTools].reverse().find(t => STRUCTURE_MUTATING_TOOLS.has(t.tool));
  let readParams: Record<string, any> = { mode: "paragraph_range", paragraph_index: 0, count: 15 };
  if (lastExecuted?.params) {
    const pIdx = lastExecuted.params.paragraph_index;
    if (typeof pIdx === "number") {
      readParams = { mode: "paragraph_range", paragraph_index: Math.max(0, pIdx - 2), count: 15 };
    }
  }

  const readResult = await executeSingleTool({
    id: `auto-read-${Date.now()}`,
    tool: "read_document",
    params: readParams,
  });
  appendToolCallCard("read_document", readParams);
  appendToolResultCard("read_document", readResult.success, readResult.result);
  results.push(readResult);
  return results;
}

async function sendAgentMessageStream(
  payload: {
    messages: ChatMessage[];
    documentContext: string;
    documentStructure?: DocumentStructure;
    selection: string;
    insertMode: InsertMode;
  },
  onDelta: (text: string) => void,
  maxIterations = 10
): Promise<{
  reply: string;
  actionPlan: ActionPlan | null;
  retrievalCount?: number;
  citations?: Array<{ fileName: string }>;
  sessionId?: string;
}> {
  // Share a single AbortController across the whole multi-turn loop so the
  // Stop button can tear down every pending /v1/chat/agent-stream and
  // /v1/chat/agent-continue request, not just the in-flight one.
  currentAbortController = new AbortController();
  const abortSignal = currentAbortController.signal;

  let sessionId: string | null = null;
  let toolResults: Array<{ toolCallId: string; toolName: string; result: string; success: boolean }> = [];
  let iteration = 0;
  let finalReply = "";
  let finalActionPlan: ActionPlan | null = null;
  let finalRetrievalCount = 0;
  let finalCitations: Array<{ fileName: string }> = [];

  while (iteration < maxIterations) {
    if (abortSignal.aborted) {
      throw new DOMException("Agent loop stopped by user", "AbortError");
    }
    iteration++;

    let res: Response;
    if (!sessionId) {
      // Initial request
      res = await fetch(`${agentBase}/v1/chat/agent-stream`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...payload, enableReAct: true, maxIterations }),
        signal: abortSignal,
      });
    } else {
      // Continue with tool results — re-read fresh document structure
      const freshDocStructure = await getStructuredContext();
      res = await fetch(`${agentBase}/v1/chat/agent-continue`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sessionId, toolResults, documentStructure: freshDocStructure }),
        signal: abortSignal,
      });
      toolResults = [];
    }

    const events = await parseSSEResponse(res, onDelta);

    let hasToolCall = false;
    for (const event of events) {
      if (event.type === "session") {
        sessionId = event.sessionId;
      }

      if (event.type === "usage") {
        applyUsageEvent(event);
      }

      if (event.type === "tool_call") {
        hasToolCall = true;

        // Autopilot: execute ALL tools inline, regardless of type.
        if (state.autopilot) {
          const executedTools: Array<{ tool: string; params: Record<string, any> }> = [];
          for (const tool of event.tools) {
            appendToolCallCard(tool.tool, tool.params);
            const result = await executeSingleTool(tool);
            appendToolResultCard(tool.tool, result.success, result.result);
            toolResults.push(result);
            executedTools.push({ tool: tool.tool, params: tool.params });
          }
          await autoReadAfterMutation(executedTools, toolResults);
          continue;
        }

        for (const tool of event.tools) {
          const perceptionTools = ["read_document", "get_selection_info", "get_document_stats", "get_paragraph_format", "get_document_tables", "read_table"];
          if (!perceptionTools.includes(tool.tool)) {
            // Action tool in agent loop — treat as action_plan
            const plan: ActionPlan = {
              actions: event.tools.map((t) => ({
                action: t.tool,
                params: t.params,
                description: `Agent: ${t.tool}`,
                toolCallId: t.id,
              })),
              explanation: `Agent 循环中的操作: ${event.tools.map((t) => t.tool).join(", ")}`,
            };
            return { reply: finalReply || plan.explanation, actionPlan: plan, retrievalCount: finalRetrievalCount, citations: finalCitations, sessionId: sessionId ?? undefined };
          }

          appendToolCallCard(tool.tool, tool.params);
          const result = await executeSingleTool(tool);
          appendToolResultCard(tool.tool, result.success, result.result);
          toolResults.push(result);
        }
      }

      // Iterable tool calls: execute action tools directly and continue loop (no user confirmation needed)
      if (event.type === "iterable_tool_call") {
        hasToolCall = true;
        const executedTools: Array<{ tool: string; params: Record<string, any> }> = [];
        for (const tool of event.tools) {
          appendToolCallCard(tool.tool, tool.params);
          const result = await executeSingleTool(tool);
          appendToolResultCard(tool.tool, result.success, result.result);
          toolResults.push(result);
          executedTools.push({ tool: tool.tool, params: tool.params });
        }
        // Auto-read after structure-mutating tools to refresh paragraph indices
        await autoReadAfterMutation(executedTools, toolResults);
      }

      if (event.type === "action_plan") {
        // Autopilot: run the plan inline and keep looping — never pause for
        // confirmation mid-run.
        if (state.autopilot && event.plan.actions.length > 0) {
          hasToolCall = true;
          const planResults = await executeActionPlan(event.plan);
          for (const r of planResults) toolResults.push(r);
          continue;
        }
        return { reply: finalReply || event.plan.explanation, actionPlan: event.plan, retrievalCount: finalRetrievalCount, citations: finalCitations, sessionId: sessionId ?? undefined };
      }

      if (event.type === "task_complete") {
        return { reply: event.summary || finalReply, actionPlan: null, retrievalCount: finalRetrievalCount, citations: finalCitations, sessionId: sessionId ?? undefined };
      }

      if (event.type === "done") {
        finalReply = event.reply ?? finalReply;
        if (state.autopilot && event.actionPlan && event.actionPlan.actions.length > 0) {
          hasToolCall = true;
          const planResults = await executeActionPlan(event.actionPlan);
          for (const r of planResults) toolResults.push(r);
        } else {
          finalActionPlan = event.actionPlan ?? finalActionPlan;
        }
        if (event.retrievalCount !== undefined) finalRetrievalCount = event.retrievalCount;
        if (event.citations) finalCitations = event.citations;
        if (event.sessionId) sessionId = event.sessionId;
      }
    }

    if (!hasToolCall || toolResults.length === 0) {
      break;
    }
  }

  return { reply: finalReply.trim(), actionPlan: finalActionPlan, retrievalCount: finalRetrievalCount, citations: finalCitations, sessionId: sessionId ?? undefined };
}

async function continueAgentLoop(
  sessionId: string,
  toolResults: Array<{ toolCallId: string; toolName: string; result: string; success: boolean }>,
  onDelta: (text: string) => void,
  maxIterations = 5
): Promise<{
  reply: string;
  actionPlan: ActionPlan | null;
  sessionId: string;
  retrievalCount?: number;
  citations?: Array<{ fileName: string }>;
}> {
  // A new AbortController is created here so the Stop button can also tear
  // down post-confirmation continuations. Replaces any leftover controller
  // from the initial sendAgentMessageStream run.
  currentAbortController = new AbortController();
  const abortSignal = currentAbortController.signal;

  let iteration = 0;
  let finalReply = "";
  let finalActionPlan: ActionPlan | null = null;
  let finalRetrievalCount = 0;
  let finalCitations: Array<{ fileName: string }> = [];
  let currentToolResults = [...toolResults];
  console.log("[continueAgentLoop] Starting with sessionId:", sessionId, "toolResults:", JSON.stringify(toolResults.map(r => ({ toolCallId: r.toolCallId, toolName: r.toolName, success: r.success }))));

  while (iteration < maxIterations) {
    if (abortSignal.aborted) {
      throw new DOMException("Agent loop stopped by user", "AbortError");
    }
    iteration++;
    console.log("[continueAgentLoop] Iteration:", iteration);

    // Re-read fresh document structure after tool execution so LLM sees updated state
    const freshDocStructure = await getStructuredContext();

    const res = await fetch(`${agentBase}/v1/chat/agent-continue`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId, toolResults: currentToolResults, documentStructure: freshDocStructure }),
      signal: abortSignal,
    });

    currentToolResults = [];

    const events = await parseSSEResponse(res, onDelta);

    let hasToolCall = false;
    for (const event of events) {
      if (event.type === "usage") {
        applyUsageEvent(event);
      }

      if (event.type === "tool_call") {
        hasToolCall = true;

        // Autopilot: execute ALL tools inline, regardless of type.
        if (state.autopilot) {
          const executedTools: Array<{ tool: string; params: Record<string, any> }> = [];
          for (const tool of event.tools) {
            appendToolCallCard(tool.tool, tool.params);
            const result = await executeSingleTool(tool);
            appendToolResultCard(tool.tool, result.success, result.result);
            currentToolResults.push(result);
            executedTools.push({ tool: tool.tool, params: tool.params });
          }
          await autoReadAfterMutation(executedTools, currentToolResults);
          continue;
        }

        const perceptionTools = ["read_document", "get_selection_info", "get_document_stats", "get_paragraph_format", "get_document_tables", "read_table"];
        for (const tool of event.tools) {
          if (!perceptionTools.includes(tool.tool)) {
            const plan: ActionPlan = {
              actions: event.tools.map((t) => ({
                action: t.tool,
                params: t.params,
                description: `Agent: ${t.tool}`,
                toolCallId: t.id,
              })),
              explanation: `Agent 循环中的操作: ${event.tools.map((t) => t.tool).join(", ")}`,
            };
            return { reply: finalReply || plan.explanation, actionPlan: plan, sessionId, retrievalCount: finalRetrievalCount, citations: finalCitations };
          }
          appendToolCallCard(tool.tool, tool.params);
          const result = await executeSingleTool(tool);
          appendToolResultCard(tool.tool, result.success, result.result);
          currentToolResults.push(result);
        }
      }

      // Iterable tool calls: execute action tools directly and continue loop
      if (event.type === "iterable_tool_call") {
        hasToolCall = true;
        const executedTools: Array<{ tool: string; params: Record<string, any> }> = [];
        for (const tool of event.tools) {
          appendToolCallCard(tool.tool, tool.params);
          const result = await executeSingleTool(tool);
          appendToolResultCard(tool.tool, result.success, result.result);
          currentToolResults.push(result);
          executedTools.push({ tool: tool.tool, params: tool.params });
        }
        // Auto-read after structure-mutating tools to refresh paragraph indices
        await autoReadAfterMutation(executedTools, currentToolResults);
      }

      if (event.type === "action_plan") {
        // Autopilot: run the plan inline and keep looping — never pause for
        // confirmation mid-run.
        if (state.autopilot && event.plan.actions.length > 0) {
          hasToolCall = true;
          const planResults = await executeActionPlan(event.plan);
          for (const r of planResults) currentToolResults.push(r);
          continue;
        }
        return { reply: finalReply || event.plan.explanation, actionPlan: event.plan, sessionId, retrievalCount: finalRetrievalCount, citations: finalCitations };
      }

      if (event.type === "task_complete") {
        return { reply: event.summary || finalReply, actionPlan: null, sessionId, retrievalCount: finalRetrievalCount, citations: finalCitations };
      }

      if (event.type === "done") {
        finalReply = event.reply ?? finalReply;
        if (state.autopilot && event.actionPlan && event.actionPlan.actions.length > 0) {
          hasToolCall = true;
          const planResults = await executeActionPlan(event.actionPlan);
          for (const r of planResults) currentToolResults.push(r);
        } else {
          finalActionPlan = event.actionPlan ?? finalActionPlan;
        }
        if (event.retrievalCount !== undefined) finalRetrievalCount = event.retrievalCount;
        if (event.citations) finalCitations = event.citations;
      }
    }

    if (!hasToolCall || currentToolResults.length === 0) {
      break;
    }
  }

  return { reply: finalReply.trim(), actionPlan: finalActionPlan, sessionId, retrievalCount: finalRetrievalCount, citations: finalCitations };
}

async function sendMessageStream(
  payload: {
    messages: ChatMessage[];
    documentContext: string;
    documentStructure?: DocumentStructure;
    selection: string;
    insertMode: InsertMode;
  },
  onDelta: (text: string) => void
): Promise<{
  reply: string;
  actionPlan: ActionPlan | null;
  retrievalCount?: number;
  citations?: Array<{ fileName: string }>;
  sessionId?: string;
}> {
  // Share the current abort controller so the Stop button can tear this down.
  currentAbortController = new AbortController();

  try {
    const res = await fetch(`${agentBase}/v1/chat/stream`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: currentAbortController.signal,
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      throw new Error(err);
    }

    if (!res.body) {
      throw new Error("流式响应不可用");
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";
    let fullReply = "";
    let donePayload: { retrievalCount?: number; citations?: Array<{ fileName: string }>; actionPlan?: ActionPlan | null } = {};

    while (true) {
      if (currentAbortController.signal.aborted) {
        break;
      }
      const { done, value } = await reader.read();
      if (done) {
        break;
      }

      buffer += decoder.decode(value, { stream: true });
      const events = buffer.split("\n\n");
      buffer = events.pop() ?? "";

      for (const event of events) {
        const line = event
          .split("\n")
          .find((item) => item.trim().startsWith("data:"));
        if (!line) {
          continue;
        }

        const raw = line.slice(5).trim();
        if (!raw) {
          continue;
        }

        let data: StreamEvent;
        try {
          data = JSON.parse(raw) as StreamEvent;
        } catch {
          continue;
        }

        if (data.type === "delta" && data.delta) {
          fullReply += data.delta;
          onDelta(fullReply);
        }

        if (data.type === "usage") {
          applyUsageEvent(data);
        }

        if (data.type === "error") {
          throw new Error(data.error || "流式响应出错");
        }

        if (data.type === "done") {
          fullReply = data.reply ?? fullReply;
          donePayload = {
            retrievalCount: data.retrievalCount,
            citations: data.citations,
            actionPlan: data.actionPlan,
          };
        }
      }
    }

    return { reply: fullReply.trim(), actionPlan: donePayload.actionPlan ?? null, ...donePayload };
  } finally {
    // Controller lifecycle is owned by setRunningState() — nothing to release here.
  }
}

/**
 * Stop any in-flight agent run.
 *
 * Two layers of cancellation:
 *  1. `currentAbortController.abort()` — immediately tears down the active
 *     fetch/SSE stream so the UI unblocks within a few hundred ms.
 *  2. `POST /v1/chat/abort` with the pending sessionId — removes the session
 *     on the server, so any stale agent-continue call (e.g. one the abort
 *     raced past) cannot keep iterating and re-hit the model.
 *
 * The UI swap to "send" is handled by sendMessage's catch/finally (it
 * detects AbortError and re-enters idle state). When the user clicks Stop
 * outside of a sendMessage flow (e.g. during confirmActionPlan), this
 * function also force-clears the running flag.
 */
async function stopAgentRun(): Promise<void> {
  const sessionIdToKill = state.pendingSessionId;
  state.pendingSessionId = null;

  // 1) Abort the in-flight fetch (and any pending /v1/chat/agent-continue
  //    calls in the same loop). Safe to call even if the controller is null.
  currentAbortController?.abort();

  // 2) Best-effort server-side cleanup. Network failure here is non-fatal:
  //    the local abort already breaks the loop, and the session will expire
  //    after its 30-minute TTL if the server call fails.
  if (sessionIdToKill) {
    try {
      await fetch(`${agentBase}/v1/chat/abort`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sessionId: sessionIdToKill }),
      });
    } catch {
      // ignore — local abort is enough
    }
  }

  // 3) Force UI back to idle if sendMessage's catch/finally hasn't already
  //    done so (e.g. user clicked Stop during confirmActionPlan).
  setRunningState(false);
  setStatus(chatStatus, "已停止生成");
}

// ─── Send Message ──────────────────────────────────────────────────────────

async function sendMessage(): Promise<void> {
  if (state.isAgentRunning) {
    setStatus(chatStatus, "正在生成中，请先点 ⏹ 停止");
    return;
  }

  const text = userInput.value.trim();
  if (!text && !state.pendingAttachment) {
    setStatus(chatStatus, "请输入消息");
    return;
  }

  const insertMode = insertModeSelect.value as InsertMode;

  // Ensure we have an active session
  if (!state.activeSessionId) {
    createNewSession();
  }

  const session = getActiveSession();
  if (!session) {
    createNewSession();
  }

  userInput.value = "";
  resizeComposerInput();

  // Build display message (with attachment indicator)
  let displayText = text;
  if (state.pendingAttachment) {
    const attachInfo = `[附件: ${state.pendingAttachment.fileName}]`;
    displayText = text ? `${text}\n${attachInfo}` : attachInfo;
  }
  appendMessage("user", displayText);

  // Build the actual prompt sent to LLM (with file content injected)
  let llmPrompt = text;
  if (state.pendingAttachment) {
    const fileBlock = `\n\n--- 以下是附加文件 "${state.pendingAttachment.fileName}" 的内容 ---\n${state.pendingAttachment.content}\n--- 文件内容结束 ---`;
    llmPrompt = llmPrompt ? `${llmPrompt}${fileBlock}` : `请阅读以下附加文件内容：${fileBlock}`;
    hideAttachmentPreview();
  }
  if (insertMode !== "chat_only") {
    llmPrompt = `${llmPrompt}\n\n请根据文档结构和用户意图选择合适的工具来操作文档。`;
  }

  // Store original user message in session
  const activeSession = getActiveSession()!;
  // Batch A: snapshot the conversation before this send so the user can
  // roll the round back (see renderCheckpointPanel).
  addCheckpoint(activeSession, llmPrompt);
  activeSession.messages.push({ role: "user", content: llmPrompt });
  trimSessionMessages(activeSession);
  autoTitleFromFirstMessage(activeSession);
  void generateSessionTitle(activeSession);

  setStatus(chatStatus, "思考中（流式返回）...");

  // Get structured document context
  const wordStructure = await getStructuredContext();
  const wordContext = await getWordContext();

  // Show selection info if available
  if (wordStructure.selection.text) {
    appendMessage("system", `已读取选区（${wordStructure.selection.text.length} 字）`);
  }

  setRunningState(true);

  try {
    const assistantEl = createAssistantMessage("");

    // Use Agent Loop for smart_action mode, fallback to legacy stream for others
    const useAgentLoop = insertMode === "smart_action";

    // Build messages for LLM: full session history with enriched last user message
    const messagesForLlm: ChatMessage[] = activeSession.messages.map((m) => ({ ...m }));
    const lastIdx = messagesForLlm.length - 1;
    if (lastIdx >= 0 && messagesForLlm[lastIdx].role === "user") {
      messagesForLlm[lastIdx] = { role: "user", content: llmPrompt };
    } else {
      messagesForLlm.push({ role: "user", content: llmPrompt });
    }

    const maxIter = cachedProviderConfig?.maxIterations ?? (Number($<HTMLInputElement>("maxIterations").value) || 10);
    const data = useAgentLoop
      ? await sendAgentMessageStream(
          {
            messages: messagesForLlm,
            documentContext: wordContext.documentContext,
            documentStructure: wordStructure,
            selection: wordContext.selection,
            insertMode,
          },
          (partial) => {
            updateAssistantMessage(assistantEl, partial);
          },
          maxIter
        )
      : await sendMessageStream(
          {
            messages: messagesForLlm,
            documentContext: wordContext.documentContext,
            documentStructure: wordStructure,
            selection: wordContext.selection,
            insertMode,
          },
          (partial) => {
            updateAssistantMessage(assistantEl, partial);
          }
        );

    state.pendingSessionId = data.sessionId ?? null;
    console.log("[sendMessage] pendingSessionId set to:", state.pendingSessionId, "actionPlan:", data.actionPlan ? `actions=${data.actionPlan.actions.length}` : "null");

    state.lastReply = data.reply;
    activeSession.messages.push({ role: "assistant", content: data.reply });
    trimSessionMessages(activeSession);
    activeSession.updatedAt = Date.now();

    const citationText = data.citations?.length
      ? `\n\n来源: ${data.citations.map((c) => c.fileName).join(", ")}（命中 ${data.retrievalCount ?? data.citations.length} 段）`
      : "";

    // Handle action plan from LLM
    if (data.actionPlan && data.actionPlan.actions.length > 0) {
      updateAssistantMessage(assistantEl, `${data.reply}${citationText}\n\n📋 操作计划：${data.actionPlan.explanation}`);
      showActionPlanPreview(data.actionPlan);
    } else {
      updateAssistantMessage(assistantEl, `${data.reply}${citationText}`);

      // Fallback: if no action plan, handle simple insert modes
      if (insertMode === "replace_selection") {
        const cleanText = cleanupMarkdownForWord(data.reply);
        showSimplePreview(cleanText, "replace");
      } else if (insertMode === "append_end") {
        const cleanText = cleanupMarkdownForWord(data.reply);
        showSimplePreview(cleanText, "append_end");
      } else {
        setStatus(chatStatus, "已完成");
      }
    }

    saveSessionsToStorage();
  } catch (error) {
    if (isAbortError(error)) {
      // Aborted by user via ⏹ Stop. The server-side session was already
      // removed by /v1/chat/abort, so no further tool calls will run.
      setStatus(chatStatus, "已停止生成");
      state.pendingSessionId = null;
    } else {
      const message = error instanceof Error ? error.message : "未知错误";
      setStatus(chatStatus, `请求失败: 无法连接本地 Agent (${message})`);
    }
  } finally {
    removeTypingBubbles();
    setRunningState(false);
  }
}

// ─── Retry Last Message ─────────────────────────────────────────────────────

async function retryLastMessage(): Promise<void> {
  const session = getActiveSession();
  if (!session || session.messages.length === 0) {
    setStatus(chatStatus, "没有可重试的消息");
    return;
  }

  // Find the last user message
  const lastUserIdx = session.messages.map((m, i) => m.role === "user" ? i : -1).filter((i) => i >= 0).pop();
  if (lastUserIdx === undefined || lastUserIdx < 0) {
    setStatus(chatStatus, "没有可重试的用户消息");
    return;
  }

  const lastUserContent = session.messages[lastUserIdx].content;

  // Remove the last assistant message if it exists after the last user message
  if (lastUserIdx + 1 < session.messages.length && session.messages[lastUserIdx + 1].role === "assistant") {
    session.messages.splice(lastUserIdx + 1);
  }

  // Set the input and send
  userInput.value = lastUserContent;
  await sendMessage();
}

// ─── Bind Actions ────────────────────────────────────────────────────────────

function bindActions(): void {
  // Config
  $("saveConfig").addEventListener("click", () => {
    void saveProviderConfig();
  });

  $("refreshModels").addEventListener("click", () => {
    void refreshModelList();
  });

  // Thinking mode: toggle sub-options visibility
  $("enableThinking").addEventListener("change", () => {
    updateThinkingOptionsVisibility($<HTMLInputElement>("enableThinking").checked);
  });

  // Config presets
  $("presetSelect").addEventListener("change", () => {
    const id = $<HTMLSelectElement>("presetSelect").value;
    if (id) void loadPresetById(id);
  });
  $("loadPreset").addEventListener("click", () => {
    const id = $<HTMLSelectElement>("presetSelect").value;
    if (id) void loadPresetById(id);
  });
  $("savePreset").addEventListener("click", () => {
    void saveCurrentAsPreset();
  });
  $("deletePreset").addEventListener("click", () => {
    void deleteSelectedPreset();
  });

  // Knowledge base
  $("uploadFile").addEventListener("click", () => {
    void uploadKnowledgeFile();
  });

  $("kbFile").addEventListener("change", syncKbFileLabel);

  $("refreshKb").addEventListener("click", () => {
    void loadKbStats();
    void loadKbFileList();
  });

  $("clearKb").addEventListener("click", () => {
    void clearKb();
  });

  $("exportKb").addEventListener("click", () => {
    void exportKb();
  });

  $("importKbFile").addEventListener("change", () => {
    void importKb();
  });

  // Menu toggle (sidebar)
  const sidebar = document.getElementById("menuSidebar")!;
  const overlay = document.getElementById("sidebarOverlay")!;
  const sidebarClose = document.getElementById("sidebarClose")!;

  function openSidebar() {
    sidebar.classList.add("open");
    overlay.classList.add("active");
  }
  function closeSidebar() {
    sidebar.classList.remove("open");
    overlay.classList.remove("active");
  }

  $("menuToggle").addEventListener("click", openSidebar);
  sidebarClose.addEventListener("click", closeSidebar);
  overlay.addEventListener("click", closeSidebar);

  // Esc closes the settings sheet (matches most island/sheet UIs).
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && sidebar.classList.contains("open")) {
      closeSidebar();
    }
  });

  // Chat
  $("sendMsg").addEventListener("click", () => {
    void sendMessage();
  });

  // Stop button: tear down any in-flight agent loop and clear server session.
  $("stopBtn").addEventListener("click", () => {
    void stopAgentRun();
  });

  // Empty-state starter chips prefill the composer (never auto-send).
  document.querySelectorAll<HTMLButtonElement>(".starter").forEach((starter) => {
    starter.addEventListener("click", () => {
      userInput.value = starter.dataset.prompt ?? "";
      resizeComposerInput();
      userInput.focus();
    });
  });

  // Allow Enter to send (Shift+Enter for newline)
  userInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      void sendMessage();
    }
  });

  // Composer grows with the draft, up to the CSS max-height.
  userInput.addEventListener("input", resizeComposerInput);

  // Insert actions
  $("retryLast").addEventListener("click", () => {
    void retryLastMessage();
  });

  // Autopilot toggle — shows a confirmation modal before engaging auto-execute mode
  const autopilotCheckbox = $<HTMLInputElement>("autopilotToggle");
  const autopilotModal = $<HTMLDivElement>("autopilotModal");
  const autopilotModalConfirm = $<HTMLButtonElement>("autopilotConfirm");
  const autopilotModalCancel = $<HTMLButtonElement>("autopilotCancel");

  autopilotCheckbox.addEventListener("change", () => {
    if (autopilotCheckbox.checked) {
      // Show the custom modal instead of native confirm()
      autopilotModal.style.display = "";

      const onConfirm = () => {
        autopilotModal.style.display = "none";
        autopilotModalConfirm.removeEventListener("click", onConfirm);
        autopilotModalCancel.removeEventListener("click", onCancel);
        state.autopilot = true;
        setStatus(chatStatus, "🚀 Autopilot 已开启 — 所有操作将自动执行");
      };

      const onCancel = () => {
        autopilotModal.style.display = "none";
        autopilotModalConfirm.removeEventListener("click", onConfirm);
        autopilotModalCancel.removeEventListener("click", onCancel);
        autopilotCheckbox.checked = false;
        setStatus(chatStatus, "");
      };

      autopilotModalConfirm.addEventListener("click", onConfirm);
      autopilotModalCancel.addEventListener("click", onCancel);
    } else {
      state.autopilot = false;
      setStatus(chatStatus, "Autopilot 已关闭");
    }
  });

  // Simple preview (legacy)
  $("previewConfirm").addEventListener("click", () => {
    void confirmSimplePreview();
  });

  $("previewCancel").addEventListener("click", () => {
    hideSimplePreview();
    setStatus(chatStatus, "已取消插入");
  });

  // Action plan preview
  $("actionPlanConfirm").addEventListener("click", () => {
    void confirmActionPlan();
  });

  $("actionPlanCancel").addEventListener("click", () => {
    hideActionPlanPreview();
    setStatus(chatStatus, "已取消操作计划");
  });

  // Session management
  $("newSession").addEventListener("click", () => {
    createNewSession();
  });

  $("clearSession").addEventListener("click", () => {
    clearCurrentSession();
  });

  $("deleteSession").addEventListener("click", () => {
    deleteCurrentSession();
  });
  $("checkpointBtn").addEventListener("click", () => {
    renderCheckpointPanel();
    checkpointPanel.style.display = checkpointPanel.style.display === "none" ? "" : "none";
  });
  $("checkpointClose").addEventListener("click", () => {
    checkpointPanel.style.display = "none";
  });
  checkpointList.addEventListener("click", (e) => {
    const target = (e.target as HTMLElement).closest("button") as HTMLButtonElement | null;
    if (!target) return;
    const id = target.dataset.id ?? "";
    if (target.classList.contains("cp-undo")) {
      void rollbackCheckpointUndo(id);
    } else if (target.classList.contains("cp-restore")) {
      restoreCheckpointMessages(id);
    }
  });
  $("refreshStats").addEventListener("click", () => {
    void loadStatsView();
  });

  sessionSelect.addEventListener("change", () => {
    const selectedId = sessionSelect.value;
    if (selectedId && selectedId !== state.activeSessionId) {
      switchToSession(selectedId);
    }
  });
}

// ─── Batch A: auto session titles ───────────────────────────────────────────

/**
 * Generates a short semantic title for the session dropdown (batch A).
 * Best-effort and one-shot per session — the truncation fallback set by
 * autoTitleFromFirstMessage stays when this fails or no provider is set.
 */
async function generateSessionTitle(session: Session): Promise<void> {
  if (session.titleGenerated) return;
  session.titleGenerated = true;
  const firstUserMsg = session.messages.find((m) => m.role === "user");
  if (!firstUserMsg) return;
  try {
    const res = await fetch(`${agentBase}/v1/chat/title`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message: firstUserMsg.content.slice(0, 400) }),
    });
    if (!res.ok) return;
    const data = await res.json();
    const title = typeof data?.title === "string" ? data.title.trim() : "";
    if (title) {
      session.title = title;
      saveSessionsToStorage();
      renderSessionSelect();
    }
  } catch {
    // keep the truncation fallback title
  }
}

// ─── Batch A: rollback checkpoints ─────────────────────────────────────────

const checkpointPanel = $<HTMLDivElement>("checkpointPanel");
const checkpointList = $<HTMLDivElement>("checkpointList");

/** Tools that change the document (their execution maps to one undo step). */
const UNDOABLE_TOOLS = new Set([
  "insert_after_heading", "insert_at_cursor", "replace_selection", "insert_at_end",
  "insert_at_start", "insert_after_paragraph", "delete_paragraph", "find_and_replace",
  "find_and_replace_v2", "replace_paragraph", "set_paragraph_style", "merge_paragraphs",
  "apply_rich_format", "delete_table", "insert_table", "insert_mermaid_image",
]);

function addCheckpoint(session: Session, userMessage: string): void {
  const list = (session.checkpoints ??= []);
  list.push({
    id: generateSessionId(),
    time: Date.now(),
    label: userMessage.trim().slice(0, 40).replace(/\n/g, " ") || "（空消息）",
    messages: JSON.parse(JSON.stringify(session.messages)) as ChatMessage[],
    undoCount: 0,
  });
  while (list.length > 10) list.shift();
  saveSessionsToStorage();
}

function trackCheckpointUndo(toolName: string, params: Record<string, any>, success: boolean): void {
  if (!success) return;
  const session = getActiveSession();
  const cp = session?.checkpoints?.[session.checkpoints.length - 1];
  if (!cp || cp.undone) return;
  if (toolName === "undo_last_action") {
    // Agent undid its own edits — fewer steps left to roll back.
    cp.undoCount = Math.max(0, cp.undoCount - (Number(params?.count) || 1));
    return;
  }
  if (UNDOABLE_TOOLS.has(toolName)) {
    cp.undoCount += 1;
    saveSessionsToStorage();
  }
}

async function rollbackCheckpointUndo(cpId: string): Promise<void> {
  const session = getActiveSession();
  const cp = session?.checkpoints?.find((c) => c.id === cpId);
  if (!cp || cp.undone || cp.undoCount <= 0) return;
  setRunningState(true);
  try {
    await executeAction({ action: "undo_last_action", params: { count: cp.undoCount }, description: "回退检查点文档改动" });
    cp.undone = true;
    cp.undoCount = 0;
    saveSessionsToStorage();
    setStatus(chatStatus, "已撤销本轮 agent 的文档改动");
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(chatStatus, `撤销失败: ${message}`);
  } finally {
    setRunningState(false);
    renderCheckpointPanel();
  }
}

function restoreCheckpointMessages(cpId: string): void {
  const session = getActiveSession();
  const cp = session?.checkpoints?.find((c) => c.id === cpId);
  if (!session || !cp) return;
  session.messages = JSON.parse(JSON.stringify(cp.messages)) as ChatMessage[];
  renderSessionMessages(session.messages);
  saveSessionsToStorage();
  checkpointPanel.style.display = "none";
  setStatus(chatStatus, "已回到该检查点（对话已回退）");
}

function renderCheckpointPanel(): void {
  const session = getActiveSession();
  const list = session?.checkpoints ?? [];
  if (list.length === 0) {
    checkpointList.innerHTML = "<p class='status'>本次会话还没有检查点（发送一条消息后自动生成）。</p>";
    return;
  }
  checkpointList.innerHTML = list
    .slice()
    .reverse()
    .map((cp) => {
      const time = new Date(cp.time).toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" });
      const undoLabel = cp.undone ? "已撤销" : `${cp.undoCount} 处`;
      const undoDisabled = cp.undone || cp.undoCount <= 0 ? "disabled" : "";
      return `
        <div class="checkpoint-item">
          <div class="checkpoint-title">${escapeHTML(`${time} · ${cp.label}`)}</div>
          <div class="checkpoint-actions">
            <button class="secondary cp-undo" data-id="${cp.id}" ${undoDisabled}>撤销文档改动（${undoLabel}）</button>
            <button class="secondary cp-restore" data-id="${cp.id}">回到发送前</button>
          </div>
        </div>`;
    })
    .join("");
}

// ─── Batch A: usage stats ──────────────────────────────────────────────────

const statTodayTokens = $<HTMLSpanElement>("statTodayTokens");
const statTotalTokens = $<HTMLSpanElement>("statTotalTokens");
const statLlmCalls = $<HTMLSpanElement>("statLlmCalls");
const statToolCalls = $<HTMLSpanElement>("statToolCalls");
const statDetail = $<HTMLParagraphElement>("statDetail");
const statToolList = $<HTMLDivElement>("statToolList");
const accountInfo = $<HTMLParagraphElement>("accountInfo");

function formatTokens(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`;
  if (n >= 1_000) return `${(n / 1_000).toFixed(1)}k`;
  return String(n);
}

function reportToolEvent(name: string, success: boolean, ms: number): void {
  void fetch(`${agentBase}/v1/stats/report`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ events: [{ name, success, ms }] }),
  }).catch(() => {
    // statistics must never break tool execution
  });
}

async function loadStatsView(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/stats`);
    if (!res.ok) return;
    const s = await res.json();
    statTodayTokens.textContent = formatTokens((s.today?.promptTokens ?? 0) + (s.today?.completionTokens ?? 0));
    statTotalTokens.textContent = formatTokens((s.total?.promptTokens ?? 0) + (s.total?.completionTokens ?? 0));
    statLlmCalls.textContent = String(s.total?.llmCalls ?? 0);
    statToolCalls.textContent = String(s.total?.toolCalls ?? 0);
    statDetail.textContent =
      `今日：入 ${formatTokens(s.today?.promptTokens ?? 0)} / 出 ${formatTokens(s.today?.completionTokens ?? 0)} tokens · 工具 ${s.today?.toolCalls ?? 0} 次` +
      `；累计工具失败 ${s.total?.toolFailures ?? 0} 次`;

    const ranking: Array<{ name: string; calls: number; avgMs: number }> = s.toolRanking ?? [];
    statToolList.innerHTML =
      ranking.length === 0
        ? "<p class='status'>还没有工具调用记录。</p>"
        : ranking
            .map((t) => `<div class="tool-rank-row"><strong>${escapeHTML(t.name)}</strong><span>${t.calls} 次 · 均 ${t.avgMs}ms</span></div>`)
            .join("");
    void renderAccountInfo();
  } catch {
    // stats unavailable — leave the panel as-is
  }
}

async function renderAccountInfo(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/provider/config`);
    if (!res.ok) return;
    const cfg = await res.json();
    accountInfo.textContent =
      `当前模型：${cfg.model ?? "未配置"} · API Key ${cfg.hasApiKey ? "已配置（已隐藏）" : "未配置"} · 方案 ${cfg.presetCount ?? 0} 个`;
  } catch {
    accountInfo.textContent = "无法读取账户信息";
  }
}

// ─── Initialize ──────────────────────────────────────────────────────────────

function initApp(): void {
  installAuthFetch();
  loadSessionsFromStorage();

  // Ensure at least one session exists
  if (state.sessions.length === 0) {
    createNewSession();
  } else {
    // Switch to the first session
    switchToSession(state.sessions[0].id);
  }

  bindActions();
  setupAttachmentHandlers();
  void loadProviderConfig();
  void loadPresets();
  void loadKbStats();
  void loadKbFileList();
  void loadStatsView();
  setStatus(chatStatus, "就绪：可直接开始对话");
  resizeComposerInput();
}

// Boot shim: the legacy Office.js host waits for Office.onReady; the desktop
// WebView2 host and plain browser initialize immediately.
if (typeof Office !== "undefined") {
  Office.onReady(() => initApp());
} else {
  initApp();
}
