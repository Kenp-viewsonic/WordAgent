import cors from "cors";
import express from "express";
import multer from "multer";
import path from "node:path";
import mammoth from "mammoth";
import { z } from "zod";
import { createSession, getSession, appendToSession, deleteSession } from "./agent-session.js";
import { callOpenAICompatible, describeDocumentStructure, listOpenAICompatibleModels, LlmHttpError, parseActionPlanFromToolCalls, streamOpenAICompatible, WORD_TOOLS } from "./llm.js";
import { executeTurn, buildAssistantMessage, detectDoomLoop, fingerprintToolCall, classifyToolCalls, TurnResult } from "./tool-runtime.js";
import { keywordRetrieve, splitToChunks } from "./retrieval.js";
import { appendChunks, clearAllChunks, deleteChunksByFile, getKbFileList, getKbStats, importChunks, loadChunks, loadProviderConfig, saveProviderConfig, loadPresets, savePreset, deletePreset } from "./store.js";
import { ActionPlan, ChatMessage, DocumentStructure, ProviderConfig, ConfigPreset, RetrievalChunk } from "./types.js";
import { logHttpRequest, logToolCall, logToolResult, logIterationEnd, logError } from "./agent-logger.js";
import crypto from "node:crypto";
import os from "node:os";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { appRoot, configRoot, ensureDataRoot, tokenFile } from "./paths.js";

/**
 * Safely parse a tool_call arguments string.
 *
 * DeepSeek (especially deepseek-v4-flash) occasionally returns malformed
 * arguments JSON — e.g. internal tokens like `<｜｜DSML｜｜` leaking into the
 * output, or truncated streaming fragments.  A bare JSON.parse in that case
 * throws and crashes the entire SSE response ("Unexpected end of JSON input"
 * on the client).  This helper never throws: on parse failure it returns
 * `{ raw_arguments: <original string> }` so the tool call can still be
 * surfaced to the user / logged for debugging.
 */
function safeParseToolArgs(raw: string | undefined | null): Record<string, any> {
  if (!raw) return {};
  try {
    return JSON.parse(raw) as Record<string, any>;
  } catch {
    return { raw_arguments: raw };
  }
}

const app = express();
const upload = multer({ limits: { fileSize: Number(process.env.MAX_UPLOAD_MB ?? 20) * 1024 * 1024 } });

const host = process.env.AGENT_BIND_HOST ?? "127.0.0.1";
const port = Number(process.env.PORT ?? 8787);

// ─── Security hardening (WP1 / D5): token auth + CORS whitelist ─────────────
// Token is generated on first start and persisted under
// %LOCALAPPDATA%\OfficeAgent\agent.token. Every /v1/* call must carry it in
// the X-Agent-Token header; /health stays open so the tray can probe the
// service without secrets.
function bootstrapAgentToken(): string {
  try {
    const existing = readFileSync(tokenFile, "utf-8").trim();
    if (existing) return existing;
  } catch {
    // no token yet — generate below
  }
  const token = crypto.randomBytes(32).toString("hex");
  mkdirSync(configRoot, { recursive: true });
  writeFileSync(tokenFile, token, { encoding: "utf-8", mode: 0o600 });
  try {
    // Best-effort ACL lock-down to the current user; %LOCALAPPDATA% is
    // already user-private, this tightens it further.
    spawnSync(
      "icacls",
      [configRoot, "/inheritance:r", "/grant:r", `${os.userInfo().username}:(OI)(CI)F`],
      { windowsHide: true }
    );
  } catch {
    // icacls unavailable — profile ACLs still apply
  }
  return token;
}

const agentToken = bootstrapAgentToken();

function isTokenValid(provided: string | undefined): boolean {
  if (!provided) return false;
  const a = Buffer.from(provided);
  const b = Buffer.from(agentToken);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

// CORS: fixed whitelist — never reflect arbitrary origins (was `origin: true`,
// which let any web page read /v1/kb/export etc.).
const corsOrigins = new Set(
  (process.env.AGENT_CORS_ORIGINS ??
    "http://127.0.0.1:8787,http://localhost:8787,http://localhost:3001,http://127.0.0.1:3001")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
);
app.use(cors({ origin: (origin, cb) => cb(null, !origin || corsOrigins.has(origin)) }));
app.use(express.json({ limit: "2mb" }));

// Serve the UI (built into public/ by scripts/package.ps1). index.html goes
// through the injector below so the same-origin page receives the agent token.
const publicDir = path.join(appRoot, "public");
app.use(express.static(publicDir, { index: false }));

// Auth gate for the API surface (D5). /health is intentionally open.
app.use((req, res, next) => {
  if (!req.path.startsWith("/v1/")) return next();
  if (isTokenValid(req.header("x-agent-token"))) return next();
  res.status(401).json({ error: "unauthorized", message: "missing or invalid X-Agent-Token" });
});

void ensureDataRoot();

app.get("/health", (_req, res) => {
  res.json({ ok: true, service: "office-agent-local", host, port });
});

function isUtf8TextExt(fileName: string): boolean {
  const ext = path.extname(fileName).toLowerCase();
  return [
    ".txt",
    ".md",
    ".markdown",
    ".json",
    ".csv",
    ".log",
    ".xml",
    ".html",
    ".htm",
    ".ts",
    ".tsx",
    ".js",
    ".jsx",
    ".py",
    ".java",
    ".c",
    ".cpp",
    ".h",
    ".hpp",
    ".go",
    ".rs",
    ".sql",
    ".yaml",
    ".yml",
  ].includes(ext);
}

async function extractUploadText(file: Express.Multer.File): Promise<string> {
  const lowerName = file.originalname.toLowerCase();
  if (lowerName.endsWith(".docx")) {
    const result = await mammoth.extractRawText({ buffer: file.buffer });
    return result.value;
  }

  if (file.mimetype.startsWith("text/") || isUtf8TextExt(file.originalname)) {
    return file.buffer.toString("utf-8");
  }

  throw new Error(`暂不支持该文件类型: ${file.originalname} (${file.mimetype || "unknown"})`);
}

const configSchema = z.object({
  baseUrl: z.string().url(),
  apiKey: z.string().min(1).optional(),
  model: z.string().min(1),
  temperature: z.number().min(0).max(2).optional(),
  maxTokens: z.number().int().positive().optional(),
  firstTokenTimeout: z.number().int().min(5).max(120).optional(),
  overallTimeout: z.number().int().min(30).max(600).optional(),
  enableThinking: z.boolean().optional(),
  includeReasoningContent: z.boolean().optional(),
  thinkingEffort: z.enum(["medium", "high"]).optional(),
  thinkingFormat: z.enum(["deepseek", "openai"]).optional(),
  maxIterations: z.number().int().min(1).max(50).optional(),
});

app.post("/v1/provider/config", async (req, res) => {
  const result = configSchema.safeParse(req.body);
  if (!result.success) {
    return res.status(400).json({ error: "Invalid provider config", details: result.error.flatten() });
  }

  const current = await loadProviderConfig();
  const resolvedApiKey = result.data.apiKey || current?.apiKey;
  if (!resolvedApiKey) {
    return res.status(400).json({ error: "API Key is required for first-time save." });
  }

  // Merge: preserve existing thinking config when not provided in request
  const merged: ProviderConfig = {
    ...current,
    ...result.data,
    apiKey: resolvedApiKey,
  };
  await saveProviderConfig(merged);
  return res.json({ ok: true });
});

app.get("/v1/provider/config", async (_req, res) => {
  const config = await loadProviderConfig();
  if (!config) {
    return res.status(404).json({ error: "Provider config not found" });
  }

  return res.json({
    baseUrl: config.baseUrl,
    model: config.model,
    temperature: config.temperature,
    maxTokens: config.maxTokens,
    firstTokenTimeout: config.firstTokenTimeout,
    overallTimeout: config.overallTimeout,
    hasApiKey: Boolean(config.apiKey),
    enableThinking: config.enableThinking ?? false,
    includeReasoningContent: config.includeReasoningContent ?? true,
    thinkingEffort: config.thinkingEffort ?? "high",
    thinkingFormat: config.thinkingFormat ?? "deepseek",
    maxIterations: config.maxIterations ?? 10,
  });
});

// ─── Config Presets ─────────────────────────────────────────────────────────

const presetSchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1),
  config: z.object({
    baseUrl: z.string().url(),
    apiKey: z.string().min(1).optional(),
    model: z.string().min(1),
    temperature: z.number().min(0).max(2).optional(),
    maxTokens: z.number().int().positive().optional(),
    firstTokenTimeout: z.number().int().min(5).max(120).optional(),
    overallTimeout: z.number().int().min(30).max(600).optional(),
    enableThinking: z.boolean().optional(),
    includeReasoningContent: z.boolean().optional(),
    thinkingEffort: z.enum(["medium", "high"]).optional(),
    thinkingFormat: z.enum(["deepseek", "openai"]).optional(),
    maxIterations: z.number().int().min(1).max(50).optional(),
  }),
});

app.get("/v1/presets", async (_req, res) => {
  const presets = await loadPresets();
  // Mask API keys in response
  const safe = presets.map((p) => ({
    ...p,
    config: { ...p.config, apiKey: "***" },
  }));
  return res.json({ ok: true, presets: safe });
});

app.post("/v1/presets", async (req, res) => {
  const result = presetSchema.safeParse(req.body);
  if (!result.success) {
    return res.status(400).json({ error: "Invalid preset", details: result.error.flatten() });
  }
  // If apiKey not provided in preset, inherit from current config
  const presetData = result.data;
  if (!presetData.config.apiKey) {
    const current = await loadProviderConfig();
    if (current?.apiKey) {
      presetData.config.apiKey = current.apiKey;
    } else {
      return res.status(400).json({ error: "API Key is required. Please save a config with API Key first, or provide apiKey in the preset." });
    }
  }
  const presets = await savePreset(presetData as ConfigPreset);
  const safe = presets.map((p) => ({
    ...p,
    config: { ...p.config, apiKey: "***" },
  }));
  return res.json({ ok: true, presets: safe });
});

app.delete("/v1/presets/:id", async (req, res) => {
  const presets = await deletePreset(req.params.id);
  const safe = presets.map((p) => ({
    ...p,
    config: { ...p.config, apiKey: "***" },
  }));
  return res.json({ ok: true, presets: safe });
});

app.post("/v1/presets/:id/activate", async (req, res) => {
  const presets = await loadPresets();
  const preset = presets.find((p) => p.id === req.params.id);
  if (!preset) {
    return res.status(404).json({ error: "Preset not found" });
  }
  await saveProviderConfig(preset.config);
  return res.json({ ok: true });
});

app.get("/v1/provider/models", async (req, res) => {
  let config = await loadProviderConfig();

  // Fallback: if no saved config, try query parameters from the frontend
  if (!config) {
    const baseUrl = typeof req.query.baseUrl === "string" ? req.query.baseUrl.trim() : "";
    const apiKey = typeof req.query.apiKey === "string" ? req.query.apiKey.trim() : "";
    if (!baseUrl || !apiKey) {
      return res.status(400).json({ error: "尚未保存模型配置。请先填写 Base URL 和 API Key 后刷新模型列表。" });
    }
    config = { baseUrl, apiKey, model: "" };
  }

  try {
    const models = await listOpenAICompatibleModels(config);
    return res.json({ ok: true, models, currentModel: config.model });
  } catch (error) {
    if (error instanceof LlmHttpError) {
      const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
      return res.status(safeStatus).json({ error: error.details || error.message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    return res.status(502).json({ error: message });
  }
});

app.get("/v1/kb/stats", async (_req, res) => {
  const stats = await getKbStats();
  return res.json({ ok: true, ...stats });
});

app.post("/v1/kb/upload", upload.single("file"), async (req, res) => {
  const file = req.file;
  if (!file) {
    return res.status(400).json({ error: "No file uploaded" });
  }

  let text: string;
  try {
    text = await extractUploadText(file);
  } catch (error) {
    const message = error instanceof Error ? error.message : "文件解析失败";
    return res.status(400).json({ error: message });
  }

  const chunks = splitToChunks(file.originalname, text);
  if (chunks.length === 0) {
    return res.status(400).json({ error: "文件未提取到可索引文本，请检查文件内容或编码。" });
  }

  await appendChunks(chunks);

  return res.json({ ok: true, fileName: file.originalname, chunkCount: chunks.length });
});

app.get("/v1/kb/files", async (_req, res) => {
  const files = await getKbFileList();
  return res.json({ ok: true, files });
});

app.delete("/v1/kb/files/:fileName", async (req, res) => {
  const fileName = decodeURIComponent(req.params.fileName);
  if (!fileName) {
    return res.status(400).json({ error: "fileName is required" });
  }

  const removed = await deleteChunksByFile(fileName);
  if (removed === 0) {
    return res.status(404).json({ error: `File not found: ${fileName}` });
  }

  return res.json({ ok: true, fileName, removedChunks: removed });
});

app.delete("/v1/kb/clear", async (_req, res) => {
  const removed = await clearAllChunks();
  return res.json({ ok: true, removedChunks: removed });
});

app.get("/v1/kb/export", async (_req, res) => {
  const chunks = await loadChunks();
  return res.json({ ok: true, chunks, exportedAt: new Date().toISOString() });
});

const importSchema = z.object({
  chunks: z.array(z.object({
    id: z.string(),
    fileName: z.string(),
    text: z.string(),
  })),
  mode: z.enum(["merge", "replace"]).optional(),
});

app.post("/v1/kb/import", async (req, res) => {
  const result = importSchema.safeParse(req.body);
  if (!result.success) {
    return res.status(400).json({ error: `Invalid import payload: ${formatZodError(result.error)}`, details: result.error.flatten() });
  }

  const { chunks, mode = "merge" } = result.data;

  try {
    const { importedChunks, totalChunks } = await importChunks(chunks, mode as "merge" | "replace");
    return res.json({ ok: true, importedChunks, totalChunks, mode });
  } catch (error) {
    const message = error instanceof Error ? error.message : "导入失败";
    return res.status(500).json({ error: message });
  }
});

// --- Chat Schema with document structure and insert mode ---

const documentStructureSchema = z.object({
  title: z.string(),
  totalParagraphs: z.number(),
  totalCharacters: z.number(),
  paragraphs: z.array(z.object({
    index: z.number(),
    text: z.string(),
    style: z.string(),
    headingLevel: z.number().optional(),
    isTable: z.boolean(),
    isList: z.boolean(),
    charCount: z.number().optional(),
    font: z.object({
      name: z.string().optional(),
      size: z.number().optional(),
      color: z.string().optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
    }).optional(),
  })),
  selection: z.object({
    text: z.string(),
    startParagraphIndex: z.number().optional(),
    endParagraphIndex: z.number().optional(),
  }),
});

const chatSchema = z.object({
  messages: z.array(z.object({ role: z.enum(["system", "user", "assistant", "tool"]), content: z.string().min(1), tool_call_id: z.string().optional(), name: z.string().optional() })),
  documentContext: z.string().optional(),
  documentStructure: documentStructureSchema.optional(),
  selection: z.string().optional(),
  insertMode: z.enum(["chat_only", "smart_action", "replace_selection", "append_end"]).optional(),
});

async function buildChatContext(payload: z.infer<typeof chatSchema>) {
  const provider = await loadProviderConfig();
  if (!provider) {
    throw new Error("Provider config missing. Please configure base_url/api_key/model first.");
  }

  const messages = payload.messages as ChatMessage[];
  const userQuery = [...messages].reverse().find((msg) => msg.role === "user")?.content ?? "";

  const chunks = await loadChunks();
  const retrieved = keywordRetrieve(chunks, userQuery, 4);

  const contextParts: string[] = [];
  if (payload.documentContext) {
    contextParts.push(`文档上下文:\n${payload.documentContext}`);
  }
  if (payload.selection) {
    contextParts.push(`当前选区:\n${payload.selection}`);
  }

  // Build document structure description for the system prompt
  let documentStructureDescription: string | undefined;
  if (payload.documentStructure) {
    documentStructureDescription = describeDocumentStructure(payload.documentStructure);
  }

  // NOTE: documentContext is only useful for chat_only mode where the LLM
  // cannot call read_document.  In agent mode the model can fetch the parts
  // it needs on demand — sending the full text in every request wastes
  // context window and confuses the model ("I need to scan everything").
  const insertMode = payload.insertMode || "smart_action";
  const dynamicContext: { documentContext?: string; selection?: string } = {};
  if (insertMode === "chat_only" && payload.documentContext) {
    dynamicContext.documentContext = payload.documentContext;
  }
  if (payload.selection) dynamicContext.selection = payload.selection;

  const enrichedMessages: ChatMessage[] = messages;

  return { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode, dynamicContext };
}

// --- Non-streaming chat endpoint ---

function formatZodError(err: z.ZodError): string {
  return err.issues.map((i) => `${i.path.join(".")}: ${i.message}`).join("; ");
}

app.post("/v1/chat", async (req, res) => {
  const parsed = chatSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: `Invalid chat payload: ${formatZodError(parsed.error)}`, details: parsed.error.flatten() });
  }

  try {
    const { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode, dynamicContext } = await buildChatContext(parsed.data);
    const result = await callOpenAICompatible(
      provider,
      enrichedMessages,
      retrieved,
      WORD_TOOLS,
      insertMode,
      documentStructureDescription,
      dynamicContext
    );
    return res.json({
      ok: true,
      reply: result.reply,
      actionPlan: result.actionPlan,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    });
  } catch (error) {
    if (error instanceof Error && error.message.startsWith("Provider config missing")) {
      return res.status(400).json({ error: error.message });
    }

    if (error instanceof LlmHttpError) {
      const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
      return res.status(safeStatus).json({ error: error.details || error.message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    return res.status(502).json({ error: message });
  }
});

// --- Streaming chat endpoint ---

app.post("/v1/chat/stream", async (req, res) => {
  const parsed = chatSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: `Invalid chat payload: ${formatZodError(parsed.error)}`, details: parsed.error.flatten() });
  }

  const abortController = new AbortController();
  req.on("aborted", () => {
    abortController.abort(new Error("client_aborted"));
  });
  res.on("close", () => {
    // Close can happen both on normal completion and client disconnect.
    if (!res.writableEnded && !abortController.signal.aborted) {
      abortController.abort(new Error("client_closed"));
    }
  });

  try {
    const { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode, dynamicContext } = await buildChatContext(parsed.data);

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();

    // Send an immediate event to confirm stream startup to the client.
    res.write(`data: ${JSON.stringify({ type: "start", ts: Date.now() })}\n\n`);

    let result: { reply: string; actionPlan: ActionPlan | null };
    try {
      result = await streamOpenAICompatible(
        provider,
        enrichedMessages,
        retrieved,
        (delta) => {
          res.write(`data: ${JSON.stringify({ type: "delta", delta })}\n\n`);
        },
        abortController.signal,
        WORD_TOOLS,
        insertMode,
        documentStructureDescription,
        dynamicContext
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      const shouldFallbackToNonStream = message.includes("first token not received");

      if (!shouldFallbackToNonStream) {
        throw error;
      }

      // Some OpenAI-compatible providers accept stream=true but don't emit SSE chunks.
      // Fall back to non-stream completion so client can still receive a final answer.
      result = await callOpenAICompatible(provider, enrichedMessages, retrieved, WORD_TOOLS, insertMode, documentStructureDescription, dynamicContext);
      res.write(`data: ${JSON.stringify({ type: "fallback", reason: "no_stream_delta" })}\n\n`);
    }

    res.write(`data: ${JSON.stringify({
      type: "done",
      reply: result.reply,
      actionPlan: result.actionPlan,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    })}\n\n`);
    return res.end();
  } catch (error) {
    if (abortController.signal.aborted) {
      return;
    }

    if (!res.headersSent) {
      if (error instanceof Error && error.message.startsWith("Provider config missing")) {
        return res.status(400).json({ error: error.message });
      }

      if (error instanceof LlmHttpError) {
        const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
        return res.status(safeStatus).json({ error: error.details || error.message });
      }

      const message = error instanceof Error ? error.message : "Unknown error";
      return res.status(502).json({ error: message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    res.write(`data: ${JSON.stringify({ type: "error", error: message })}\n\n`);
    return res.end();
  }
});

// --- Agent Loop Endpoints ---

const agentStreamSchema = chatSchema.extend({
  enableReAct: z.boolean().optional(),
  maxIterations: z.number().int().min(1).optional(),
});

async function runAgentIteration(
  provider: ProviderConfig,
  messages: ChatMessage[],
  retrieved: RetrievalChunk[],
  insertMode: string,
  documentStructureDescription: string | undefined,
  dynamicContext: { documentContext?: string; selection?: string } | undefined,
  onEvent: (event: any) => void,
  signal: AbortSignal,
  iteration: number,
  traceId: string,
  maxTokensOverride?: number,
): Promise<{ reply: string; actionPlan: ActionPlan | null; done: boolean; doneReason?: string; reasoningContent?: string; finishReason?: string; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> }> {
  const turn = await executeTurn({
    provider,
    messages,
    retrieved,
    tools: WORD_TOOLS,
    insertMode,
    documentStructureDescription,
    dynamicContext,
    signal,
    onDelta: (delta) => onEvent({ type: "delta", delta }),
    maxTokensOverride,
    turnNumber: iteration,
  });

  // Log each tool call issued (for any turn that had them).
  if (turn.toolCalls) {
    for (const tc of turn.toolCalls) {
      const params = safeParseToolArgs(tc.function.arguments);
      await logToolCall(undefined, { iteration, toolCallId: tc.id, toolName: tc.function.name, params });
    }
  }

  switch (turn.doneReason) {
    case "task_complete": {
      const summary = turn.taskCompleteSummary || "任务已完成";
      onEvent({ type: "task_complete", summary });
      await logIterationEnd(traceId, { iteration, done: true, reason: `task_complete: ${summary}` });
      return {
        reply: summary,
        actionPlan: null,
        done: true,
        doneReason: turn.doneReason,
        reasoningContent: turn.reasoningContent,
        finishReason: turn.finishReason,
        toolCalls: turn.toolCalls,
      };
    }
    case "tool_calls_pending": {
      const tcs = turn.toolCalls || [];
      const classification = classifyToolCalls(tcs);
      onEvent({
        type: "tool_call",
        tools: tcs.map((tc) => ({ id: tc.id, tool: tc.function.name, params: safeParseToolArgs(tc.function.arguments) })),
        classification,
      });
      await logIterationEnd(traceId, { iteration, done: false, reason: "tool_calls_pending" });
      return {
        reply: turn.reply,
        actionPlan: turn.actionPlan,
        done: false,
        doneReason: turn.doneReason,
        reasoningContent: turn.reasoningContent,
        finishReason: turn.finishReason,
        toolCalls: tcs,
      };
    }
    case "length_truncated": {
      // Surface as a server-side info event so the trace records the bump.
      onEvent({ type: "length_truncated", nextMaxTokens: turn.nextMaxTokens });
      await logIterationEnd(traceId, { iteration, done: false, reason: `length_truncated; bumped to maxTokens=${turn.nextMaxTokens}` });
      return {
        reply: turn.reply,
        actionPlan: null,
        done: false,
        doneReason: turn.doneReason,
        reasoningContent: turn.reasoningContent,
        finishReason: turn.finishReason,
        toolCalls: undefined,
      };
    }
    case "no_content": {
      await logIterationEnd(traceId, { iteration, done: true, reason: "no_content" });
      return {
        reply: turn.reply || "模型没有返回可用内容。",
        actionPlan: null,
        done: true,
        doneReason: turn.doneReason,
        reasoningContent: turn.reasoningContent,
        finishReason: turn.finishReason,
        toolCalls: undefined,
      };
    }
    case "error": {
      await logIterationEnd(traceId, { iteration, done: true, reason: `error: ${turn.error ?? "unknown"}` });
      throw new Error(turn.error || "LLM error");
    }
    case "final_reply":
    default: {
      await logIterationEnd(traceId, { iteration, done: true, reason: "final_reply" });
      return {
        reply: turn.reply,
        actionPlan: turn.actionPlan,
        done: true,
        doneReason: turn.doneReason,
        reasoningContent: turn.reasoningContent,
        finishReason: turn.finishReason,
        toolCalls: undefined,
      };
    }
  }
}

app.post("/v1/chat/agent-stream", async (req, res) => {
  const parsed = agentStreamSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: `Invalid chat payload: ${formatZodError(parsed.error)}`, details: parsed.error.flatten() });
  }

  const abortController = new AbortController();
  req.on("aborted", () => abortController.abort(new Error("client_aborted")));
  res.on("close", () => {
    if (!res.writableEnded && !abortController.signal.aborted) {
      abortController.abort(new Error("client_closed"));
    }
  });

  try {
    const { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode, dynamicContext } = await buildChatContext(parsed.data);
    const maxIterations = parsed.data.maxIterations ?? 10;

    const httpTraceId = await logHttpRequest("/v1/chat/agent-stream", {
      method: "POST",
      body: {
        messageCount: parsed.data.messages?.length ?? 0,
        hasDocumentContext: !!parsed.data.documentContext,
        hasDocumentStructure: !!parsed.data.documentStructure,
        hasSelection: !!parsed.data.selection,
        insertMode: parsed.data.insertMode,
        maxIterations,
      },
    });

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();

    res.write(`data: ${JSON.stringify({ type: "start", ts: Date.now() })}
\n`);

    // Create session with initial messages and cache document context for agent-continue
    const sessionId = createSession(enrichedMessages);
    const createdSession = getSession(sessionId);
    if (createdSession) {
      createdSession.documentStructureDescription = documentStructureDescription;
      createdSession.insertMode = insertMode;
      createdSession.dynamicContext = dynamicContext;
    }
    res.write(`data: ${JSON.stringify({ type: "session", sessionId })}
\n`);

    let currentMessages = [...enrichedMessages];
    let finalReply = "";
    let finalActionPlan: ActionPlan | null = null;
    let iteration = 0;
    let currentMaxTokensOverride: number | undefined = undefined;
    const recentToolFingerprints: string[] = [];

    while (iteration < maxIterations) {
      iteration++;

      let iterationResult: { reply: string; actionPlan: ActionPlan | null; done: boolean; doneReason?: string; reasoningContent?: string; finishReason?: string; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> };
      try {
        iterationResult = await runAgentIteration(
          provider,
          currentMessages,
          retrieved,
          insertMode,
          documentStructureDescription,
          dynamicContext,
          (event) => res.write(`data: ${JSON.stringify(event)}
\n`),
          abortController.signal,
          iteration,
          httpTraceId,
          currentMaxTokensOverride,
        );
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        const shouldFallbackToNonStream = message.includes("first token not received");

        if (!shouldFallbackToNonStream) {
          throw error;
        }

        // Fallback to non-stream (stream timed out before first token).
        const fallbackResult = await callOpenAICompatible(
          provider,
          currentMessages,
          retrieved,
          WORD_TOOLS,
          insertMode,
          documentStructureDescription,
          dynamicContext
        );

        if (fallbackResult.toolCalls && fallbackResult.toolCalls.length > 0) {
          const fbTcs = fallbackResult.toolCalls;
          const fbClass = classifyToolCalls(fbTcs);
          res.write(`data: ${JSON.stringify({ type: "tool_call", tools: fbTcs.map((tc) => ({ id: tc.id, tool: tc.function.name, params: safeParseToolArgs(tc.function.arguments) })), classification: fbClass })}
\n`);
          iterationResult = { reply: fallbackResult.reply, actionPlan: parseActionPlanFromToolCalls(fallbackResult.toolCalls), done: false, toolCalls: fallbackResult.toolCalls, reasoningContent: fallbackResult.reasoningContent };
        } else {
          res.write(`data: ${JSON.stringify({ type: "fallback", reason: "no_stream_delta" })}
\n`);
          iterationResult = { reply: fallbackResult.reply, actionPlan: fallbackResult.actionPlan, done: true, reasoningContent: fallbackResult.reasoningContent };
        }
      }

      finalReply = iterationResult.reply;
      finalActionPlan = iterationResult.actionPlan;

      // Doom loop detection: same (tool_name + params) called N times in a row?
      if (iterationResult.toolCalls && iterationResult.toolCalls.length > 0) {
        for (const tc of iterationResult.toolCalls) {
          if (tc.function.name !== "task_complete") {
            recentToolFingerprints.push(fingerprintToolCall(tc));
          }
        }
        if (detectDoomLoop(recentToolFingerprints)) {
          res.write(`data: ${JSON.stringify({ type: "doom_loop", fingerprints: recentToolFingerprints.slice(-4) })}
\n`);
          await logIterationEnd(httpTraceId, { sessionId, iteration, done: true, reason: "doom_loop" });
          finalReply = "检测到 Agent 陷入死循环（连续多次调用相同工具且参数相同），已自动终止。请检查任务描述或重新发起。";
          finalActionPlan = null;
          break;
        }
      }

      // Append assistant message via the unified helper.
      // buildAssistantMessage handles all three branches (tool_calls,
      // text-parsed actionPlan, plain reply) in one place.
      // reasoning_content is always included for tool-call rounds (DeepSeek
      // requires it, otherwise 400).  For non-tool-call rounds we respect
      // the includeReasoningContent config.
      const wantReasoning = provider.includeReasoningContent ?? true;
      const assistantMsg = buildAssistantMessage({
        reply: iterationResult.reply,
        actionPlan: iterationResult.actionPlan,
        reasoningContent: iterationResult.reasoningContent,
        finishReason: iterationResult.finishReason,
        toolCalls: iterationResult.toolCalls,
        done: iterationResult.done,
        doneReason: (iterationResult.doneReason ?? (iterationResult.done ? "final_reply" : "tool_calls_pending")) as TurnResult["doneReason"],
      });
      if (assistantMsg) {
        if (!assistantMsg.tool_calls && !wantReasoning) {
          assistantMsg.reasoning_content = undefined;
        }
        currentMessages.push(assistantMsg);
      }

      if (iterationResult.done) {
        break;
      }

      // length_truncated → don't terminate; just bump maxTokens and loop again.
      if (iterationResult.doneReason === "length_truncated") {
        currentMaxTokensOverride = provider.maxTokens
          ? Math.min(provider.maxTokens * 2, 16384)
          : 8192;
        // Don't push the await_tool_result event — we're staying on the
        // server side and looping again.
        continue;
      }

      // Otherwise: tool_calls_pending → wait for client to execute and
      // call /v1/chat/agent-continue.
      res.write(`data: ${JSON.stringify({ type: "await_tool_result", iteration })}
\n`);
      break;
    }

    // Save final state to session before ending
    const session = getSession(sessionId);
    if (session) {
      session.messages = currentMessages;
      session._doomFingerprints = recentToolFingerprints;
    }

    res.write(`data: ${JSON.stringify({
      type: "done",
      reply: finalReply,
      actionPlan: finalActionPlan,
      iteration,
      sessionId,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    })}
\n`);
    return res.end();
  } catch (error) {
    if (abortController.signal.aborted) {
      await logError(undefined, { endpoint: "/v1/chat/agent-stream", error: "client aborted" });
      return;
    }

    const errMsg = error instanceof Error ? error.message : "Unknown error";
    const errStack = error instanceof Error ? error.stack : undefined;
    await logError(undefined, { endpoint: "/v1/chat/agent-stream", error: errMsg, stack: errStack });

    if (!res.headersSent) {
      if (error instanceof Error && error.message.startsWith("Provider config missing")) {
        return res.status(400).json({ error: error.message });
      }
      if (error instanceof LlmHttpError) {
        const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
        return res.status(safeStatus).json({ error: error.details || error.message });
      }
      const message = error instanceof Error ? error.message : "Unknown error";
      return res.status(502).json({ error: message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    res.write(`data: ${JSON.stringify({ type: "error", error: message })}
\n`);
    return res.end();
  }
});

const agentContinueSchema = z.object({
  sessionId: z.string().min(1),
  toolResults: z.array(z.object({
    toolCallId: z.string(),
    toolName: z.string(),
    result: z.string(),
    success: z.boolean(),
  })),
  documentStructure: documentStructureSchema.optional(),
});

/**
 * Abort an in-flight agent session.
 *
 * The frontend Stop button calls this with the current `pendingSessionId` so
 * that any subsequent `/v1/chat/agent-continue` request — even one that
 * raced past the AbortController — will hit `Session not found` and bail.
 * The session map's per-session TTL (30 min) already cleans up stragglers,
 * but for an explicit stop we want the kill to be immediate.
 *
 * This endpoint does NOT abort already-running LLM streams; the in-flight
 * `/v1/chat/agent-stream` and `/v1/chat/agent-continue` requests tear
 * themselves down via `req.on("aborted")` / `res.on("close")` when the
 * client closes the connection.
 */
app.post("/v1/chat/abort", async (req, res) => {
  const sessionId = typeof req.body?.sessionId === "string" ? req.body.sessionId : "";
  if (sessionId) {
    const removed = deleteSession(sessionId);
    await logHttpRequest("/v1/chat/abort", {
      method: "POST",
      body: { sessionId, removed },
    });
  }
  return res.json({ ok: true });
});

app.post("/v1/chat/agent-continue", async (req, res) => {
  const parsed = agentContinueSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: `Invalid continue payload: ${formatZodError(parsed.error)}`, details: parsed.error.flatten() });
  }

  const abortController = new AbortController();
  req.on("aborted", () => abortController.abort(new Error("client_aborted")));
  res.on("close", () => {
    if (!res.writableEnded && !abortController.signal.aborted) {
      abortController.abort(new Error("client_closed"));
    }
  });

  try {
    const session = getSession(parsed.data.sessionId);
    if (!session) {
      await logError(undefined, { endpoint: "/v1/chat/agent-continue", error: `Session not found: ${parsed.data.sessionId}` });
      return res.status(404).json({ error: "Session not found or expired" });
    }

    const httpTraceId = await logHttpRequest("/v1/chat/agent-continue", {
      method: "POST",
      sessionId: parsed.data.sessionId,
      body: {
        toolResultCount: parsed.data.toolResults?.length ?? 0,
        toolResults: parsed.data.toolResults?.map((tr) => ({ toolCallId: tr.toolCallId, toolName: tr.toolName, success: tr.success, resultLength: tr.result?.length ?? 0 })),
        hasDocumentStructure: !!parsed.data.documentStructure,
      },
    });

    // Log each tool result
    for (const tr of parsed.data.toolResults) {
      await logToolResult(undefined, {
        sessionId: parsed.data.sessionId,
        toolCallId: tr.toolCallId,
        toolName: tr.toolName,
        result: tr.result,
        success: tr.success,
      });
    }

    // Use fresh document structure from request if provided, otherwise fall back to session-cached context
    let documentStructureDescription = session.documentStructureDescription;
    let insertMode = session.insertMode || "smart_action";
    let dynamicContext = session.dynamicContext || {};

    if (parsed.data.documentStructure) {
      documentStructureDescription = describeDocumentStructure(parsed.data.documentStructure);
      // Update session cache with fresh data
      session.documentStructureDescription = documentStructureDescription;
    }

    const { provider, retrieved } = await buildChatContext({
      messages: session.messages,
      documentContext: undefined,
      documentStructure: undefined,
      selection: undefined,
      insertMode: insertMode as any,
    });

    // Ensure the session has at least one assistant message with matching
    // tool_calls before we append tool results.  We don't just check the
    // LAST assistant because DeepSeek sometimes returns a reasoning-only
    // response (no tool_calls) that ends up as a trailing assistant message
    // — the next agent-continue would then incorrectly fail.
    const lastAssistantIdx = [...session.messages].reverse().findIndex((m) => m.role === "assistant");
    if (lastAssistantIdx === -1) {
      return res.status(400).json({ error: "无法继续 Agent 循环：会话中没有 assistant 消息，无法附加工具结果。" });
    }
    const lastAssistant = session.messages[session.messages.length - 1 - lastAssistantIdx];

    // Collect all tool_call IDs across *all* assistant messages in the session
    const allToolCallIds = new Set<string>();
    for (const m of session.messages) {
      if (m.role === "assistant" && m.tool_calls) {
        for (const tc of m.tool_calls) {
          allToolCallIds.add(tc.id);
        }
      }
    }

    // Check if the incoming tool results reference known tool_call IDs
    const unknownIds = parsed.data.toolResults.filter((tr) => !allToolCallIds.has(tr.toolCallId));
    if (unknownIds.length === parsed.data.toolResults.length) {
      return res.status(400).json({
        error: `无法继续 Agent 循环：工具结果引用的 tool_call_id（${unknownIds.map((t) => t.toolCallId).join(", ")}）不在会话的任何 assistant 消息中。请重新发起任务。`,
      });
    }

    // If the last assistant has no tool_calls but some tool results match
    // earlier assistant messages, accept them.  The stale/duplicate results
    // are harmless — the LLM will see repeated tool output and can ignore it.
    // Filter out results for tool_call_ids that are NOT in the last assistant
    // (stale results from previous iterations).
    const validResults = parsed.data.toolResults.filter((tr) =>
      lastAssistant.tool_calls?.some((tc: any) => tc.id === tr.toolCallId)
    );

    // Append valid tool results to session messages
    for (const tr of validResults) {
      session.messages.push({
        role: "tool",
        content: tr.result,
        tool_call_id: tr.toolCallId,
        name: tr.toolName,
      });
    }

    // If ALL results were stale, we still proceed — the LLM can respond
    // based on the existing conversation without new tool data.

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();

    res.write(`data: ${JSON.stringify({ type: "start", ts: Date.now() })}
\n`);

    // Track the iteration for this continue cycle (always 1 since agent-continue
    // does max one LLM call per request; the loop is driven by the frontend).
    const continueIteration = (session._continueIteration ? session._continueIteration + 1 : 1);
    session._continueIteration = continueIteration;

    let result: { reply: string; actionPlan: ActionPlan | null; reasoningContent?: string; finishReason?: string; doneReason?: string; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>; error?: string; nextMaxTokens?: number };
    try {
      const turn = await executeTurn({
        provider,
        messages: session.messages,
        retrieved,
        tools: WORD_TOOLS,
        insertMode,
        documentStructureDescription,
        dynamicContext,
        signal: abortController.signal,
        onDelta: (delta) => res.write(`data: ${JSON.stringify({ type: "delta", delta })}\n\n`),
        turnNumber: continueIteration,
      });
      result = {
        reply: turn.reply,
        actionPlan: turn.actionPlan,
        reasoningContent: turn.reasoningContent,
        finishReason: turn.finishReason,
        toolCalls: turn.toolCalls,
        doneReason: turn.doneReason,
        error: turn.error,
        nextMaxTokens: turn.nextMaxTokens,
      };
    } catch (error) {
      // Hard error that executeTurn didn't handle (e.g. AbortError).
      throw error;
    }

    // Length truncated → retry internally with expanded maxTokens.
    // The frontend never sees length_truncated as a terminal event.
    if (result.doneReason === "length_truncated") {
      const bumped = result.nextMaxTokens ?? Math.min((provider.maxTokens ?? 4096) * 2, 16384);
      // Push the reasoning output already received so the frontend updates.
      res.write(`data: ${JSON.stringify({ type: "length_truncated_retry", nextMaxTokens: bumped })}\n\n`);
      const retryTurn = await executeTurn({
        provider: { ...provider, maxTokens: bumped },
        messages: session.messages,
        retrieved,
        tools: WORD_TOOLS,
        insertMode,
        documentStructureDescription,
        dynamicContext,
        signal: abortController.signal,
        onDelta: (delta) => res.write(`data: ${JSON.stringify({ type: "delta", delta })}\n\n`),
        turnNumber: continueIteration,
        maxTokensOverride: bumped,
      });
      result = {
        reply: retryTurn.reply,
        actionPlan: retryTurn.actionPlan,
        reasoningContent: retryTurn.reasoningContent,
        finishReason: retryTurn.finishReason,
        toolCalls: retryTurn.toolCalls,
        doneReason: retryTurn.doneReason,
        error: retryTurn.error,
        nextMaxTokens: retryTurn.nextMaxTokens,
      };
    }

    // Doom-loop guard: track recent tool fingerprints across agent-continue
    // invocations on this session.
    if (!session._doomFingerprints) session._doomFingerprints = [];
    if (result.toolCalls && result.toolCalls.length > 0) {
      for (const tc of result.toolCalls) {
        if (tc.function.name !== "task_complete") {
          session._doomFingerprints.push(fingerprintToolCall(tc));
        }
      }
      if (detectDoomLoop(session._doomFingerprints)) {
        res.write(`data: ${JSON.stringify({ type: "doom_loop", fingerprints: session._doomFingerprints.slice(-4) })}\n\n`);
        await logIterationEnd(httpTraceId, { sessionId: session.id, iteration: continueIteration, done: true, reason: "doom_loop" });
        res.write(`data: ${JSON.stringify({
          type: "done",
          reply: "检测到 Agent 陷入死循环（连续多次调用相同工具且参数相同），已自动终止。请检查任务描述或重新发起。",
          actionPlan: null,
          sessionId: session.id,
          retrievalCount: retrieved.length,
          citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
        })}\n\n`);
        return res.end();
      }
    }

    // Save the new assistant message to session for next iteration.
    // Always persist tool_calls so that when the client later sends tool
    // results via agent-continue, the preceding assistant message is present.
    // Orphaned tool_calls are cleaned up by sanitizeMessages() before each LLM call.
    //
    // reasoning_content is always included for tool-call rounds (DeepSeek
    // requires it, otherwise 400).  For non-tool-call rounds we respect the
    // includeReasoningContent config.
    const wantReasoning = provider.includeReasoningContent ?? true;
    const turnResult: TurnResult = {
      reply: result.reply,
      actionPlan: result.actionPlan,
      reasoningContent: result.reasoningContent,
      finishReason: result.finishReason,
      toolCalls: result.toolCalls,
      done: result.doneReason !== "tool_calls_pending",
      doneReason: (result.doneReason ?? "final_reply") as TurnResult["doneReason"],
    };
    const assistantMsg = buildAssistantMessage(turnResult);
    if (assistantMsg) {
      if (!assistantMsg.tool_calls && !wantReasoning) {
        assistantMsg.reasoning_content = undefined;
      }
      session.messages.push(assistantMsg);
    }

    // Branch on the unified doneReason
    switch (result.doneReason) {
      case "tool_calls_pending": {
        const tcs = result.toolCalls || [];
        const classification = classifyToolCalls(tcs);
        await logIterationEnd(httpTraceId, { sessionId: session.id, iteration: continueIteration, done: false, reason: "tool_calls_pending" });
        res.write(`data: ${JSON.stringify({
          type: "tool_call",
          tools: tcs.map((tc) => ({ id: tc.id, tool: tc.function.name, params: safeParseToolArgs(tc.function.arguments) })),
          classification,
        })}\n\n`);
        break;
      }
      case "task_complete": {
        const summary = turnResult.taskCompleteSummary || "任务已完成";
        await logIterationEnd(httpTraceId, { sessionId: session.id, iteration: continueIteration, done: true, reason: `task_complete: ${summary}` });
        res.write(`data: ${JSON.stringify({ type: "task_complete", summary })}\n\n`);
        break;
      }
      case "length_truncated":
        // Already handled above with an internal retry. If control reaches
        // here, the retry was also truncated — treat as final reply.
        await logIterationEnd(httpTraceId, { sessionId: session.id, iteration: continueIteration, done: true, reason: "length_truncated_after_retry" });
        break;
      case "no_content":
        await logIterationEnd(httpTraceId, { sessionId: session.id, iteration: continueIteration, done: true, reason: "no_content" });
        break;
      case "error":
        await logIterationEnd(httpTraceId, { sessionId: session.id, iteration: continueIteration, done: true, reason: `error: ${result.error ?? "unknown"}` });
        throw new Error(result.error || "LLM error");
      case "final_reply":
      default:
        await logIterationEnd(httpTraceId, { sessionId: session.id, iteration: continueIteration, done: true, reason: "final_reply" });
        break;
    }

    res.write(`data: ${JSON.stringify({
      type: "done",
      reply: result.reply,
      actionPlan: (result.toolCalls && result.toolCalls.length > 0) ? null : result.actionPlan,
      sessionId: session.id,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    })}
\n`);
    return res.end();
  } catch (error) {
    if (abortController.signal.aborted) {
      await logError(undefined, { endpoint: "/v1/chat/agent-continue", error: "client aborted" });
      return;
    }

    const errMsg = error instanceof Error ? error.message : "Unknown error";
    const errStack = error instanceof Error ? error.stack : undefined;
    await logError(undefined, { endpoint: "/v1/chat/agent-continue", error: errMsg, stack: errStack });

    if (!res.headersSent) {
      if (error instanceof Error && error.message.startsWith("Provider config missing")) {
        return res.status(400).json({ error: error.message });
      }
      if (error instanceof LlmHttpError) {
        const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
        return res.status(safeStatus).json({ error: error.details || error.message });
      }
      const message = error instanceof Error ? error.message : "Unknown error";
      return res.status(502).json({ error: message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    res.write(`data: ${JSON.stringify({ type: "error", error: message })}
\n`);
    return res.end();
  }
});

// SPA fallback: serve index.html for non-API routes, with the agent token
// injected as a bootstrap script. Same-origin only (SOP): cross-origin pages
// can trigger requests but cannot read the response body, and every /v1/*
// request is rejected without the token anyway.
let indexHtmlCache: string | null = null;
function loadIndexHtml(): string {
  if (indexHtmlCache === null) {
    try {
      const raw = readFileSync(path.join(publicDir, "index.html"), "utf-8");
      const bootstrap = `<script>window.__AGENT_TOKEN__=${JSON.stringify(agentToken)};</script>`;
      indexHtmlCache = raw.includes("</head>")
        ? raw.replace("</head>", `${bootstrap}\n  </head>`)
        : bootstrap + raw;
    } catch {
      indexHtmlCache = "";
    }
  }
  return indexHtmlCache;
}

app.get("*", (_req, res) => {
  const html = loadIndexHtml();
  if (!html) {
    res.status(404).json({ error: "Not found" });
    return;
  }
  res.type("html").send(html);
});

app.listen(port, host, () => {
  console.log(`Local agent listening on http://${host}:${port}`);
});
