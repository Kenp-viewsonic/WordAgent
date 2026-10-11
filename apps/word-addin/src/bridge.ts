/**
 * Host bridge for the desktop line (WP1: auth plumbing; WP2/WP3: WebView2
 * tool-execution protocol). See plans/com-shell-migration-kickoff.md §3/§6.4.
 *
 * In the WebView2 task pane (and in a plain browser during WP1) the UI talks
 * to the agent over http://127.0.0.1:8787. The agent requires the
 * `X-Agent-Token` header on every /v1/* call; the token is bootstrapped into
 * the same-origin page as `window.__AGENT_TOKEN__`.
 *
 * WP3: document tools execute in the COM host (Word OM) via request/response
 * messages over `window.chrome.webview` — this file owns that protocol.
 */

type WindowWithToken = Window & {
  __AGENT_TOKEN__?: string;
  Word?: unknown;
  chrome?: { webview?: WebViewLike };
};

type WebViewLike = {
  postMessage(message: unknown): void;
  addEventListener(type: "message", handler: (e: { data: unknown }) => void): void;
};

type BridgeEnvelope = {
  type: string;
  id: number;
  success?: boolean;
  result?: string;
  data?: unknown;
  error?: string;
};

type PendingCall = {
  resolve: (value: BridgeEnvelope) => void;
  reject: (reason: Error) => void;
};

function hostWindow(): WindowWithToken {
  return window as WindowWithToken;
}

/** True when running inside the legacy Office.js host. */
export function isOfficeHost(): boolean {
  return typeof hostWindow().Word !== "undefined";
}

/** True when running inside the desktop WebView2 host (WP2+). */
export function isWebViewHost(): boolean {
  return !!hostWindow().chrome?.webview;
}

/** Document tools require one of the Word hosts; chat works everywhere. */
export function documentToolsAvailable(): boolean {
  return isOfficeHost() || isWebViewHost();
}

/**
 * Wrap window.fetch so same-origin agent calls carry the token header.
 * Cross-origin requests are never touched — the token must not leak.
 */
export function installAuthFetch(): void {
  const token = hostWindow().__AGENT_TOKEN__;
  if (!token) return;

  const originalFetch = window.fetch.bind(window);

  window.fetch = ((input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url =
      typeof input === "string" ? input : input instanceof URL ? input.href : input.url;

    if (isSameOrigin(url)) {
      const headers = new Headers(
        init?.headers ?? (typeof input === "object" && "headers" in input ? input.headers : undefined)
      );
      headers.set("X-Agent-Token", token);
      init = { ...init, headers };
    }

    return originalFetch(input as RequestInfo, init);
  }) as typeof window.fetch;
}

function isSameOrigin(url: string): boolean {
  if (url.startsWith("/")) return true;
  try {
    return new URL(url, window.location.href).origin === window.location.origin;
  } catch {
    return false;
  }
}

// ─── WP3: WebView2 tool-execution protocol ───────────────────────────────────
// Request:  { type: "tool-request", id, method, payload }
// Response: { type: "tool-response", id, success, result?, data?, error? }

let bridgeSeq = 0;
const bridgePending = new Map<number, PendingCall>();

function initBridgeListener(): void {
  const webview = hostWindow().chrome?.webview;
  if (!webview) return;
  webview.addEventListener("message", (e) => {
    const msg = e.data as BridgeEnvelope | undefined;
    if (!msg || msg.type !== "tool-response" || !bridgePending.has(msg.id)) return;
    const pending = bridgePending.get(msg.id)!;
    bridgePending.delete(msg.id);
    if (msg.success) {
      pending.resolve(msg);
    } else {
      pending.reject(new Error(msg.error || "文档工具执行失败"));
    }
  });
}

function bridgeCall(method: string, payload: unknown): Promise<BridgeEnvelope> {
  const webview = hostWindow().chrome?.webview;
  if (!webview) {
    return Promise.reject(new Error("当前环境不支持文档工具（WebView2 桥不可用）"));
  }
  initBridgeListener();
  return new Promise((resolve, reject) => {
    const id = ++bridgeSeq;
    bridgePending.set(id, { resolve, reject });
    webview.postMessage({ type: "tool-request", id, method, payload });
    setTimeout(() => {
      if (bridgePending.has(id)) {
        bridgePending.delete(id);
        reject(new Error(`文档工具执行超时（${method}）`));
      }
    }, 120000);
  });
}

export async function bridgeExecuteAction(action: {
  action: string;
  params: Record<string, any>;
  description: string;
  toolCallId?: string;
}): Promise<string> {
  const res = await bridgeCall("executeAction", { action });
  return res.result ?? "";
}

export async function bridgeGetStructuredContext(): Promise<unknown> {
  const res = await bridgeCall("getStructuredContext", {});
  return res.data;
}

export async function bridgeGetWordContext(): Promise<{ documentContext: string; selection: string }> {
  const res = await bridgeCall("getWordContext", {});
  return (res.data as { documentContext: string; selection: string }) ?? { documentContext: "", selection: "" };
}

export async function bridgeSimpleInsert(mode: string, text: string): Promise<string> {
  const res = await bridgeCall("simpleInsert", { mode, text });
  return res.result ?? "";
}
