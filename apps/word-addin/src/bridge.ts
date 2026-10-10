/**
 * Host bridge for the desktop line (WP1: auth plumbing; WP2/WP3: WebView2
 * tool-execution protocol). See plans/com-shell-migration-kickoff.md §3/§6.4.
 *
 * In the WebView2 task pane (and in a plain browser during WP1) the UI talks
 * to the agent over http://127.0.0.1:8787. The agent requires the
 * `X-Agent-Token` header on every /v1/* call; the token is bootstrapped into
 * the same-origin page as `window.__AGENT_TOKEN__`.
 */

type WindowWithToken = Window & {
  __AGENT_TOKEN__?: string;
  Word?: unknown;
  chrome?: { webview?: unknown };
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
