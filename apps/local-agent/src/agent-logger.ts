import fs from "node:fs/promises";
import path from "node:path";
import { dataRoot } from "./paths.js";

/**
 * Agent iteration logger.
 *
 * Writes a structured JSON-lines log of every LLM request/response cycle to
 * `data/agent-trace.jsonl`. Each line is a self-contained JSON object so the
 * file can be tailed / grepped / parsed incrementally without loading the
 * whole file.
 *
 * Logged events:
 *  - request_start  : the exact messages array + payload sent to the LLM API
 *  - response_raw   : raw HTTP status + body (or stream summary)
 *  - response_parsed: parsed reply / tool_calls / actionPlan
 *  - tool_call      : tool call request issued to the add-in
 *  - tool_result    : tool result returned from the add-in
 *  - iteration_end  : iteration outcome (done / continue / error)
 *  - error          : any error thrown during the cycle
 *
 * A monotonically increasing `traceId` ties all events of a single
 * request/response cycle together.
 */

const traceFile = path.join(dataRoot, "agent-trace.jsonl");

let counter = 0;

function nextTraceId(): string {
  counter += 1;
  return `${Date.now()}-${counter.toString(36)}`;
}

function ts(): string {
  return new Date().toISOString();
}

async function appendLog(entry: Record<string, any>): Promise<void> {
  try {
    await fs.mkdir(dataRoot, { recursive: true });
    const line = JSON.stringify({ ...entry, _ts: ts() }) + "\n";
    await fs.appendFile(traceFile, line, "utf8");
  } catch {
    // Logging must never break the request flow.
  }
}

export interface RequestStartLog {
  endpoint: string;
  sessionId?: string;
  iteration?: number;
  model: string;
  messages: unknown[];
  tools?: unknown;
  temperature?: number;
  maxTokens?: number;
  stream: boolean;
  dynamicContext?: unknown;
  retrievedChunks?: Array<{ id: string; fileName: string; text: string }>;
}

export async function logRequestStart(log: RequestStartLog): Promise<string> {
  const traceId = nextTraceId();
  await appendLog({
    event: "request_start",
    traceId,
    endpoint: log.endpoint,
    sessionId: log.sessionId,
    iteration: log.iteration,
    model: log.model,
    temperature: log.temperature,
    maxTokens: log.maxTokens,
    stream: log.stream,
    messages: log.messages,
    tools: log.tools,
    dynamicContext: log.dynamicContext,
    retrievedChunks: log.retrievedChunks,
    messageCount: Array.isArray(log.messages) ? log.messages.length : 0,
  });
  return traceId;
}

export async function logResponseRaw(
  traceId: string,
  info: { status: number; contentType?: string; bodyPreview?: string; streamChunkCount?: number; fullTextLength?: number; fullReasoningLength?: number }
): Promise<void> {
  await appendLog({
    event: "response_raw",
    traceId,
    status: info.status,
    contentType: info.contentType,
    bodyPreview: info.bodyPreview,
    streamChunkCount: info.streamChunkCount,
    fullTextLength: info.fullTextLength,
    fullReasoningLength: info.fullReasoningLength,
  });
}

export async function logResponseParsed(
  traceId: string,
  parsed: { reply: string; actionPlan: unknown | null; toolCalls?: unknown }
): Promise<void> {
  await appendLog({
    event: "response_parsed",
    traceId,
    reply: parsed.reply,
    replyLength: parsed.reply?.length ?? 0,
    actionPlan: parsed.actionPlan,
    toolCalls: parsed.toolCalls,
  });
}

export async function logToolCall(
  traceId: string | undefined,
  info: { sessionId?: string; iteration?: number; toolCallId: string; toolName: string; params: unknown }
): Promise<void> {
  await appendLog({
    event: "tool_call",
    traceId,
    sessionId: info.sessionId,
    iteration: info.iteration,
    toolCallId: info.toolCallId,
    toolName: info.toolName,
    params: info.params,
  });
}

export async function logToolResult(
  traceId: string | undefined,
  info: { sessionId?: string; iteration?: number; toolCallId: string; toolName: string; result: string; success: boolean }
): Promise<void> {
  await appendLog({
    event: "tool_result",
    traceId,
    sessionId: info.sessionId,
    iteration: info.iteration,
    toolCallId: info.toolCallId,
    toolName: info.toolName,
    result: info.result,
    resultLength: info.result?.length ?? 0,
    success: info.success,
  });
}

export async function logIterationEnd(
  traceId: string,
  info: { sessionId?: string; iteration: number; done: boolean; reason: string }
): Promise<void> {
  await appendLog({
    event: "iteration_end",
    traceId,
    sessionId: info.sessionId,
    iteration: info.iteration,
    done: info.done,
    reason: info.reason,
  });
}

export async function logError(
  traceId: string | undefined,
  info: { sessionId?: string; iteration?: number; endpoint?: string; error: string; stack?: string }
): Promise<void> {
  await appendLog({
    event: "error",
    traceId,
    sessionId: info.sessionId,
    iteration: info.iteration,
    endpoint: info.endpoint,
    error: info.error,
    stack: info.stack,
  });
}

/**
 * Convenience: log a top-level HTTP request boundary (before any LLM call).
 * Useful for correlating "Unexpected end of JSON input" style errors with
 * the exact incoming payload.
 */
export async function logHttpRequest(
  endpoint: string,
  info: { method: string; body?: unknown; sessionId?: string }
): Promise<string> {
  const traceId = nextTraceId();
  await appendLog({
    event: "http_request",
    traceId,
    endpoint,
    method: info.method,
    sessionId: info.sessionId,
    body: info.body,
  });
  return traceId;
}
