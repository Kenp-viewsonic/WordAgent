/**
 * Tool Runtime — unified turn execution pipeline.
 *
 * Inspired by anomalyco/opencode's ToolRuntime.runTools() pattern.
 *
 * Each "turn" is one LLM call. The runtime:
 *   1. Streams the LLM response (via streamOpenAICompatible).
 *   2. Captures `finish_reason` from the final chunk — this is the ONLY
 *      signal that drives loop continuation. No more reasoningOnly hacks.
 *   3. Surfaces `reasoning_content` separately for the frontend to display.
 *   4. On `finish_reason === "length"`, auto-expands `maxTokens` for the
 *      next turn (up to 2×) so a long reasoning chain doesn't permanently
 *      starve tool-call output.
 *   5. Classifies the turn outcome and returns a single `TurnResult` —
 *      the caller (`agent-stream` loop or `agent-continue` handler) only
 *      needs to forward SSE events and persist messages.
 *
 * No tool calls are dispatched server-side — all tools (perception and
 * action alike) are executed by the Word Add-in frontend. This eliminates
 * the old `isPerceptionOnlyPlan` / `isIterablePlan` categorization that
 * doubled the branching logic.
 */
import { ActionPlan, ChatMessage, ProviderConfig, RetrievalChunk, ToolDefinition } from "./types.js";
import { streamOpenAICompatible, parseActionPlanFromToolCalls, callOpenAICompatible, LlmHttpError } from "./llm.js";

/**
 * Why the turn is done. Drives the caller's response.
 *
 * - `tool_calls_pending`: LLM wants the frontend to execute tool calls.
 *                        Caller should push `tool_call` SSE event and wait
 *                        for the next `/v1/chat/agent-continue`.
 * - `task_complete`: LLM explicitly ended the task. Caller should push
 *                    `task_complete` SSE event with the summary.
 * - `final_reply`: LLM produced a text reply without tool calls.
 *                  Caller pushes `done` event.
 * - `length_truncated`: Output was truncated mid-stream. We auto-expanded
 *                       `maxTokens` for the next turn — caller should
 *                       loop back into `executeTurn`.
 * - `no_content`: LLM produced nothing usable. Caller pushes `done`.
 * - `error`: Network / HTTP error. Caller handles.
 */
export type TurnDoneReason =
  | "tool_calls_pending"
  | "task_complete"
  | "final_reply"
  | "length_truncated"
  | "no_content"
  | "error";

export interface TurnResult {
  // What the LLM said
  reply: string;
  actionPlan: ActionPlan | null;
  reasoningContent: string | undefined;
  finishReason: string | undefined;
  toolCalls: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> | undefined;

  // Loop control
  done: boolean;
  doneReason: TurnDoneReason;

  // For task_complete calls
  taskCompleteSummary?: string;

  // For error turns
  error?: string;

  // If `length_truncated` was bumped, this is the new value used by the
  // caller's next `executeTurn` invocation. Undefined otherwise.
  nextMaxTokens?: number;
}

export interface ExecuteTurnParams {
  provider: ProviderConfig;
  messages: ChatMessage[];
  retrieved: RetrievalChunk[];
  tools: ToolDefinition[];
  insertMode: string;
  documentStructureDescription?: string;
  dynamicContext?: { documentContext?: string; selection?: string };
  signal: AbortSignal;
  onDelta: (text: string) => void;

  /**
   * Override the provider's configured `maxTokens` for this turn only.
   * Used by the auto-expand logic on `length_truncated`. Caller persists
   * the bumped value and passes it again on the next turn.
   */
  maxTokensOverride?: number;

  /**
   * Current turn number (1-based). For logging.
   */
  turnNumber?: number;
}

const PERCEPTION_TOOLS = new Set([
  "read_document",
  "get_selection_info",
  "get_document_stats",
  "get_paragraph_format",
  "get_document_tables",
  "read_table",
]);

const ACTION_TOOLS = new Set([
  "find_and_replace",
  "find_and_replace_v2",
  "insert_at_cursor",
  "insert_after_heading",
  "insert_at_end",
  "insert_at_start",
  "insert_after_paragraph",
  "delete_paragraph",
  "insert_table",
  "delete_table",
  "insert_mermaid_image",
  "replace_paragraph",
  "set_paragraph_style",
  "merge_paragraphs",
  "apply_rich_format",
  "undo_last_action",
]);

/**
 * Lightweight categorization that the FRONTEND uses to decide whether to
 * auto-iterate (perception only) or pause for user confirmation (action).
 * The server no longer uses these to gate the loop — it just relays the
 * categorization as a runtime attribute on the tool_call event.
 */
export function classifyToolCalls(
  toolCalls: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>,
): { allPerception: boolean; hasAction: boolean; hasControl: boolean } {
  let allPerception = true;
  let hasAction = false;
  let hasControl = false;
  for (const tc of toolCalls) {
    const name = tc.function.name;
    if (PERCEPTION_TOOLS.has(name)) {
      // still perception
    } else if (ACTION_TOOLS.has(name)) {
      allPerception = false;
      hasAction = true;
    } else {
      // task_complete, reply_only, etc.
      allPerception = false;
      hasControl = true;
    }
  }
  return { allPerception, hasAction, hasControl };
}

/**
 * Execute one LLM turn. Single entry point for both `/v1/chat/agent-stream`
 * (initial turn) and `/v1/chat/agent-continue` (subsequent turns).
 */
export async function executeTurn(params: ExecuteTurnParams): Promise<TurnResult> {
  const {
    provider,
    messages,
    retrieved,
    tools,
    insertMode,
    documentStructureDescription,
    dynamicContext,
    signal,
    onDelta,
    maxTokensOverride,
  } = params;

  // Apply the maxTokens override for this turn only.
  const effectiveProvider = maxTokensOverride !== undefined
    ? { ...provider, maxTokens: maxTokensOverride }
    : provider;

  // ── 1. Call the LLM ────────────────────────────────────────────────
  let llmResult;
  try {
    llmResult = await streamOpenAICompatible(
      effectiveProvider,
      messages,
      retrieved,
      onDelta,
      signal,
      tools,
      insertMode,
      documentStructureDescription,
      dynamicContext,
    );
  } catch (error) {
    // ── Stream failed — fall back to non-streaming callOpenAICompatible ──
    // (Preserves prior fallback for "first token not received" cases.)
    const errMsg = error instanceof Error ? error.message : String(error);
    if (errMsg.includes("first token not received")) {
      try {
        llmResult = await callOpenAICompatible(
          effectiveProvider,
          messages,
          retrieved,
          tools,
          insertMode,
          documentStructureDescription,
          dynamicContext,
        );
      } catch (fallbackErr) {
        return {
          reply: "",
          actionPlan: null,
          reasoningContent: undefined,
          finishReason: undefined,
          toolCalls: undefined,
          done: true,
          doneReason: "error",
          error: fallbackErr instanceof Error ? fallbackErr.message : String(fallbackErr),
        };
      }
    } else if (error instanceof LlmHttpError) {
      return {
        reply: "",
        actionPlan: null,
        reasoningContent: undefined,
        finishReason: undefined,
        toolCalls: undefined,
        done: true,
        doneReason: "error",
        error: error.message,
      };
    } else {
      return {
        reply: "",
        actionPlan: null,
        reasoningContent: undefined,
        finishReason: undefined,
        toolCalls: undefined,
        done: true,
        doneReason: "error",
        error: errMsg,
      };
    }
  }

  // ── 2. Categorize the turn ──────────────────────────────────────────
  const toolCalls = llmResult.toolCalls;
  const finishReason = llmResult.finishReason;

  // Case A: tool_calls present — frontend needs to execute them.
  if (toolCalls && toolCalls.length > 0) {
    // task_complete is a special control tool — terminates the loop.
    const taskCompleteCall = toolCalls.find((tc) => tc.function.name === "task_complete");
    if (taskCompleteCall) {
      let summary = "任务已完成";
      try {
        const args = JSON.parse(taskCompleteCall.function.arguments);
        if (args?.summary) summary = String(args.summary);
      } catch {
        // Use the raw arguments text if JSON is malformed — the frontend
        // can still display something meaningful.
        summary = taskCompleteCall.function.arguments || summary;
      }
      return {
        reply: summary,
        actionPlan: null,
        reasoningContent: llmResult.reasoningContent,
        finishReason,
        toolCalls,
        done: true,
        doneReason: "task_complete",
        taskCompleteSummary: summary,
      };
    }

    // Any other tool_calls → frontend executes, then calls agent-continue.
    return {
      reply: llmResult.reply,
      actionPlan: llmResult.actionPlan,
      reasoningContent: llmResult.reasoningContent,
      finishReason,
      toolCalls,
      done: false,
      doneReason: "tool_calls_pending",
    };
  }

  // Case B: no tool_calls.
  // finish_reason drives the loop decision — NOT reasoning_content.
  if (finishReason === "length") {
    // Truncated mid-stream. The reasoning chain or text was cut off
    // before the LLM could produce tool_calls. Auto-expand maxTokens
    // and let the caller re-invoke executeTurn.
    const current = effectiveProvider.maxTokens ?? 4096;
    const bumped = Math.min(current * 2, 16384);
    return {
      reply: llmResult.reply,
      actionPlan: null,
      reasoningContent: llmResult.reasoningContent,
      finishReason,
      toolCalls: undefined,
      done: false,
      doneReason: "length_truncated",
      nextMaxTokens: bumped,
    };
  }

  // No tool_calls and finish_reason !== "length" → final reply.
  if (!llmResult.reply || llmResult.reply === "模型没有返回可用内容。") {
    return {
      reply: llmResult.reply || "模型没有返回可用内容。",
      actionPlan: null,
      reasoningContent: llmResult.reasoningContent,
      finishReason,
      toolCalls: undefined,
      done: true,
      doneReason: "no_content",
    };
  }

  return {
    reply: llmResult.reply,
    actionPlan: llmResult.actionPlan,
    reasoningContent: llmResult.reasoningContent,
    finishReason,
    toolCalls: undefined,
    done: true,
    doneReason: "final_reply",
  };
}

/**
 * Build the assistant message that should be appended to the session
 * after a turn completes. Mirrors what the LLM returned, with
 * reasoning_content and tool_calls preserved for the next round.
 *
 * `actionPlan` is only used as a fallback when there are no tool_calls —
 * in that case we synthesize a tool_call-shaped assistant message so the
 * conversation history still has tool_calls to anchor against.
 */
export function buildAssistantMessage(result: TurnResult): ChatMessage | null {
  // Has real tool_calls → use them as-is.
  if (result.toolCalls && result.toolCalls.length > 0) {
    return {
      role: "assistant",
      content: result.reply || "",
      reasoning_content: result.reasoningContent,
      tool_calls: result.toolCalls,
    };
  }

  // Action plan from text parsing (no streaming tool_calls) → synthesize.
  if (result.actionPlan && result.actionPlan.actions.length > 0) {
    const synthetic = result.actionPlan.actions.map((action, index) => ({
      id: action.toolCallId || `text-parsed-${Date.now()}-${index}`,
      type: "function" as const,
      function: {
        name: action.action,
        arguments: JSON.stringify(action.params),
      },
    }));
    return {
      role: "assistant",
      content: result.reply,
      reasoning_content: result.reasoningContent,
      tool_calls: synthetic,
    };
  }

  // Plain text reply.
  if (result.reply) {
    return {
      role: "assistant",
      content: result.reply,
      reasoning_content: result.reasoningContent,
    };
  }

  return null;
}

/**
 * Doom-loop guard. Inspects the recent history of (toolName + params-hash)
 * fingerprints. If the most recent N fingerprints are all identical, we
 * conclude the LLM is stuck in a loop and the caller should terminate.
 *
 * OpenCode uses the same threshold-based approach (DOOM_LOOP_THRESHOLD = 3).
 */
export function detectDoomLoop(
  recentFingerprints: string[],
  threshold = 4,
): boolean {
  if (recentFingerprints.length < threshold) return false;
  const tail = recentFingerprints.slice(-threshold);
  return tail.every((fp) => fp === tail[0]);
}

export function fingerprintToolCall(tc: { function: { name: string; arguments: string } }): string {
  // Normalize arguments — whitespace and property order shouldn't matter
  // for doom-loop detection.
  let argsNorm = tc.function.arguments;
  try {
    const parsed = JSON.parse(tc.function.arguments);
    argsNorm = JSON.stringify(sortKeysDeep(parsed));
  } catch {
    // keep raw
  }
  return `${tc.function.name}|${argsNorm}`;
}

function sortKeysDeep(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortKeysDeep);
  if (value && typeof value === "object") {
    const sorted: Record<string, unknown> = {};
    for (const k of Object.keys(value as Record<string, unknown>).sort()) {
      sorted[k] = sortKeysDeep((value as Record<string, unknown>)[k]);
    }
    return sorted;
  }
  return value;
}