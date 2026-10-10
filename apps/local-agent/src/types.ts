export type ProviderConfig = {
  baseUrl: string;
  apiKey: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
  firstTokenTimeout?: number;
  overallTimeout?: number;
  /** Enable thinking / reasoning mode (DeepSeek, Agnes, etc.) */
  enableThinking?: boolean;
  /** Whether to include reasoning_content when sending history back to the API.
   *  Defaults to true because tool-call rounds always require it. */
  includeReasoningContent?: boolean;
  /** Thinking effort level. "medium" and "high" map sensibly across providers. */
  thinkingEffort?: "medium" | "high";
  /** Parameter format: "deepseek" uses `thinking` + `reasoning_effort`;
   *  "openai" uses `chat_template_kwargs.enable_thinking`. */
  thinkingFormat?: "deepseek" | "openai";
  /**
   * Maximum iterations the agent loop may run before forcibly stopping.
   * Per-preset so different models can be tuned independently (e.g.
   * DeepSeek for deep research, Agnes for shallow edits). 1..50.
   */
  maxIterations?: number;
};

export type ConfigPreset = {
  id: string;
  name: string;
  config: ProviderConfig;
  createdAt: number;
  updatedAt: number;
};

export type ChatMessage = {
  role: "system" | "user" | "assistant" | "tool";
  content: string;
  tool_call_id?: string;
  name?: string;
  /** Reasoning / thinking chain content (DeepSeek, Agnes, etc.).
   *  Required for tool-call rounds in DeepSeek; optional otherwise. */
  reasoning_content?: string;
  tool_calls?: Array<{
    id: string;
    type: "function";
    function: { name: string; arguments: string };
  }>;
};

export type RetrievalChunk = {
  id: string;
  fileName: string;
  text: string;
};

// --- Document Structure ---

export type DocumentParagraph = {
  index: number;
  text: string;
  style: string;
  headingLevel?: number;
  isTable: boolean;
  isList: boolean;
  charCount?: number;
  font?: {
    name?: string;
    size?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
  };
};

export type DocumentStructure = {
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

// --- Word Actions ---

export type WordAction = {
  action: string;
  params: Record<string, any>;
  description: string;
  toolCallId?: string;
};

export type ActionPlan = {
  actions: WordAction[];
  explanation: string;
};

// --- Action Execution Result ---

export type ActionResult = {
  action: string;
  success: boolean;
  message: string;
  details?: string;
  data?: Record<string, any>;
};

// --- Tool Call Result (for Agent Loop) ---

export type ToolCallResult = {
  toolCallId: string;
  toolName: string;
  result: string;
  success: boolean;
  data?: Record<string, any>;
};

// --- Agent Session (for multi-turn tool calling) ---

export type AgentSession = {
  id: string;
  messages: ChatMessage[];
  createdAt: number;
  updatedAt: number;
  /** Cached document structure description for agent-continue context preservation */
  documentStructureDescription?: string;
  /** Cached insert mode for agent-continue */
  insertMode?: string;
  /** Cached dynamic context (documentContext / selection) for agent-continue */
  dynamicContext?: { documentContext?: string; selection?: string };
  /** Internal counter for agent-continue iteration tracking */
  _continueIteration?: number;
  /** Rolling list of recent (tool_name + normalized params) fingerprints
   *  used by the doom-loop detector. Cleared when the session is reset. */
  _doomFingerprints?: string[];
};

// --- Tool Definitions (OpenAI Function Calling format) ---

export type ToolFunction = {
  name: string;
  description: string;
  parameters: Record<string, any>;
};

export type ToolDefinition = {
  type: "function";
  function: ToolFunction;
};

// --- LLM Response ---

export type LlmResponse = {
  reply: string;
  actionPlan: ActionPlan | null;
  toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>;
};

// --- Agent Continue Payload ---

export type AgentContinuePayload = {
  sessionId: string;
  toolCallId: string;
  result: string;
};

// --- Agent Stream Options ---

export type AgentStreamOptions = {
  enableReAct?: boolean;
  maxIterations?: number;
};

// --- Insert Mode ---

export type InsertMode = "chat_only" | "smart_action" | "replace_selection" | "append_end";

// --- Read Document Modes ---

export type ReadDocumentMode = "paragraph_range" | "heading_context" | "selection" | "cursor_surrounding";

// --- Rich Format Options ---

export type RichFormatOptions = {
  font?: { name?: string; size?: number; color?: string; bold?: boolean; italic?: boolean };
  hyperlink?: { text: string; url: string };
};

// --- Find and Replace V2 Options ---

export type FindReplaceV2Options = {
  find_text: string;
  replace_text: string;
  replace_all?: boolean;
  match_case?: boolean;
  match_whole_word?: boolean;
};
