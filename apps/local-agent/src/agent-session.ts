import { ChatMessage, AgentSession, ToolCallResult } from "./types.js";

const sessions = new Map<string, AgentSession>();

const SESSION_TTL_MS = 30 * 60 * 1000; // 30 minutes

function generateId(): string {
  return `sess-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function cleanupExpiredSessions(): void {
  const now = Date.now();
  for (const [id, session] of sessions) {
    if (now - session.updatedAt > SESSION_TTL_MS) {
      sessions.delete(id);
    }
  }
}

export function createSession(initialMessages: ChatMessage[]): string {
  cleanupExpiredSessions();
  const id = generateId();
  const session: AgentSession = {
    id,
    messages: [...initialMessages],
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
  sessions.set(id, session);
  return id;
}

export function getSession(id: string): AgentSession | undefined {
  const session = sessions.get(id);
  if (session) {
    session.updatedAt = Date.now();
  }
  return session;
}

export function appendToSession(
  id: string,
  assistantMessage: ChatMessage,
  toolResults?: ToolCallResult[]
): void {
  const session = sessions.get(id);
  if (!session) return;

  session.messages.push(assistantMessage);

  if (toolResults && toolResults.length > 0) {
    for (const tr of toolResults) {
      session.messages.push({
        role: "tool",
        content: tr.result,
        tool_call_id: tr.toolCallId,
        name: tr.toolName,
      });
    }
  }

  session.updatedAt = Date.now();
}

export function deleteSession(id: string): boolean {
  return sessions.delete(id);
}

// Periodic cleanup every 10 minutes
setInterval(cleanupExpiredSessions, 10 * 60 * 1000);
