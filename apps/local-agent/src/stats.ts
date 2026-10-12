import fs from "node:fs/promises";
import path from "node:path";
import { dataRoot, ensureDataRoot } from "./paths.js";

/**
 * Local usage statistics (batch A): LLM token usage + document-tool call
 * aggregation. Stored per calendar day under the user's data dir — never
 * leaves the machine. All mutations are serialized through one write chain so
 * concurrent hooks cannot lose updates.
 */

type DayStats = {
  promptTokens: number;
  completionTokens: number;
  llmCalls: number;
  toolCalls: number;
  toolFailures: number;
  toolByName: Record<string, { calls: number; failures: number; totalMs: number }>;
};

type StatsFile = {
  days: Record<string, DayStats>;
};

const MAX_DAYS = 180;

function emptyDay(): DayStats {
  return {
    promptTokens: 0,
    completionTokens: 0,
    llmCalls: 0,
    toolCalls: 0,
    toolFailures: 0,
    toolByName: {},
  };
}

function todayKey(): string {
  // Local date as YYYY-MM-DD
  return new Date().toLocaleDateString("sv-SE");
}

function statsFile(): string {
  return path.join(dataRoot, "usage-stats.json");
}

async function loadFile(): Promise<StatsFile> {
  try {
    const raw = await fs.readFile(statsFile(), "utf-8");
    const parsed = JSON.parse(raw) as StatsFile;
    if (parsed && typeof parsed.days === "object") return parsed;
  } catch {
    // first run or unreadable — start fresh
  }
  return { days: {} };
}

let writeChain: Promise<void> = Promise.resolve();

function mutate(fn: (day: DayStats, all: StatsFile) => void): void {
  writeChain = writeChain
    .then(async () => {
      await ensureDataRoot();
      const all = await loadFile();
      const key = todayKey();
      const day = (all.days[key] ??= emptyDay());
      fn(day, all);
      // Retention: keep the newest MAX_DAYS entries
      const keys = Object.keys(all.days).sort();
      while (keys.length > MAX_DAYS) {
        delete all.days[keys.shift() as string];
      }
      await fs.writeFile(statsFile(), JSON.stringify(all, null, 2), "utf-8");
    })
    .catch(() => {
      // stats must never break chat/tool flows
    });
}

/** Hook: called from llm.ts whenever a provider response reports usage. */
export function recordLlmUsage(promptTokens?: number, completionTokens?: number): void {
  const p = Number(promptTokens ?? 0) || 0;
  const c = Number(completionTokens ?? 0) || 0;
  if (p === 0 && c === 0) return;
  mutate((day) => {
    day.promptTokens += p;
    day.completionTokens += c;
    day.llmCalls += 1;
  });
}

export type ToolEvent = {
  name: string;
  success: boolean;
  ms?: number;
};

/** Hook: UI reports executed document tools (POST /v1/stats/report). */
export function recordToolEvents(events: ToolEvent[]): void {
  if (!Array.isArray(events) || events.length === 0) return;
  mutate((day) => {
    for (const e of events) {
      if (!e || typeof e.name !== "string") continue;
      day.toolCalls += 1;
      if (e.success === false) day.toolFailures += 1;
      const slot = (day.toolByName[e.name] ??= { calls: 0, failures: 0, totalMs: 0 });
      slot.calls += 1;
      if (e.success === false) slot.failures += 1;
      slot.totalMs += Number(e.ms ?? 0) || 0;
    }
  });
}

/** Aggregated view for GET /v1/stats. */
export async function getStatsSummary(): Promise<Record<string, unknown>> {
  const all = await loadFile();
  const today = all.days[todayKey()] ?? emptyDay();

  const totals = emptyDay();
  for (const day of Object.values(all.days)) {
    totals.promptTokens += day.promptTokens;
    totals.completionTokens += day.completionTokens;
    totals.llmCalls += day.llmCalls;
    totals.toolCalls += day.toolCalls;
    totals.toolFailures += day.toolFailures;
    for (const [name, slot] of Object.entries(day.toolByName)) {
      const t = (totals.toolByName[name] ??= { calls: 0, failures: 0, totalMs: 0 });
      t.calls += slot.calls;
      t.failures += slot.failures;
      t.totalMs += slot.totalMs;
    }
  }

  const toolRanking = Object.entries(totals.toolByName)
    .map(([name, s]) => ({
      name,
      calls: s.calls,
      failures: s.failures,
      avgMs: s.calls > 0 ? Math.round(s.totalMs / s.calls) : 0,
    }))
    .sort((a, b) => b.calls - a.calls)
    .slice(0, 10);

  return {
    today: {
      promptTokens: today.promptTokens,
      completionTokens: today.completionTokens,
      llmCalls: today.llmCalls,
      toolCalls: today.toolCalls,
      toolFailures: today.toolFailures,
    },
    total: {
      promptTokens: totals.promptTokens,
      completionTokens: totals.completionTokens,
      llmCalls: totals.llmCalls,
      toolCalls: totals.toolCalls,
      toolFailures: totals.toolFailures,
    },
    toolRanking,
    daysTracked: Object.keys(all.days).length,
  };
}
