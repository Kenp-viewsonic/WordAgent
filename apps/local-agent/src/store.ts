import fs from "node:fs/promises";
import path from "node:path";
import { ProviderConfig, ConfigPreset, RetrievalChunk } from "./types.js";
import { dataRoot, ensureDataRoot } from "./paths.js";

const configFile = path.join(dataRoot, "provider-config.json");
const presetsFile = path.join(dataRoot, "config-presets.json");
const kbFile = path.join(dataRoot, "kb-chunks.json");

// ─── Single Config (legacy) ─────────────────────────────────────────────────

export async function saveProviderConfig(config: ProviderConfig): Promise<void> {
  await ensureDataRoot();
  await fs.writeFile(configFile, JSON.stringify(config, null, 2), "utf-8");
}

export async function loadProviderConfig(): Promise<ProviderConfig | null> {
  try {
    const raw = await fs.readFile(configFile, "utf-8");
    return JSON.parse(raw) as ProviderConfig;
  } catch {
    return null;
  }
}

// ─── Config Presets ─────────────────────────────────────────────────────────

export async function loadPresets(): Promise<ConfigPreset[]> {
  try {
    const raw = await fs.readFile(presetsFile, "utf-8");
    return JSON.parse(raw) as ConfigPreset[];
  } catch {
    return [];
  }
}

export async function savePreset(preset: ConfigPreset): Promise<ConfigPreset[]> {
  const presets = await loadPresets();
  const idx = presets.findIndex((p) => p.id === preset.id);
  if (idx >= 0) {
    presets[idx] = { ...preset, updatedAt: Date.now() };
  } else {
    presets.push({ ...preset, createdAt: Date.now(), updatedAt: Date.now() });
  }
  await ensureDataRoot();
  await fs.writeFile(presetsFile, JSON.stringify(presets, null, 2), "utf-8");
  return presets;
}

export async function deletePreset(id: string): Promise<ConfigPreset[]> {
  const presets = await loadPresets();
  const filtered = presets.filter((p) => p.id !== id);
  await ensureDataRoot();
  await fs.writeFile(presetsFile, JSON.stringify(filtered, null, 2), "utf-8");
  return filtered;
}

export async function appendChunks(chunks: RetrievalChunk[]): Promise<void> {
  await ensureDataRoot();
  const current = await loadChunks();
  const merged = [...current, ...chunks];
  await fs.writeFile(kbFile, JSON.stringify(merged, null, 2), "utf-8");
}

export async function loadChunks(): Promise<RetrievalChunk[]> {
  try {
    const raw = await fs.readFile(kbFile, "utf-8");
    return JSON.parse(raw) as RetrievalChunk[];
  } catch {
    return [];
  }
}

export async function deleteChunksByFile(fileName: string): Promise<number> {
  const current = await loadChunks();
  const remaining = current.filter((chunk) => chunk.fileName !== fileName);
  const removed = current.length - remaining.length;
  await ensureDataRoot();
  await fs.writeFile(kbFile, JSON.stringify(remaining, null, 2), "utf-8");
  return removed;
}

export async function clearAllChunks(): Promise<number> {
  const current = await loadChunks();
  const removed = current.length;
  await ensureDataRoot();
  await fs.writeFile(kbFile, "[]", "utf-8");
  return removed;
}

export async function getKbFileList(): Promise<Array<{ fileName: string; chunkCount: number }>> {
  const chunks = await loadChunks();
  const map = new Map<string, number>();
  for (const chunk of chunks) {
    map.set(chunk.fileName, (map.get(chunk.fileName) ?? 0) + 1);
  }
  return [...map.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([fileName, chunkCount]) => ({ fileName, chunkCount }));
}

export async function getKbStats(): Promise<{
  chunkCount: number;
  fileCount: number;
  files: string[];
}> {
  const chunks = await loadChunks();
  const fileSet = new Set(chunks.map((chunk) => chunk.fileName));
  return {
    chunkCount: chunks.length,
    fileCount: fileSet.size,
    files: [...fileSet].sort((a, b) => a.localeCompare(b)).slice(-20),
  };
}

export async function importChunks(chunks: RetrievalChunk[], mode: "merge" | "replace"): Promise<{ importedChunks: number; totalChunks: number }> {
  await ensureDataRoot();
  const existing = mode === "replace" ? [] : await loadChunks();
  const merged = [...existing, ...chunks];
  await fs.writeFile(kbFile, JSON.stringify(merged, null, 2), "utf-8");
  return { importedChunks: chunks.length, totalChunks: merged.length };
}
