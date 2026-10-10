import { RetrievalChunk } from "./types.js";

function normalizeText(input: string): string {
  return input
    .toLowerCase()
    .replace(/\r\n/g, "\n")
    .replace(/[\u3000\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildQueryTokens(query: string): string[] {
  const normalized = normalizeText(query);
  const latinWords = normalized.match(/[a-z0-9_\-\.]{2,}/g) ?? [];
  const cjkChars = normalized.match(/[\u4e00-\u9fff]/g) ?? [];

  const cjkBigrams: string[] = [];
  for (let i = 0; i < cjkChars.length - 1; i += 1) {
    cjkBigrams.push(cjkChars[i] + cjkChars[i + 1]);
  }

  const tokenSet = new Set<string>([...latinWords, ...cjkBigrams]);
  return [...tokenSet].filter((token) => token.length >= 2);
}

export function splitToChunks(fileName: string, content: string, size = 700): RetrievalChunk[] {
  const normalized = content.replace(/\r\n/g, "\n");
  const chunks: RetrievalChunk[] = [];

  for (let i = 0, cursor = 0; cursor < normalized.length; i += 1, cursor += size) {
    const text = normalized.slice(cursor, cursor + size).trim();
    if (!text) {
      continue;
    }
    chunks.push({
      id: `${fileName}-${i}`,
      fileName,
      text,
    });
  }

  return chunks;
}

export function keywordRetrieve(chunks: RetrievalChunk[], query: string, limit = 4): RetrievalChunk[] {
  const tokens = buildQueryTokens(query);

  if (tokens.length === 0) {
    return chunks.slice(-limit);
  }

  const ranked = [...chunks]
    .map((chunk) => {
      const lower = chunk.text.toLowerCase();
      const nameLower = chunk.fileName.toLowerCase();
      const score = tokens.reduce((acc, token) => {
        const contentHit = lower.includes(token) ? 1 : 0;
        const fileNameHit = nameLower.includes(token) ? 0.5 : 0;
        return acc + contentHit + fileNameHit;
      }, 0);
      return { chunk, score };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
    .map((item) => item.chunk);

  return ranked.length > 0 ? ranked : chunks.slice(-limit);
}
