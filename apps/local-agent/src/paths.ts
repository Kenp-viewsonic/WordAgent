import fs from "node:fs";
import fsp from "node:fs/promises";
import os from "node:os";
import path from "node:path";

/**
 * Central path resolution for the local agent.
 *
 * Must work in every runtime shape we ship (see §6.5 of
 * plans/com-shell-migration-kickoff.md):
 *  - dev (tsx / node dist/server.js): resources next to the module
 *  - packaged agent.exe (Node SEA): resources next to the executable
 *  - data always lives in %LOCALAPPDATA%\OfficeAgent\data because the
 *    install directory (Program Files) is not writable.
 */

// __dirname exists in the esbuild CJS bundle / SEA runtime; under tsx (ESM)
// fall back to the entry script directory.
function moduleDir(): string {
  // @ts-ignore __dirname is not defined in ESM
  if (typeof __dirname === "string") {
    // @ts-ignore
    return __dirname as string;
  }
  const entry = process.argv[1];
  return entry ? path.dirname(entry) : process.cwd();
}

const sea = (process as { sea?: unknown }).sea;
export const isPackaged: boolean =
  (typeof sea === "object" && sea !== null) ||
  path.basename(process.execPath).toLowerCase() === "agent.exe";

/** Resource root: where the bundled `public/` UI lives. */
export const appRoot: string =
  process.env.OFFICE_AGENT_HOME ??
  (isPackaged ? path.dirname(process.execPath) : moduleDir());

const localAppData =
  process.env.LOCALAPPDATA ?? path.join(os.homedir(), "AppData", "Local");

/** Install-state root: token file, WebView2 profile, logs. */
export const configRoot: string = path.join(localAppData, "OfficeAgent");

/** Writable data root (provider config, KB chunks, trace log). */
export const dataRoot: string =
  process.env.OFFICE_AGENT_DATA_DIR ?? path.join(configRoot, "data");

export const tokenFile: string = path.join(configRoot, "agent.token");
export const traceFile: string = path.join(dataRoot, "agent-trace.jsonl");

const managedFiles = [
  "provider-config.json",
  "config-presets.json",
  "kb-chunks.json",
];

/** Create the data root and migrate legacy ./data contents once. */
export async function ensureDataRoot(): Promise<void> {
  await fsp.mkdir(dataRoot, { recursive: true });
  await migrateLegacyData();
}

let migrated = false;

async function migrateLegacyData(): Promise<void> {
  if (migrated) return;
  migrated = true;
  const legacyDirs = [path.join(process.cwd(), "data"), path.join(appRoot, "data")];
  for (const legacyDir of legacyDirs) {
    if (path.resolve(legacyDir) === path.resolve(dataRoot)) continue;
    for (const name of managedFiles) {
      const target = path.join(dataRoot, name);
      if (fs.existsSync(target)) continue;
      const source = path.join(legacyDir, name);
      try {
        await fsp.copyFile(source, target);
        console.log(`[paths] migrated legacy ${name} from ${legacyDir}`);
      } catch {
        // no legacy file — nothing to migrate
      }
    }
  }
}
