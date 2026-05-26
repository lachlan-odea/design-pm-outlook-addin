import type { JsonBinConfig, Project, Workspace } from "./types";

const JSONBIN_BASE = "https://api.jsonbin.io/v3/b";

function headers(cfg: JsonBinConfig, write = false): HeadersInit {
  const h: Record<string, string> = { "X-Master-Key": cfg.apiKey };
  if (cfg.accessKey) h["X-Access-Key"] = cfg.accessKey;
  if (write) h["Content-Type"] = "application/json";
  return h;
}

export async function fetchWorkspace(cfg: JsonBinConfig): Promise<Workspace> {
  if (!cfg.binId || !cfg.apiKey) throw new Error("JSONBin is not configured.");
  const res = await fetch(`${JSONBIN_BASE}/${cfg.binId}/latest`, {
    headers: headers(cfg),
  });
  if (!res.ok) {
    throw new Error(`JSONBin GET failed: ${res.status} ${res.statusText}`);
  }
  const json = (await res.json()) as { record: Workspace };
  if (!json.record || !Array.isArray(json.record.projects)) {
    throw new Error("Bin doesn't look like a Design PM workspace.");
  }
  return json.record;
}

export async function putWorkspace(
  cfg: JsonBinConfig,
  workspace: Workspace
): Promise<void> {
  const res = await fetch(`${JSONBIN_BASE}/${cfg.binId}`, {
    method: "PUT",
    headers: headers(cfg, true),
    body: JSON.stringify(workspace),
  });
  if (!res.ok) {
    throw new Error(`JSONBin PUT failed: ${res.status} ${res.statusText}`);
  }
}

export async function appendProject(
  cfg: JsonBinConfig,
  project: Project
): Promise<Workspace> {
  // Read-then-write. JSONBin doesn't support atomic append, so this is
  // last-write-wins. For a 9-person design team that's acceptable.
  const ws = await fetchWorkspace(cfg);
  const next: Workspace = { ...ws, projects: [project, ...ws.projects] };
  await putWorkspace(cfg, next);
  return next;
}
