import type { JsonBinConfig, Project, Workspace } from "./types";

const JSONBIN_BASE = "https://api.jsonbin.io/v3/b";

// Bin ID sentinel that routes reads/writes to localStorage instead of JSONBin,
// so the add-in can be exercised end-to-end while JSONBin is unreachable.
export const MOCK_BIN_ID = "mock";
const MOCK_STORAGE_KEY = "design-pm-mock-workspace";

function isMock(cfg: JsonBinConfig): boolean {
  return cfg.binId.trim().toLowerCase() === MOCK_BIN_ID;
}

function seedMockWorkspace(): Workspace {
  return {
    designers: [
      { id: "d-lachlan", name: "Lachlan O'Dea", initials: "LO", color: "#7c3aed", pin: "0000" },
      { id: "d-sample", name: "Sample Designer", initials: "SD", color: "#22d3ee", pin: "0000" },
    ],
    projects: [],
    currentDesignerId: "d-lachlan",
  };
}

function loadMockWorkspace(): Workspace {
  const raw = localStorage.getItem(MOCK_STORAGE_KEY);
  if (raw) {
    try {
      const parsed = JSON.parse(raw) as Workspace;
      if (parsed && Array.isArray(parsed.projects) && Array.isArray(parsed.designers)) {
        return parsed;
      }
    } catch {
      // fall through to seed
    }
  }
  const seeded = seedMockWorkspace();
  localStorage.setItem(MOCK_STORAGE_KEY, JSON.stringify(seeded));
  return seeded;
}

function saveMockWorkspace(ws: Workspace): void {
  localStorage.setItem(MOCK_STORAGE_KEY, JSON.stringify(ws));
}

function headers(cfg: JsonBinConfig, write = false): HeadersInit {
  const h: Record<string, string> = { "X-Master-Key": cfg.apiKey };
  if (cfg.accessKey) h["X-Access-Key"] = cfg.accessKey;
  if (write) h["Content-Type"] = "application/json";
  return h;
}

export async function fetchWorkspace(cfg: JsonBinConfig): Promise<Workspace> {
  if (isMock(cfg)) return loadMockWorkspace();
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
  if (isMock(cfg)) {
    saveMockWorkspace(workspace);
    return;
  }
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
