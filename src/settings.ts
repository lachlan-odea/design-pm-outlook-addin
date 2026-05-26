import type { JsonBinConfig } from "./types";

// Office.context.roamingSettings persists per-mailbox per-add-in and syncs across
// the user's Outlook clients. Falls back to localStorage outside of Office (e.g. when
// developing the pane in a browser).

const KEY = "designpm.jsonbin.config.v1";

function inOffice(): boolean {
  return typeof Office !== "undefined" && !!Office.context?.roamingSettings;
}

export function loadConfig(): JsonBinConfig {
  let raw: unknown;
  if (inOffice()) {
    raw = Office.context.roamingSettings.get(KEY);
  } else {
    try {
      const s = localStorage.getItem(KEY);
      raw = s ? JSON.parse(s) : null;
    } catch {
      raw = null;
    }
  }
  if (raw && typeof raw === "object") {
    const o = raw as Record<string, unknown>;
    return {
      binId: typeof o.binId === "string" ? o.binId : "",
      apiKey: typeof o.apiKey === "string" ? o.apiKey : "",
      accessKey: typeof o.accessKey === "string" ? o.accessKey : "",
    };
  }
  return { binId: "", apiKey: "", accessKey: "" };
}

export async function saveConfig(cfg: JsonBinConfig): Promise<void> {
  if (inOffice()) {
    Office.context.roamingSettings.set(KEY, cfg);
    await new Promise<void>((resolve, reject) => {
      Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message ?? "saveAsync failed"));
      });
    });
  } else {
    localStorage.setItem(KEY, JSON.stringify(cfg));
  }
}

export function isConfigured(cfg: JsonBinConfig): boolean {
  return !!cfg.binId && !!cfg.apiKey;
}
