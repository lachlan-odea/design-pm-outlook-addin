import "./styles.css";
import { isConfigured, loadConfig, saveConfig } from "./settings";
import { appendProject, fetchWorkspace } from "./jsonbin";
import { readCurrentEmail, type EmailContext } from "./email";
import {
  PRODUCT_AREAS,
  type Designer,
  type JsonBinConfig,
  type Priority,
  type Project,
  type Workspace,
} from "./types";

let config: JsonBinConfig = { binId: "", apiKey: "", accessKey: "" };
let workspace: Workspace | null = null;
let email: EmailContext | null = null;

Office.onReady(async () => {
  config = loadConfig();
  bindUi();
  if (!isConfigured(config)) {
    showView("settings");
    return;
  }
  await bootForm();
});

function bindUi() {
  $("#settings-btn")!.addEventListener("click", () => {
    populateSettingsForm();
    showView("settings");
  });
  $("#settings-save")!.addEventListener("click", onSettingsSave);
  $("#settings-test")!.addEventListener("click", onSettingsTest);
  $("#form-reset")!.addEventListener("click", () => {
    if (email && workspace) fillForm(email, workspace);
  });
  $("#form-submit")!.addEventListener("click", onSubmit);
}

async function bootForm() {
  showView("loading");
  try {
    [workspace, email] = await Promise.all([fetchWorkspace(config), readCurrentEmail()]);
    fillForm(email, workspace);
    showView("form");
  } catch (err) {
    showSettingsError(formatError(err));
    populateSettingsForm();
    showView("settings");
  }
}

function fillForm(ctx: EmailContext, ws: Workspace) {
  populateAssignees(ws.designers);
  populateBrands(ws);
  populateProductAreas();

  setVal("f-title", ctx.subject || "Untitled brief");
  setVal(
    "f-overview",
    truncate(ctx.bodyText.replace(/​/g, "").trim(), 600) ||
      "Brief sourced from email."
  );
  setVal("f-owner", ctx.senderName || ctx.senderEmail);
  setVal("f-client", "");
  setVal("f-commenced", todayISO());
  setVal("f-due", ctx.detectedDueDate ?? "");
  setVal("f-priority", ctx.detectedPriority ?? "Normal");
  setVal("f-brand", "");
  setVal("f-area", "");
  setVal("f-brief", ctx.firstUrl ?? "");

  // Default assignee: signed-in user from the workspace if their name matches
  // the local part of their email, otherwise leave unassigned.
  const myEmail = Office.context.mailbox.userProfile?.emailAddress?.toLowerCase();
  let meId = "";
  if (myEmail) {
    const handle = myEmail.split("@")[0].replace(/[._]/g, " ");
    const match = ws.designers.find((d) => d.name.toLowerCase().includes(handle));
    if (match) meId = match.id;
  }
  setVal("f-assignee", meId);

  setBanner(null);
}

function populateAssignees(designers: Designer[]) {
  const sel = el<HTMLSelectElement>("f-assignee");
  sel.innerHTML = "";
  sel.appendChild(option("", "— Unassigned —"));
  designers.forEach((d) => sel.appendChild(option(d.id, d.name)));
}

function populateBrands(ws: Workspace) {
  const list = el<HTMLDataListElement>("brand-options");
  const brands = new Set<string>();
  ws.projects.forEach((p) => p.brand && brands.add(p.brand));
  list.innerHTML = "";
  [...brands].sort().forEach((b) => {
    const opt = document.createElement("option");
    opt.value = b;
    list.appendChild(opt);
  });
}

function populateProductAreas() {
  const sel = el<HTMLSelectElement>("f-area");
  sel.innerHTML = "";
  sel.appendChild(option("", "— select —"));
  PRODUCT_AREAS.forEach((p) => sel.appendChild(option(p, p)));
  sel.appendChild(option("__other__", "Other…"));
  sel.addEventListener("change", () => {
    if (sel.value !== "__other__") return;
    const custom = prompt("Product area:");
    if (custom) {
      const opt = option(custom, custom);
      sel.insertBefore(opt, sel.querySelector('option[value="__other__"]'));
      sel.value = custom;
    } else {
      sel.value = "";
    }
  });
}

function populateSettingsForm() {
  setVal("cfg-bin-id", config.binId);
  setVal("cfg-api-key", config.apiKey);
  setVal("cfg-access-key", config.accessKey ?? "");
  showSettingsError(null);
}

async function onSettingsSave() {
  const next: JsonBinConfig = {
    binId: getVal("cfg-bin-id").trim(),
    apiKey: getVal("cfg-api-key").trim(),
    accessKey: getVal("cfg-access-key").trim() || undefined,
  };
  if (!next.binId || !next.apiKey) {
    showSettingsError("Bin ID and master key are required.");
    return;
  }
  try {
    await saveConfig(next);
    config = next;
    await bootForm();
  } catch (err) {
    showSettingsError(formatError(err));
  }
}

async function onSettingsTest() {
  showSettingsError(null);
  const trial: JsonBinConfig = {
    binId: getVal("cfg-bin-id").trim(),
    apiKey: getVal("cfg-api-key").trim(),
    accessKey: getVal("cfg-access-key").trim() || undefined,
  };
  try {
    await fetchWorkspace(trial);
    showSettingsError("✓ Connected. Bin looks good.");
  } catch (err) {
    showSettingsError(formatError(err));
  }
}

async function onSubmit() {
  if (!workspace) return;
  const title = getVal("f-title").trim();
  if (!title) {
    showFormError("Title is required.");
    return;
  }
  const project: Project = {
    id: `p-${Date.now()}`,
    title,
    overview: getVal("f-overview").trim(),
    owner: getVal("f-owner").trim(),
    client: getVal("f-client").trim(),
    brand: getVal("f-brand").trim(),
    productArea: getVal("f-area") === "__other__" ? "" : getVal("f-area"),
    briefUrl: getVal("f-brief").trim(),
    dueDate: getVal("f-due"),
    priority: getVal("f-priority") as Priority,
    assigneeId: getVal("f-assignee") || null,
    milestones: [],
    comments: [],
    createdAt: getVal("f-commenced")
      ? new Date(`${getVal("f-commenced")}T00:00:00Z`).toISOString()
      : new Date().toISOString(),
    source: "outlook",
  };

  setSubmitting(true);
  showFormError(null);
  try {
    workspace = await appendProject(config, project);
    setBanner(`✓ Added "${project.title}" to Design PM.`);
  } catch (err) {
    showFormError(formatError(err));
  } finally {
    setSubmitting(false);
  }
}

function setSubmitting(busy: boolean) {
  const btn = el<HTMLButtonElement>("form-submit");
  btn.disabled = busy;
  btn.textContent = busy ? "Adding…" : "Add project";
}

function showView(name: "loading" | "settings" | "form") {
  ["loading", "settings", "form"].forEach((n) => {
    const node = document.getElementById(`view-${n}`);
    if (!node) return;
    node.classList.toggle("hidden", n !== name);
  });
}

function showSettingsError(msg: string | null) {
  const node = document.getElementById("settings-error")!;
  if (!msg) {
    node.classList.add("hidden");
    node.textContent = "";
    return;
  }
  node.classList.remove("hidden");
  node.textContent = msg;
}

function showFormError(msg: string | null) {
  const node = document.getElementById("form-error")!;
  if (!msg) {
    node.classList.add("hidden");
    node.textContent = "";
    return;
  }
  node.classList.remove("hidden");
  node.textContent = msg;
}

function setBanner(msg: string | null) {
  const node = document.getElementById("form-banner")!;
  if (!msg) {
    node.classList.add("hidden");
    node.textContent = "";
    return;
  }
  node.classList.remove("hidden");
  node.textContent = msg;
}

function $(sel: string): HTMLElement | null {
  return document.querySelector(sel);
}
function el<T extends HTMLElement = HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing element #${id}`);
  return node as T;
}
function getVal(id: string): string {
  return el<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(id).value;
}
function setVal(id: string, value: string) {
  el<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>(id).value = value;
}
function option(value: string, label: string): HTMLOptionElement {
  const o = document.createElement("option");
  o.value = value;
  o.textContent = label;
  return o;
}
function truncate(s: string, max: number): string {
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}
function todayISO(): string {
  return new Date().toISOString().slice(0, 10);
}
function formatError(err: unknown): string {
  if (err instanceof Error) return err.message;
  return String(err);
}
