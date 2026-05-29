# Design PM — Outlook add-in

Reads the open Outlook message and creates a project in your Design PM JSONBin
in one click. Pairs with the [Design PM React app](../project%20manangement%20tool/).

## What it does

- Adds a "Brief to Design PM" button to the Outlook ribbon on read messages.
- The task pane pre-fills a project from the email:
  - **Title** ← subject
  - **Overview** ← body (truncated to 600 chars)
  - **Owner** ← sender name / email
  - **Commenced** ← today
  - **Due date** ← best-effort parse of "due by …", "deadline …", ISO dates, dd/mm/yyyy
  - **Priority** ← detects "urgent", "asap", "important", "low priority" etc.
  - **Brief URL** ← first http(s) link in the body
  - **Assignee / Brand / Content type** ← dropdowns sourced live from the workspace bin
- One click writes the new project into the same JSONBin that the React app reads,
  so it appears in everyone's Design PM immediately. The project is flagged with
  `source: "outlook"`.

## Architecture

```
Outlook (Win/Mac/Web) ─▶ Office.js ─▶ task pane (this add-in)
                                          │
                                          ▼ GET /v3/b/{id}/latest
                                       JSONBin  ◀── React PM app
                                          ▲ PUT /v3/b/{id}
                                          │
                                       project appended
```

The add-in reads the workspace, pushes a new project to the front of
`projects[]`, then PUTs the whole record back. JSONBin doesn't support
atomic append, so this is last-write-wins — fine for a team of nine.

## Local development

Prereqs: Node 20+, Outlook desktop or Outlook on the web, a JSONBin bin
already initialised by the React app.

```bash
cd design-pm-outlook-addin
npm install
npm run certs:install   # one-time: installs a trusted CA cert for localhost
npm run manifest:dev    # writes manifest.dev.xml pointing at https://localhost:3000
npm run dev             # starts the HTTPS dev server on :3000
```

`npm run certs:install` runs Microsoft's `office-addin-dev-certs` tool, which
adds a CA cert to your Windows Trusted Root store. After this Outlook (and your
browser) trust `https://localhost:3000` without any warnings — required for
Outlook on the web to load the task pane during install.

To check the manifest before sideloading:

```bash
npm run manifest:validate
```

### Sideload the add-in

**Outlook on the web** (easiest for testing):
1. Open Outlook on the web, click any message.
2. Click **... → Get Add-ins → My add-ins → Add a custom add-in → Add from file…**
3. Pick `manifest.dev.xml` from this folder.

**Outlook desktop (Windows)**:
1. Go to **File → Manage Add-ins** (opens browser).
2. Same flow as above.

The **Design PM** group with **Brief to Design PM** appears on the ribbon when
viewing an email. Click it → enter your JSONBin bin ID + master key → done.

### Troubleshooting

**"Add-in installation failed"** in Outlook on the web — almost always one of:
- `npm run dev` isn't running. The server must be up while Outlook fetches the
  manifest URLs at install time.
- The dev cert isn't trusted yet. Run `npm run certs:install` once, then
  restart the dev server. Confirm by opening `https://localhost:3000/icons/icon-64.png`
  in the same browser; it should load with no warning.
- The manifest has a schema issue. Run `npm run manifest:validate` and fix
  anything it flags.

## Deploying to GitHub Pages

1. Push this folder to a GitHub repo (e.g. `design-pm-outlook-addin`).
2. In repo **Settings → Pages**, set **Source: GitHub Actions**.
3. Push to `main`. The `.github/workflows/deploy.yml` workflow:
   - Builds the site with `VITE_BASE=/design-pm-outlook-addin/`.
   - Writes `manifest.prod.xml` with `BASE_URL=https://<owner>.github.io/<repo>`.
   - Publishes everything to Pages.
4. After deploy, download `manifest.prod.xml` from the deployed site
   (`https://<owner>.github.io/<repo>/manifest.xml`) and sideload it in Outlook
   for production users — or upload it to **Microsoft 365 admin → Integrated apps**
   for centralised deployment across your tenant.

## Settings storage

JSONBin credentials live in `Office.context.roamingSettings` (per-user,
per-add-in, synced across Outlook sessions). When developing outside Office
(plain browser), the add-in falls back to `localStorage`.

## File guide

| File | Purpose |
|---|---|
| `manifest.xml` | Templated Office Add-in manifest. `{{BASE_URL}}` is replaced by `scripts/build-manifest.mjs`. |
| `index.html` | Task pane entry. Loads Office.js + `src/main.ts`. |
| `commands.html` | Required by manifest's `FunctionFile`; currently no-op. |
| `src/main.ts` | UI controller — wires the form and submit flow. |
| `src/email.ts` | Reads the current Outlook message and best-effort parses dates / priority. |
| `src/jsonbin.ts` | Read/write the workspace bin. |
| `src/settings.ts` | Persist JSONBin creds via `roamingSettings`. |
| `src/types.ts` | Mirror of the PM tool's `Workspace` / `Project` types. |
| `scripts/generate-icons.mjs` | Generates solid purple PNG placeholder icons. Replace with branded artwork. |
| `scripts/build-manifest.mjs` | Replaces `{{BASE_URL}}` in `manifest.xml` for dev / prod. |
| `.github/workflows/deploy.yml` | Builds + publishes to GitHub Pages. |

## Replacing the icons

The default purple squares come from `scripts/generate-icons.mjs`. Drop your own
PNGs into `public/icons/` named `icon-16.png`, `icon-32.png`, `icon-64.png`,
`icon-80.png`, `icon-128.png` — or run `npm run icons -- --force` to overwrite.

## Privacy / security note

The master JSONBin key is stored in `roamingSettings`, which is encrypted at
rest by Microsoft 365 but is recoverable by anyone who controls the user's
mailbox. If you wire this up to a sensitive workspace, gate the bin with an
**access key** (the third field in the settings panel) and treat the master
key as a deploy-time secret.
