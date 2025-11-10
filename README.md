# Word Add-in – Maximum Precision Tracked Revisions

This repository contains a Word (web/desktop) task pane add-in that applies rich-text diffs directly inside Word using native track-changes. It is built with React, Office.js, and Webpack, and focuses on applying deletion and insertion operations with single-character accuracy—even when Word’s internal text runs differ from plain text.

---

## Key Capabilities

- **Diff-driven editing** – Accepts structured diffs (`[{ op: 'delete'|'insert'|'equal', text: string }]`) and maps them onto the current Word selection.
- **Robust deletion logic** – Multi-layer search (unique context, progressive context sizes, extended patterns, fallbacks) to ensure the exact text fragment is removed once and only once.
- **Insertion handling** – Keeps Word’s tracking mode on for insertions; applies blue formatting so additions are visually distinct until accepted.
- **Cleanup helpers** – Utility functions to accept all/range/filtered tracked changes, converting temporary blue insertions back to black text while keeping tracking enabled.
- **Detailed logging** – Console output traces every strategy (with emojis and offsets) for debugging inside the task pane developer console.

---

## Repository Layout

| Path | Purpose |
|------|---------|
| `src/minimal/App.jsx` | React UI for the task pane; previews diffs, triggers Word actions. |
| `src/minimal/wordUtils.js` | Office.js utilities (tracked diff application, tracking state, acceptance helpers). |
| `src/minimal/diff.js` | Diff helper functions used to compute operations. |
| `src/minimal/tests.js` | Local test harness for diff/Word utility behavior. |
| `src/main.jsx` | Entry point; waits for `Office.onReady` before rendering React. |
| `taskpane.html` | Host page emitted by Webpack/HtmlWebpackPlugin (no manual script tags required). |
| `webpack.config.js` | Webpack configuration; externalizes Office.js and configures dev server. |

---

## Prerequisites

- Node.js 16+ (Node 18 recommended).
- npm 8+.
- Microsoft 365 account with Word (web or desktop) that allows sideloaded add-ins.
- Office.js dev certificates (`npx office-addin-dev-certs install`) for HTTPS dev server.

---

## Installation & Local Development

```bash
npm install
npm run dev-server
```

`npm run dev-server` launches Webpack dev server on https://localhost:3000 with hot module reload. The task pane consumes bundles from this endpoint.

### Sideloading the Add-in

1. **Word on the web**  
   - Open https://office.com → Word → create/open a document.  
   - Insert → Office Add-ins → Upload My Add-in → choose your manifest.  
   - Accept dev-certificate warnings if prompted.

2. **Word desktop**  
   - Use the same manifest and load it via Insert → My Add-ins → Shared Folder (or run relevant npm script if provided).  
   - Ensure Word trusts the localhost HTTPS certificate.

> Tip: Keep browser dev tools open on the task pane for real-time logs (`wordUtils.js` prints extensive diagnostics).

---

## Working with Diffs

1. Select the text in Word you want to replace.
2. Provide diffs (programmatically or through UI inputs in `App.jsx`).
3. Click **Preview Changes**.  
   - Word helpers normalize text, build position maps, and apply deletions (Phase 1) then insertions (Phase 2), logging each step.  
   - Deletions show as red tracked strikethrough; insertions show as tracked underlines formatted in blue.
4. Use the acceptance buttons (All / Range / Filtered) to accept changes and convert blue insertions back to black text while keeping tracking enabled.

---

## Core Utility Highlights (`wordUtils.js`)

- `normalizeText` strips carriage returns and normalizes Unicode (NFC), aligning Word’s search behavior with diff text.
- `findUniqueContext` expands context windows until a unique pattern is found, guaranteeing accurate targeting.
- `replaceSelectionWithNativeTrackedRevisions` orchestrates operations:
  - Builds position maps and verifies uniqueness.
  - Processes deletions with multi-level strategies (unique pattern, progressive contexts, extended context fallbacks, last-resort searches).
  - Processes insertions with tracking ON and temporary formatting.
  - Provides rich console logging at every decision point.
- `accept*` helpers accept tracked changes, recolor blue additions, and ensure tracking mode remains `TrackAll`.

---

## Scripts

| Command | Description |
|---------|-------------|
| `npm run dev-server` | Start Webpack dev server (default workflow). |
| `npm run build` | Production build (if defined in `package.json`). |
| `npm run lint` / `npm run test` | Add as needed for linting/tests (see package file). |

---

## Troubleshooting

| Symptom | Likely Cause | Remedy |
|---------|--------------|--------|
| `TypeError: r is not a function` | Office.js not ready or bundled | Ensure `Office.onReady` wraps React render (`src/main.jsx`); `webpack.config.js` should externalize `office`. |
| 404 for `src/main.jsx` | Manual script tag in `taskpane.html` | Remove manual `<script>`; let HtmlWebpackPlugin inject assets. |
| HMR warning “Cannot apply update” | Word task pane limitations | Reload task pane (right-click pane → Reload). |
| Insertions appear untracked | Tracking turned off before insertion | Current code keeps tracking ON; verify no local edits reintroduce toggles. |
| Deletion skipped with “Pattern not found” | Context mismatch | Review console logs; ensure selection text matches diff source or reduce context size. |

---

## Deployment

1. Build production assets: `npm run build`.
2. Host `dist/` over HTTPS (Azure Static Apps, AWS S3 + CloudFront, etc.).
3. Update manifest `SourceLocation` to production URL.
4. Deploy manifest via Microsoft 365 admin center or shared catalog.

---

## Contributing

1. Fork the repo, create a feature branch (`git checkout -b feature/<name>`).
2. Make changes, add relevant tests/logging updates.
3. Ensure lint/test scripts pass.
4. Open a pull request with context and screenshots (if UI changes).

---

## Support

- Office Add-ins docs: <https://learn.microsoft.com/office/dev/add-ins/>
- Office dev tooling: <https://github.com/OfficeDev/office-toolbox>
- LINE support: n/a — this add-in focuses on Word operations.  
- For project-specific issues, open a GitHub issue with reproduction steps and relevant console logs.

Happy editing! 🎯
