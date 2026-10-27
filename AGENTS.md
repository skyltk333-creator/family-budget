# Repository Guidelines

## Project Structure & Module Organization

This repository is a dependency-free, Korean-language family budget dashboard that processes Shinhan Card and Toss Bank XLSX statements in the browser.

- `index.html`: page structure, upload controls, dashboard, and transaction table.
- `style.css`: shared styling and responsive layouts, including the 720px breakpoint.
- `app.js`: XLSX ZIP/XML parsing, bank-specific readers, classification, deduplication, rendering, and CSV export.
- `README.md`: local setup, Cloudflare Pages deployment, and supported-file limitations.

All application files live at the repository root. There are no separate asset or test directories, package manifest, or build pipeline.

## Build, Test, and Development Commands

- `python3 -m http.server 8000`: serve the repository locally; open `http://localhost:8000` or the forwarded Codespaces port.
- `node --check app.js`: check JavaScript syntax when Node.js is available; this does not validate browser behavior.

No dependency installation or build command is required. Cloudflare Pages serves the root directory with framework preset None and an empty build command.

## Coding Style & Naming Conventions

Use plain JavaScript with strict mode, native browser APIs, and no external runtime dependencies. Match nearby formatting: JavaScript currently uses compact expressions and one-space indentation inside blocks; HTML and CSS are compact. Avoid unrelated reformatting.

Use camelCase for functions and variables, uppercase names for shared constants such as `CATS`, and existing lowercase DOM IDs. Keep visible interface text in Korean. Render imported text through `textContent` rather than HTML interpolation.

## Testing Guidelines

There is no automated test framework or coverage threshold. Validate changes manually in Chrome or Edge with support for `DecompressionStream('deflate-raw')`.

Start with the `가상 데이터 체험` button, which loads 90 synthetic transactions. Check month selection, search, category filters, category edits, recalculated totals, CSV export, and reset. For parser changes, use synthetic XLSX statements for both banks and exercise invalid-file errors. Check narrow-screen layouts for UI changes.

## Commit & Pull Request Guidelines

Git history contains only `Add files via upload`, so no established commit convention exists. Use concise, imperative subjects describing the change. Pull requests should explain the behavior changed, list manual verification, link relevant issues, and include screenshots for visual changes.

## Privacy & Configuration

Keep statement processing local and transaction data in browser memory. Never commit real financial statements, exported transaction CSVs, or screenshots containing personal financial information. Use synthetic fixtures and preserve the application's no-network-request behavior.
