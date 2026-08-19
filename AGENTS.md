# Repository Guidelines

- Google Apps Script runtime files and `appsscript.json` belong in `src/`.
- Keep development documentation, tests, and Node.js tooling outside `src/`.
- Do not change the existing `.clasp.json` Script ID unless explicitly requested.
- Run `clasp status` before `clasp push` and verify that only files under `src/` are listed.
