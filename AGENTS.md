# SSM Audit Agent Instructions

- Keep this application fully offline except for one anonymous GitHub release check when the HTML opens or refreshes.
- Keep all workbook processing local and session-only. Never persist, upload, log, or learn from a user's audit workbook.
- Treat audit targets as specimens only. Do not derive defaults, rules, profiles, or examples from them.
- Treat completed-project comparison references as temporary precedent only. Never persist them or use them to derive audit rules, defaults, or learned behavior.
- The governing sources are the SSM SOP, the approved Exto upload contract, and explicitly approved golden reference logic.
- Keep the audit rule engine behavior aligned with SSManagement. When changing a rule, deliberately update and test both applications.
- Never include confidential registry names, filenames, values, screenshots, or workbook contents in source, fixtures, commits, releases, or release notes. The source is public, so this applies to git history as well: commit nothing that would need removing later.
- Use synthetic fixtures for committed tests. Confidential workbooks may be used only for local, uncommitted verification.
- The source repository is public (made public 2026-08-19). Every release is published twice with the same versioned single-file HTML and gzip artifacts: on `noahfgarrett/SSM-Audit-Releases` (the updater's source of truth, with `latest.json`) and as a GitHub release on `noahfgarrett/SSM-Audit` itself (tag `vX.Y.Z`) so the source repo's Releases page carries the file too. Never publish a bare tag without the release.
- The updater must use only `noahfgarrett/SSM-Audit-Releases`; it must never consume SSManagement or SSM Builder releases.
- For every release, add brief user-facing changelog bullets such as:
  - Fixed an issue where ...
  - Added a new feature to ...
- Never mention disabled features, disabled checks, or disabled audit rules in changelogs, patch notes, update notes, or release notes.
- Keep release version values synchronized across `package.json`, `src/boot.js`, `src/changelog.json`, and release assets.
- Every release must also update `latest.json` on the `main` branch of `noahfgarrett/SSM-Audit-Releases` (`{version, downloadUrl, releaseNotes, publishedAt}`; `downloadUrl` must end with `SSM-Audit-v<version>.html`). The updater reads it from raw.githubusercontent.com when the GitHub API check is rate-limited or blocked, so a stale manifest means old copies stop seeing new releases.
