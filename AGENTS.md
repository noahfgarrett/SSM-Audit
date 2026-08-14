# SSM Audit Agent Instructions

- Keep this application fully offline except for one anonymous GitHub release check when the HTML opens or refreshes.
- Keep all workbook processing local and session-only. Never persist, upload, log, or learn from a user's audit workbook.
- Treat audit targets as specimens only. Do not derive defaults, rules, profiles, or examples from them.
- Treat completed-project comparison references as temporary precedent only. Never persist them or use them to derive audit rules, defaults, or learned behavior.
- The governing sources are the SSM SOP, the approved Exto upload contract, and explicitly approved golden reference logic.
- Keep the audit rule engine behavior aligned with SSManagement. When changing a rule, deliberately update and test both applications.
- Never include confidential registry names, filenames, values, screenshots, or workbook contents in source, fixtures, commits, releases, or release notes.
- Use synthetic fixtures for committed tests. Confidential workbooks may be used only for local, uncommitted verification.
- Keep the source repository private. Publish only versioned single-file HTML and gzip artifacts to the public release repository.
- The updater must use only `noahfgarrett/SSM-Audit-Releases`; it must never consume SSManagement or SSM Builder releases.
- For every release, add brief user-facing changelog bullets such as:
  - Fixed an issue where ...
  - Added a new feature to ...
- Never mention disabled features, disabled checks, or disabled audit rules in changelogs, patch notes, update notes, or release notes.
- Keep release version values synchronized across `package.json`, `src/boot.js`, `src/changelog.json`, and release assets.
