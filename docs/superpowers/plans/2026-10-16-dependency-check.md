# Dependency Check Implementation Plan

> **For agentic workers:** executed inline (user asked to build immediately after the spec). Steps use checkbox syntax.

**Goal:** Dependencies tab that derives expected dependencies from EasyPower / Cable / PMD (xlsx or CSV) with SSManagement's partition-fold rule, compares them to the registry, and exports the result.

**Architecture:** Pure `expected-dependencies.js` computes `{rows, notInRegistry, counts}` inside the existing review worker; findings swap the old relationship rule for an expected-dependency rule; a new `ui/dependencies.js` screen renders and exports via `dependency-export.js`.

**Tech Stack:** Vanilla ES modules bundled by `build/build.mjs` (manifest order), vendored SheetJS (xlsx-js-style), `node --test`.

Spec: `docs/superpowers/specs/2026-10-16-dependency-check-design.md`

---

### Task 1: Core rule — `src/audit/expected-dependencies.js`
- [ ] Write `tests/expected-dependencies.test.mjs` (synthetic): cross-group → expected; same group → none;
      electrical+same building → none; electrical across buildings → expected; blank group part → none;
      non-registry hop skipped → upstream expected + `skipped`; feeder in Closest Parent → note;
      remote Dependency Project ignored; loop → flagged; counts; notInRegistry lists source tags + affected rows.
- [ ] Run `node --test tests/expected-dependencies.test.mjs` → FAIL (module missing).
- [ ] Implement `auditExpectedDependencies(snapshot, references)`.
- [ ] Run → PASS. Commit.

### Task 2: Findings + session result
- [ ] Tests: `auditEngineeringFindings` emits `reference.engineering-expected-dependency` for Missing rows only;
      old `reference.engineering-relationship` gone; `auditSessionResult(...).dependencyCheck` present when
      EasyPower/Cable/PMD loaded. Rewrite the three old relationship tests + locality test to the new rule.
- [ ] Implement in `engineering-references.js`, `review.js`; update id in `actions.js:165`.
- [ ] Run engineering + review tests → PASS. Commit.

### Task 3: CSV input
- [ ] Test: worker path reads a CSV EasyPower/cable file, sheet named after the file, UTF-8 kept.
- [ ] Implement in `src/io/import-worker.js`; slots accept `.csv`; toast text.
- [ ] Run → PASS. Commit.

### Task 4: Export workbook — `src/audit/dependency-export.js`
- [ ] Test: Summary / Dependencies / Not in registry sheets, header rows, frozen header, autofilter, shading on Missing.
- [ ] Implement `buildDependencyCheckWorkbook(check, name)`.
- [ ] Run → PASS. Commit.

### Task 5: UI tab
- [ ] Test (existing UI-test harness): nav item + badge; empty state shows 3 slots; table rows + counts; status filter.
- [ ] Implement `src/ui/dependencies.js`, nav entry in `ui/audit.js`, `boot.js` route, manifest entries, CSS,
      export reusable `engineeringUploadsHtml(kinds)` / `wireEngineeringUploads(navigate, rerender)`.
- [ ] Run → PASS. Commit.

### Task 6: Docs + verification
- [ ] Guide section + README line for the Dependencies tab (no release notes, no version bump).
- [ ] `npm test` (all pass), `npm run build`, browser check with a synthetic registry + CSV EasyPower at desktop and 900px.
- [ ] Self-review diff; commit. Nothing pushed.
