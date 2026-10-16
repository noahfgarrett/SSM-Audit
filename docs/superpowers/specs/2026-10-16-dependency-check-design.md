# Dependency Check — design

Date: 2026-10-16 · Status: approved · Scope: SSM Audit only (no SSManagement change, no release)

## Goal

Upload EasyPower, Cable Schedule and/or PMD exports (Excel **or CSV**), derive the dependencies those
documents imply using the same boundary rule SSManagement uses when it builds a registry, and compare
them tag-by-tag with the uploaded registry. Make the result easy to view and to export.

## User-facing behaviour

### New sidebar tab: **Dependencies**
- Sits between *SSM Hierarchy* and *Compare Projects*. Enabled once an audit result exists.
- Badge: number of equipment with a missing expected dependency.
- If no EasyPower / Cable Schedule / PMD is loaded, the tab shows those three upload slots
  (same slots as the upload screen; `.xlsx`, `.xls`, `.csv`). Uploading re-runs the check in place.

### Three statuses
| Status | Meaning |
|---|---|
| **Match** | Every dependency the documents imply is in the row's Dependencies. |
| **Missing** | At least one expected dependency is absent. Note "listed as Closest Parent" when the feeder sits there instead. |
| **Not in registry** | A tag from the documents that the registry does not contain. Rows downstream of it are compared against the nearest upstream tag that *is* in the registry and carry a warning listing the skipped tag(s). |

Only *dependencies* are reported. A feed that stays inside one group (see rule) is an expected parent and
is not listed. Registry dependencies that no document supports are not reported.

### Screen
- Count strip: `Compared · Match · Missing · Not in registry`; each count filters the table.
  *Compared* = registry rows with at least one expected dependency (Match + Missing).
- Search (tag) + Discipline + Building filters.
- One table, one row per item: **Equipment ID · Status · Expected dependencies · Registry dependencies ·
  Missing · Note · Source**. Rows expand to show the source chain
  (e.g. `EasyPower Sheet1 row 12: TEST-SWBD → TEST-MCC → TEST-PMP`).
- **Download dependency check** button → one Excel workbook (below).

### Export workbook
1. **Summary** — counts by status, and by discipline × status.
2. **Dependencies** — the table above; frozen header, auto-filter, Missing / Not-in-registry cells shaded.
3. **Not in registry** — source tag · document · sheet · row · registry rows affected.

### Findings
*Missing* rows become audit findings under a new rule **Expected dependency missing**
(`reference.engineering-expected-dependency`, category *dependencies*, severity *warning*), replacing the
old **Source connection not represented in registry** rule (`reference.engineering-relationship`), which is
removed along with its references. *Not in registry* is already covered by the existing
**Source tag not found in registry** rule. The main audit, its Excel report and Actions are otherwise unchanged.

## The rule (mirrors SSManagement's partition fold, `compiler/fold.js`)

1. **Edges** — each source connection becomes *child ← feeder*:
   EasyPower: consecutive chain members (Starting Source → Downstream n → Final Source → ID Name);
   Cable Schedule: Panel (From) → Load (To); PMD: Panel → Instrument Tag.
2. **Group** of a registry tag = normalized `Building | Discipline | System Name` from the registry.
   If any part is blank the group is *unknown* and nothing is expected (a crossing must be proven).
3. For each registry tag that is a child in any edge, walk upstream (breadth-first over all feeders) until
   reaching tags that are in the registry, recording every skipped (non-registry) tag. Depth cap 25;
   a revisited tag stops the walk and marks the row *loop in source chain*.
4. A resolved upstream tag **U** is an *expected dependency* of **T** when both groups are known and differ,
   **except** when both are ELECTRICAL and share a Building (the electrical feed chain nests).
5. Compare with the row's local Dependencies (dependencies in another Dependency Project are ignored, as
   today). Missing = expected − present. Matching uses `auditNormId` (exact, case/space-normalized).

## Architecture

| Unit | Responsibility |
|---|---|
| `src/audit/expected-dependencies.js` (new, pure) | `auditExpectedDependencies(snapshot, references)` → `{rows, notInRegistry, counts}`. No DOM, no state. |
| `src/audit/engineering-references.js` | Remove the relationship branch/rule; add the expected-dependency rule; `auditEngineeringFindings(snapshot, references, check)` emits findings from `check.rows` with status `missing`. |
| `src/audit/review.js` | `auditSessionResult` computes the check once, passes it to the findings, attaches it as `result.dependencyCheck`. Runs in the existing review worker. |
| `src/io/import-worker.js` | CSV: read with UTF-8 codepage; name the single sheet after the file. |
| `src/ui/audit.js` | Slots accept `.csv`; export `engineeringUploadsHtml(kinds)` / `wireEngineeringUploads(navigate, rerender)`; nav entry; actions policy id update. |
| `src/ui/dependencies.js` (new) | `renderDependencyCheck(navigate)` — counts, filters, table, expandable chains, empty state with slots. |
| `src/audit/dependency-export.js` (new) | `buildDependencyCheckWorkbook(check, name)`; download via existing compact writer. |
| `src/boot.js`, `build/manifest.mjs` | Route `dependencies` screen; register new modules. |

The four parity-mirrored engine files (`model.js`, `engine.js`, `rev21-contract.js`, `vf-item-masters.js`) are not changed.

## Error handling
- No recognizable headers → existing "check the document type and headers" message (CSV included).
- Empty/unreadable CSV → "The file has no rows" / read error toast.
- Loops → walk stops; row flagged.
- No audit result → tab disabled.

## Testing (synthetic `TEST-*` fixtures only)
- Rule: cross-group → expected; same group → none; electrical same building → none; electrical across
  buildings → expected; unknown group → none; skipped hop → warning + upstream used; listed-as-parent note;
  other-project dependency ignored; loop.
- CSV read through the worker path; sheet named after the file.
- Findings: new rule replaces old; counts.
- Export: three sheets, headers, frozen row, filter.
- UI: tab renders counts/table/empty state; nav badge.
- Full `npm test`, `npm build`, visual check in the browser with a synthetic registry.
