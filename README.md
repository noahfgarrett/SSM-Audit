# SSM Audit

SSM Audit is a standalone, offline-first review tool for completed Exto Cx Registry workbooks. Load a registry and it checks every equipment row against the SSM SOP, the approved Rev21 upload lists, and commissioning logic — then explains each finding in plain language, shows the registry as a tree, and exports an Excel report built for working the list down.

## What it checks

Every check has a plain-language statement and belongs to one of three sources:

- **Registry Integrity** — the registry agrees with itself and with the approved Rev21 lists: one row per tag, tags that are codes rather than typed descriptions, parents that can be found, UPNs and System Names from the approved values (letter-code UPNs such as RR/SEC/MISC included), I&C rows on the approved controls discipline, Dependency Project only when a dependency really is external, and a note on classifications not in the Rev21 dropdown so site-specific codes can be reviewed.
- **SSM SOP** — how equipment nests and depends: a child stays in its parent's UPN, discipline, and building; anything crossing them is a dependency, not a parent; instruments nest under equipment in the UPN their tag carries; a VFD sits under what it drives with its panel and PLC as dependencies; FMS hardwired I/O sits under its VFD; LCPs sit with their skid; control valves and room sensors sit under the equipment they serve; organizational headers carry a Blank Item Master, no dependencies, and are never themselves a dependency; every system has a row at its top; a parent duplicated as a dependency is noted (routine in electrical, worth a look elsewhere); L2 milestones name their UPN; Item Masters use the VF standard (site-prefixed names are flagged with the VF equivalent proposed).
- **Commissioning Logic** — expected relationships: driven and control equipment trace back to power, RIOs name their controller, heat trace follows its transformer, VESDA depends on its fire alarm panel, drives sit under the equipment they run.

Findings carry one of four levels. They describe what *kind* of problem it is — not how the registry got here. Most audited registries are already uploaded, so the levels are about the data itself:

| Level | Meaning |
|---|---|
| **Invalid** | The row contradicts the registry or the approved lists — a duplicate tag, a blank or self-referencing parent, a parent that does not exist, a loop, or a UPN / System Name / discipline that is not in the Rev21 upload template. On a fresh upload these are the rows Exto would reject; on an uploaded registry they are the rows that cannot be right. |
| **Rule broken** | Valid data that breaks an SSM SOP rule — a parent in another UPN, a dependency that cannot be found, an instrument or LCP nested off its UPN, a VESDA without its fire alarm panel. |
| **Check this** | A strong pattern says something is off; an engineer should look. |
| **Note** | Worth knowing, not necessarily wrong — a same-system dependency in a mechanical discipline, a milestone not yet assigned, an Item Master still on a site prefix. |

The full rulebook is in the app under **Rules**, with live counts once a registry is loaded.

## The workspace

- **Dashboard** — the overview an audit opens on: level totals, milestone readiness per L2 phase, hierarchy health, dependencies, a checks overview that also shows which checks passed, and structure stats. Every number opens the findings narrowed to it, and the whole dashboard scopes to whatever filters are active.
- **Audit findings** — every issue, most serious first. Level chips, one **Filters** panel for discipline / L2 milestone / UPN / building / source / topic / individual checks (only the active filters show on the page, as chips), grouping by check or phase, and search across tags and reasons. Open any row for *Why this was flagged → What must be true → the relationship diagram (for parent, dependency, and loop findings) → Found / Expected → What to do → the registry row*, with **Show in hierarchy** and Previous/Next.
- **SSM hierarchy** — Building → Discipline → System → equipment, exactly as uploaded, with finding counts on every branch, a **Findings only** view, header rows marked, and search that opens the matching branches.
- **Compare projects** — line a registry up against a completed project by UPN.
- **Rules** — every check in plain language, grouped by source and topic, each with an **Example** button: a small made-up registry showing exactly what gets flagged (cell shaded, offending characters marked), why, what the audit says — run live on that mock data — and what right looks like.

## The Excel report

**Export** builds a workbook made for actioning:

- **Dashboard** — one large overall progress bar, the level totals, and a wide progress bar with a live percentage per discipline and per L2 phase.
- **Index** — one row per phase, hyperlinked to its tab; every tab links back.
- **One tab per L2 phase** — every piece of equipment in that phase as an indented tree, each nest level in its own colour with the ID always in black, Closest Parent and Dependencies side by side, one line per finding with the reason and what to do, and the cell the finding is about shaded light red. Tick the **Actioned** box (☐ → ☑) on an equipment's first line: the row turns green and the Index and Dashboard update.
- **All Findings** and **Rules**.

## Privacy

The workbook is read in browser memory on the user's device. It is not uploaded, saved into a profile, or used to learn future logic. Closing or refreshing the HTML clears the audit session. The only network request is one anonymous check for a newer release when the page opens.

## Development

```bash
npm test
npm run build
```

The build produces `SSM-Audit.html`, a single offline HTML file. The audit engine (`src/audit/model.js`, `src/audit/engine.js`, `src/exto/rev21-contract.js`, `src/exto/vf-item-masters.js`) is shared byte-for-byte with SSManagement's embedded audit; `tests/parity.test.mjs` enforces it. Change a rule in one place, mirror it to the other, and run both suites.

## Release Model

- Private source: `noahfgarrett/SSM-Audit`
- Public artifacts only: `noahfgarrett/SSM-Audit-Releases`
- Release assets: `SSM-Audit-vX.Y.Z.html` and `SSM-Audit-vX.Y.Z.html.gz`
