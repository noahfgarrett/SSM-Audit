# SSM Audit

SSM Audit is a standalone, offline-first review tool for completed Exto Cx Registry workbooks. It checks hierarchy, dependencies, metadata, milestones, Item Masters, headers, and upload readiness against the SSM SOP and approved Exto Rev21 values.

## Privacy

The selected workbook is read in browser memory on the user's device. It is not uploaded, saved into a profile, or used to learn future logic. Closing or refreshing the HTML clears the audit session.

## Development

```bash
npm test
npm run build
```

The build produces `SSM-Audit.html`, a single offline HTML file. On open or refresh, it makes one anonymous request to `noahfgarrett/SSM-Audit-Releases` to check for a newer version.

## Release Model

- Private source: `noahfgarrett/SSM-Audit`
- Public artifacts only: `noahfgarrett/SSM-Audit-Releases`
- Release assets: `SSM-Audit-vX.Y.Z.html` and `SSM-Audit-vX.Y.Z.html.gz`
