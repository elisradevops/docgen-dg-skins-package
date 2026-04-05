# Skins for Docgen JSON

Can be outputed in various types:

- JSON - default
- HTML

## Additional skins

- `time-machine-report` - builds the Historical Query Compare report structure used by DocGen (summary table, compare table, and per-work-item differences).

## Release pipeline notes

- GitHub Actions publish workflow runs on `main` pushes.
- `npm publish` runs before version bump/tag.
- If current `package.json` version already exists on npm, publish and bump are skipped.
- On failed historical runs, prefer triggering a new run from latest `main` instead of rerunning an old workflow execution.
