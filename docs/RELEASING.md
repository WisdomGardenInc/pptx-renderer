# Release Process

## Pre-Release Checklist

- [ ] Update changelog/release notes
- [ ] Verify tests pass: `pnpm test`
- [ ] Verify build passes: `pnpm build`
- [ ] Verify package entries and real browser paths: `pnpm test:package && pnpm test:browser`
- [ ] Verify PDF.js contracts: `pnpm test:pdfjs-contract` and the CI PDF.js 6 matrix job
- [ ] Verify lint/types/exports: `pnpm lint && pnpm typecheck && pnpm knip && pnpm publint`
- [ ] Verify bundle size: `pnpm size`
- [ ] Confirm `README` and docs are up to date
- [ ] Confirm security-impacting changes are documented

## Versioning

Use semantic versioning:

- `MAJOR`: breaking API changes
- `MINOR`: backward-compatible features
- `PATCH`: backward-compatible fixes

## Suggested Steps

1. Create release branch or prepare release commit on `main`.
2. Bump version in `package.json`.
3. Run full verification (`pnpm test`, `pnpm build`, `pnpm test:package`, `pnpm test:browser`).
4. Tag release (`vX.Y.Z`) — pushing the tag publishes to npm.
5. Publish release notes with migration notes if needed.

## Command Scenario (GitHub + npm)

Use this sequence for an actual release:

1. Verify and build:
   - `pnpm test`
   - `pnpm build`
   - `pnpm test:package`
   - `pnpm test:browser`
   - `pnpm size`
2. Commit release metadata:
   - `git add -A`
   - `git commit -m "chore: release vX.Y.Z"`
3. (Optional) Flatten history to one initial commit:
   - `/bin/zsh -lc 'NEW_ROOT=$(git commit-tree HEAD^{tree} -m "First commit") && git reset --hard "$NEW_ROOT"'`
4. Create annotated tag:
   - `git tag -a vX.Y.Z -m "vX.Y.Z"`
5. Push branch and tag:
   - `git push origin main`
   - `git push origin vX.Y.Z`
   - Pushing the tag starts the `Release` workflow, which re-runs lint, types,
     unit tests, build, package and size checks and then publishes to npm.
     Nothing else is needed; do not publish by hand as well.
6. Create GitHub Release:
   - `gh release create vX.Y.Z --title "vX.Y.Z" --notes-file CHANGELOG.md`

## npm Trusted Publishing

`.github/workflows/release.yml` publishes with trusted publishing (OIDC), so no
npm token is stored in the repository. npm exchanges the workflow's short-lived
GitHub identity for a publish credential and attests provenance from that run.

One-time setup on npmjs.com, under the package's Settings -> Trusted publishing:

| Field             | Value                             |
| ----------------- | --------------------------------- |
| Organization/user | `WisdomGardenInc`                 |
| Repository        | `pptx-renderer`                   |
| Workflow filename | `release.yml`                     |
| Environment       | (leave empty unless one is added) |

Constraints worth knowing:

- A trusted publisher can only be attached to a package that already exists, so
  the very first version of a new package name has to be published manually with
  `npm login && npm publish --access public`.
- Renaming the workflow file, the repository, or the npm scope breaks the trust
  relationship until the publisher entry is updated to match.
- `package.json`'s `repository` field must point at the same repository as the
  publisher entry.
- The workflow pins Node 24 and upgrades npm, because trusted publishing needs
  npm >= 11.5.1 and Node >= 22.14.

## Release Notes Template

- Summary
- Added
- Changed
- Fixed
- Security
- Migration Notes
