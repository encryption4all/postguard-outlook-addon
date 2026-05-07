# <p align="center"><img src="./img/pg_logo.svg" height="128px" alt="PostGuard" /></p>

> For full documentation, visit [docs.postguard.eu](https://docs.postguard.eu/repos/postguard-outlook-addon).

Identity-based email encryption add-in for Microsoft Outlook. Users can send and receive encrypted email using [Yivi](https://yivi.app) identity verification, without needing to exchange keys. This is one of the main end-user clients for PostGuard, alongside the Thunderbird add-on.

Targets the new Outlook on Windows and Outlook for Mac as a taskpane mail add-in (Compose + Read).

## Development

Requires Node.js 20 or later.

```bash
npm install
npm run dev-server     # https://localhost:3000 with the dev cert
npm start              # sideload manifest.xml into Outlook
```

Build, validate and lint:

```bash
npm run build          # production webpack bundle into dist/
npm run validate       # check manifest.xml against the Office Add-in schema
npm run lint           # ESLint (flat config) + Prettier
```

CI on every PR runs lint (`--max-warnings=0`), `tsc --noEmit`, the production build, and `npm run validate`.

## Releasing

Releases are automated via [release-please](https://github.com/googleapis/release-please) using [Conventional Commits](https://www.conventionalcommits.org/). The flow:

1. Merge PRs to `master` with conventional commit messages (`feat:`, `fix:`, …).
2. The `Release` workflow runs and release-please opens — or updates — a release PR titled `chore(main): release X.Y.Z`.
3. Merging that release PR creates the `vX.Y.Z` tag, builds and pushes the production Docker image to `ghcr.io/encryption4all/postguard-outlook-addon:X.Y.Z` (and `:latest`), and uploads `dist/manifest.xml` to the GitHub Release as a sideloadable asset.

Non-release pushes to `master` build a staging image tagged `:edge` (and `:sha-<commit>`) hosted at `addin.staging.postguard.eu`.

If you want to cut a release whose commits are all `chore:` (release-please skips those by default for `0.x` versions), push a commit to master with a `Release-As: X.Y.Z` footer to force the next release.

## License

MIT
