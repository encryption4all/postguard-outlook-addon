# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

PostGuard end-to-end email encryption as an Office Add-in for the new Outlook on Windows + macOS. It is a taskpane-only mail add-in (`<Host Name="Mailbox">`) that runs in both Compose and Read modes; encryption/decryption uses the `@e4a/pg-js` SDK with Yivi-based identity-based encryption (IBE). There is no backend in this repo — the PKG and Cryptify services are external.

## Commands

- `npm run build` / `npm run build:dev` — production / development webpack build into `dist/`.
- `npm run watch` — webpack in watch mode.
- `npm run dev-server` — webpack-dev-server on `https://localhost:3000` with the dev cert from `office-addin-dev-certs`.
- `npm start` — `office-addin-debugging start manifest.xml`. Sideloads the manifest and launches the configured Outlook host. Use `npm stop` to unload.
- `npm run validate` — validates `manifest.xml` against the Office Add-in schema. Run after manifest edits.
- `npm run lint` / `npm run lint:fix` / `npm run prettier` — `office-addin-lint` wrappers (ESLint + Prettier with the office-addins config).
- `npm run signin` / `npm run signout` — manage the M365 dev account used by the debugging tools.

There are no automated tests in this project.

## Build-time configuration

Three URLs are baked into the bundle via webpack `DefinePlugin` (see `webpack.config.js`):

- `PKG_URL` — PostGuard Key Generation server.
- `CRYPTIFY_URL` — Cryptify file-share service.
- `POSTGUARD_WEBSITE_URL` — used by the SDK envelope for the browser fallback link.

These are read from `.env` (copy `.env.example`) or fall back to staging defaults. They are accessed through `src/lib/pkg-client.ts` — do not read `process.env` elsewhere.

The webpack config also rewrites `https://localhost:3000/` → `https://addin.postguard.eu/` inside `manifest.xml` when building in non-development mode, so the *same* manifest is used for dev sideloading and production hosting.

## Architecture

### Entry points and bundles

Webpack builds three entries:

- `polyfill` — `core-js` + `regenerator-runtime`, prepended to both HTML pages.
- `taskpane` — `src/taskpane/taskpane.ts` plus the HTML template; this is where almost all UI logic lives.
- `commands` — `src/commands/commands.ts`. Required by the manifest's `<FunctionFile>` but currently a no-op. Every user action goes through the taskpane.

### Taskpane dispatch

`src/taskpane/taskpane.ts` is the single runtime entry. After `Office.onReady`, it inspects the item via `isComposeMode()` (which probes for `subject.setAsync` because compose items have no `itemId` until first save) and routes to either `mountComposeView()` or `mountReadView()`. All views are sibling `<section>`s inside `taskpane.html`; `showView(name)` toggles `hidden` on each. There is no router and no framework — everything is plain TS + `getElementById`.

### Compose flow (`compose-view.ts`)

State (`encrypt` toggle, per-recipient `Policy`, sender `signAttributes`) is held in a module-local `state` object. The "Encrypt & Send" button:

1. Refreshes recipients from Office.js (Outlook compose has no recipient-changed event, so we re-pull on every action).
2. Builds a MIME blob with `buildMime` from `@e4a/pg-js`, including all readable attachments (cloud attachments are skipped — Office.js cannot read their bytes).
3. Switches to the Yivi view, instantiates `PostGuard`, and calls `pg.encrypt({ sign: pg.sign.yivi(...), recipients, data })` mounting the Yivi widget at `#yivi-web-form`.
4. Calls `pg.email.createEnvelope(...)` which yields `{ subject, htmlBody, attachment }`. We `setSubject` / `setBody` on the draft, remove the original plaintext attachments, and add the encrypted blob as `postguard.encrypted`.

BCC is unsupported and the UI hard-blocks Encrypt & Send when BCC is present (the PostGuard envelope cannot represent BCC because all recipients are encrypted *to* and visible to each other in the policy).

### Read flow (`read-view.ts`)

Two ciphertext sources are tried in order: a `postguard.encrypted` attachment, then an ASCII-armored block (`-----BEGIN POSTGUARD MESSAGE-----`) inside the HTML body (for forward compatibility with text-only emails). On Decrypt, `pg.open({ data }).decrypt({ element, recipient })` runs the Yivi disclosure flow and returns `{ plaintext, sender }`. Plaintext is rendered into a sandboxed iframe via `iframe.srcdoc`; sender attributes become badges.

Outlook does not allow an add-in to mutate the displayed message in read mode, so the decrypted view is *only* visible inside the taskpane. A persistent `notificationMessages` banner is added to the message itself.

### `src/lib/` boundaries

- `office-helpers.ts` — promisified wrappers around the callback-based Office.js mailbox API. Anything touching `Office.context.mailbox.item` should go through here.
- `auth.ts` + `graph-client.ts` — Graph SSO via `OfficeRuntime.auth.getAccessToken({ forMSGraphAccess: true })`. **Currently dormant**: the manifest has no `<WebApplicationInfo>` block, so `getGraphToken` will reject and Graph-dependent features (sent-copy, in-place message replacement) are skipped. Wiring requires registering an Azure AD app and adding the manifest block.
- `mime.ts` — minimal MIME helpers, deliberately not a full RFC 5322 parser.
- `storage.ts` — `Office.context.roamingSettings` (~32KB JSON budget per mailbox).
- `pkg-client.ts` — exports the build-time URLs and a `clientHeaders()` helper that stamps `X-PostGuard-Client-Version`.
- `attributes.ts`, `i18n.ts`, `types.ts`, `encoding.ts` — data shapes and small utilities. `i18n.t()` is an inline lookup (no `browser.i18n` in Office Add-ins).

### WASM loading

`@e4a/pg-js` (≥ 0.10.0) inlines the `pg-wasm` binary as a base64 string at *its* prebuild (see `postguard-js/scripts/generate-wasm-base64.mjs` and `src/util/wasm.ts`) and calls `init({ module_or_path: decodeBase64(WASM_BASE64) })` at runtime. There is no separate `index_bg.wasm` file to ship.

However, the wasm-bindgen-generated `__wbg_init` function inside the inlined shim *also* contains a dead default-value branch — `if (module_or_path === undefined) module_or_path = new URL("index_bg.wasm", import.meta.url)` — that's never taken at runtime but webpack 5 statically analyzes and tries to resolve. We work around it in `webpack.config.js` with a `parser: { url: false }` rule scoped to `node_modules/@e4a/pg-js/`. Tracked upstream at [encryption4all/postguard#153](https://github.com/encryption4all/postguard/issues/153) and [encryption4all/postguard-js#30](https://github.com/encryption4all/postguard-js/issues/30); remove the rule once those ship.

Older `pg-js` releases (≤ ~0.9.x) used the `new URL("index_bg.wasm", ...)` lookup as the *real* load path and required a postinstall hook to copy the wasm next to the bundle. That hook has been removed. If you ever pin to one of those older versions you'll need to restore it.

Webpack still has `experiments.asyncWebAssembly` + `syncWebAssembly` and a `\.wasm$` asset rule. They're harmless with the inlined-base64 SDK and would be needed again if the SDK ever switches back to URL-based loading.

## Conventions in this codebase

- TypeScript `strict: true`; `noEmitOnError: true`. Babel does the actual TS transform via `@babel/preset-typescript`.
- The `@e4a/pg-js` types are loose in places — the code uses `as never` casts at SDK boundaries deliberately. Don't try to "fix" these without verifying against the SDK source.
- `console`/global lint warnings are silenced by `office-addin-lint` defaults; keep error surfaces user-visible via `showError` / `setStatus` instead.

## Outlook Add-in quirks

`docs/outlook-quirks.md` is a running log of platform behaviors that surprised us during development — Smart Alerts / launchevent dispatch oddities, cross-runtime state-sharing issues (`customProperties` vs `internetHeaders`), CSS interactions with `[hidden]`, debugging via `--devtools`, etc. Read it before debugging anything that "should work" but doesn't, and add to it whenever you discover a new surprise.
