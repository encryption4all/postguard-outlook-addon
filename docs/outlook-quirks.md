# Outlook Add-in quirks and gotchas

A running log of behaviors in the Outlook Add-in / Office.js platform that surprised
us — what we expected vs. what actually happens, and how we worked around it.
Bias: new Outlook on Windows (`platform === "OfficeOnline"`, OWA-in-WebView2),
Mailbox 1.12+. Add new entries as you find them.

---

## Manifest

### `<Version>` must be ≥ 1.0

`office-addin-manifest validate` rejects manifests whose top-level `<Version>` is
under `1.0` with a misleading "Manifest Version Too Low" error (it sounds like it's
talking about the schema version). Bump to `1.0.0.0` even for pre-release add-ins.

### `VersionOverridesV1_1` fully replaces `V1_0` for clients that support it

Nesting a `V1_1` block inside a `V1_0` block is the documented pattern, but the inner
block is **not additive**. Any extension point that needs to exist in both worlds must
be declared in both. Each level also has its own `<Resources>` block (no inheritance).
The current manifest duplicates the read/compose taskpane points across both levels
on purpose.

### Validator error messages are unreliable

`office-addin-manifest validate` prints `"Mailbox add-in not containing ItemSend
event is valid"` regardless of whether you actually have an `OnMessageSend`
LaunchEvent. Don't read it as confirmation that your event is wired up; only
runtime testing tells you that.

---

## Event-based activation (`OnMessageSend` / Smart Alerts)

### Outlook auto-probes `launchevent.html?et=…` even without a manifest declaration

If new Outlook decides your add-in might have a launch event handler, it will
fetch `<source-location>/launchevent.html?et=<event>` on Send, **regardless of what
the manifest says**. If the file 404s or the JS doesn't call `event.completed()`,
Outlook hangs ~15 s and shows a "PostGuard timed out / Don't send" dialog that's
indistinguishable from a real Smart Alerts block.

Fix: always ship a `launchevent.html` that loads, registers a handler, and is fast
to call `event.completed()`. We have a 12-second client-side fallback timeout in
`launchevent.ts` so the user is never stuck.

### `SendMode="SoftBlock"` is what makes Send actually wait for the handler

Without it, Outlook fires the event but doesn't block the send while the handler
runs. Smart Alerts UX (the dialog with the "Take Action" button) requires
`<LaunchEvent Type="OnMessageSend" SendMode="SoftBlock" />`.

### Stable `office.js` requires `Office.onReady()` to dispatch events

The Microsoft sample (`outlook-encrypt-decrypt-messages`) just calls
`Office.actions.associate(...)` at the top of the file with no `Office.onReady`
anywhere. That works because the sample loads `lib/beta/hosted/office.js`. The
stable `lib/1/hosted/office.js` throws "Office.js has not fully loaded. Your app
must call Office.onReady()" when the framework tries to dispatch an event and
no `onReady` was ever registered. Wrap your `associate` call in an `Office.onReady`
callback.

### `Office.actions.associate` registration must complete before the event fires

The launchevent runtime is a fresh WebView2 launched on demand. If anything in
the launchevent script load is slow (heavy polyfill chunks, dynamic imports)
the registration may not be in place by the time Outlook dispatches. Drop the
`polyfill` chunk from `launchevent.html` — WebView2 doesn't need core-js.

### `commandId` on `event.completed({ allowEvent: false })` opens a manifest button

Pass the id of one of your `<Control xsi:type="Button">` definitions and Outlook
adds a "Take Action" button to the Smart Alerts dialog that runs that command —
in our case, opens the PostGuard taskpane. No need to wire any "open this
taskpane" plumbing yourself.

---

## Cross-runtime state: launchevent ↔ taskpane

### The launchevent runtime is a *separate* WebView2 from the taskpane

It has its own JS context, its own `window`, its own DevTools target. Anything
the taskpane writes to module-local state, `Office.context.roamingSettings`
in-memory, etc., is invisible to the launchevent handler. Use APIs that
persist on the message itself.

### `customProperties` does not propagate cross-runtime in new Outlook

`item.loadCustomPropertiesAsync` works in both runtimes, but a property written
from the taskpane (with all the right `saveAsync` calls) reads back as missing
in the OnMessageSend handler. We never definitively diagnosed why — possibly the
local-state write doesn't make it into the server save that the handler later
reads against. Either way: don't trust customProperties for cross-runtime
state in new Outlook on Windows.

### `internetHeaders` is what does work — but names must start with `x-`

`item.internetHeaders.setAsync({ "x-pg-encrypt-on-send": "true" })` in compose
mode is reliably visible to `item.internetHeaders.getAsync(...)` in the
OnMessageSend handler. Note the `x-` prefix is enforced by Office.js for
custom headers.

### `sessionData` is compose-only — *not* visible in launch events

Despite docs implying it's a general per-session state mechanism, `sessionData`
methods are not available in launch event handlers. Don't reach for it for
taskpane↔send-handler communication.

### `displayDialogAsync` from the launchevent runtime needs the add-in domain in `<AppDomains>`

A dialog opened from the regular taskpane is allowed without an `<AppDomains>`
entry as long as it's same-origin with `<SourceLocation>`. From the *launchevent*
runtime that rule does not apply: the runtime is hosted inside an
`outlook.office.com` iframe, so from its point of view the dialog URL is
cross-origin and Office checks the manifest's `<AppDomains>`. Symptoms:

- **Outlook on the web** rejects with `code=12011` (`BlockedNavigation`) and a
  message about "different security zones".
- **New Outlook for Mac** rejects with the generic E_FAIL (HRESULT
  `-2147467259`, message "An internal error has occurred."). Same root cause,
  uglier wrapper.

Fix: list the add-in's own host in `<AppDomains>` (`https://addin.postguard.eu`,
`https://addin.staging.postguard.eu`, and `https://localhost:3000` for the dev
sideload). It's only the launchevent dialog path that needs this; the taskpane
keeps working without the explicit entry.

### `displayDialogAsync` from the launchevent runtime is broken on Outlook for Mac

On new Outlook for Mac the same call rejects with `code=-2147467259`
("An internal error has occurred.") regardless of `<AppDomains>`, dialog
options (`displayInIframe`, `promptBeforeOpen`, omitted), dialog size (we
went from 9% up to 40%), runtime declaration (with and without
`<Override type="javascript">`), and Office.js channel (`/lib/1/` vs
`/lib/beta/`). Mailbox 1.13 is supported. The same manifest works on
Outlook on the web and new Outlook on Windows.

We filed [OfficeDev/office-js#6677](https://github.com/OfficeDev/office-js/issues/6677);
related stale reports are #3138, #3085, and #5681.

Workaround: detect `Office.context.platform === Office.PlatformType.Mac`
in `onMessageSendHandler` and `block()` the OnSend with a Smart Alert
pointing the user at the taskpane "Encrypt & Send" button, which uses
the same dialog API but from the taskpane runtime where it works. The
branch is in `src/launchevent/launchevent.ts`; remove it once #6677
ships a fix.

### `window.location.href` in the launchevent runtime is *not* the add-in origin on every host

On Outlook on Web and new Outlook on Windows, the launchevent runtime loads
`launchevent.html` from your `<bt:Url id="WebViewRuntime.Url">`, so
`window.location.href` is the add-in origin and you can derive other URLs from
it. New Outlook for Mac instead uses the `<Override type="javascript"
resid="JSRuntime.Url"/>` branch and runs `launchevent.js` directly — there
`window.location` resolves to an Office-internal URL, not the add-in. Passing
that to `displayDialogAsync` fails with the unhelpful `An internal error has
occurred.` Use a build-time-injected `process.env.ADDIN_PUBLIC_URL` (see
`webpack.config.js` DefinePlugin) for any absolute URL the launchevent
handler hands back to Office.

---

## Office.js compose-mode quirks

### Compose drafts have no `itemId` until first save

Probing `subject.setAsync` is the canonical way to detect compose mode (the
property only exists in compose). `itemId` is undefined for new drafts, so
anything that needs a server-side identity (custom-property persistence,
internet-header server flush) requires a `saveAsync` first.

### `addFileAttachmentFromBase64Async` returns when the attachment is *queued*, not committed

The callback fires when Office has accepted the attachment locally, before the
server upload completes. If the user clicks Send right away, Outlook can race
the upload and surface that as a Smart-Alerts-style "PostGuard timed out"
dialog. Call `item.saveAsync` after attaching to force a server flush.

### Internet header writes need a *trailing* `saveAsync` to flush

Same pattern as attachments. Our `persistEncryptOnSend` does
`saveAsync → internetHeaders.setAsync → saveAsync`. The first save guarantees
an `itemId` so the header has somewhere to attach; the second pushes the
header change to the server so the OnMessageSend handler reads it.

### `RecipientsChanged` *does* exist (Mailbox 1.7+)

The original code had a comment claiming Outlook has no recipient-changed
event — that was true for old Outlook but obsolete since Mailbox 1.7. Subscribe
with `item.addHandlerAsync(Office.EventType.RecipientsChanged, …)`.

### Cloud attachments can't be read via Office.js

`attachmentType === Cloud` items don't return bytes from
`getAttachmentContentAsync`. Skip them with a defensive check; otherwise the
async result is a no-op or a failure depending on Outlook version.

### Tenant DLP can scrub attachment bytes while keeping metadata

For attachments whose extension violates a tenant DLP policy (e.g. `.exe`
on a corporate M365 tenant), `item.attachments` lists the file with its
real `name` and `size` but `getAttachmentContentAsync` returns
`format=base64` with `content.length === 0`. The bytes were scrubbed
client-side before our add-in could read them. There's no client-side
recovery — the bytes literally don't exist in our context.

Detect this with a `size > 0 && content.length === 0` check and fail
loudly instead of encrypting nothing. The user typically has to zip /
rename the offending file to get it past the policy.

### There is no `sendAsync` for compose

Office.js can save a draft programmatically but cannot send it. This is the
specific reason the one-click encrypt-on-Send flow is hard: the OnMessageSend
handler can encrypt and `event.completed({ allowEvent: true })` to release the
send Outlook had pending, but it can't *initiate* a send from a "regular"
taskpane button. (See option 1 vs option 2 in our design notes.)

### Body size affects Send latency

`body.setAsync` resolves locally, but if the body is multi-MB (e.g., an inline
base64 ciphertext block) the server-side commit can take long enough to look
like a hang. Either trim the body before setAsync, or call `saveAsync` after
to surface the wait inside the taskpane's "Saving…" UI rather than at Send.

---

## Display & rendering

### Author CSS overrides the `[hidden]` attribute

A rule like `.pg-view { display: block }` wins over the user-agent
`[hidden] { display: none }` rule. Setting `el.hidden = true` becomes a no-op
visually. Either scope the display rule with `:not([hidden])` or add an
explicit `.pg-view[hidden] { display: none }`. We hit this hard during initial
bring-up — the loading spinner stayed forever even though `showView()` ran.

### New Outlook draws its own taskpane header

The host shows a header bar with the add-in name + icon, sourced from the
manifest's `DisplayName` and `IconUrl`. If you also draw your own header in
the taskpane HTML, you get a double header. Don't draw your own.

---

## Platform & runtime

### New Outlook on Windows reports `platform === "OfficeOnline"`

It's WebView2 wrapping OWA. Don't gate desktop-specific code on
`platform === Windows` — you'll exclude new Outlook.

### Office Add-ins do not have `browser.i18n`

There's no built-in localization API. Ship strings inline and look them up
yourself. We use `i18n.t()` which falls back to English.

---

## Build & tooling

### `@microsoft/app-manifest@1.0.5` is CJS but requires ESM-only `strip-bom@^5`

`office-addin-debugging` blows up at load with `ERR_REQUIRE_ESM` after a fresh
`npm install`. Pin the dep with an `overrides` block in `package.json`:

```json
"overrides": { "strip-bom": "^4.0.0" }
```

### `@e4a/pg-wasm` ships a dead `new URL("index_bg.wasm", import.meta.url)` branch

Wasm-bindgen-generated init has a default-value branch that's never reachable
when callers always pass `module_or_path` (which `pg-js` does). Webpack 5
statically analyzes the `new URL(...)` and tries to resolve the file at build
time, failing because pg-js inlines the wasm as base64 and ships no separate
wasm file. Workaround: `parser: { url: false }` scoped to the pg-js module.
Tracked at [encryption4all/postguard#153](https://github.com/encryption4all/postguard/issues/153)
and [encryption4all/postguard-js#30](https://github.com/encryption4all/postguard-js/issues/30).

### Webpack-dev-server HMR can't apply entry config changes

Adding/removing a webpack `entry` (e.g., when we added `launchevent`)
invalidates HMR. The console shows `[HMR] Cannot apply update. Need to do a
full reload!`. Restart the dev server (`Ctrl+C`, `npm run dev-server`) when
touching `entry` or `plugins`.

### `office-addin-debugging` aggressively caches the sideloaded manifest

After manifest changes, `npm stop && npm start` to re-sideload. New Outlook
itself may also need a full close (kill any lingering `olk.exe` in Task Manager)
to re-fetch the manifest.

---

## Debugging

### `olk.exe --devtools` is required to right-click → Inspect

New Outlook only exposes the "Inspect" context-menu item when launched with
`--devtools`. The exe lives at
`C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_<version>_x64__8wekyb3d8bbwe\olk.exe`;
launch via `Start-Process -ArgumentList "--devtools"` or a shortcut.
[Reference](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook#debug-your-add-in).

### The launchevent runtime needs its own DevTools target

The right-click→Inspect on the taskpane only attaches to *that* WebView. To see
the OnMessageSend handler's console, set
`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=9222"` before
launching Outlook, then attach via `edge://inspect/#devices` with `localhost:9222`
configured. The launchevent target appears briefly on Send.

### `?et=…` in the URL is the dispatched event type

When Outlook fetches `launchevent.html?et=<something>` you can read the value
to know which event was dispatched without the runtime having to associate
multiple handlers. Useful for debugging when a dispatch isn't reaching the
handler you expect.

### Personal photo / `pg_logo.png` 404s are OWA noise

The host issues a stream of `imageB2/.../resize` and `pg_logo.png` requests
during compose / read rendering. They're CSP-blocked or 404, none of them are
from our add-in, and they're safe to ignore when triaging real issues in the
console.

### `XML-parsefout: geen hoofdelement gevonden` on `taskpane.html?et=` is OWA noise

When an add-in declares an `OnMessageSend` LaunchEvent, OWA probes the taskpane
URL with an empty `?et=` query string (`https://localhost:3000/taskpane.html?et=`)
and runs the response through an XML parser. HTML isn't valid XML so Firefox
logs `XML-parsefout: geen hoofdelement gevonden` ("XML parse error: no root
element found") to the console. The parse failure is internal to OWA's probe
logic — the actual taskpane mount uses the response normally and your `Office.onReady`
code runs fine after this error. Safe to ignore.

---

## Things we still don't fully understand

- Why customProperties propagated for the MS encrypt-attachments sample but not
  for us. Possibly Mailbox version, possibly a subtle save-ordering difference,
  possibly tenant config.
- Whether tenant-level Smart Alerts policy (e.g., on `caesar.nl`) is forcing
  Outlook to probe `launchevent.html` even when our manifest doesn't declare a
  LaunchEvent. The current behavior is consistent with that hypothesis but we
  haven't confirmed.
- The exact relationship between "draft saved with `saveAsync`" and "header
  visible to OnMessageSend". The double-`saveAsync` in `persistEncryptOnSend`
  works empirically; whether both are strictly needed is unclear.
