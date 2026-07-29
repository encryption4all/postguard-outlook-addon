# Changelog

## [0.5.0](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.4.0...v0.5.0) (2026-07-29)


### Features

* **metrics:** tag cryptify uploads with X-Cryptify-Source: outlook ([#96](https://github.com/encryption4all/postguard-outlook-addon/issues/96)) ([5629fa5](https://github.com/encryption4all/postguard-outlook-addon/commit/5629fa584270d123c86f8ea2fe3c4a672bd00901))
* show decrypted email in a popup dialog instead of the taskpane ([#110](https://github.com/encryption4all/postguard-outlook-addon/issues/110)) ([588b9ae](https://github.com/encryption4all/postguard-outlook-addon/commit/588b9ae942a6154ffbf1ab5c3e129ee2bd415650))


### Bug Fixes

* **compose:** enforce values on required decryption attributes ([#57](https://github.com/encryption4all/postguard-outlook-addon/issues/57)) ([#92](https://github.com/encryption4all/postguard-outlook-addon/issues/92)) ([928d925](https://github.com/encryption4all/postguard-outlook-addon/commit/928d92547fc0e5a57551933bdabd71016b77c796))
* dialog client-version header uses pg4ol and live package version ([#105](https://github.com/encryption4all/postguard-outlook-addon/issues/105)) ([5300dc4](https://github.com/encryption4all/postguard-outlook-addon/commit/5300dc423107d55b4c2d0a0dc74b54d177f9ca5e)), closes [#103](https://github.com/encryption4all/postguard-outlook-addon/issues/103)
* escape single quotes in Graph OData $filter (OData injection) ([#119](https://github.com/encryption4all/postguard-outlook-addon/issues/119)) ([355a960](https://github.com/encryption4all/postguard-outlook-addon/commit/355a960b4d6aae924f3bae2a201e8d8967021d64))
* point Cryptify at storage.postguard.eu, not the dead fileshare host ([#133](https://github.com/encryption4all/postguard-outlook-addon/issues/133)) ([2e9acc4](https://github.com/encryption4all/postguard-outlook-addon/commit/2e9acc4ac1b83b0623334f361bf407e91d8eb54c)), closes [#132](https://github.com/encryption4all/postguard-outlook-addon/issues/132)
* **read-view:** ensure decrypted body meets WCAG 2.2 AA contrast ([#58](https://github.com/encryption4all/postguard-outlook-addon/issues/58)) ([#109](https://github.com/encryption4all/postguard-outlook-addon/issues/109)) ([1f7ea91](https://github.com/encryption4all/postguard-outlook-addon/commit/1f7ea91e80793c553a1db0d2aed2e6c2269a069a))
* **security:** stop leaking stack traces into user-facing UI ([#113](https://github.com/encryption4all/postguard-outlook-addon/issues/113)) ([#120](https://github.com/encryption4all/postguard-outlook-addon/issues/120)) ([dac4e9f](https://github.com/encryption4all/postguard-outlook-addon/commit/dac4e9fe29b9d13725104b1a5e2d3888ee6596bf))
* **settings:** gate prefill PII console.log behind non-production check ([#118](https://github.com/encryption4all/postguard-outlook-addon/issues/118)) ([7f47df0](https://github.com/encryption4all/postguard-outlook-addon/commit/7f47df0f929c1efe30dec345707b9fc479634187))
* sync From-address selector to encrypt flow ([#90](https://github.com/encryption4all/postguard-outlook-addon/issues/90)) ([c4ca35c](https://github.com/encryption4all/postguard-outlook-addon/commit/c4ca35cc878d5a1ddaaee80c7f0372b1142490fc))
* **yivi-dialog:** guard encryption callbacks against silent/leaked throws ([#108](https://github.com/encryption4all/postguard-outlook-addon/issues/108)) ([d413cf5](https://github.com/encryption4all/postguard-outlook-addon/commit/d413cf5077a4b4d40f48264c850b4a3a5bb0870d)), closes [#78](https://github.com/encryption4all/postguard-outlook-addon/issues/78)

## [0.4.0](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.3.0...v0.4.0) (2026-05-16)


### Features

* persist upload recoveryToken via pg-js onUploadInit ([c6baad5](https://github.com/encryption4all/postguard-outlook-addon/commit/c6baad5ecf9a3d2fe4d0d84947ea89a400a8c2e9))
* persist upload recoveryToken via pg-js onUploadInit ([04af0a6](https://github.com/encryption4all/postguard-outlook-addon/commit/04af0a67a4dec200156c4e2fec1bfd77fbdf8a18))
* surface UploadSessionExpiredError distinctly, bump pg-js to 1.7.0 ([9c0b01b](https://github.com/encryption4all/postguard-outlook-addon/commit/9c0b01b80944c6f36b331dba46dad5417a0b4289))
* surface UploadSessionExpiredError distinctly, bump pg-js to 1.7.0 ([a0a51f2](https://github.com/encryption4all/postguard-outlook-addon/commit/a0a51f2b58810fe943177c8edff61699830eea01)), closes [#82](https://github.com/encryption4all/postguard-outlook-addon/issues/82)


### Bug Fixes

* localise upload-session-expired Smart Alert text ([e3dae43](https://github.com/encryption4all/postguard-outlook-addon/commit/e3dae432ed09b94d8dc0e3ee689c2e6d34b5056b))

## [0.3.0](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.2.0...v0.3.0) (2026-05-11)


### Features

* **compose:** default encryption off and add global toggle + status banner ([2685299](https://github.com/encryption4all/postguard-outlook-addon/commit/2685299aba05520a69d507fd96e9274de847268c))
* **compose:** per-draft x-pg-encrypt-on-send is the send-time authority ([79a93f6](https://github.com/encryption4all/postguard-outlook-addon/commit/79a93f666744442129e20cd135438b51a37a7470))
* **launchevent:** set status banner on compose open via OnNewMessageCompose ([b640ab8](https://github.com/encryption4all/postguard-outlook-addon/commit/b640ab88b6aa0e4dacdfb4e3e8b28d6fefc24bd5))


### Bug Fixes

* **compose:** force-repaint status banner via remove + replace ([f6a1e08](https://github.com/encryption4all/postguard-outlook-addon/commit/f6a1e082b8a2140018a2973f1d126dc7dcbf605a))
* **launchevent:** address dobby review on PR [#67](https://github.com/encryption4all/postguard-outlook-addon/issues/67) ([f0d968d](https://github.com/encryption4all/postguard-outlook-addon/commit/f0d968d42b1d401a40928bc726ba7dfab97104c5))
* **launchevent:** block the send on any failure once encryption is committed ([9cf7adc](https://github.com/encryption4all/postguard-outlook-addon/commit/9cf7adc9a3290a9dc3f462e9181a61c5516b5d31))

## [0.2.0](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.1.6...v0.2.0) (2026-05-10)


### Features

* **sign:** prefill sender attributes in Settings; sign with optional fallback ([e56f75f](https://github.com/encryption4all/postguard-outlook-addon/commit/e56f75f98d15d5d276fd93755dedbe91588f4627))
* **taskpane:** add Dutch (nl) translations and reorder Sign above Manage Access ([d35d281](https://github.com/encryption4all/postguard-outlook-addon/commit/d35d28165e1411162253d5662b64c9013104befa))
* **taskpane:** Dutch i18n + reorder Sign above Manage Access ([#51](https://github.com/encryption4all/postguard-outlook-addon/issues/51), [#52](https://github.com/encryption4all/postguard-outlook-addon/issues/52)) ([1604d7e](https://github.com/encryption4all/postguard-outlook-addon/commit/1604d7e0c674fab73b2dcde3696107645d8e0023))


### Bug Fixes

* **compose:** narrow attr.v before calling extraAttribute ([a700f28](https://github.com/encryption4all/postguard-outlook-addon/commit/a700f28c5cacfe6d67f1c4e8013ba4bc51df7cbe))
* **compose:** narrow attr.v before extraAttribute call ([d69e083](https://github.com/encryption4all/postguard-outlook-addon/commit/d69e083ef97a9a7f87fade46e54d3bf17a744502))
* **compose:** use optional sign attributes via Yivi ([#49](https://github.com/encryption4all/postguard-outlook-addon/issues/49), [#56](https://github.com/encryption4all/postguard-outlook-addon/issues/56)) ([c750895](https://github.com/encryption4all/postguard-outlook-addon/commit/c750895f841ed2adc488199585c8da1854ef1352))
* **launchevent,dialog:** forward optional sign attrs through send pipeline ([fccf152](https://github.com/encryption4all/postguard-outlook-addon/commit/fccf152cf041643a61805131570dee4bc6314764))
* **launchevent,yivi:** enlarge dialog and center QR widget ([02bf2d8](https://github.com/encryption4all/postguard-outlook-addon/commit/02bf2d8abef614acc4dbd308b434828fdad8f23a))
* **launchevent:** prompt before opening Yivi dialog by default ([#48](https://github.com/encryption4all/postguard-outlook-addon/issues/48)) ([202ea22](https://github.com/encryption4all/postguard-outlook-addon/commit/202ea229ac5c78d7b72b71d4568f22f73feb80b0))
* **launchevent:** prompt before opening Yivi dialog by default ([#48](https://github.com/encryption4all/postguard-outlook-addon/issues/48)) ([bf27cbe](https://github.com/encryption4all/postguard-outlook-addon/commit/bf27cbeaf78243d53d626fbdc0a8cb39f4dae403))
* **settings:** cache prefills in module-level state to dodge stale get ([2843a57](https://github.com/encryption4all/postguard-outlook-addon/commit/2843a578294883f4166feaa5055b4c4e1e0f546a))
* **settings:** explicit Save button for sender-attribute prefills ([7503625](https://github.com/encryption4all/postguard-outlook-addon/commit/7503625aab87c2cdcfe01988b39bcde3912c242d))
* **settings:** persist prefills on change, serialize+retry saveAsync ([c25889b](https://github.com/encryption4all/postguard-outlook-addon/commit/c25889bb10d9c28949baa1c8b3bee8864161b87e))
* **settings:** Save returns to the prior view after persisting ([a69309b](https://github.com/encryption4all/postguard-outlook-addon/commit/a69309b0158fe8da5b3bf831537be4b2a2202eae))
* **settings:** wire listeners once, refresh values directly without cloning ([a38cf20](https://github.com/encryption4all/postguard-outlook-addon/commit/a38cf20b6b57c2299dd3813f2913975dca240b4c))
* **taskpane:** move Settings entry to a labeled footer button ([d8e777d](https://github.com/encryption4all/postguard-outlook-addon/commit/d8e777d72cacc1802b004db411869b71bec49354))
* **yivi-dialog:** forward sign-attribute values to pg.sign.yivi ([8a9d936](https://github.com/encryption4all/postguard-outlook-addon/commit/8a9d9363dd6798732b6357d5d5432580370ae3e4))
* **yivi,deps:** use @privacybydesign/yivi-css for Yivi widget styling ([fe33bee](https://github.com/encryption4all/postguard-outlook-addon/commit/fe33bee40bb3b9e71b149f17ad2b28ceaa4fe412))
* **yivi:** center QR inside Yivi host by styling SDK classes ([8d2c0c5](https://github.com/encryption4all/postguard-outlook-addon/commit/8d2c0c53a62cf54a84a1eef9ae4dbcbedea62a7e))

## [0.1.6](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.1.5...v0.1.6) (2026-05-07)


### Bug Fixes

* **launchevent,yivi-dialog:** surface real error message instead of "[object Object]" ([498779a](https://github.com/encryption4all/postguard-outlook-addon/commit/498779a133ad2cc2bd158c6d5dfee67b6f3cdee8))
* **launchevent,yivi-dialog:** surface real error message instead of "[object Object]" ([a875e03](https://github.com/encryption4all/postguard-outlook-addon/commit/a875e03bc399e0ed1b47bed7a66d45bbf5093580))

## [0.1.5](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.1.4...v0.1.5) (2026-05-07)


### Bug Fixes

* **a11y:** WCAG 2.2 AA fixes for taskpane and yivi dialog ([7ee8eee](https://github.com/encryption4all/postguard-outlook-addon/commit/7ee8eee5742ef1d78d9a31fadc523507a472673b))

## [0.1.4](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.1.3...v0.1.4) (2026-05-07)


### Miscellaneous Chores

* release 0.1.4 ([2b8e41a](https://github.com/encryption4all/postguard-outlook-addon/commit/2b8e41a3a10fecb360ee0be5b249324a4733212f))

## [0.1.3](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.1.2...v0.1.3) (2026-05-06)


### Bug Fixes

* **launchevent:** Mac fallback to taskpane; retry + Safari hint elsewhere ([21e812a](https://github.com/encryption4all/postguard-outlook-addon/commit/21e812ad42a1a1d1b36455ff0784b1fa0348ea75))
* **launchevent:** Mac fallback to taskpane; retry pattern + Safari hint ([d114e87](https://github.com/encryption4all/postguard-outlook-addon/commit/d114e8759dee4dd1762c653f15218f8ad46103cd))
* **launchevent:** only deflect Mac when message isn't already encrypted ([56264bd](https://github.com/encryption4all/postguard-outlook-addon/commit/56264bd5d7f4056ef45963115ac2b81ff32ae0ac))
* **taskpane:** broaden isComposeMode for Outlook for Mac ([144411a](https://github.com/encryption4all/postguard-outlook-addon/commit/144411a88b47263c1bce1e6182b761b55715cbb1))
* **taskpane:** drop the hidden attribute on the Encrypt & Send button ([2d83c81](https://github.com/encryption4all/postguard-outlook-addon/commit/2d83c810a49571c5496a86d93c80506c2f6cbf49))
* **taskpane:** show Encrypt & Send button only on Outlook for Mac ([551446d](https://github.com/encryption4all/postguard-outlook-addon/commit/551446df47c81266319a3e2776c91d382b3c0f7d))
* **ui:** add focus-visible and active states to interactive elements ([d78c52c](https://github.com/encryption4all/postguard-outlook-addon/commit/d78c52cb753ecfafa54c132752032a01b154c320))
* **ui:** add focus-visible and active states to interactive elements ([26db5ca](https://github.com/encryption4all/postguard-outlook-addon/commit/26db5ca41d8b84c42aadc90c6e56a57902e4372f))

## [0.1.2](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.1.1...v0.1.2) (2026-05-05)


### Bug Fixes

* **launchevent:** always prompt on Apple WebKit; drop retry and Safari hint ([7f24d77](https://github.com/encryption4all/postguard-outlook-addon/commit/7f24d7785e84e979542c5b0fd4ab153f713cf008))
* **launchevent:** always prompt on WebKit, drop retry and Safari hint ([21c76d0](https://github.com/encryption4all/postguard-outlook-addon/commit/21c76d0a7cbc849bf11f5c99d3718e225b19929d))
* **launchevent:** skip optimistic attempt on Outlook for Mac ([bfc73ce](https://github.com/encryption4all/postguard-outlook-addon/commit/bfc73cec558726a5fb127bda1373b547338650b9))

## [0.1.1](https://github.com/encryption4all/postguard-outlook-addon/compare/v0.1.0...v0.1.1) (2026-05-05)


### Bug Fixes

* allowlist add-in domain in &lt;AppDomains&gt; for launchevent dialogs ([a55edd6](https://github.com/encryption4all/postguard-outlook-addon/commit/a55edd6a781320f3a8691574a7ad7464c1b42e10))
* **launchevent:** always show Office popup prompt so dialogs open ([a3cfeb1](https://github.com/encryption4all/postguard-outlook-addon/commit/a3cfeb11f667b98171ea1d4f82d90ad78410d891))
* **launchevent:** bake add-in origin in at build time, not runtime ([14b2f1c](https://github.com/encryption4all/postguard-outlook-addon/commit/14b2f1c7fdd1cddddd49afbe60b7d109c9e077f2))
* **launchevent:** branch promptBeforeOpen on Apple WebKit, not platform ([ac52cd7](https://github.com/encryption4all/postguard-outlook-addon/commit/ac52cd76fda6715ad83f37e10dae173f2b9ec3f8))
* **launchevent:** derive Yivi dialog URL from runtime origin ([485c718](https://github.com/encryption4all/postguard-outlook-addon/commit/485c718465ed61730a6619c401a9e854b6419f0d))
* **launchevent:** derive Yivi dialog URL from runtime origin ([b1b189d](https://github.com/encryption4all/postguard-outlook-addon/commit/b1b189d1d8c6f4d3ecf17659b1bc6e9c31e90232))
* **launchevent:** drop displayInIframe/promptBeforeOpen on Mac ([f96a9fc](https://github.com/encryption4all/postguard-outlook-addon/commit/f96a9fc61e69a962dbf4083cca3deef536fa6ff6))
* **launchevent:** drop promptBeforeOpen: false so Office's prompt fires ([c0d45d2](https://github.com/encryption4all/postguard-outlook-addon/commit/c0d45d2f1e2fc8b5c854311f07868513084d4e8c))
* **launchevent:** floor dialog size at 30% of screen ([6b144f0](https://github.com/encryption4all/postguard-outlook-addon/commit/6b144f01f60994b426c5d1b07f7d0881012e3edb))
* **launchevent:** keep Office's prompt on every platform ([c5163bb](https://github.com/encryption4all/postguard-outlook-addon/commit/c5163bb79d5323e9111f01b2ae29249fdfb48cc5))
* **launchevent:** keep Office's prompt on every platform ([359617c](https://github.com/encryption4all/postguard-outlook-addon/commit/359617ce384d723cdbfd851ee4366232dd0a40cf))
* **launchevent:** keep Office's prompt on Mac, suppress on Web/Windows ([8fc3f55](https://github.com/encryption4all/postguard-outlook-addon/commit/8fc3f55924179bcac061b004841ca7bb091e0fd0))
* **launchevent:** only show Office's popup prompt on Mac ([c273736](https://github.com/encryption4all/postguard-outlook-addon/commit/c273736b46feb34ce89801fd6031b70e80071229))
* **launchevent:** re-add MIN_DIALOG_PCT for usable dialog size ([694e4d1](https://github.com/encryption4all/postguard-outlook-addon/commit/694e4d16f386d8250b152dbe77113ef5fe848a15))
* **launchevent:** surface displayDialogAsync diagnostics in Smart Alert ([9d884ce](https://github.com/encryption4all/postguard-outlook-addon/commit/9d884ce170fc2fac44da392766244a5dab8a19a5))
* **launchevent:** try-without-prompt, fall back to prompt; Safari hint ([184ba9a](https://github.com/encryption4all/postguard-outlook-addon/commit/184ba9a0f20b12486192dc668e99b543274521c2))
* **launchevent:** use build-time ADDIN_PUBLIC_URL for Yivi dialog URL ([91baeef](https://github.com/encryption4all/postguard-outlook-addon/commit/91baeefada929a3559f4623498d466599d2b50a9))
* **launchevent:** use iframe-mode dialog on Mac to bypass popup blocker ([9466f09](https://github.com/encryption4all/postguard-outlook-addon/commit/9466f09db5d0bab7727448658d0a9e8d3db9d614))
* **launchevent:** use promptBeforeOpen on Mac for popup gesture ([c1d69fc](https://github.com/encryption4all/postguard-outlook-addon/commit/c1d69fcd8dfda526fcc06202be602c3031795b94))
* **manifest:** allowlist add-in domains for dialogs from launchevent ([0a74f9a](https://github.com/encryption4all/postguard-outlook-addon/commit/0a74f9a90b7f7285ee3ed4d551a0a99225c92e41))
* **nginx:** keep inherited mime.types so HTML serves as text/html ([7c81b05](https://github.com/encryption4all/postguard-outlook-addon/commit/7c81b05b0ad1404e72461dcc9579d0ef4c1164c8))
* **nginx:** keep inherited mime.types so HTML serves as text/html ([1d0b8ee](https://github.com/encryption4all/postguard-outlook-addon/commit/1d0b8ee6d204336e3a494e3b878593725b1caffc))
* use pg4ol metric client id to match PKG convention ([35e8424](https://github.com/encryption4all/postguard-outlook-addon/commit/35e84244b0cc21cd70848d8f005d55ff44ebb48c))
* use pg4ol metric client id to match PKG convention ([11ae4ba](https://github.com/encryption4all/postguard-outlook-addon/commit/11ae4ba47efc34f06cd2fc101a3d6f4782d16488))


### Reverts

* **launchevent:** drop exploratory diagnostics; keep AppDomains fix ([066d228](https://github.com/encryption4all/postguard-outlook-addon/commit/066d2280b7f27a89217f9a5ffe8b56338afd6bfe))
* **launchevent:** drop MIN_DIALOG_PCT floor ([3f63e3d](https://github.com/encryption4all/postguard-outlook-addon/commit/3f63e3df0756f219d3c896c9a10a9f9e7023a79b))
* **launchevent:** drop MIN_DIALOG_PCT floor ([79be563](https://github.com/encryption4all/postguard-outlook-addon/commit/79be56305446e83a197323687ff4d084850e267a))
