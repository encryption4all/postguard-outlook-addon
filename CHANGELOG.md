# Changelog

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
