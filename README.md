# PostGuard addon for Outlook

| :exclamation: We are still in alpha development, thus use at your own risk and not for production purposes! |
| ----------------------------------------------------------------------------------------------------------- |

The PostGuard addon utilizes E2E identity-based encryption to secure e-mails. This allows for easy-to-use encryption without the burden of key management.

Anyone can encrypt without prior setup using this system. For decryption, a user requests a decryption key from trusted third party. To do so, the user must authenticate using [Yivi](https://yivi.app/), a privacy-friendly decentralized identity platform based on the Idemix protocol. Any combination of attributes in the Yivi ecosystem can be used to encrypt e-mails, which allows for detailed access control over the e-mail's content.

Examples include:

- Sending e-mails to health care professionals, using their professional registration number.
- Sending e-mails to people that can prove that they have a certain role within an organisation.
- Sending e-mails to people that can prove that they are over 18.
- Sending e-mails to people that can prove that they live within a country, city, postal code or street.
- Or any combination of the previous examples.

For more information, see [our website](https://postguard.eu/) and [our Github
organisation](https://github.com/encryption4all/).

## Prerequisites

Node and a package manager `npm`, are required to build the addon. Building the addon was tested on:

- node `v16.13.0`

## Building

Install the dependencies and start the addon using:

```
npm start
```

To build the addon, use:

```
npm run build
```

## WebAssembly

Postguard's [cryptographic core](https://github.com/encryption4all/postguard/tree/main/pg-core) is implemented in Rust, which is compiled down to WebAssembly in [irmaseal-wasm-bindings](https://github.com/encryption4all/postguard/tree/main/pg-wasm) using `wasm-pack`.

## Funding

PostGuard is being developed by a multidisciplinary (cyber security, UX) team from
[iHub](https://ihub.ru.nl/) at Radboud University, in the Netherlands, funded
by NWO.
