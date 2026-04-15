# <p align="center"><img src="./img/pg_logo.svg" height="128px" alt="PostGuard" /></p>

> For full documentation, visit [docs.postguard.eu](https://docs.postguard.eu/repos/postguard-outlook-addon).

Identity-based email encryption add-in for Microsoft Outlook. Users can send and receive encrypted email using [Yivi](https://yivi.app) identity verification, without needing to exchange keys. This is one of the main end-user clients for PostGuard, alongside the Thunderbird add-on.

## Development

Requires Node.js 20 or later.

```bash
npm install
npm run dev-server
```

The dev server runs on port 3000. For production builds:

```bash
npm run build
```

## Releasing

There are no automated releases currently. To release a new version:

1. Update the version in `package.json`.
2. Run `npm run build`.
3. Deploy through the Office admin center.

## License

MIT
