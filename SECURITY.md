# Security Policy

## Reporting a vulnerability

Please report security problems privately. Use **Report a vulnerability** on this repository's [Security tab](https://github.com/betoxf/teams-chat-backup/security/advisories/new), or email [support@teamschatbackup.com](mailto:support@teamschatbackup.com) with "Security" in the subject.

Include the extension version, browser, steps to reproduce and what an attacker could gain. We aim to reply within 3 business days. Please don't open a public issue for a vulnerability until a fix has shipped.

Only the latest version on the Chrome Web Store is supported.

## Security boundaries

These are the guarantees the code is written to keep. A change that breaks one is a security bug.

- **No backup server.** Chats, names, images and files are stored in the browser's IndexedDB (`db.js`) and in files you choose to download. The extension has no server that receives chat content.
- **Access only when you click.** The manifest requests `activeTab` and `scripting`, not host permissions. `content.js` is injected into the current Teams tab only after you open the popup there.
- **Network requests are limited to:**
  - the Teams and SharePoint/OneDrive URLs already present in the chat, to download images and attachments you can already open, using your existing Teams session;
  - `api.revenuecat.com`, only to check Pro status. The request contains your random License ID and nothing else from your chats.
- **The RevenueCat key is a public client key.** RevenueCat Web Billing keys (`rcb_…`) are designed to ship inside client apps and can only read the status of a License ID. This repository uses a placeholder; the store build uses the real public key.
- **No analytics, telemetry or remote code.** The Content Security Policy is `script-src 'self'`, and every library is bundled in `libraries/`.
- **Imports are treated as untrusted.** Files imported into the Chat Viewer (`.tcb`, `.json`, `.zip`) are parsed as data, and message text is HTML-escaped before it is displayed. A way to run script from an imported file is a security bug.
- **`downloads` is optional.** It is requested only when you save a file.

## Out of scope

- What Microsoft Teams shows you or lets you open. The extension only reads what your account can already see.
- Someone with access to your unlocked computer and browser profile. Local backups are as private as the profile they live in.
