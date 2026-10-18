# Architecture

Teams Chat Backup is a Manifest V3 extension written in plain JavaScript. There is no build step and no server.

## Components

```
 Toolbar popup (popup.html / popup.js)
   │  chrome.scripting.executeScript  (activeTab, after you click)
   ▼
 Teams tab: content.js ── reads the DOM, scrolls, downloads images/files from Teams
   │  chrome.runtime messages (messages, images as base64, PDF data)
   ▼
 popup.js ── db.js ──► IndexedDB (in your browser)
   │                       ▲
   │                       │
   ▼                       │
 background.js         Chat Viewer (viewer.html / viewer.js)
   ├─ file downloads      search, filters, edit/star/delete, export/import .zip
   └─ Pro status ──► api.revenuecat.com  (License ID only)
```

## Backup flow

1. You open the popup on a Teams tab. `popup.js` checks the URL and injects `content.js` into that tab.
2. `content.js` lists the chats in the sidebar. You choose one or more.
3. For each chat, `content.js`:
   - opens it if it isn't the current one;
   - scrolls the message list to the very top until Teams stops loading older messages (longer waits in Long Chat mode);
   - walks down the list and reads each message: author, ISO timestamp, text, images, reactions, attachments and system messages;
   - converts images to base64 so they survive after Teams' blob URLs expire;
   - for PDF backups, captures the chat with html2canvas and stitches pages with jsPDF.
4. `popup.js` saves the result with `db.js`. If the chat was backed up before, new messages are merged and your edits, stars and deletions are kept.

## Storage (`db.js`)

IndexedDB database with stores for chat metadata, messages, files (PDF, text) and the recycle bin. Deleting a chat moves it to the recycle bin first. Everything stays in the browser profile until you delete it or uninstall the extension, which is why the viewer offers **Export All**.

## Chat Viewer

`viewer.html` reads IndexedDB directly and renders chats in a Teams-like layout: search with highlighting, jump to date, starred filter, edit history, attachments tab, PDF/HTML/TXT downloads, and `.zip` export/import (JSZip).

## Free and Pro

The free tier limits how much of a chat is saved and how many chats can be backed up. Pro is a one-time purchase:

- `license.js` creates a random License ID and builds the checkout link for it.
- After checkout, `background.js` asks RevenueCat whether that License ID has the Pro entitlement and caches the answer in `chrome.storage`, so Pro keeps working offline.
- Restoring a purchase on another computer works with the same License ID or the purchase email.

## Selectors and languages

Teams changes its markup often. `content.js` tries many selectors for the message list and messages, and identifies parts of a message by structure (`time` elements, author and body containers) rather than by text, so it works in every Teams display language.

## Tests

`tests/` uses Node's built-in `node:test`:

- `offline-pro.test.mjs` runs `background.js` against a mocked `chrome` API to check Pro activation, caching and offline behavior.
- `pdf-text.test.mjs` checks PDF text handling.
