# Contributing

Thanks for helping. Bug reports, Teams layout breakages and small focused fixes are the most useful contributions.

By contributing you agree that your contribution is licensed under the project's [license](LICENSE) and that Bullfighter Studio may also ship it in the Chrome Web Store build.

## Set up

1. Clone the repository. There are no dependencies and no build step.
2. Open `chrome://extensions`, turn on **Developer mode**, click **Load unpacked** and select the folder.
3. After editing, click the reload icon on the extension card and reopen Teams.

## Change standards

- Keep it plain JavaScript that runs as-is in Manifest V3. No bundler, no remote scripts.
- Don't add permissions. If a change really needs one, explain why in the PR.
- Chat content must never leave the browser. See [SECURITY.md](SECURITY.md).
- Teams changes its markup often. When you touch selectors in `content.js`, keep the existing fallbacks and don't depend on display text, so every Teams language keeps working.
- Match the style of the surrounding code and keep diffs focused.

## Validate

```bash
node --test tests/*.mjs
```

Then test by hand in Teams: back up a short chat, a long chat and a channel; open the result in the Chat Viewer; export and re-import a `.zip`.

## Pull requests

- One topic per PR, with a short description of the behavior change and how you tested it.
- Add screenshots for visible UI changes.
- CI must pass.
