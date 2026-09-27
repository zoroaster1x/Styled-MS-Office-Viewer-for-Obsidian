# Styled MS Office Viewer {{VERSION}}

Read Microsoft Office and OpenDocument files inside Obsidian with the styling they were saved with. Workbooks keep their fills, fonts, borders, merged cells and frozen panes; Word documents are laid out as pages; presentations open as slides with a thumbnail rail, a slideshow mode and speaker notes. Strictly read only: the plugin never writes to your documents.

Supports `.xlsx` `.xlsm` `.xltx` `.xltm` `.csv` `.tsv` `.ods` `.docx` `.docm` `.dotx` `.dotm` `.odt` `.rtf` `.doc` `.pptx` `.pptm` `.ppsx` `.ppsm` `.potx` `.potm` `.odp`. Requires Obsidian 1.5.0 or newer. Desktop and mobile.

Release assets are built and attested by GitHub Actions. Verify a downloaded `main.js` with:

```bash
gh attestation verify main.js --repo zoroaster1x/Styled-MS-Office-Viewer-for-Obsidian
```

## Changes since {{PREVIOUS}}

{{CHANGES}}

## Install

1. Download `main.js`, `manifest.json` and `styles.css` from the assets below.
2. Put them in `<Vault>/.obsidian/plugins/styled-ms-office-viewer/`.
3. In Obsidian: Settings, Community plugins, reload the installed plugins, enable **Styled MS Office Viewer**.

## Documentation

Formats, settings, benchmarks and the honest known limits live in the [README](https://github.com/zoroaster1x/Styled-MS-Office-Viewer-for-Obsidian#readme).

## Funding

If this plugin saves you time, consider supporting its development. Every contribution goes toward maintenance, new format coverage and the long tail of edge cases that make real documents painful to render.

**Monero (XMR):**

```
8BdxmQSniku4dBJXWPXeXvgjztmj5nmvWQqeCrVvCtYciusbAyo4rqrGCefTfQ4gGaVZmLN7VgLiYUYyBdYFEwHn1UWPjWs
```

> **Tip:** You can easily purchase Litecoin using [Cake Wallet](https://cakewallet.com/) and then, within the app, create a Monero wallet and exchange the Litecoin into it, pointed at the address above.

Crypto isn't your thing? Starring the repository, filing clear bug reports with a sample file, and telling other Obsidian users about the plugin all help just as much.

## License

GPL-3.0-or-later. Copyright (C) 2026 Zoroaster1x.
