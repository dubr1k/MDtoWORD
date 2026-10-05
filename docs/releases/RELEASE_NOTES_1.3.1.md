# MDtoWORD 1.3.1

## Два варианта ГОСТ / Two GOST variants

- **ГОСТ 7.32 / GOST 7.32** (`preset="gost"`): existing behaviour preserved — A4, black Times New Roman 14 pt, margins left/right/top/bottom 30/15/20/20 mm, 1.5 spacing, 1.25 cm first-line indent, native footnotes and centred footer page numbers.
- **ГОСТ — пользовательский (12 pt) / GOST — custom (12 pt)** (`preset="gost_user"`): a separate user adaptation, **not a claim of full GOST 7.32 compliance**. A4, black Times New Roman 12 pt, margins 30/15/15/15 mm, 1.5 spacing, justified body and no headers/footers.
- Both variants are selectable in the Russian/English GUI, MCP `markdown_to_word` and `preview_markdown`, and the documented Python API. The `default` preset is unchanged. Explicit font, size, page and footnote options retain priority. Explicit templates keep their own styles, margins, headers and footers.
- Markdown `[^n]` references remain real, automatically numbered, bottom-of-page Word footnotes by default. Supply GOST-formatted bibliographic entries in the Markdown yourself: the converter preserves them but does not generate or validate a bibliography. Ordinary Markdown links remain hyperlinks. Explicit `footnotes="section"` remains available.

## Release verification

GitHub Actions now publishes only after tests, metadata checks and **both** macOS arm64 and Windows x64 builds succeed. Published ZIPs are recursively checked for CRC integrity, SHA-256, executable/runtime/icon presence and first-party source/privacy exclusions.

Download `MDtoWORD-macOS-arm64.zip` or `MDtoWORD-Windows-x64.zip` with the corresponding `.sha256`. The macOS app is ad-hoc signed, not notarized; the Windows executable is not Authenticode-signed.

## CI runtime fix

The first 1.3.0 tag run was stopped before publication because the Ubuntu headless GUI tests lacked `libEGL.so.1`. The workflow now explicitly installs the Qt system runtime before running the full suite. The unpublished 1.3.0 tag is kept immutable; 1.3.1 is the complete two-platform release.
