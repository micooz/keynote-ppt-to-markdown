---
name: keynote-ppt-to-markdown
description: "Convert Keynote/PPT presentations to Markdown with embedded images and speaker notes. Uses AppleScript for .key files on macOS."
---

# Keynote/PPT to Markdown Skill

Convert presentations to Markdown with slide images and speaker notes.

## Agent Usage

Execute the CLI to convert:

```bash
# Convert Keynote file (macOS only)
node src/index.js <presentation.key> [output]

# Convert PowerPoint file
node src/index.js <presentation.pptx> [output]

# Or use npx
npx keynote-ppt-to-markdown <file> [output]
```

## Output Structure

```
output/
├── presentation.md   # Markdown with images and notes
└── images/
    ├── 001.png
    ├── 002.png
    └── ...
```

## Notes

- **Keynote (.key)**: Requires macOS with Keynote installed (uses AppleScript)
- **PowerPoint (.pptx)**: Works on any platform
- Dependencies: adm-zip, pptx2json, xml2js

## Source Files

- `src/index.js` - Main CLI entry (compiled from TypeScript)
- `src/utils.js` - PPTX parsing utilities
- `src/export_slides_to_images.applescript` - Keynote slide export
- `src/convert_keynote_to_pptx.applescript` - Keynote to PPTX conversion
