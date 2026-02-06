---
name: keynote-ppt-to-markdown
description: "Convert Keynote/PPT presentations to Markdown with speaker notes and slide images. Use when you need to extract content from presentation files for documentation, AI processing, or content repurposing."
---

# Keynote/PPT to Markdown Skill

Convert Keynote (.key) and PowerPoint (.pptx) presentations to structured Markdown with images and speaker notes.

## Script Usage

Run the conversion script directly:

```bash
# Convert Keynote or PPTX file
./scripts/convert.sh <presentation.key or .pptx> [output-directory]

# Examples
./scripts/convert.sh presentation.key
./scripts/convert.sh presentation.pptx ./output
./slides_convert.sh slides.key /path/to/output
```

## Output Structure

```
output-directory/
├── presentation.md    # Markdown with embedded images and speaker notes
└── images/
    ├── slide-1.png
    ├── slide-2.png
    └── ...
```

## Programmatic Usage (Node.js)

```javascript
const { runCli } = require('./dist/index.js');

async function convert() {
  await runCli();
}
```

## Notes

- Keynote conversion requires macOS with Keynote installed
- PowerPoint (.pptx) works on any platform
- Dependencies: npm install (adm-zip, pptx2json, xml2js)
