---
name: keynote-ppt-to-markdown
description: "Convert Keynote/PPT to Markdown with images and speaker notes. Agent skill with built-in scripts."
---

# Keynote/PPT to Markdown Skill

Convert presentations to Markdown with embedded images and speaker notes.

## Agent Usage

Execute the conversion script:

```bash
./scripts/convert.sh <presentation.key or .pptx> [output-directory]
```

## Output

```
output/
├── presentation.md   # Markdown with images and notes
└── images/
    ├── slide-001.png
    ├── slide-002.png
    └── ...
```

## Requirements

- Node.js >= 18
- macOS for Keynote (.key) files
- PowerPoint (.pptx) works on any platform
