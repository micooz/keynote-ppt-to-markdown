# keynote-ppt-to-markdown

[English](README.md) | [中文](README_en.md)

Convert Keynote/PPT presentations to Markdown with embedded images and speaker notes.

## Features

- Convert Keynote (.key) and PowerPoint (.pptx) to Markdown
- Preserve speaker notes
- Export slide images
- AppleScript for macOS/Keynote support

## Usage

```bash
# Using npx
npx keynote-ppt-to-markdown <presentation.key or .pptx> [output-directory]

# Using Node.js directly
node dist/index.js <file> [output]
```

## Output

```
output/
├── presentation.md    # Markdown with images and notes
└── images/
    ├── 001.png
    ├── 002.png
    └── ...
```

## Requirements

- Node.js >= 18
- macOS for Keynote (.key) files
- PowerPoint (.pptx) works on any platform

## Development

```bash
npm install
npm run build    # Compile TypeScript
npm run dev      # Watch mode
```

## License

MIT
