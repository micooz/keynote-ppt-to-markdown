---
name: keynote-ppt-to-markdown
description: "Convert Keynote/PPT presentations to Markdown format with speaker notes and slide images. Use when you need to extract content from presentation files for documentation, AI processing, or content repurposing."
---

# Keynote/PPT to Markdown Skill

Convert Keynote (.key) and PowerPoint (.pptx) presentations to structured Markdown with images and speaker notes.

## Installation

```bash
npm install -g keynote-ppt-to-markdown
```

## Usage

### CLI Command

```bash
ppt2md <presentation-path> [output-directory]
```

### Parameters

- `<presentation-path>`: Required. Path to Keynote (.key) or PowerPoint (.pptx) file
- `[output-directory]`: Optional. Output directory (defaults to current directory)

### Examples

```bash
# Convert Keynote file
ppt2md presentation.key

# Convert PowerPoint file to specific output directory
ppt2md presentation.pptx ./output

# Use npx directly
npx keynote-ppt-to-markdown slides.key ./docs
```

## Output Structure

```
output-directory/
├── presentation.md    # Main Markdown file with embedded images
└── images/
    ├── slide-1.png
    ├── slide-2.png
    └── ...
```

## Notes

- Keynote conversion requires macOS with Keynote installed
- PowerPoint (.pptx) conversion works on any platform
- Images are extracted from each slide and embedded as Markdown image references
- Speaker notes are preserved in the Markdown output

## Programmatic Usage

```typescript
import { convert } from './src/index';

const result = await convert({
  input: 'presentation.key',
  output: './output',
  includeNotes: true,
  includeImages: true
});
```
