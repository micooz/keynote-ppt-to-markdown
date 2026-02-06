# keynote-ppt-to-markdown

[English](README.md) | [中文](README_en.md)

Convert Keynote/PPT presentations to Markdown with embedded images and speaker notes.

## Features

- Convert PowerPoint (.pptx) to Markdown
- Preserve speaker notes
- Export slide images
- Pure Python implementation (no dependencies)

## Usage

```bash
python3 scripts/convert.py <presentation.pptx> [-o output-directory]
```

### Examples

```bash
# Convert PPTX file
python3 scripts/convert.py presentation.pptx

# Output to specific directory
python3 scripts/convert.py slides.pptx -o ./docs
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

## Notes

- **PowerPoint (.pptx)**: Works on any platform (pure Python)
- **Keynote (.key)**: Requires macOS with Keynote installed. Export as PPTX first.

## Requirements

- Python 3.6+

## License

MIT
