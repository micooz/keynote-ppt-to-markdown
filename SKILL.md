---
name: keynote-ppt-to-markdown
description: "Convert Keynote/PPT presentations to Markdown with embedded images and speaker notes. Pure Python implementation."
---

# Keynote/PPT to Markdown Skill

Convert presentations to structured Markdown with slide images and speaker notes.

## Usage

```bash
python3 scripts/convert.py <presentation.pptx> [-o output-directory]
```

### Examples

```bash
# Convert PPTX file
python3 scripts/convert.py presentation.pptx

# Convert to specific output directory
python3 scripts/convert.py slides.pptx -o ./output

# Specify output
python3 scripts/convert.py slides.pptx --output /path/to/output
```

## Output

```
output/
├── presentation.md   # Markdown with images and notes
└── images/
    ├── 001.png
    ├── 002.png
    └── ...
```

## Notes

- **PowerPoint (.pptx)**: Works on any platform (pure Python)
- **Keynote (.key)**: Requires macOS with Keynote installed (use Keynote to export as PPTX first)
- Dependencies: Python 3.6+
