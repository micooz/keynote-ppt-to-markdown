# keynote-ppt-to-markdown

[中文](README.md) | [English](README_en.md)

A command-line tool to convert Keynote/PPT presentations to Markdown format, supporting the export of speaker notes and slide images.

## Features

- Supports Keynote and PowerPoint presentation conversion
- Preserves speaker notes
- Exports slide images
- Generates structured Markdown documents

## Usage

Run directly:

```bash
npx keynote-ppt-to-markdown <presentation_path> [output_directory]
```

Or, install first and then run:

```bash
npm install -g keynote-ppt-to-markdown
ppt2md <presentation_path> [output_directory]
```

### Arguments

- `<presentation_path>`: Required, path to the Keynote or PowerPoint file
- `[output_directory]`: Optional, specify the output directory, defaults to the current directory

### Examples

```bash
# Convert a Keynote file
npx keynote-ppt-to-markdown presentation.key

# Convert a PowerPoint file and specify the output directory
npx keynote-ppt-to-markdown presentation.pptx ./output
```

## Output

The converted output includes:

- `presentation.md`: A Markdown file containing images of each slide and speaker notes
- `images/`: A directory containing all slide images

## Development

### Install Dependencies

```bash
npm install
```

### Build

```bash
npm run build
```

### Run

```bash
npm run dev
npx keynote-ppt-to-markdown <presentation_path> [output_directory]
```

## License

MIT

## Contributing

Issues and Pull Requests are welcome!

## Acknowledgements

Cursor helped me complete almost all the code, thanks! 