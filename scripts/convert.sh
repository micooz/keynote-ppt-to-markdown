#!/bin/bash
# Keynote/PPT to Markdown Converter
# Usage: ./convert.sh <presentation.key or .pptx> [output-directory]

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

INPUT_FILE="$1"
OUTPUT_DIR="${2:-.}"

if [ -z "$INPUT_FILE" ]; then
    echo "Usage: $0 <presentation.key or .pptx> [output-directory]"
    exit 1
fi

if [ ! -f "$INPUT_FILE" ]; then
    echo "Error: Input file not found: $INPUT_FILE"
    exit 1
fi

# Install dependencies if needed
if [ ! -d "node_modules" ]; then
    echo "Installing dependencies..."
    npm install --silent
fi

# Run the conversion
echo "Converting $INPUT_FILE to Markdown..."
node dist/index.js "$INPUT_FILE" "$OUTPUT_DIR"

echo "Done!"
