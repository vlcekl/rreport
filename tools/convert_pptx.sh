#!/bin/bash

# convert_pptx.sh
# Converts a PPTX file to a Quarto-compatible Markdown file (.qmd)
# Extracts images, moves them to the appropriate assets folder, and injects YAML front matter.

# Ensure we exit on errors
set -e

# --- Configuration ---
SOURCE_FILE="$1"
# Ensure we have an input, or find the newest one
if [ -z "$SOURCE_FILE" ]; then
    echo "No input file provided. Searching for newest .pptx in import/pptx/..."
    
    # Use ls -t to sort by time, grep for pptx, head -n 1 to get the newest
    NEWEST_PPTX=$(ls -t import/pptx/*.pptx 2>/dev/null | head -n 1)
    
    if [ -z "$NEWEST_PPTX" ]; then
        echo "Error: No .pptx files found in import/pptx/."
        echo "Usage: ./tools/convert_pptx.sh [path_to_pptx_file]"
        exit 1
    fi
    
    SOURCE_FILE="$NEWEST_PPTX"
    echo "Auto-selected: $SOURCE_FILE"
fi

if [ ! -f "$SOURCE_FILE" ]; then
    echo "Error: File '$SOURCE_FILE' not found."
    exit 1
fi
    echo "Error: File '$SOURCE_FILE' not found."
    exit 1
fi

# Determine filenames and directories
BASENAME=$(basename "$SOURCE_FILE" .pptx)
OUTPUT_DIR="present"
ASSETS_DIR="$OUTPUT_DIR/assets/$BASENAME"
OUTPUT_FILE="$OUTPUT_DIR/$BASENAME.qmd"
TEMP_QMD="out.qmd" # pptx2md default output when using --qmd without explicit path in some versions/cases, let's explicitly set it.
TEMP_QMD_EXPLICIT="${BASENAME}_temp.qmd"

echo "=========================================="
echo "Starting Conversion: $BASENAME.pptx"
echo "=========================================="

# --- 1. Environment Check ---
VENV_DIR=".venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "Virtual environment not found. Creating..."
    python3 -m venv "$VENV_DIR"
fi

echo "Activating virtual environment..."
source "$VENV_DIR/bin/activate"

# Ensure tool is installed
if ! command -v pptx2md &> /dev/null; then
    echo "Installing pptx2md..."
    pip install pptx2md
fi

# --- 2. Conversion ---
echo "Running pptx2md..."
# Explicitly name the output so we know where it is, and use the --qmd flag.
# It automatically extracts images to an 'img' directory relative to where it runs.
pptx2md "$SOURCE_FILE" --qmd -o "$TEMP_QMD_EXPLICIT"

# --- 3. Asset Management ---
# pptx2md usually creates an "img" directory for extracted assets.
if [ -d "img" ]; then
    echo "Found extracted images. Moving to $ASSETS_DIR..."
    mkdir -p "$ASSETS_DIR"
    cp -r img/* "$ASSETS_DIR"/
    rm -rf img
    
    # Update image paths in the generated qmd.
    # It usually writes them as ![...](img/...) or <img src="img/..."/>
    echo "Updating image paths in the markdown..."
    sed -i "s|\](img/|\](assets/$BASENAME/|g" "$TEMP_QMD_EXPLICIT"
    sed -i "s|src=\"img/|src=\"assets/$BASENAME/|g" "$TEMP_QMD_EXPLICIT"
else
    echo "No images found to extract/move."
fi

# --- 4. YAML Front Matter Injection ---
echo "Injecting Quarto YAML front matter..."

# Create a new file with the YAML header
cat <<EOF > "$OUTPUT_FILE"
---
title: "Converted: $BASENAME"
format: 
  revealjs:
    theme: default
    slide-number: false
    preview-links: auto
  pptx:
    # reference-doc: ../templates/template.pptx
    slide-level: 2
execute:
  echo: false
  warning: false
  message: false
---

EOF

# Append the converted content (skipping any generic title it might have injected at the very top if we wanted, 
# but simply appending is safest).
cat "$TEMP_QMD_EXPLICIT" >> "$OUTPUT_FILE"

# Clean up temp file
rm "$TEMP_QMD_EXPLICIT"

echo "=========================================="
echo "Conversion Complete!"
echo "Output: $OUTPUT_FILE"
echo "Assets: $ASSETS_DIR"
echo "=========================================="
echo "Next step: Run '/refine_quarto' with your agent to clean up the formatting!"

deactivate
