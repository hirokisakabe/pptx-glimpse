#!/bin/bash
set -euo pipefail

# Script to generate LibreOffice VRT reference images.
# Assumed to be executed inside a Docker container.
#
# Usage:
#   bash vrt/libreoffice/update_snapshots.sh

FIXTURE_DIR="/workspace/vrt/libreoffice/fixtures"
OUTPUT_DIR="/workspace/vrt/libreoffice/snapshots"
TEMP_DIR="/tmp/libreoffice-render"
TARGET_WIDTH=960

mkdir -p "$OUTPUT_DIR" "$TEMP_DIR"

found=0

for pptx_file in "$FIXTURE_DIR"/*.pptx; do
    [ -f "$pptx_file" ] || continue
    found=1

    basename=$(basename "$pptx_file" .pptx)
    name="$basename"

    if [[ "$name" == editor-validity-* ]]; then
        echo "Skipping editor validity fixture: $pptx_file"
        continue
    fi

    echo "Rendering: $pptx_file"

    # Convert to PNG with LibreOffice headless
    libreoffice --headless --convert-to png \
        --outdir "$TEMP_DIR" "$pptx_file" 2>/dev/null

    # Find output file
    png_file="$TEMP_DIR/${basename}.png"
    if [ ! -f "$png_file" ]; then
        echo "  WARNING: No PNG output for $pptx_file"
        continue
    fi

    # Resize to 960px width with Pillow and save
    python3 -c "
from PIL import Image
import sys

img = Image.open('$png_file')
ratio = $TARGET_WIDTH / img.width
new_height = int(img.height * ratio)
img = img.resize(($TARGET_WIDTH, new_height), Image.LANCZOS)
output_path = '$OUTPUT_DIR/${name}-slide1.png'
img.save(output_path)
print(f'  Saved: {output_path} ({$TARGET_WIDTH}x{new_height})')
"
done

if [ "$found" -eq 0 ]; then
    echo "ERROR: No fixtures found in $FIXTURE_DIR"
    echo "Run fixture generation first."
    exit 1
fi

rm -rf "$TEMP_DIR"
echo "Done!"
