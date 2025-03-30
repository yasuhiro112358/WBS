#!/bin/bash
SOURCE_DIR="./exported"
TARGET_DIR="./utf8"

echo "[INFO] Converting files to UTF-8..."
echo "[INFO] Source directory: $SOURCE_DIR"
echo "[INFO] Target directory: $TARGET_DIR"

mkdir -p "$TARGET_DIR"
for file in "$SOURCE_DIR"/*; do
    filename=$(basename "$file")
    
    if [[ "$filename" == *.frx ]]; then
        echo "[INFO] $filename is a binary file (.frx) - skipped"
        continue
    fi

    tmpfile="$(mktemp)"
    if iconv -f SHIFT_JIS -t UTF-8 "$file" > "$tmpfile"; then
        mv "$tmpfile" "$TARGET_DIR/$filename"
        echo "[INFO] $filename converted to UTF-8"
    else
        echo "[ERROR] Failed to convert $filename"
        rm -f "$tmpfile"
    fi
done
