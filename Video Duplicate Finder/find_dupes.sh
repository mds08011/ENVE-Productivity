#!/bin/bash

# Define the search directory
SEARCH_DIR="/media/mds08011/Elements/"

# Check if directory exists
if [ ! -d "$SEARCH_DIR" ]; then
    echo "Error: Directory $SEARCH_DIR not found."
    exit 1
fi

echo "Scanning for duplicate videos in $SEARCH_DIR..."
echo "This may take a while depending on the number of files..."

# 1. Find video files (case-insensitive)
# 2. Calculate MD5 checksum for each
# 3. Sort by checksum
# 4. Filter to show only duplicates
find "$SEARCH_DIR" -type f \( -iname "*.mp4" -o -iname "*.wmv" -o -iname "*.mkv" -o -iname "*.mov" -o -iname "*.avi" -o -iname "*.flv" \) -print0 | \
xargs -0 md5sum | \
sort | \
uniq -w32 -dD
