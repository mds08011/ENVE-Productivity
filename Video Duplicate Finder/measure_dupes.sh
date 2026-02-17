#!/bin/bash

# 1. Define the target directory
SEARCH_DIR="/media/mds08011/Elements/"
TEMP_FILE="/tmp/video_hashes.txt"

# File extensions to look for
EXTENSIONS="-iname *.mp4 -o -iname *.wmv -o -iname *.mkv -o -iname *.mov -o -iname *.avi -o -iname *.flv -o -iname *.m4v -o -iname *.mpg"

echo "---------------------------------------------------"
echo "Starting scan in: $SEARCH_DIR"
echo "This will calculate MD5 hashes. It may take time."
echo "---------------------------------------------------"

# 2. Find files and calculate hashes
# We use a temporary file to store the list: [Hash] [FilePath]
find "$SEARCH_DIR" -type f \( $EXTENSIONS \) -exec md5sum {} + | sort > "$TEMP_FILE"

# 3. Analyze the list
current_hash=""
total_wasted_bytes=0

echo "Duplicate files found (these are safe to delete):"

while read -r hash filepath; do
    # If the hash matches the previous line, it is a duplicate
    if [ "$hash" == "$current_hash" ]; then
        
        # Get the size of this specific duplicate file in bytes
        file_size=$(stat -c%s "$filepath")
        
        # Add to total
        total_wasted_bytes=$((total_wasted_bytes + file_size))
        
        # Print the path for the user to see
        echo "  [DUPLICATE] $filepath"
    else
        # This is a new unique file (the 'original'), so we reset the hash 
        # and do NOT add its size to the wasted total.
        current_hash="$hash"
    fi
done < "$TEMP_FILE"

# 4. cleanup
rm "$TEMP_FILE"

# 5. Convert bytes to human-readable format
if command -v numfmt &> /dev/null; then
    human_size=$(numfmt --to=iec --suffix=B $total_wasted_bytes)
else
    # Fallback if numfmt isn't installed
    human_size="$((total_wasted_bytes / 1024 / 1024)) MB"
fi

echo "---------------------------------------------------"
echo "Scan Complete."
echo "Total space taken by duplicates: $human_size"
echo "---------------------------------------------------"
