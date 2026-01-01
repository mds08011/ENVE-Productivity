#!/bin/bash

# ================= CONFIGURATION =================
# 1. Source: Your External HDD
SOURCE_DIR="/media/mds08011/Elements/"

# 2. Destination: Your pCloud Drive 
# (Update "Civil_Projects_Backup" to your desired folder name on pCloud)
DEST_DIR="$HOME/pCloudDrive/Civil_Projects_Backup"

# 3. Dry Run Mode (Safety First)
# Set to true to test what WOULD happen without actually copying anything.
# Set to false to actually run the upload.
DRY_RUN=true

# ================= EXCLUSION LIST =================
# Based on your provided Exclude.txt
# This list defines which file extensions to SKIP during upload.
EXCLUDED_EXTENSIONS=(
    ".dwg" ".bak" ".shp" ".shx" ".dbf" ".prj" ".cpg" ".sbn" ".sbx"
    ".rvt" ".rte" ".rfa" ".DGN" ".DWF" ".DXF" ".IFC" ".SAT" ".SKP"
    ".jpg" ".heic" ".mov" ".aec" ".atx" ".gdbtable" ".las" ".laz"
    ".pts" ".e57" ".gdb" ".mdb" ".geodatabase" ".kml" ".kmz" 
    ".nwc" ".nwd" ".obj" ".3ds" ".fbx" ".tiff" ".tif" ".ipt" 
    ".iam" ".idw" ".nc" ".sda" ".pln" ".str" ".ovr" ".xml" 
    ".tfw" ".mdx" ".trc" ".gif" ".wav"
)

# ================= LOGIC =================

# 1. Build the exclusion parameters for rsync
rsync_excludes=()
for ext in "${EXCLUDED_EXTENSIONS[@]}"; do
    # Add wildcard to extension (e.g., .dwg -> *.dwg)
    rsync_excludes+=( --exclude="*$ext" )
    # Also attempt to exclude uppercase version just in case (e.g., *.DWG)
    # This is a basic uppercase conversion for common safety
    rsync_excludes+=( --exclude="*${ext^^}" )
done

# 2. Prepare the command options
# -a : Archive mode (recursive, preserves timestamps/permissions)
# -v : Verbose (shows progress)
# -h : Human readable numbers
# --delete : (OPTIONAL) Deletes files in Cloud if you deleted them on HDD. 
#            I omitted it for safety, but you can add it if you want true mirroring.
CMD_OPTS="-avh --progress"

if [ "$DRY_RUN" = true ]; then
    CMD_OPTS="$CMD_OPTS --dry-run"
    echo "!!! DRY RUN MODE ACTIVE !!!"
    echo "No files will be transferred. This is just a test."
    echo "------------------------------------------------"
fi

echo "Starting Sync..."
echo "Source: $SOURCE_DIR"
echo "Dest:   $DEST_DIR"
echo "Excluding ${#EXCLUDED_EXTENSIONS[@]} file types..."

# 3. Run rsync
rsync $CMD_OPTS "${rsync_excludes[@]}" "$SOURCE_DIR" "$DEST_DIR"

echo "------------------------------------------------"
if [ "$DRY_RUN" = true ]; then
    echo "Dry run complete. If the output looks correct:"
    echo "Edit this script and change DRY_RUN=true to DRY_RUN=false"
fi