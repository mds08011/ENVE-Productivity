#!/bin/bash

# --- CONFIGURATION ---
SEARCH_DIR="/media/mds08011/Elements/"
OUTPUT_FILE="storage_report.txt"

# EXCLUDE LIST
# Logic: These files will be counted as "EXCLUDED" (Left on HDD, not uploaded).
# archives (zip/7z) and video (mp4/mov/avi) have been REMOVED from this list so they will count towards your Cloud Backup.
EXCLUDE_LIST="dwg bak shp shx dbf prj cpg sbn sbx rvt rte rfa dgn dwf dxf ifc sat skp aec atx gdbtable las laz pts e57 gdb mdb geodatabase kml kmz nwc nwd obj 3ds fbx tiff tif ipt iam idw nc sda pln str ovr xml tfw mdx trc sid ecw jp2 sv$ rcp rcs"

echo "=========================================="
echo "Analyzing: $SEARCH_DIR"
echo "Target Cloud Cap: 2048 GB (2TB)"
echo "Rules: Keeping Archives & Video. Excluding CAD/GIS/PointClouds."
echo "Scanning files (this may take time)..."

find "$SEARCH_DIR" -type f -printf "%s %f\n" | awk -v exclude_str="$EXCLUDE_LIST" '
BEGIN {
    # Load exclude list into an associative array
    split(exclude_str, arr, " ")
    for (i in arr) {
        excludes[tolower(arr[i])] = 1
    }
}
{
    size = $1
    # Extract extension safely
    n = split($NF, a, ".")
    if (n > 1) {
        ext = tolower(a[n])
    } else {
        ext = "(no_extension)"
    }

    # Aggregate stats
    count[ext]++
    sum[ext] += size
    total_size += size

    # Check against Exclusion List
    if (ext in excludes) {
        excluded_size += size
        status[ext] = "[EXCLUDE]"
    } else {
        keep_size += size
        status[ext] = "[UPLOAD]"
    }
}
END {
    # Print Header
    printf "%-10s %-10s %-15s %-15s\n", "ACTION", "EXT", "COUNT", "SIZE (GB)"
    print "------------------------------------------------------"

    # Sort and print
    for (e in sum) {
        gb = sum[e] / 1024 / 1024 / 1024
        # Show extensions larger than 10MB to keep list clean
        if (gb > 0.01) {
             printf "%-10s %-10s %-15d %-15.2f\n", status[e], e, count[e], gb
        }
    }
    
    # Calculate Final Totals
    total_gb = total_size / 1024 / 1024 / 1024
    excluded_gb = excluded_size / 1024 / 1024 / 1024
    upload_gb = (total_size - excluded_size) / 1024 / 1024 / 1024
    
    # Footer Summary
    print "SUMMARY_MARKER" 
    print "------------------------------------------------------"
    printf "TOTAL DRIVE SIZE:      %10.2f GB\n", total_gb
    printf "EXCLUDED (CAD/GIS):   -%10.2f GB\n", excluded_gb
    print "------------------------------------------------------"
    printf "CLOUD UPLOAD SIZE:     %10.2f GB\n", upload_gb
    print "------------------------------------------------------"
    
    if (upload_gb <= 2048) {
        print "RESULT: SUCCESS. The Video & Archives fit in 2TB."
    } else {
        diff = upload_gb - 2048
        printf "RESULT: OVER LIMIT by %.2f GB.\n", diff
    }
}' | sort -k4 -rn | awk '/SUMMARY_MARKER/{flag=1; next} !flag {print} flag {print}' > "$OUTPUT_FILE"

# --- TERMINAL DISPLAY ---
echo ""
echo "Analysis Complete."
echo "Top 20 File Types by Size:"
echo "------------------------------------------------------"
head -n 25 "$OUTPUT_FILE"
echo ""
echo "------------------------------------------------------"
tail -n 8 "$OUTPUT_FILE"
echo "------------------------------------------------------"
echo "Full detailed report: $OUTPUT_FILE"