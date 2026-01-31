#!/bin/bash

# --- CONFIGURATION ---
SEARCH_DIR="/media/mds08011/Elements/"

# UPDATED EXCLUDE LIST
# Now includes: jpg, out, tmp, hdf, msg, ts, dng, grs, heic, hbn, nef, rpc, stp
EXCLUDE_LIST="dwg bak shp shx dbf prj cpg sbn sbx rvt rte rfa dgn dwf dxf ifc sat skp aec atx gdbtable las laz pts e57 gdb mdb geodatabase kml kmz nwc nwd obj 3ds fbx tiff tif ipt iam idw nc sda pln str ovr xml tfw mdx trc sid ecw jp2 sv$ rcp rcs jpg out tmp hdf msg ts dng grs heic hbn nef rpc stp"

echo "=========================================="
echo "Analyzing: $SEARCH_DIR"
echo "Target Cloud Cap: 2048 GB (2TB)"
echo "Mode: Terminal Output Only"
echo "Scanning files..."

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

    # Sort logic handled by piping to sort command below
    for (e in sum) {
        gb = sum[e] / 1024 / 1024 / 1024
        # Threshold: Only show types > 0.5 GB to keep terminal clean
        if (gb > 0.5) {
             printf "%-10s %-10s %-15d %-15.2f\n", status[e], e, count[e], gb
        }
    }
    
    # Calculate Final Totals
    total_gb = total_size / 1024 / 1024 / 1024
    excluded_gb = excluded_size / 1024 / 1024 / 1024
    upload_gb = (total_size - excluded_size) / 1024 / 1024 / 1024
    
    # Footer Summary (Printed last so it stays visible)
    print "SUMMARY_MARKER" 
    print "------------------------------------------------------"
    printf "TOTAL DRIVE SIZE:      %10.2f GB\n", total_gb
    printf "EXCLUDED FILES:       -%10.2f GB\n", excluded_gb
    print "------------------------------------------------------"
    printf "CLOUD UPLOAD SIZE:     %10.2f GB\n", upload_gb
    print "------------------------------------------------------"
    
    if (upload_gb <= 2048) {
        print "RESULT: SUCCESS. Fits in 2TB Cloud Storage."
    } else {
        diff = upload_gb - 2048
        printf "RESULT: OVER LIMIT by %.2f GB.\n", diff
    }
}' | sort -k4 -rn | awk '/SUMMARY_MARKER/{flag=1; next} !flag {print} flag {print}'