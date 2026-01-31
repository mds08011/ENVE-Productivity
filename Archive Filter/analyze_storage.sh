#!/bin/bash

# --- CONFIGURATION ---
SEARCH_DIR="/media/mds08011/Elements/"

# UPDATED EXCLUDE LIST
# Added: pst, dll, cab, xml
EXCLUDE_LIST="dwg bak shp shx dbf prj cpg sbn sbx rvt rte rfa dgn dwf dxf ifc sat skp aec atx gdbtable las laz pts e57 gdb mdb geodatabase kml kmz nwc nwd obj 3ds fbx tiff tif ipt iam idw nc sda pln str ovr xml tfw mdx trc sid ecw jp2 sv$ rcp rcs jpg out tmp hdf msg ts dng grs heic hbn nef rpc stp skb jpeg sqlite sqllite png mms rpt adf h5 rws x_t qsb mpk hof flt fit prc pack dem pst dll cab"

echo "=========================================="
echo "Analyzing: $SEARCH_DIR"
echo "Target Cloud Cap: 2048 GB (2TB)"
echo "Mode: Full List (No Item Limit)"
echo "Scanning files..."

# 1. Gather data
# 2. Sort by Category (Upload/Exclude) then by Size
find "$SEARCH_DIR" -type f -printf "%s %f\n" | awk -v exclude_str="$EXCLUDE_LIST" '
BEGIN {
    # Load exclude list
    split(exclude_str, arr, " ")
    for (i in arr) {
        excludes[tolower(arr[i])] = 1
    }
}
{
    size = $1
    # Extract extension
    n = split($NF, a, ".")
    if (n > 1) {
        ext = tolower(a[n])
    } else {
        ext = "(no_extension)"
    }

    # Data aggregation
    count[ext]++
    sum[ext] += size
    total_size += size

    if (ext in excludes) {
        excluded_size += size
        type[ext] = "EXCLUDE"
    } else {
        keep_size += size
        type[ext] = "UPLOAD"
    }
}
END {
    # We print raw data here to be sorted/formatted by the next awk command
    for (e in sum) {
        gb = sum[e] / 1024 / 1024 / 1024
        # Threshold: Only show types > 10MB (0.01 GB) to avoid flooding terminal with empty files
        if (gb > 0.01) { 
            print type[e], e, count[e], gb
        }
    }
    
    # Pass totals as a special footer line
    total_gb = total_size / 1024 / 1024 / 1024
    excluded_gb = excluded_size / 1024 / 1024 / 1024
    upload_gb = (total_size - excluded_size) / 1024 / 1024 / 1024
    print "TOTALS", total_gb, excluded_gb, upload_gb
}' | sort -k1,1r -k4,4rn | awk '
BEGIN {
    # Setup format strings
    fmt_head = "%-10s %-15s %-15s\n"
    fmt_row  = "%-10s %-15d %-15.2f\n"
}
{
    if ($1 == "TOTALS") {
        # Capture totals for the end
        t_total = $2
        t_excl = $3
        t_up = $4
        next
    }

    # Store lines in arrays to print separately
    if ($1 == "UPLOAD") {
        upload_rows[u_count++] = sprintf(fmt_row, $2, $3, $4)
    } else {
        exclude_rows[e_count++] = sprintf(fmt_row, $2, $3, $4)
    }
}
END {
    # --- PRINT UPLOAD LIST ---
    print "\n=========================================="
    print "   KEEP / UPLOAD LIST"
    print "=========================================="
    printf fmt_head, "EXT", "COUNT", "SIZE (GB)"
    print "------------------------------------------"
    # Print ALL uploads found
    for (i=0; i<u_count; i++) {
        printf "%s", upload_rows[i]
    }

    # --- PRINT EXCLUDE LIST ---
    print "\n=========================================="
    print "   EXCLUSION LIST"
    print "=========================================="
    printf fmt_head, "EXT", "COUNT", "SIZE (GB)"
    print "------------------------------------------"
    # Print ALL excludes found
    for (i=0; i<e_count; i++) {
        printf "%s", exclude_rows[i]
    }

    # --- FINAL SUMMARY ---
    print "\n=========================================="
    print "             FINAL ANALYSIS"
    print "=========================================="
    printf "TOTAL DRIVE SIZE:      %10.2f GB\n", t_total
    printf "EXCLUDED FILES:       -%10.2f GB\n", t_excl
    print "------------------------------------------"
    printf "PROPOSED CLOUD UPLOAD: %10.2f GB\n", t_up
    print "------------------------------------------"

    if (t_up <= 2048) {
        print "RESULT: SUCCESS! Fits in 2TB."
    } else {
        diff = t_up - 2048
        printf "RESULT: OVER LIMIT by %.2f GB.\n", diff
    }
}'