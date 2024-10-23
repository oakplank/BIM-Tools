import pandas as pd
import os
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QFileDialog
import sys

def compare_excel_files(file_paths):
    """
    Compare multiple Excel files and track changes between consecutive versions based on 'Part Location'.
    Handles multiple entries per 'Part Location'.
    
    Args:
        file_paths: List of paths to Excel files to compare
    Returns:
        String path to the generated comparison report
    """
    if len(file_paths) < 2:
        raise ValueError("At least two Excel files are required for comparison")

    # Create a timestamp for the output file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = os.path.splitext(os.path.basename(file_paths[0]))[0]
    
    # Create a 'reports' directory in the same folder as the first Excel file
    report_dir = os.path.join(os.path.dirname(file_paths[0]), 'reports')
    os.makedirs(report_dir, exist_ok=True)
    
    # Create the full output path
    output_file = os.path.join(report_dir, f"{base_filename}_comparison_{timestamp}.txt")

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(f"Excel File Comparison Report - Sheet: Glazing Master\n")
        f.write(f"Base File: {file_paths[0]}\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        # Compare files sequentially
        for i in range(1, len(file_paths)):
            prev_file = file_paths[i-1]
            curr_file = file_paths[i]

            # Read Excel files - specifically the "Glazing Master" sheet
            try:
                df_prev = pd.read_excel(prev_file, sheet_name="Glazing Master")
                df_curr = pd.read_excel(curr_file, sheet_name="Glazing Master")
            except Exception as e:
                f.write(f"Error reading 'Glazing Master' sheet from files {prev_file} or {curr_file}: {str(e)}\n")
                continue

            # Ensure 'Part Location' exists
            if 'Part Location' not in df_prev.columns or 'Part Location' not in df_curr.columns:
                f.write(f"'Part Location' column missing in one of the files {prev_file} or {curr_file}\n")
                continue

            # Group by 'Part Location'
            grouped_prev = df_prev.groupby('Part Location')
            grouped_curr = df_curr.groupby('Part Location')

            # Get all unique Part Locations across both files
            all_part_locations = set(grouped_prev.groups.keys()).union(set(grouped_curr.groups.keys()))

            f.write(f"\n{'='*80}\n")
            f.write(f"COMPARISON {i}: {os.path.basename(prev_file)} -> {os.path.basename(curr_file)}\n")
            f.write(f"{'='*80}\n")

            # Initialize flags for sections
            added_flag = False
            removed_flag = False
            changes_flag = False

            for part_loc in all_part_locations:
                prev_rows = grouped_prev.get_group(part_loc) if part_loc in grouped_prev.groups else pd.DataFrame()
                curr_rows = grouped_curr.get_group(part_loc) if part_loc in grouped_curr.groups else pd.DataFrame()

                # If Part Location only exists in current file -> Added
                if part_loc not in grouped_prev.groups:
                    if not added_flag:
                        f.write("\nNEW PART LOCATIONS ADDED:\n")
                        f.write("-" * 40 + "\n")
                        added_flag = True
                    for _, row in curr_rows.iterrows():
                        f.write(f"\nPart Location: {part_loc}\n")
                        for col in df_curr.columns:
                            f.write(f"  {col}: {row[col]}\n")
                    continue

                # If Part Location only exists in previous file -> Removed
                if part_loc not in grouped_curr.groups:
                    if not removed_flag:
                        f.write("\nPART LOCATIONS REMOVED:\n")
                        f.write("-" * 40 + "\n")
                        removed_flag = True
                    for _, row in prev_rows.iterrows():
                        f.write(f"\nPart Location: {part_loc}\n")
                        for col in df_prev.columns:
                            f.write(f"  {col}: {row[col]}\n")
                    continue

                # Both files have the Part Location, check for changes
                if not prev_rows.empty and not curr_rows.empty:
                    # Convert rows to dictionaries for comparison
                    prev_dicts = prev_rows.drop(columns=['Part Location']).to_dict(orient='records')
                    curr_dicts = curr_rows.drop(columns=['Part Location']).to_dict(orient='records')

                    # Track matched rows
                    matched_indices_prev = set()
                    matched_indices_curr = set()
                    changes = []

                    for p_idx, p_row in enumerate(prev_dicts):
                        match_found = False
                        for c_idx, c_row in enumerate(curr_dicts):
                            if c_idx in matched_indices_curr:
                                continue
                            if p_row == c_row:
                                matched_indices_prev.add(p_idx)
                                matched_indices_curr.add(c_idx)
                                match_found = True
                                break
                        if not match_found:
                            changes.append({'previous': p_row, 'current': None})

                    for c_idx, c_row in enumerate(curr_dicts):
                        if c_idx not in matched_indices_curr:
                            # Check if this current row matches any previous row
                            match_found = False
                            for p_idx, p_row in enumerate(prev_dicts):
                                if p_idx in matched_indices_prev:
                                    continue
                                if c_row == p_row:
                                    matched_indices_prev.add(p_idx)
                                    matched_indices_curr.add(c_idx)
                                    match_found = True
                                    break
                            if not match_found:
                                changes.append({'previous': None, 'current': c_row})

                    if changes:
                        # Further check if any row in current matches any row in previous
                        # If at least one match exists, no concern
                        # Else, highlight discrepancies
                        total_matches = len(matched_indices_prev)  # Number of exact matches
                        if total_matches == 0:
                            if not changes_flag:
                                f.write("\nPART LOCATION CHANGES:\n")
                                f.write("-" * 40 + "\n")
                                changes_flag = True
                            f.write(f"\nPart Location: {part_loc}\n")
                            for change in changes:
                                if change['previous'] and not change['current']:
                                    f.write("  Row Removed:\n")
                                    for col, val in change['previous'].items():
                                        f.write(f"    {col}: {val}\n")
                                elif not change['previous'] and change['current']:
                                    f.write("  Row Added:\n")
                                    for col, val in change['current'].items():
                                        f.write(f"    {col}: {val}\n")
                                elif change['previous'] and change['current']:
                                    f.write("  Row Changed:\n")
                                    for col in df_prev.columns:
                                        prev_val = change['previous'].get(col, '')
                                        curr_val = change['current'].get(col, '')
                                        if prev_val != curr_val:
                                            f.write(f"    {col}: Previous='{prev_val}' | Current='{curr_val}'\n")
            # If no changes, indicate
            if not (added_flag or removed_flag or changes_flag):
                f.write("No differences found\n")

            f.write(f"\n{'='*80}\n")

    return output_file

if __name__ == "__main__":
    # Create Qt Application
    app = QApplication(sys.argv)

    # Open file dialog
    file_paths, _ = QFileDialog.getOpenFileNames(
        None,
        "Select Excel files to compare",
        "",
        "Excel Files (*.xlsx *.xls)"
    )

    if file_paths:
        # Sort the file_paths based on filename to ensure correct order
        # Assuming the filenames are structured to allow correct sorting (e.g., version numbers)
        file_paths_sorted = sorted(file_paths, key=lambda x: os.path.basename(x))

        try:
            output_file = compare_excel_files(file_paths_sorted)
            print(f"\nComparison report generated at:\n{output_file}")
            
            # Optional: Open the containing folder
            os.startfile(os.path.dirname(output_file))
        except Exception as e:
            print(f"Error: {str(e)}")
    else:
        print("No files selected")

    # Clean up the Qt Application
    app.quit()
