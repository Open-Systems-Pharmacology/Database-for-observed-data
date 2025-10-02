import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import sys
import pandas as pd

def _align_frames(df1: pd.DataFrame, df2: pd.DataFrame):
    # Preserve original left-to-right ordering by taking columns from df1 then any new ones from df2
    ordered_union = list(df1.columns) + [c for c in df2.columns if c not in df1.columns]
    all_cols = ordered_union
    df1 = df1.reindex(columns=all_cols)
    df2 = df2.reindex(columns=all_cols)

    # Standardize dtypes where they differ (convert both to string/object for safety)
    for col in all_cols:
        if col in df1.columns and col in df2.columns:
            if df1[col].dtype != df2[col].dtype:
                df1[col] = df1[col].astype(str)
                df2[col] = df2[col].astype(str)

    # Align row count
    max_len = max(len(df1), len(df2))
    df1 = df1.reindex(range(max_len))
    df2 = df2.reindex(range(max_len))
    return df1, df2

def compare_excel(file1, file2):
    xl1 = pd.ExcelFile(file1)
    xl2 = pd.ExcelFile(file2)
    diff_report = []

    sheets1 = set(xl1.sheet_names)
    sheets2 = set(xl2.sheet_names)
    all_sheets = sorted(sheets1.union(sheets2))

    for sheet in all_sheets:
        if sheet not in sheets1:
            diff_report.append(f"Sheet '{sheet}' only in {file2}")
            continue
        if sheet not in sheets2:
            diff_report.append(f"Sheet '{sheet}' only in {file1}")
            continue

        df1 = xl1.parse(sheet)
        df2 = xl2.parse(sheet)

        # Quick equality check first (fast path)
        if df1.equals(df2):
            continue

        # Track structural differences
        cols1 = set(df1.columns)
        cols2 = set(df2.columns)
        added_cols = sorted(list(cols2 - cols1))
        removed_cols = sorted(list(cols1 - cols2))
        row_diff_note = ""
        if len(df1) != len(df2):
            row_diff_note = f" (row count: {len(df1)} -> {len(df2)})"

        diff_report.append(f"Changes in sheet '{sheet}':")
        if added_cols:
            diff_report.append(f"  Columns only in NEW: {', '.join(added_cols)}")
        if removed_cols:
            diff_report.append(f"  Columns only in OLD: {', '.join(removed_cols)}")
        if row_diff_note:
            diff_report.append(f"  Row count changed{row_diff_note}")

        # Align frames so DataFrame.compare will not raise
        aligned_df1, aligned_df2 = _align_frames(df1, df2)

        # Perform cell-level comparison
        diff = aligned_df1.compare(
            aligned_df2,
            keep_shape=True,
            keep_equal=False,
            result_names=("OLD", "NEW")
        )

        # Remove rows & columns that are entirely empty after comparison
        diff = diff.dropna(how='all')
        diff = diff.dropna(axis=1, how='all')

        if diff.empty:
            # Structural change only (e.g., columns added/removed) but no overlapping cell value changes.
            diff_report.append("  (No individual cell value changes; only structural differences.)")
            continue

        # Make output friendlier
        diff = diff.fillna('')
        # Convert zero-based index to Excel 1-based + header row (assuming original header row = row 1)
        diff.index = diff.index + 2
        diff.index.name = "EXCEL ROW"
        diff_report.append(diff.to_markdown())

    return "\n\n".join(diff_report) if diff_report else "No differences found."

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python compare_excels.py <old.xlsx> <new.xlsx> <out_report.txt>")
        sys.exit(1)
    file1, file2, out = sys.argv[1], sys.argv[2], sys.argv[3]
    result = compare_excel(file1, file2)
    with open(out, "w", encoding="utf-8") as f:
        f.write(result)