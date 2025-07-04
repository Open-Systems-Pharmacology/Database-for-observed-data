import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import sys
import pandas as pd

def compare_excel(file1, file2):
    xl1 = pd.ExcelFile(file1)
    xl2 = pd.ExcelFile(file2)
    diff_report = []

    sheets1 = set(xl1.sheet_names)
    sheets2 = set(xl2.sheet_names)
    all_sheets = sheets1.union(sheets2)

    for sheet in all_sheets:
        if sheet not in sheets1:
            diff_report.append(f"Sheet '{sheet}' only in {file2}")
            continue
        if sheet not in sheets2:
            diff_report.append(f"Sheet '{sheet}' only in {file1}")
            continue
        df1 = xl1.parse(sheet)
        df2 = xl2.parse(sheet)
        if not df1.equals(df2):
            diff_report.append(f"Changes in sheet '{sheet}':")
            diff = df1.compare(df2, keep_shape=True, keep_equal=False, result_names=("OLD", "NEW"))
            diff = diff.dropna(how='all')  # Only keep rows with actual differences
            diff = diff.dropna(axis=1, how='all')# Only columns with differences
            diff = diff.fillna('')               # Replace NaN with empty string
            diff.index = diff.index + 2  # Shift index to match Excel row numbers
            diff.index.name = "EXCEL ROW"    # Set index header
            diff_report.append(diff.to_markdown())
    return "\n\n".join(diff_report) if diff_report else "No differences found."

if __name__ == "__main__":
    file1, file2, out = sys.argv[1], sys.argv[2], sys.argv[3]
    result = compare_excel(file1, file2)
    with open(out, "w") as f:
        f.write(result)