#Build Text files with frequency, resistance and reactance extraction from each worksheet of excel and store by making folders of sheetname
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pathlib import Path
import re

# === Ask user for input and output locations via dialog boxes ===
root = tk.Tk()
root.withdraw()  # hide main window

# Prompt for input Excel file
input_path = filedialog.askopenfilename(
    title="Select the Excel file",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not input_path:
    print("❌ No input file selected. Exiting.")
    exit(1)

# Prompt for output directory
output_dir = filedialog.askdirectory(
    title="Select the output folder"
)
if not output_dir:
    print("❌ No output directory selected. Exiting.")
    exit(1)

INPUT_FILE = input_path
OUTPUT_ROOT = Path(output_dir)

# Ensure the output root exists
OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

# Load the workbook
xls = pd.ExcelFile(INPUT_FILE)

for sheet_name in xls.sheet_names:
    # Parse sheet
    df = xls.parse(sheet_name)
    
    # Find the frequency column (first column containing 'freq', case‐insensitive)
    freq_cols = [c for c in df.columns if re.search(r"freq", c, re.IGNORECASE)]
    if not freq_cols:
        print(f"⚠️  No frequency column found in sheet '{sheet_name}', skipping.")
        continue
    freq_col = freq_cols[0]
    
    # Identify all Primary columns
    primaries = [c for c in df.columns if re.match(r"Primary", c, re.IGNORECASE)]
    if not primaries:
        print(f"⚠️  No Primary columns in '{sheet_name}', skipping.")
        continue
    
    # Prepare sheet‐specific output folder
    sheet_dir = OUTPUT_ROOT / sheet_name
    sheet_dir.mkdir(parents=True, exist_ok=True)
    
    for primary in primaries:
        # find its column index and grab the *next* column as secondary
        idx = df.columns.get_loc(primary)
        if idx + 1 >= len(df.columns):
            continue
        secondary = df.columns[idx + 1]
        if not re.match(r"Secondary", secondary, re.IGNORECASE):
            # skip if the next column isn’t a Secondary
            continue
        
        # Subset down to the three columns
        out_df = df[[freq_col, primary, secondary]].copy()
        out_df.columns = ["Frequency", "Primary", "Secondary"]
        
        # Build a safe filename
        safe_name = re.sub(r"[^\w\-]+", "_", primary)
        out_file = sheet_dir / f"{safe_name}.txt"
        
        # Write tab-delimited
        out_df.to_csv(out_file, sep="\t", index=False)
        print(f"✅ Wrote {out_file}")

print("All done!")
