#Multiply reactance value by -1 or convert resistance and reactance from kOhm to Ohm


import pandas as pd
import tkinter as tk
from tkinter import filedialog


def process_sheets(sheets, primary_factor, secondary_factor):
    """
    Multiply any column with 'Primary' in its name by primary_factor
    and any column with 'Secondary' in its name by secondary_factor.
    """
    for name, df in sheets.items():
        # Clean column names
        df.columns = df.columns.str.strip()
        # Apply factors to matching columns
        for col in df.columns:
            if 'Primary' in col:
                df[col] = df[col] * primary_factor
            if 'Freq'in col:
                df[col] = df[col] * secondary_factor
            if 'Secondary' in col:
                df[col] = df[col] * secondary_factor
        sheets[name] = df
    return sheets


def main():
    # 1) Ask user to pick input and output files via dialog boxes
    root = tk.Tk()
    root.withdraw()

    input_file = filedialog.askopenfilename(
        title="Select the input Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not input_file:
        print("❌ No input file selected. Exiting.")
        return

    output_file = filedialog.asksaveasfilename(
        title="Save converted Excel file as",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not output_file:
        print("❌ No output location selected. Exiting.")
        return

    # 2) Load all sheets from the selected workbook
    sheets = pd.read_excel(input_file, sheet_name=None)

    # 3) Prompt user for the desired operation
    print("\nChoose an operation:")
    print("  1) Multiply Primary columns by 1000")
    print("  2) Multiply both Primary and Secondary columns by 1000")
    print("  3) Multiply Secondary columns by -1")
    print("  4) Multiply Primary by 1000 and Secondary by -1000 (original both-at-once)")
    choice = input("Enter 1, 2, 3 or 4: ").strip()

    # Determine the multiplication factors based on choice
    if choice == '1':
        pf, sf = 1000, 1
    elif choice == '2':
        pf, sf = 1000, 1000
    elif choice == '3':
        pf, sf = 1, -1
    else:
        pf, sf = 1000, -1000

    # 4) Apply transformations
    sheets = process_sheets(sheets, pf, sf)

    # 5) Save the processed sheets with retry on PermissionError
    while True:
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for name, df in sheets.items():
                    df.to_excel(writer, sheet_name=name, index=False)
            break
        except PermissionError:
            print(f"❌ Permission denied when writing to {output_file}. It may be open in another program.")
            retry = input("Close the file and press Enter to retry, or type 'new' to choose a different output path: ").strip().lower()
            if retry == 'new':
                output_file = filedialog.asksaveasfilename(
                    title="Save converted Excel file as",
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")]
                )
                if not output_file:
                    print("❌ No output location selected. Exiting.")
                    return

    # 6) Display a preview of each sheet in the terminal
    print(f"\n✅ Converted data saved to: {output_file}\n")
    for name, df in sheets.items():
        print(f"--- Sheet: {name} ---")
        print(df.head().to_string(index=False))
        print()


if __name__ == "__main__":
    main()
