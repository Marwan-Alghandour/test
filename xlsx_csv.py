import os
import pandas as pd

# Replace with the full path to your main parent folder containing the 12 folders
root_dir = r"D:\SalesHistory"

converted_count = 0

for subdir, dirs, files in os.walk(root_dir):
    for file in files:
        # Target .xlsx files and ignore temporary Excel lock files (~$)
        if file.endswith(".xlsx") and not file.startswith("~$"):
            file_path = os.path.join(subdir, file)

            try:
                # Read all sheets from Excel file
                sheets = pd.read_excel(file_path, sheet_name=None)

                # Merge all sheets into one DataFrame
                df = pd.concat(sheets.values(), ignore_index=True)

                # Define the output CSV path (same name, .csv extension)
                csv_path = os.path.splitext(file_path)[0] + ".csv"

                # Save to CSV (utf-8-sig handles special characters correctly)
                df.to_csv(csv_path, index=False, encoding="utf-8-sig")

                converted_count += 1
                print(f"Converted: {file}")
            except Exception as e:
                print(f"Error converting {file}: {e}")

print(f"\nDone! Successfully converted {converted_count} files.")