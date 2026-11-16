
import os
import gc
import pandas as pd
from concurrent.futures import ProcessPoolExecutor, as_completed
from multiprocessing import freeze_support

# =====================================================
# CONFIGURATION
# =====================================================

ROOT_DIR = r"D:\SalesHistory"

MAX_WORKERS = 2

ENGINE = "calamine"
CSV_ENCODING = "utf-8-sig"

SKIP_EXISTING = True

# =====================================================
# CONVERT ONE EXCEL FILE
# =====================================================

def convert_excel(file_path):

    csv_path = os.path.splitext(file_path)[0] + ".csv"
    temp_csv = csv_path + ".tmp"

    # Skip existing CSV
    if SKIP_EXISTING and os.path.exists(csv_path):
        return ("skipped", file_path, 0, "Already exists")

    try:
        print(f"[START] {file_path}", flush=True)

        total_rows = 0

        # Open workbook
        excel = pd.ExcelFile(
            file_path,
            engine=ENGINE
        )

        sheet_names = excel.sheet_names

        if not sheet_names:
            excel.close()
            open(temp_csv, "w", encoding=CSV_ENCODING).close()
            os.replace(temp_csv, csv_path)
            return ("success", file_path, 0, "No sheets")

        # Open CSV only once.
        # This also ensures the UTF-8 BOM is written only once.
        with open(
            temp_csv,
            "w",
            encoding=CSV_ENCODING,
            newline=""
        ) as output:

            for index, sheet_name in enumerate(sheet_names, 1):

                print(
                    f"[SHEET {index}/{len(sheet_names)}] "
                    f"{os.path.basename(file_path)}: {sheet_name}",
                    flush=True
                )

                # Read one sheet at a time
                df = pd.read_excel(
                    excel,
                    sheet_name=sheet_name
                )

                # Write headers only for the first sheet
                df.to_csv(
                    output,
                    index=False,
                    header=(index == 1),
                    lineterminator="\n"
                )

                total_rows += len(df)

                del df

        excel.close()

        # Publish CSV only after successful completion
        os.replace(temp_csv, csv_path)

        print(
            f"[DONE] {os.path.basename(file_path)} "
            f"| Rows: {total_rows:,}",
            flush=True
        )

        return ("success", file_path, total_rows, "")

    except Exception as e:

        # Remove incomplete temporary CSV
        try:
            if os.path.exists(temp_csv):
                os.remove(temp_csv)
        except OSError:
            pass

        return (
            "failed",
            file_path,
            0,
            f"{type(e).__name__}: {e}"
        )


# =====================================================
# FIND EXCEL FILES
# =====================================================

def find_excel_files(root_dir):

    excel_files = []

    ignored_dirs = {
        "$RECYCLE.BIN",
        "SYSTEM VOLUME INFORMATION"
    }

    def walk_error(error):
        print(f"[SCAN ERROR] {error}", flush=True)

    for subdir, dirs, files in os.walk(
        root_dir,
        onerror=walk_error
    ):

        dirs[:] = [
            d for d in dirs
            if d.upper() not in ignored_dirs
        ]

        for file in files:

            if (
                file.lower().endswith(".xlsx")
                and not file.startswith("~$")
            ):

                file_path = os.path.join(subdir, file)

                if SKIP_EXISTING:
                    csv_path = os.path.splitext(file_path)[0] + ".csv"

                    if os.path.exists(csv_path):
                        continue

                excel_files.append(file_path)

    return excel_files


# =====================================================
# MAIN PARALLEL PROCESSING
# =====================================================

def main():

    print("=" * 60)
    print("HIGH PERFORMANCE EXCEL TO CSV CONVERTER")
    print("=" * 60)

    print("\nScanning folders...", flush=True)

    files = find_excel_files(ROOT_DIR)

    total_files = len(files)

    print(f"Files to convert: {total_files}", flush=True)
    print(f"Parallel workers: {MAX_WORKERS}", flush=True)

    if total_files == 0:
        print("No Excel files found.")
        return

    converted = 0
    skipped = 0
    failed = 0
    total_rows = 0

    # Parallel processing
    with ProcessPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:

        futures = {
            executor.submit(convert_excel, file): file
            for file in files
        }

        for future in as_completed(futures):

            try:
                status, file_path, rows, error = future.result()

                if status == "success":
                    converted += 1
                    total_rows += rows

                elif status == "skipped":
                    skipped += 1

                else:
                    failed += 1
                    print(
                        f"\n[ERROR] {file_path}\n{error}",
                        flush=True
                    )

            except Exception as e:
                failed += 1
                print(f"[WORKER ERROR] {e}", flush=True)

            completed = converted + skipped + failed

            print(
                f"\nPROGRESS: {completed}/{total_files} "
                f"| Converted: {converted} "
                f"| Failed: {failed} "
                f"| Skipped: {skipped}",
                flush=True
            )

    print("\n" + "=" * 60)
    print("ALL CONVERSIONS COMPLETED")
    print("=" * 60)

    print(f"Total files      : {total_files}")
    print(f"Converted       : {converted}")
    print(f"Skipped         : {skipped}")
    print(f"Failed          : {failed}")
    print(f"Total rows      : {total_rows:,}")

    print("=" * 60)


# =====================================================
# ENTRY POINT (IMPORTANT FOR WINDOWS EXE)
# =====================================================

if __name__ == "__main__":

    freeze_support()

    main()