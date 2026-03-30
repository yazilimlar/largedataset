import os
import re
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

# ============================================================
# Excel Folder to CSV Converter
# - User selects source folder
# - Script reads all Excel files in that folder
# - Each sheet is exported as a separate CSV
# - User selects destination folder
# ============================================================

SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xls")


def clean_name(name: str) -> str:
    """Make a safe filename."""
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name


def choose_folder(title: str) -> str:
    """Open folder picker dialog."""
    folder = filedialog.askdirectory(title=title)
    return folder.strip() if folder else ""


def choose_yes_no(title: str, message: str) -> bool:
    return messagebox.askyesno(title, message)


def get_engine(file_path: str) -> str | None:
    """
    Pick a pandas engine based on file extension.
    .xlsx / .xlsm -> openpyxl
    .xls -> xlrd (requires xlrd installed)
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return "openpyxl"
    elif ext == ".xls":
        return "xlrd"
    return None


def export_excel_to_csvs(source_folder: str, dest_folder: str, overwrite: bool = False):
    excel_files = [
        f for f in os.listdir(source_folder)
        if os.path.isfile(os.path.join(source_folder, f))
        and f.lower().endswith(SUPPORTED_EXTENSIONS)
        and not f.startswith("~$")
    ]

    if not excel_files:
        messagebox.showinfo("No Files Found", "No Excel files were found in the selected folder.")
        return

    total_files = 0
    total_sheets = 0
    success_exports = 0
    errors = []

    for file_name in excel_files:
        file_path = os.path.join(source_folder, file_name)
        base_name = clean_name(os.path.splitext(file_name)[0])
        engine = get_engine(file_path)

        try:
            # Load workbook metadata first
            xl = pd.ExcelFile(file_path, engine=engine)
            sheet_names = xl.sheet_names
            total_files += 1

            for sheet_name in sheet_names:
                total_sheets += 1
                safe_sheet = clean_name(sheet_name)
                output_name = f"{base_name}__{safe_sheet}.csv"
                output_path = os.path.join(dest_folder, output_name)

                if os.path.exists(output_path) and not overwrite:
                    errors.append(f"Skipped existing file: {output_name}")
                    continue

                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
                    df.to_csv(output_path, index=False, encoding="utf-8-sig")
                    success_exports += 1
                except Exception as e:
                    errors.append(
                        f"Failed sheet '{sheet_name}' in file '{file_name}': {str(e)}"
                    )

        except Exception as e:
            errors.append(f"Failed workbook '{file_name}': {str(e)}")

    summary = (
        f"Completed.\n\n"
        f"Excel files found: {len(excel_files)}\n"
        f"Workbooks processed: {total_files}\n"
        f"Sheets found: {total_sheets}\n"
        f"CSV files created: {success_exports}\n"
        f"Issues: {len(errors)}"
    )

    if errors:
        log_path = os.path.join(dest_folder, "excel_to_csv_errors.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(errors))
        summary += f"\n\nError log saved to:\n{log_path}"

    messagebox.showinfo("Finished", summary)


def main():
    root = tk.Tk()
    root.withdraw()
    root.update()

    messagebox.showinfo(
        "Excel Folder to CSV Converter",
        "You will be asked to choose:\n\n"
        "1. The folder containing Excel files\n"
        "2. The folder where CSV files should be saved"
    )

    source_folder = choose_folder("Select folder containing Excel files")
    if not source_folder:
        messagebox.showwarning("Cancelled", "No source folder selected.")
        return

    dest_folder = choose_folder("Select folder where CSV files should be saved")
    if not dest_folder:
        messagebox.showwarning("Cancelled", "No destination folder selected.")
        return

    overwrite = choose_yes_no(
        "Overwrite Existing Files?",
        "If a CSV with the same name already exists,\n"
        "do you want to overwrite it?"
    )

    try:
        export_excel_to_csvs(source_folder, dest_folder, overwrite=overwrite)
    except Exception:
        traceback_text = traceback.format_exc()
        messagebox.showerror("Unexpected Error", traceback_text)


if __name__ == "__main__":
    main()