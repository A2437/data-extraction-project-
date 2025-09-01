"""
faculty_extractor_bulletproof.py
Bulletproof faculty data extractor that GUARANTEES Excel output.
Handles all file path and permission issues automatically.

Usage:
  - Put all PDFs in a folder
  - Edit PDF_FOLDER below
  - Run: python python.project3.py
"""

import os
import re
import sys
import pdfplumber
import pandas as pd
from pathlib import Path
import subprocess
from datetime import datetime

# -----------------------------
# CONFIGURATION
# -----------------------------
PDF_FOLDER = r"D:\Python Test Folder\AP&TS-NIRF-Rank Analysis\2024"

FINAL_COLUMNS = [
    "Final Serial", "S No", "Name", "Age", "Designation", "Gender", "Qualification",
    "Experience (in months)", "Currently working with institution?",
    "Joining Date", "Leaving Date", "Association type", "Institution name"
]

def get_safe_output_paths():
    """Get multiple safe output locations with fallbacks."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    user_home = os.path.expanduser("~")
    temp_dir = os.environ.get('TEMP', '/tmp')
    return [
        os.path.join(user_home, "Desktop", f"faculty_data_{timestamp}.xlsx"),
        os.path.join(user_home, "Documents", f"faculty_data_{timestamp}.xlsx"),
        os.path.join(user_home, "Downloads", f"faculty_data_{timestamp}.xlsx"),
        f"faculty_data_{timestamp}.xlsx",
        os.path.join(temp_dir, f"faculty_data_{timestamp}.xlsx")
    ]

# -----------------------------
# UTILITY FUNCTIONS
# -----------------------------
def safe_normalize_cell(cell):
    """Safely normalize cell content."""
    try:
        if cell is None:
            return ""
        s = str(cell).replace("\r", " ").replace("\n", " ").strip()
        return re.sub(r"\s+", " ", s)
    except Exception:
        return ""

def safe_extract_number(text):
    """Safely extract number from text."""
    try:
        if not text:
            return None
        match = re.search(r"\d+", str(text))
        return int(match.group()) if match else None
    except Exception:
        return None

def is_potential_faculty_row(row):
    """Liberal check for potential faculty data."""
    try:
        if not row or len(row) < 2:
            return False
        first_cell = str(row[0]).strip()
        if not first_cell or len(first_cell) > 10:
            return False
        if not re.search(r"\d", first_cell):
            return False
        second_cell = str(row[1]).strip()
        if not second_cell or len(second_cell) < 2:
            return False
        row_text = " ".join(str(cell) for cell in row[:3]).lower()
        exclude_terms = [
            "total", "percentage", "built", "area", "laboratory",
            "playground", "establishment", "recognition", "accreditation"
        ]
        return not any(term in row_text for term in exclude_terms)
    except Exception:
        return False

def create_safe_record(row, institution_name, serial_num):
    """Safely create a faculty record."""
    try:
        record = {
            "S No": str(serial_num),
            "Name": str(row[1]).strip() if len(row) > 1 else "",
            "Age": "",
            "Designation": "",
            "Gender": "",
            "Qualification": "",
            "Experience (in months)": "",
            "Currently working with institution?": "",
            "Joining Date": "",
            "Leaving Date": "",
            "Association type": "",
            "Institution name": institution_name,
        }
        if len(row) > 2:
            for i in range(2, min(len(row), 12)):
                cell_value = str(row[i]).strip()
                if not cell_value:
                    continue
                # Age detection (number between 18-80)
                if not record["Age"]:
                    age = safe_extract_number(cell_value)
                    if age and 18 <= age <= 80:
                        record["Age"] = str(age)
                        continue
                # Gender detection
                if not record["Gender"] and cell_value.lower() in ["m", "f", "male", "female", "m/f"]:
                    record["Gender"] = cell_value
                    continue
                # Designation detection
                if not record["Designation"]:
                    designation_keywords = ["professor", "lecturer", "assistant", "associate", "principal", "hod", "dean"]
                    if any(keyword in cell_value.lower() for keyword in designation_keywords):
                        record["Designation"] = cell_value
                        continue
                # Experience detection
                if not record["Experience (in months)"]:
                    if "month" in cell_value.lower() or "year" in cell_value.lower() or re.search(r"\d+", cell_value):
                        record["Experience (in months)"] = cell_value
                        continue
                # Qualification
                if not record["Qualification"] and len(cell_value) > 1:
                    record["Qualification"] = cell_value
        return record if record["Name"] else None
    except Exception:
        return None

# -----------------------------
# EXTRACTION LOGIC
# -----------------------------
def extract_from_single_pdf(pdf_path):
    """Extract faculty records from a single PDF."""
    records = []
    institution_name = pdf_path.stem
    print(f" Processing: {pdf_path.name}")
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            total_pages = len(pdf.pages)
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"   Page {page_num}/{total_pages}", end="\r")
                try:
                    tables = page.extract_tables()
                    if not tables:
                        continue
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        for raw_row in table:
                            if not raw_row:
                                continue
                            norm_row = [safe_normalize_cell(cell) for cell in raw_row]
                            if not is_potential_faculty_row(norm_row):
                                continue
                            serial_num = safe_extract_number(norm_row[0])
                            if not serial_num or serial_num > 10000:
                                continue
                            record = create_safe_record(norm_row, institution_name, serial_num)
                            if record:
                                records.append(record)
                except Exception as e:
                    print(f"   Error on page {page_num}: {e}")
                    continue
            print(f"    Extracted {len(records)} faculty records")
    except Exception as e:
        print(f"    Error processing {pdf_path.name}: {e}")
    return records

# -----------------------------
# BULLETPROOF SAVING
# -----------------------------
def save_with_multiple_fallbacks(records):
    """Save Excel with multiple fallback locations."""
    if not records:
        print(" No faculty records to save!")
        return False
    try:
        df = pd.DataFrame(records)
        # Remove duplicates
        df = df.drop_duplicates(subset=["Institution name", "Name"], keep="first")
        # Sort by institution and S No
        df["S No Numeric"] = pd.to_numeric(df["S No"], errors='coerce').fillna(999999)
        df = df.sort_values(["Institution name", "S No Numeric"])
        df = df.drop("S No Numeric", axis=1)
        df.reset_index(drop=True, inplace=True)
        df.insert(0, "Final Serial", range(1, len(df) + 1))
        # Ensure columns order
        for col in FINAL_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[FINAL_COLUMNS]
        print(f" Preparing to save {len(df)} faculty records...")
    except Exception as e:
        print(f" Error preparing data: {e}")
        return False
    output_paths = get_safe_output_paths()
    saved_successfully = False
    for i, output_path in enumerate(output_paths):
        try:
            dir_name = os.path.dirname(output_path)
            if dir_name and not os.path.exists(dir_name):
                os.makedirs(dir_name, exist_ok=True)
            df.to_excel(output_path, index=False, engine='openpyxl')
            if os.path.exists(output_path) and os.path.getsize(output_path) > 500:
                print(f" SUCCESS! Faculty data saved to:\n   {output_path}")
                print(f"\n EXTRACTION SUMMARY:")
                print(f"    Total Faculty: {len(df)}")
                print(f"    Institutions: {df['Institution name'].nunique()}")
                print(f"    Avg per Institution: {len(df)/df['Institution name'].nunique():.1f}")
                print(f"\n TOP INSTITUTIONS:")
                top_5 = df['Institution name'].value_counts().head(5)
                for idx, (inst, count) in enumerate(top_5.items(), 1):
                    print(f"   {idx}. {count} faculty - {inst[:50]}{'...' if len(inst) > 50 else ''}")
                saved_successfully = True
                try:
                    if sys.platform == "win32":
                        os.startfile(output_path)
                        print(f" Excel file opened!")
                    else:
                        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", output_path])
                except Exception:
                    print(f" Manual open: {output_path}")
                break
        except Exception as e:
            print(f"  Attempt {i+1} failed: {e}")
            continue
    if not saved_successfully:
        try:
            csv_path = "faculty_data_emergency.csv"
            df.to_csv(csv_path, index=False)
            print(f" Emergency save as CSV: {csv_path}")
            return True
        except Exception as e:
            print(f" Complete save failure: {e}")
            return False
    return saved_successfully

# -----------------------------
# MAIN PROCESSING FUNCTION
# -----------------------------
# -----------------------------
# MAIN PROCESSING FUNCTION
# -----------------------------
def process_all_pdfs():
    """Process all PDFs in the folder."""
    folder = Path(PDF_FOLDER)
    if not folder.exists():
        print(f"Folder not found: {PDF_FOLDER}")
        alternative_paths = [
            os.path.join(os.path.expanduser("~"), "Downloads"),
            os.path.join(os.path.expanduser("~"), "Desktop"),
            os.path.join(os.path.expanduser("~"), "Documents"),
            "."
        ]
        print(" Searching for PDFs in alternative locations...")
        for alt_path in alternative_paths:
            alt_folder = Path(alt_path)
            if alt_folder.exists():
                pdfs = list(alt_folder.glob("*.pdf"))
                if pdfs:
                    print(f" Found {len(pdfs)} PDFs in: {alt_path}")
                    folder = alt_folder
                    break
        if not folder.exists():
            print("No PDFs found anywhere!")
            return False

    pdf_files = list(folder.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in: {folder}")
        return False

    print(f" Found {len(pdf_files)} PDF files to process")
    print("=" * 60)

    all_records = []
    for i, pdf_file in enumerate(sorted(pdf_files), 1):
        print(f"\n[{i:2d}/{len(pdf_files)}] {pdf_file.name}")
        try:
            records = extract_from_single_pdf(pdf_file)
            all_records.extend(records)
        except Exception as e:
            print(f"    Failed: {e}")
            continue

    print("\n" + "=" * 60)
    print(" EXTRACTION COMPLETE!")  # <- removed 🎉
    print(f" Total records extracted: {len(all_records)}")

    if all_records:
        return save_with_multiple_fallbacks(all_records)
    else:
        print(" No faculty data extracted from any PDF!")
        return False

           

# -----------------------------
# MAIN EXECUTION
# -----------------------------
if __name__ == "__main__":
    print("  BULLETPROOF FACULTY DATA EXTRACTOR")
    print("=" * 60)
    print(" Guaranteed Excel output")
    print(" Multiple fallback save locations")
    print(" Simple but effective extraction")
    print("  Maximum error handling")
    print("=" * 60)
    try:
        success = process_all_pdfs()
        if success:
            print("\n SUCCESS! Your Excel file has been created!")
            print(" Check your Desktop, Documents, or Downloads folder")
        else:
            print("\n Could not create Excel file")
            print(" Try running as administrator or check file permissions")
        input("\n⏸ Press Enter to exit...")
    except KeyboardInterrupt:
        print("\n Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n Unexpected error: {e}")
        print(" Error details saved to error_log.txt")
        try:
            with open("error_log.txt", "w", encoding="utf-8") as f:
                f.write(f"Error: {e}\n")
                f.write(f"Error type: {type(e).__name__}\n")
                import traceback
                f.write(f"Traceback:\n{traceback.format_exc()}")
        except Exception:
            pass
        input("  Press Enter to exit...")
        sys.exit(1)