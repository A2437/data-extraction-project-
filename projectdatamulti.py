"""
faculty_extractor_multi_college.py
Enhanced faculty data extractor that creates SEPARATE Excel files for each college.
Converts experience from months to years automatically.

FEATURES:
- Creates individual Excel file for each college (53 colleges = 53 Excel files)
- Converts experience from months to years with proper calculation
- Maintains all existing functionality and error handling
- Opens all Excel files automatically after processing

Usage:
  - Put all PDFs in a folder
  - Edit PDF_FOLDER below
  - Run: python faculty_extractor_multi_college.py
"""

import os
import re
import sys
import pdfplumber
import pandas as pd
from pathlib import Path
import subprocess
from datetime import datetime
import math

# -----------------------------
# ROBUST CONFIG WITH FALLBACKS
# -----------------------------
PDF_FOLDER = r"C:\Users\Student\Downloads\AP&TS-NIRF-Rank Analysis\AP&TS-NIRF-Rank Analysis\2024"

# Multiple fallback locations for output
def get_safe_output_directory():
    """Get safe output directory with fallbacks."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Try multiple safe locations
    possible_directories = [
        # Desktop with timestamp folder
        os.path.join(os.path.expanduser("~"), "Desktop", f"Faculty_Data_{timestamp}"),
        # Documents
        os.path.join(os.path.expanduser("~"), "Documents", f"Faculty_Data_{timestamp}"),
        # Downloads
        os.path.join(os.path.expanduser("~"), "Downloads", f"Faculty_Data_{timestamp}"),
        # Current directory
        f"Faculty_Data_{timestamp}",
        # Temp directory
        os.path.join(os.environ.get('TEMP', '/tmp'), f"Faculty_Data_{timestamp}")
    ]
    
    return possible_directories

# Final columns with Experience in Years
FINAL_COLUMNS = [
    "S No", "Name", "Age", "Designation", "Gender", "Qualification",
    "Experience (in years)", "Currently working with institution?",
    "Joining Date", "Leaving Date", "Association type", "Institution name"
]

# -----------------------------
# SAFE UTILITY FUNCTIONS
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

def convert_months_to_years(months_text):
    """Convert experience from months to years with proper formatting."""
    try:
        if not months_text:
            return ""
        
        # Extract number from text
        months = safe_extract_number(months_text)
        if not months:
            return months_text  # Return original if no number found
        
        # Convert months to years
        if months < 12:
            return f"{months} months"  # Keep as months if less than a year
        else:
            years = months / 12
            if years == int(years):
                return f"{int(years)} years"  # Whole years
            else:
                return f"{years:.1f} years"  # Decimal years
                
    except Exception:
        return str(months_text) if months_text else ""

def is_potential_faculty_row(row):
    """Liberal check for potential faculty data."""
    try:
        if not row or len(row) < 2:
            return False
        
        # Check if first cell could be a serial number
        first_cell = str(row[0]).strip()
        if not first_cell or len(first_cell) > 10:
            return False
        
        # Must have some kind of number in first cell
        if not re.search(r"\d", first_cell):
            return False
        
        # Second cell should look like a name (not empty, has letters)
        second_cell = str(row[1]).strip()
        if not second_cell or len(second_cell) < 2:
            return False
        
        # Exclude obvious non-faculty content
        row_text = " ".join(str(cell) for cell in row[:3]).lower()
        exclude_terms = [
            "total", "percentage", "built", "area", "laboratory", 
            "playground", "establishment", "recognition", "accreditation"
        ]
        
        return not any(term in row_text for term in exclude_terms)
        
    except Exception:
        return False

# -----------------------------
# ENHANCED EXTRACTION FOR INDIVIDUAL COLLEGES
# -----------------------------
def extract_from_single_pdf(pdf_path):
    """Extract faculty data from single PDF for individual college."""
    records = []
    institution_name = pdf_path.stem
    
    print(f"📄 Processing: {pdf_path.name}")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
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
                        
                        # Process each row in the table
                        for row_idx, raw_row in enumerate(table):
                            if not raw_row:
                                continue
                            
                            # Normalize the row
                            norm_row = [safe_normalize_cell(cell) for cell in raw_row]
                            
                            # Check if this could be faculty data
                            if not is_potential_faculty_row(norm_row):
                                continue
                            
                            # Extract serial number
                            serial_num = safe_extract_number(norm_row[0])
                            if not serial_num or serial_num > 10000:
                                continue
                            
                            # Create record with years conversion
                            record = create_safe_record_with_years(norm_row, institution_name, serial_num)
                            if record:
                                records.append(record)
                
                except Exception as e:
                    print(f"   Error on page {page_num}: {e}")
                    continue
            
            print(f"   ✅ Extracted {len(records)} faculty records")
            
    except Exception as e:
        print(f"   ❌ Error processing {pdf_path.name}: {e}")
    
    return records, institution_name

def create_safe_record_with_years(row, institution_name, serial_num):
    """Create faculty record with experience converted to years."""
    try:
        record = {
            "S No": str(serial_num),
            "Name": str(row[1]).strip() if len(row) > 1 else "",
            "Age": "",
            "Designation": "",
            "Gender": "",
            "Qualification": "",
            "Experience (in years)": "",  # Changed to years
            "Currently working with institution?": "",
            "Joining Date": "",
            "Leaving Date": "",
            "Association type": "",
            "Institution name": institution_name,
        }
        
        # Fill in available data safely
        if len(row) > 2:
            experience_found = False
            
            # Try to determine what each column contains
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
                
                # Designation detection (contains common job titles)
                if not record["Designation"]:
                    designation_keywords = ["professor", "lecturer", "assistant", "associate", "principal", "hod", "dean"]
                    if any(keyword in cell_value.lower() for keyword in designation_keywords):
                        record["Designation"] = cell_value
                        continue
                
                # Experience detection and conversion to years
                if not record["Experience (in years)"] and not experience_found:
                    if ("month" in cell_value.lower() or "year" in cell_value.lower() or 
                        re.search(r"\d+", cell_value)):
                        # Convert months to years
                        converted_experience = convert_months_to_years(cell_value)
                        record["Experience (in years)"] = converted_experience
                        experience_found = True
                        continue
                
                # If nothing else matches, put it in qualification
                if not record["Qualification"] and len(cell_value) > 1:
                    record["Qualification"] = cell_value
        
        # Must have a name to be valid
        return record if record["Name"] else None
        
    except Exception:
        return None

# -----------------------------
# INDIVIDUAL COLLEGE EXCEL CREATION
# -----------------------------
def save_individual_college_excel(records, institution_name, output_directory):
    """Save individual college data to separate Excel file."""
    if not records:
        return None
    
    try:
        # Create DataFrame for this college
        df = pd.DataFrame(records, columns=FINAL_COLUMNS)
        
        # Remove duplicates for this college
        df = df.drop_duplicates(subset=["Name"], keep="first")
        
        # Sort by S No (as numbers for proper ordering)
        df["S No Numeric"] = pd.to_numeric(df["S No"], errors='coerce').fillna(999999)
        df = df.sort_values(["S No Numeric"])
        df = df.drop("S No Numeric", axis=1)
        df.reset_index(drop=True, inplace=True)
        df.insert(0, "Final Serial", range(1, len(df) + 1))
        
        # Create safe filename
        safe_filename = re.sub(r'[<>:"/\\|?*]', '_', institution_name)
        excel_filename = f"{safe_filename}_Faculty_Data.xlsx"
        excel_path = os.path.join(output_directory, excel_filename)
        
        # Save to Excel
        df.to_excel(excel_path, index=False, engine='openpyxl')
        
        return excel_path, len(df)
        
    except Exception as e:
        print(f"   ❌ Error saving {institution_name}: {e}")
        return None

# -----------------------------
# MULTI-COLLEGE PROCESSING
# -----------------------------
def process_all_pdfs_separately():
    """Process all PDFs and create separate Excel files for each college."""
    
    # Validate folder
    folder = Path(PDF_FOLDER)
    if not folder.exists():
        print(f"❌ Folder not found: {PDF_FOLDER}")
        
        # Try to find alternative locations
        alternative_paths = [
            r"C:\Users\Student\Downloads",
            r"C:\Users\Student\Desktop",
            r"C:\Users\Student\Documents",
            "."
        ]
        
        print("🔍 Searching for PDFs in alternative locations...")
        for alt_path in alternative_paths:
            if os.path.exists(alt_path):
                pdfs = list(Path(alt_path).glob("*.pdf"))
                if pdfs:
                    print(f"📁 Found {len(pdfs)} PDFs in: {alt_path}")
                    folder = Path(alt_path)
                    break
        
        if not folder.exists():
            print("💥 No PDFs found anywhere!")
            return False
    
    # Get all PDF files
    pdf_files = [p for p in folder.iterdir() if p.suffix.lower() == ".pdf"]
    
    if not pdf_files:
        print(f"❌ No PDF files found in: {folder}")
        return False
    
    # Create output directory
    output_directories = get_safe_output_directory()
    output_dir = None
    
    for directory in output_directories:
        try:
            os.makedirs(directory, exist_ok=True)
            output_dir = directory
            break
        except Exception:
            continue
    
    if not output_dir:
        print("❌ Could not create output directory!")
        return False
    
    print(f"🎯 Found {len(pdf_files)} PDF files to process")
    print(f"📁 Output directory: {output_dir}")
    print("=" * 60)
    
    # Process each PDF separately
    created_files = []
    total_faculty = 0
    
    for i, pdf_file in enumerate(sorted(pdf_files), 1):
        print(f"\n[{i:2d}/{len(pdf_files)}] {pdf_file.name}")
        
        try:
            # Extract data for this college
            records, institution_name = extract_from_single_pdf(pdf_file)
            
            if records:
                # Save individual Excel file
                result = save_individual_college_excel(records, institution_name, output_dir)
                
                if result:
                    excel_path, faculty_count = result
                    created_files.append(excel_path)
                    total_faculty += faculty_count
                    print(f"   ✅ Created: {os.path.basename(excel_path)} ({faculty_count} faculty)")
                else:
                    print(f"   ❌ Failed to save Excel for {institution_name}")
            else:
                print(f"   ⚠️  No faculty data found")
                
        except Exception as e:
            print(f"   ❌ Failed: {e}")
            continue
    
    print(f"\n" + "=" * 60)
    print(f"🎉 PROCESSING COMPLETE!")
    print(f"📊 Total Excel files created: {len(created_files)}")
    print(f"👥 Total faculty records: {total_faculty}")
    print(f"📁 All files saved in: {output_dir}")
    
    # Try to open all Excel files
    if created_files:
        print(f"\n🚀 Opening {len(created_files)} Excel files...")
        
        for i, excel_file in enumerate(created_files[:10]):  # Open max 10 files to avoid overwhelming
            try:
                if sys.platform == "win32":
                    subprocess.Popen(['start', excel_file], shell=True)
                else:
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", excel_file])
            except Exception:
                pass
        
        if len(created_files) > 10:
            print(f"   (Opening first 10 files, remaining {len(created_files) - 10} files are in the folder)")
        
        # Also open the output folder
        try:
            if sys.platform == "win32":
                subprocess.Popen(['explorer', output_dir])
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", output_dir])
            print(f"📂 Output folder opened!")
        except:
            pass
        
        return True
    else:
        print("❌ No Excel files were created!")
        return False

# -----------------------------
# SUMMARY REPORT CREATION
# -----------------------------
def create_summary_report(output_dir, created_files, total_faculty):
    """Create a summary report of all processed colleges."""
    try:
        summary_data = []
        
        for excel_file in created_files:
            try:
                # Read each Excel file to get summary info
                df = pd.read_excel(excel_file)
                college_name = os.path.basename(excel_file).replace('_Faculty_Data.xlsx', '')
                
                summary_data.append({
                    'College Name': college_name,
                    'Faculty Count': len(df) - 1,  # Subtract 1 for header
                    'Excel File': os.path.basename(excel_file)
                })
                
            except Exception:
                continue
        
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_path = os.path.join(output_dir, "00_SUMMARY_REPORT.xlsx")
            summary_df.to_excel(summary_path, index=False, engine='openpyxl')
            print(f"📋 Summary report created: {summary_path}")
            
    except Exception as e:
        print(f"⚠️  Could not create summary report: {e}")

# -----------------------------
# ENHANCED MAIN EXECUTION
# -----------------------------
if __name__ == "__main__":
    print("🎓 MULTI-COLLEGE FACULTY DATA EXTRACTOR")
    print("=" * 60)
    print("✅ Creates separate Excel file for each college")
    print("🔄 Converts experience from months to years")
    print("⚡ Opens all Excel files automatically")
    print("🛡️  Maximum error handling")
    print("=" * 60)
    
    try:
        success = process_all_pdfs_separately()
        
        if success:
            print("\n🎊 SUCCESS! Individual Excel files created for each college!")
            print("📱 Check the opened folder for all Excel files")
            print("💡 Each college now has its own dedicated Excel file")
        else:
            print("\n💥 Could not create Excel files")
            print("💡 Check folder path and file permissions")
        
        input("\n⏸️  Press Enter to exit...")
        
    except KeyboardInterrupt:
        print("\n🛑 Process interrupted by user")
        sys.exit(1)
        
    except Exception as e:
        print(f"\n💥 Unexpected error: {e}")
        print("📋 Error details saved to error_log.txt")
        
        try:
            with open("error_log.txt", "w") as f:
                f.write(f"Error: {e}\n")
                f.write(f"Error type: {type(e).__name__}\n")
                import traceback
                f.write(f"Traceback:\n{traceback.format_exc()}")
        except:
            pass
        
        input("⏸️  Press Enter to exit...")
        sys.exit(1)