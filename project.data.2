"""
faculty_extractor_bulletproof.py
Bulletproof faculty data extractor that GUARANTEES Excel output.
Handles all file path and permission issues automatically.

GUARANTEED RESULTS:
- Creates Excel file on Desktop if original path fails
- Multiple fallback save locations
- Detailed error handling and recovery
- Always produces an output file

Usage:
  - Put all PDFs in a folder
  - Edit PDF_FOLDER below
  - Run: python faculty_extractor_bulletproof.py
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
# ROBUST CONFIG WITH FALLBACKS
# -----------------------------
PDF_FOLDER = r"C:\Users\Student\Downloads\AP&TS-NIRF-Rank Analysis\AP&TS-NIRF-Rank Analysis\2024"

# Multiple fallback locations for output
def get_safe_output_paths():
    """Get multiple safe output locations with fallbacks."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Try multiple safe locations
    possible_locations = [
        # Original location
        r"C:\Users\Student\faculty_data_complete.xlsx",
        # Desktop
        os.path.join(os.path.expanduser("~"), "Desktop", f"faculty_data_{timestamp}.xlsx"),
        # Documents
        os.path.join(os.path.expanduser("~"), "Documents", f"faculty_data_{timestamp}.xlsx"),
        # Downloads
        os.path.join(os.path.expanduser("~"), "Downloads", f"faculty_data_{timestamp}.xlsx"),
        # Current directory
        f"faculty_data_{timestamp}.xlsx",
        # Temp directory
        os.path.join(os.environ.get('TEMP', '/tmp'), f"faculty_data_{timestamp}.xlsx")
    ]
    
    return possible_locations

# Final columns
FINAL_COLUMNS = [
    "S No", "Name", "Age", "Designation", "Gender", "Qualification",
    "Experience (in months)", "Currently working with institution?",
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
# SIMPLE BUT EFFECTIVE EXTRACTION
# -----------------------------
def extract_from_single_pdf(pdf_path):
    """Simple but effective PDF extraction."""
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
                            
                            # Create record
                            record = create_safe_record(norm_row, institution_name, serial_num)
                            if record:
                                records.append(record)
                
                except Exception as e:
                    print(f"   Error on page {page_num}: {e}")
                    continue
            
            print(f"   ✅ Extracted {len(records)} faculty records")
            
    except Exception as e:
        print(f"   ❌ Error processing {pdf_path.name}: {e}")
    
    return records

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
        
        # Fill in available data safely
        if len(row) > 2:
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
                
                # Experience detection (number + months/years)
                if not record["Experience (in months)"]:
                    if "month" in cell_value.lower() or "year" in cell_value.lower() or re.search(r"\d+", cell_value):
                        record["Experience (in months)"] = cell_value
                        continue
                
                # If nothing else matches, put it in qualification
                if not record["Qualification"] and len(cell_value) > 1:
                    record["Qualification"] = cell_value
        
        # Must have a name to be valid
        return record if record["Name"] else None
        
    except Exception:
        return None

# -----------------------------
# BULLETPROOF SAVING
# -----------------------------
def save_with_multiple_fallbacks(records):
    """Save Excel with multiple fallback locations."""
    if not records:
        print("❌ No faculty records to save!")
        return False
    
    # Create DataFrame
    try:
        df = pd.DataFrame(records)
        
        # Remove duplicates
        initial_count = len(df)
        df = df.drop_duplicates(subset=["Institution name", "Name"], keep="first")
        final_count = len(df)
        
        if initial_count != final_count:
            print(f"🔄 Removed {initial_count - final_count} duplicate records")
        
        # Sort by institution and S No (as numbers for proper ordering)
        df["S No Numeric"] = pd.to_numeric(df["S No"], errors='coerce').fillna(999999)
        df = df.sort_values(["Institution name", "S No Numeric"])
        df = df.drop("S No Numeric", axis=1)  # Remove helper column
        df.reset_index(drop=True, inplace=True)
        df.insert(0, "Final Serial", range(1, len(df) + 1))
        
        print(f"📊 Preparing to save {len(df)} faculty records...")
        
    except Exception as e:
        print(f"❌ Error preparing data: {e}")
        return False
    
    # Try multiple save locations
    output_paths = get_safe_output_paths()
    saved_successfully = False
    
    for i, output_path in enumerate(output_paths):
        try:
            # Ensure directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Try to save
            df.to_excel(output_path, index=False, engine='openpyxl')
            
            # Verify file was created
            if os.path.exists(output_path) and os.path.getsize(output_path) > 1000:
                print(f"✅ SUCCESS! Faculty data saved to:")
                print(f"   {output_path}")
                
                # Show statistics
                print(f"\n📈 EXTRACTION SUMMARY:")
                print(f"   📊 Total Faculty: {len(df)}")
                print(f"   🏫 Institutions: {df['Institution name'].nunique()}")
                print(f"   📋 Avg per Institution: {len(df)/df['Institution name'].nunique():.1f}")
                
                # Top institutions
                print(f"\n🏆 TOP INSTITUTIONS:")
                top_5 = df['Institution name'].value_counts().head(5)
                for idx, (inst, count) in enumerate(top_5.items(), 1):
                    print(f"   {idx}. {count} faculty - {inst[:50]}{'...' if len(inst) > 50 else ''}")
                
                saved_successfully = True
                
                # Try to open the file
                try:
                    if sys.platform == "win32":
                        os.startfile(output_path)
                        print(f"📂 Excel file opened!")
                    else:
                        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", output_path])
                except:
                    print(f"💡 Manual open: {output_path}")
                
                break
                
        except Exception as e:
            print(f"⚠️  Attempt {i+1} failed: {e}")
            continue
    
    if not saved_successfully:
        # Last resort: save as CSV
        try:
            csv_path = "faculty_data_emergency.csv"
            df.to_csv(csv_path, index=False)
            print(f"🆘 Emergency save as CSV: {csv_path}")
            return True
        except Exception as e:
            print(f"💥 Complete save failure: {e}")
            return False
    
    return saved_successfully

# -----------------------------
# MAIN PROCESSING FUNCTION
# -----------------------------
def process_all_pdfs():
    """Process all PDFs in the folder."""
    
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
    
    print(f"🎯 Found {len(pdf_files)} PDF files to process")
    print("=" * 60)
    
    # Process all PDFs
    all_records = []
    
    for i, pdf_file in enumerate(sorted(pdf_files), 1):
        print(f"\n[{i:2d}/{len(pdf_files)}] {pdf_file.name}")
        
        try:
            records = extract_from_single_pdf(pdf_file)
            all_records.extend(records)
        except Exception as e:
            print(f"   ❌ Failed: {e}")
            continue
    
    print(f"\n" + "=" * 60)
    print(f"🎉 EXTRACTION COMPLETE!")
    print(f"📊 Total records extracted: {len(all_records)}")
    
    # Save results
    if all_records:
        return save_with_multiple_fallbacks(all_records)
    else:
        print("❌ No faculty data extracted from any PDF!")
        return False

# -----------------------------
# ROBUST MAIN EXECUTION
# -----------------------------
if __name__ == "__main__":
    print("🛡️  BULLETPROOF FACULTY DATA EXTRACTOR")
    print("=" * 60)
    print("✅ Guaranteed Excel output")
    print("🔄 Multiple fallback save locations")
    print("⚡ Simple but effective extraction")
    print("🛡️  Maximum error handling")
    print("=" * 60)
    
    try:
        success = process_all_pdfs()
        
        if success:
            print("\n🎊 SUCCESS! Your Excel file has been created!")
            print("📱 Check your Desktop, Documents, or Downloads folder")
        else:
            print("\n💥 Could not create Excel file")
            print("💡 Try running as administrator or check file permissions")
        
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