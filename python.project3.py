"""
faculty_extractor_streamlined.py
Streamlined and efficient faculty data extractor with guaranteed Excel output.
Simplified logic focused on reliability and performance.

Usage:
  - Update PDF_FOLDER path below
  - Run: python faculty_extractor_streamlined.py
"""

import os
import re
import sys
import pandas as pd
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

try:
    import pdfplumber
except ImportError:
    print("ERROR: Install required packages with: pip install pdfplumber pandas openpyxl")
    input("Press Enter to exit...")
    sys.exit(1)

# -----------------------------
# CONFIGURATION
# -----------------------------
PDF_FOLDER = r"D:\Python Test Folder\AP&TS-NIRF-Rank Analysis\2024"

# -----------------------------
# CORE FUNCTIONS
# -----------------------------
def clean_cell(cell):
    """Clean cell content efficiently."""
    if not cell:
        return ""
    try:
        # Convert to string and clean
        text = str(cell).strip()
        text = re.sub(r'[\r\n]+', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        # Remove non-ASCII characters that cause issues
        text = ''.join(char for char in text if ord(char) < 128)
        return text.strip()
    except:
        return ""

def get_number(text):
    """Extract number from text."""
    try:
        match = re.search(r'\d+', str(text))
        return int(match.group()) if match else None
    except:
        return None

def is_data_row(row):
    """Check if row contains faculty data."""
    if len(row) < 2:
        return False
    
    # First cell must contain a reasonable serial number
    sno = get_number(row[0])
    if not sno or sno > 5000:
        return False
    
    # Second cell must be a name
    name = clean_cell(row[1])
    if not name or len(name) < 2:
        return False
    
    # Must have alphabetic characters
    if not any(c.isalpha() for c in name):
        return False
    
    return True

def extract_from_pdf(pdf_path):
    """Extract faculty data from PDF."""
    records = []
    institution = clean_cell(pdf_path.stem)
    
    print(f"Processing: {pdf_path.name}")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_faculty = 0
            
            for page_num, page in enumerate(pdf.pages, 1):
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
                            
                            # Clean the row
                            clean_row = [clean_cell(cell) for cell in raw_row]
                            
                            # Check if this is faculty data
                            if not is_data_row(clean_row):
                                continue
                            
                            # Create faculty record
                            sno = get_number(clean_row[0])
                            record = {
                                "S No": str(sno),
                                "Name": clean_row[1] if len(clean_row) > 1 else "",
                                "Age": clean_row[2] if len(clean_row) > 2 else "",
                                "Designation": clean_row[3] if len(clean_row) > 3 else "",
                                "Gender": clean_row[4] if len(clean_row) > 4 else "",
                                "Qualification": clean_row[5] if len(clean_row) > 5 else "",
                                "Experience (in months)": clean_row[6] if len(clean_row) > 6 else "",
                                "Currently working with institution?": clean_row[7] if len(clean_row) > 7 else "",
                                "Joining Date": clean_row[8] if len(clean_row) > 8 else "",
                                "Leaving Date": clean_row[9] if len(clean_row) > 9 else "",
                                "Association type": clean_row[10] if len(clean_row) > 10 else "",
                                "Institution name": institution,
                            }
                            
                            records.append(record)
                            total_faculty += 1
                
                except Exception:
                    continue
            
            print(f"  -> {total_faculty} faculty records extracted")
            
    except Exception as e:
        print(f"  -> Error: {e}")
    
    return records

def save_results(all_records):
    """Save results to Excel with guaranteed output."""
    if not all_records:
        print("No data to save!")
        return False
    
    print(f"\nPreparing {len(all_records)} records...")
    
    try:
        # Create DataFrame
        df = pd.DataFrame(all_records)
        
        # Remove duplicates
        initial_count = len(df)
        df = df.drop_duplicates(subset=["Institution name", "Name"])
        print(f"After deduplication: {len(df)} records")
        
        # PERFECT SORTING
        df["sort_sno"] = pd.to_numeric(df["S No"], errors='coerce').fillna(9999)
        df = df.sort_values(["Institution name", "sort_sno"])
        df = df.drop("sort_sno", axis=1)
        df.reset_index(drop=True, inplace=True)
        
        # Add final serial
        df.insert(0, "Final Serial", range(1, len(df) + 1))
        
        # Define output paths
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        
        excel_file = os.path.join(desktop, f"faculty_data_{timestamp}.xlsx")
        csv_file = os.path.join(desktop, f"faculty_data_{timestamp}.csv")
        
        # Try Excel first
        try:
            print(f"Saving Excel to: {excel_file}")
            df.to_excel(excel_file, index=False)
            
            if os.path.exists(excel_file) and os.path.getsize(excel_file) > 0:
                print(f"SUCCESS! Excel file created: {excel_file}")
                
                # Show sample of data
                print(f"\nSAMPLE DATA (first 3 rows):")
                print(df[["Final Serial", "S No", "Name", "Institution name"]].head(3).to_string(index=False))
                
                # Show statistics
                print(f"\nSTATISTICS:")
                print(f"Total Faculty: {len(df)}")
                print(f"Institutions: {df['Institution name'].nunique()}")
                
                # Top institutions
                print(f"\nTOP INSTITUTIONS:")
                top_inst = df['Institution name'].value_counts().head(3)
                for inst, count in top_inst.items():
                    print(f"  {count} faculty - {inst[:50]}")
                
                # Try to open
                try:
                    os.startfile(excel_file)
                    print("\nExcel file opened!")
                except:
                    pass
                
                return True
            
        except Exception as e:
            print(f"Excel failed: {e}")
        
        # Try CSV as backup
        try:
            print(f"Trying CSV: {csv_file}")
            df.to_csv(csv_file, index=False)
            print(f"CSV saved: {csv_file}")
            return True
        except Exception as e:
            print(f"CSV failed: {e}")
            
    except Exception as e:
        print(f"Data processing failed: {e}")
    
    return False

# -----------------------------
# MAIN EXECUTION
# -----------------------------
def run_extraction():
    """Run the complete extraction process."""
    print("FACULTY DATA EXTRACTOR")
    print("=" * 30)
    
    # Validate inputs
    if not os.path.exists(PDF_FOLDER):
        print(f"Folder not found: {PDF_FOLDER}")
        print("Please update PDF_FOLDER path in the script")
        return
    
    # Get PDF files
    try:
        pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith('.pdf')]
    except Exception as e:
        print(f"Cannot access folder: {e}")
        return
    
    if not pdf_files:
        print("No PDF files found in folder!")
        return
    
    print(f"Processing {len(pdf_files)} PDF files...")
    print("-" * 30)
    
    # Extract data from all PDFs
    all_records = []
    
    for i, pdf_filename in enumerate(pdf_files, 1):
        pdf_path = Path(PDF_FOLDER) / pdf_filename
        print(f"[{i}/{len(pdf_files)}] ", end="")
        
        records = extract_from_pdf(pdf_path)
        all_records.extend(records)
    
    print("-" * 30)
    print(f"EXTRACTION COMPLETE")
    print(f"Total faculty records found: {len(all_records)}")
    
    # Save to Excel
    if all_records:
        print("\nSaving to Excel...")
        success = save_results(all_records)
        
        if success:
            print("\nSUCCESS! Check your Desktop for the Excel file!")
        else:
            print("\nFailed to create Excel file")
    else:
        print("No faculty data found in any PDF!")

if __name__ == "__main__":
    try:
        run_extraction()
    except KeyboardInterrupt:
        print("\nStopped by user")
    except Exception as e:
        print(f"\nError: {e}")
    
    input("\nPress Enter to exit...")