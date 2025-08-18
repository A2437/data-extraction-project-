import pdfplumber
import pandas as pd
import os
import webbrowser

# -----------------------------
# CONFIGURATION
# -----------------------------
PDF_PATH = r"C:\Users\Abhishek & Amma\Downloads\NIRF-Application-2025-Engineering.pdf"  # Path to client's PDF
OUTPUT_EXCEL = "Filtered_Faculty.xlsx"

# -----------------------------
# HELPER FUNCTION: Qualification check
# -----------------------------
def is_valid_qualification(q):
    """Return True if qualification is M.Tech, M.E., or Ph.D (and not B.Tech/B.E)."""
    if not q:
        return False
    q_lower = q.strip().lower()
    valid_terms = {"ph.d", "phd", "m.tech", "mtech", "m.e.", "me"}
    invalid_terms = {"b.tech", "b.e", "be"}
    has_valid = any(term in q_lower for term in valid_terms)
    has_invalid = any(term in q_lower for term in invalid_terms)
    return has_valid and not has_invalid

# -----------------------------
# SAFE SAVE FUNCTION
# -----------------------------
def safe_save_excel(df, file_name):
    """Save Excel safely, avoiding 'Permission Denied' errors."""
    if os.path.exists(file_name):
        try:
            os.remove(file_name)
        except PermissionError:
            raise PermissionError(
                f"Please close '{file_name}' in Excel before running the script."
            )
    df.to_excel(file_name, index=False)
    abs_path = os.path.abspath(file_name)
    print(f"File saved to: {abs_path}")
    return abs_path

# -----------------------------
# MAIN EXTRACTION
# -----------------------------
def extract_faculty_from_pdf(pdf_path):
    data_rows = []
    headers_found = False
    column_names = []
    in_faculty_section = False

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""

            # Detect start of Faculty Details
            if "Faculty Details" in text:
                in_faculty_section = True

            # Stop at the next known section
            if in_faculty_section and "Financial Resources" in text:
                break

            if not in_faculty_section:
                continue

            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                # Clean rows
                cleaned_table = [
                    [cell.strip() if cell else "" for cell in row]
                    for row in table
                    if any(cell and cell.strip() for cell in row)
                ]
                if not cleaned_table:
                    continue

                # First table: capture headers
                if not headers_found:
                    column_names = cleaned_table[0]
                    try:
                        age_idx = [h.lower().strip() for h in column_names].index("age")
                        qual_idx = [h.lower().strip() for h in column_names].index("qualification")
                    except ValueError:
                        continue
                    headers_found = True
                    table_data = cleaned_table[1:]
                else:
                    table_data = cleaned_table

                # Process rows
                for row in table_data:
                    if len(row) < len(column_names):
                        row += [""] * (len(column_names) - len(row))
                    try:
                        age = int(row[age_idx])
                    except ValueError:
                        continue
                    qualification = row[qual_idx]
                    if age <= 65 and is_valid_qualification(qualification):
                        data_rows.append(row)

    return pd.DataFrame(data_rows, columns=column_names) if data_rows else pd.DataFrame()

# -----------------------------
# RUN SCRIPT
# -----------------------------
if __name__ == "__main__":
    try:
        df = extract_faculty_from_pdf(PDF_PATH)
        if not df.empty:
            saved_path = safe_save_excel(df, OUTPUT_EXCEL)
            print(f"Extraction complete. {len(df)} records saved.")
            
            # Auto-open file
            try:
                os.startfile(saved_path)  # Windows
            except AttributeError:
                webbrowser.open(f"file://{saved_path}")
        else:
            print("No matching faculty found based on criteria.")
    except FileNotFoundError:
        print(f"PDF file not found: {PDF_PATH}")
    except Exception as e:
        print(f"Error: {e}")
