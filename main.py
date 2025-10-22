import pandas as pd
from docxtpl import DocxTemplate
import os
from docx2pdf import convert
from datetime import datetime

# Configuration

EXCEL_FILE = 'employees.xlsx'
TEMPLATE_DIR = "templates"
OUTPUT_DIR = "output"

# Read data
data = pd.read_excel(EXCEL_FILE)

# Make sure StartDate is in datetime format
if not pd.api.types.is_datetime64_any_dtype(data['StartDate']):
    data['StartDate'] = pd.to_datetime(data['StartDate'], errors='coerce')

# Extract available document types from the template directory
available_templates = [
    f.replace("template_", "").replace(".docx", "")
    for f in os.listdir(TEMPLATE_DIR)
    if f.startswith("template_") and f.endswith(".docx")
]

print("Available document types:")
for i, doc_type in enumerate(available_templates, start=1):
    print(f"{i}. {doc_type}")

choice = -1
while choice < 1 or choice > len(available_templates):
    try:
        choice = int(input(f"Select a document type (1-{len(available_templates)}): "))
    except ValueError:
        print("Invalid input. Please enter a number.")

selected_doc_type = available_templates[choice - 1]
print(f"Selected document type: {selected_doc_type}")

for index, row in data.iterrows():    
    # If Excel has a DocumentType column, we can skip rows that don't match
    row_doc_type = str(row.get("DocumentType", selected_doc_type)).strip().lower()
    if row_doc_type != selected_doc_type:
        continue
    template_path = os.path.join(TEMPLATE_DIR, f"template_{selected_doc_type}.docx")
    if not os.path.exists(template_path):
        print(f"Template not found: {template_path}")
        continue
    doc = DocxTemplate(template_path)

    context = {
        "Name": row.get("Name", ""),
        "Position": row.get("Position", ""),
        "StartDate": row["StartDate"].strftime("%B %d, %Y") if pd.notnull(row["StartDate"]) else "",
        "CompanyName": row.get("CompanyName", "Your Company"),
        "HRName": row.get("HRName", "HR Team"),
        "HREmail": row.get("HREmail", "hr@company.com"),
        "HRPhone": row.get("HRPhone", ""),
        "ResponseDeadline": row.get("ResponseDeadline", ""),
        "SupervisorName": row.get("SupervisorName", ""),
        "OfficeLocation": row.get("OfficeLocation", "Head Office"),
        "CurrentDate": datetime.now().strftime("%B %d, %Y"), 
    }
    doc.render(context)

    # Create subfolder for selected document type
    doc_output_dir = os.path.join(OUTPUT_DIR, selected_doc_type)
    os.makedirs(doc_output_dir, exist_ok=True)

    base_filename = f"{selected_doc_type.capitalize()}_{row['Name'].replace(' ', '_')}"
    docx_file = os.path.join(doc_output_dir, f"{base_filename}.docx")
    pdf_file = os.path.join(doc_output_dir, f"{base_filename}.pdf")

    doc.save(docx_file)
    print("Generated DOCX:", docx_file)

    try:
        convert(docx_file, pdf_file)
        print("Converted to PDF:", pdf_file)
        os.remove(docx_file)
    except Exception as e:
        print(f"Error converting to PDF for {row['Name']}: {e}")


print(f"All {selected_doc_type} documents have been generated ")