import pandas as pd
from docxtpl import DocxTemplate
import os
from docx2pdf import convert

data = pd.read_excel('employees.xlsx')

for index, row in data.iterrows():
    doc = DocxTemplate("template.docx")
    
    context = {
        "Name" : row['Name'],
        "Position" : row['Position'],
        "StartDate": pd.to_datetime(row['StartDate']).strftime('%B %d, %Y') 
    }
    doc.render(context)
    docx_file = os.path.join("output", f"Offer_Letter_{row['Name'].replace(' ', '_')}.docx")
    doc.save(docx_file)
    print(f"Generated DOCX: {docx_file}")
    pdf_file = os.path.join("output", f"Offer_Letter_{row['Name'].replace(' ', '_')}.pdf")
    convert(docx_file, pdf_file)
    print(f"Converted to PDF: {pdf_file}")
    
    os.remove(docx_file)

print("All offer letters have been generated and converted to PDF.")
