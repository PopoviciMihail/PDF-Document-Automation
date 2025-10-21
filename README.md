# PDF Document Automation

Automated offer letter generation system that creates personalized PDF documents from Excel data using Word templates.

## Features

- Reads employee data from Excel spreadsheet
- Generates personalized offer letters using Word templates
- Automatically converts documents to PDF format
- Batch processes multiple employees at once

## Project Structure

```
PDF Document Automation/
├── output/                      # Generated PDF files
├── utils/                       # Utility scripts
│   └── generateInitialExcel.py  # Creates sample Excel file
├── employees.xlsx               # Employee data source
├── main.py                      # Main automation script
├── template.docx                # Offer letter template
├── requirements.txt             # Python dependencies
└── README.md                    # Project documentation
```

## Prerequisites

- Python 3.7+
- Microsoft Word (required for PDF conversion on Windows)

## Installation

1. Clone or download this repository

2. Create a virtual environment:

```bash
python -m venv venv
```

3. Activate the virtual environment:

- Windows:

```bash
     venv\Scripts\activate
```

- Mac/Linux:

```bash
     source venv/bin/activate
```

4. Install required packages:

```bash
pip install -r requirements.txt
```

## Usage

1. Prepare your employee data in `employees.xlsx` with columns:

   - Name
   - Position
   - StartDate

2. Customize the `template.docx` file with your letter content using template variables:

   - `{{ Name }}`
   - `{{ Position }}`
   - `{{ StartDate }}`

3. Run the automation script:

```bash
python main.py
```

4. Find generated PDFs in the `output/` folder

## Template Variables

The following variables can be used in your Word template:

| Variable          | Description           | Example           |
| ----------------- | --------------------- | ----------------- |
| `{{ Name }}`      | Employee's full name  | Alice Smith       |
| `{{ Position }}`  | Job position          | Software Engineer |
| `{{ StartDate }}` | Employment start date | January 15, 2025  |

## Example Excel Data

| Name        | Position          | StartDate  |
| ----------- | ----------------- | ---------- |
| Alice Smith | Software Engineer | 2025-01-15 |
| Bob Johnson | Product Manager   | 2025-02-01 |
| Charlie Lee | Data Analyst      | 2025-01-20 |

## Dependencies

- `pandas` - Excel file processing
- `docxtpl` - Word template rendering
- `docx2pdf` - PDF conversion
- `openpyxl` - Excel file support

## Troubleshooting

**PDF conversion fails:**

- Ensure Microsoft Word is installed on Windows
- Check that Word is not running in the background

**Template not found error:**

- Verify `template.docx` exists in the project root
- Do not edit .docx files in text editors (use Microsoft Word)

**Date formatting issues:**

- Ensure StartDate column in Excel contains valid dates
