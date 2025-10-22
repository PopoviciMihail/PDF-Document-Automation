# PDF Document Automation System

Automated document generation system that creates personalized PDF documents (offer letters, policy acknowledgements, probation confirmations, etc.) from Excel data using Word templates.

## Features

- **Multi-template support** - Generate different document types from a single system
- **Interactive template selection** - Choose which document type to generate at runtime
- **Batch processing** - Process multiple employees from Excel in one run
- **Automatic PDF conversion** - Converts Word documents to PDF automatically
- **Organized output** - Generates documents in separate folders by type
- **Template variables** - Fully customizable with dynamic data from Excel

## Project Structure

```
PDF Document Automation/
├── output/                              # Generated PDF files (organized by type)
│   ├── offer_letter/                    # Offer letter PDFs
│   ├── policy_acknowledgement/          # Policy acknowledgement PDFs
│   └── probation_confirmation/          # Probation confirmation PDFs
├── templates/                           # Word document templates
│   ├── template_offer_letter.docx       # Offer letter template
│   ├── template_policy_acknowledgement.docx
│   └── template_probation_confirmation.docx
├── venv/                                # Virtual environment
├── .gitignore                           # Git ignore file
├── employees.xlsx                       # Employee data source
├── main.py                              # Main automation script
├── README.md                            # Project documentation
└── requirements.txt                     # Python dependencies
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

   - **Windows:**
     ```bash
     venv\Scripts\activate
     ```
   - **Mac/Linux:**
     ```bash
     source venv/bin/activate
     ```

4. Install required packages:

```bash
pip install -r requirements.txt
```

## Usage

### 1. Prepare Your Excel File

Create or update `employees.xlsx` with the following columns:

| Column Name      | Required | Description                  | Example           |
| ---------------- | -------- | ---------------------------- | ----------------- |
| Name             | Yes      | Employee's full name         | Alice Smith       |
| Position         | Yes      | Job title                    | Software Engineer |
| StartDate        | Yes      | Employment start date        | 2025-02-01        |
| DocumentType     | Optional | Type of document to generate | offer_letter      |
| CompanyName      | No       | Company name                 | Tech Corp Inc.    |
| HRName           | No       | HR contact name              | Sarah Johnson     |
| HREmail          | No       | HR contact email             | hr@company.com    |
| HRPhone          | No       | HR contact phone             | (555) 123-4567    |
| ResponseDeadline | No       | Deadline for response        | January 30, 2025  |
| SupervisorName   | No       | Supervisor's name            | John Doe          |
| OfficeLocation   | No       | Office location              | New York Office   |

**Note:** If you add a `DocumentType` column, only rows matching the selected document type will be processed.

### 2. Create Templates

Templates must be Word documents (.docx) placed in the `templates/` folder with the naming convention:

- `template_offer_letter.docx`
- `template_policy_acknowledgement.docx`
- `template_probation_confirmation.docx`

**Available template variables:**

- `{{ Name }}` - Employee's full name
- `{{ Position }}` - Job position
- `{{ StartDate }}` - Formatted start date (e.g., "February 01, 2025")
- `{{ CompanyName }}` - Company name
- `{{ HRName }}` - HR contact name
- `{{ HREmail }}` - HR contact email
- `{{ HRPhone }}` - HR contact phone
- `{{ ResponseDeadline }}` - Response deadline
- `{{ SupervisorName }}` - Supervisor's name
- `{{ OfficeLocation }}` - Office location
- `{{ CurrentDate }}` - Today's date (auto-generated)

### 3. Run the Script

```bash
python main.py
```

The script will:

1. Display available document types
2. Prompt you to select a document type
3. Generate PDFs for all matching employees
4. Save them in `output/[document_type]/` folder

### Example Output

```
Available document types:
1. offer_letter
2. policy_acknowledgement
3. probation_confirmation
Select a document type (1-3): 1
Selected document type: offer_letter
Generated DOCX: output/offer_letter/Offer_letter_Alice_Smith.docx
Converted to PDF: output/offer_letter/Offer_letter_Alice_Smith.pdf
Generated DOCX: output/offer_letter/Offer_letter_Bob_Johnson.docx
Converted to PDF: output/offer_letter/Offer_letter_Bob_Johnson.pdf
All offer_letter documents have been generated
```

## Template Examples

### Offer Letter Template

```
{{ CompanyName }}

{{ CurrentDate }}

Dear {{ Name }},

We are pleased to offer you the position of {{ Position }} starting on {{ StartDate }}.

Please respond by {{ ResponseDeadline }}.

Best regards,
{{ HRName }}
{{ HREmail }}
```

### Policy Acknowledgement Template

```
{{ CompanyName }}

{{ CurrentDate }}

Dear {{ Name }},

This letter confirms that you have reviewed and acknowledged our company policies.

Position: {{ Position }}
Supervisor: {{ SupervisorName }}
Office Location: {{ OfficeLocation }}

Signature: _________________________
Date: _________________________
```

## Configuration

You can modify these variables in `main.py`:

```python
EXCEL_FILE = 'employees.xlsx'    # Path to Excel file
TEMPLATE_DIR = "templates"        # Templates folder
OUTPUT_DIR = "output"             # Output folder
```

## Adding New Document Types

1. Create a new Word template: `template_your_document_type.docx`
2. Place it in the `templates/` folder
3. Use template variables like `{{ Name }}`, `{{ Position }}`, etc.
4. Run the script - it will automatically appear in the selection menu

## Dependencies

```
pandas - Excel file processing
docxtpl - Word template rendering
docx2pdf - PDF conversion
openpyxl - Excel file support
python-docx - Word document manipulation
```

## Troubleshooting

**"Template not found" error:**

- Verify template files exist in `templates/` folder
- Check naming convention: `template_[document_type].docx`
- Ensure file extensions are `.docx` not `.doc`

**PDF conversion fails:**

- Ensure Microsoft Word is installed
- Close Word if it's running in the background
- Check file permissions in output folder

**Template variables not replacing:**

- Use double curly braces: `{{ Variable }}` not `{ Variable }`
- Ensure column names in Excel match variable names exactly
- Don't edit .docx files in VS Code or text editors (use Microsoft Word)

**Date formatting issues:**

- Ensure StartDate column contains valid dates
- Use format: YYYY-MM-DD or Excel date format

## Best Practices

1. **Always create templates in Microsoft Word** - Never edit .docx files in text editors
2. **Test with one employee first** - Add a test row in Excel before batch processing
3. **Backup templates** - Keep original templates before making changes
4. **Use consistent naming** - Follow the `template_[type].docx` convention
5. **Check PDFs** - Review generated PDFs to ensure formatting is correct
