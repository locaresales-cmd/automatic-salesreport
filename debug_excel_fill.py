
import os
import sys
import openpyxl
from report_generator import fill_excel_template

# Define dummy data
dummy_data = {
    "cl_company_name": "Test Company Ltd.",
    "cl_attendee_name": "Test Person",
    "cl_attendee_role": "Manager",
    "our_attendee_name": "Our Rep",
    "overall_summary": "This is a test summary.",
    "service_overview": "Test Service Overview",
    "website_url": "https://example.com",
    "questions_from_us": [
        {"question": "Test Q1 from Us", "answer": "Test A1"}
    ],
    "questions_from_client": [
        {"question": "Test Q1 from Client", "answer": "Test A1"}
    ]
}

# Template path (from app.py)
template_path = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"
output_path = "debug_output.xlsx"

if not os.path.exists(template_path):
    print(f"Template not found at: {template_path}")
    sys.exit(1)

print("Running fill_excel_template...")
try:
    result_path = fill_excel_template(template_path, dummy_data, output_path)
    print(f"Successfully generated: {result_path}")
    
    # Verify content
    wb = openpyxl.load_workbook(result_path)
    sheet = wb.active
    
    print(f"B1 (Company Name): {sheet['B1'].value}")
    print(f"B3 (Website URL): {sheet['B3'].value}")
    print(f"C35 (Service Overview): {sheet['C35'].value}")
    print(f"I3 (First Question from Us): {sheet['I3'].value}")
    
except Exception as e:
    print(f"Error: {e}")
