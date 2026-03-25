import pypdf
import pandas as pd
import os

pdf_path = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/8ba0d12e-f2ee-4002-9533-54a0940f4eaa_営業レポートマニュアル.pdf"
excel_path = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"

print("--- PDF Inspection ---")
try:
    reader = pypdf.PdfReader(pdf_path)
    print(f"Number of pages: {len(reader.pages)}")
    for i, page in enumerate(reader.pages[:5]): # Inspect first 5 pages
        print(f"\n[Page {i+1}]")
        print(page.extract_text())
except Exception as e:
    print(f"Error reading PDF: {e}")

print("\n--- Excel Inspection ---")
try:
    xls = pd.ExcelFile(excel_path)
    print(f"Sheet names: {xls.sheet_names}")
    for sheet_name in xls.sheet_names:
        print(f"\n[Sheet: {sheet_name}]")
        df = pd.read_excel(xls, sheet_name=sheet_name, nrows=10) # Inspect first 10 rows
        print(df.to_string())
except Exception as e:
    print(f"Error reading Excel: {e}")
