import os
import sys

print("--- Checking Imports ---")
try:
    import streamlit
    print("streamlit imported")
    import langchain
    print("langchain imported")
    import langchain_openai
    print("langchain_openai imported")
    import langchain_google_genai
    print("langchain_google_genai imported")
    import openpyxl
    print("openpyxl imported")
    import pypdf
    print("pypdf imported")
    import pandas
    print("pandas imported")
except Exception as e:
    print(f"IMPORT ERROR: {e}")
    sys.exit(1)

print("\n--- Checking Local Modules ---")
try:
    from utils import extract_text_from_pdf
    print("utils imported")
    from report_generator import generate_report_content, fill_excel_template
    print("report_generator imported")
except Exception as e:
    print(f"LOCAL MODULE ERROR: {e}")
    sys.exit(1)

print("\n--- Checking File Paths ---")
DEFAULT_TEMPLATE_PATH = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"
DEFAULT_MANUAL_PATH = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/8ba0d12e-f2ee-4002-9533-54a0940f4eaa_営業レポートマニュアル.pdf"

if os.path.exists(DEFAULT_TEMPLATE_PATH):
    print(f"Template found: {DEFAULT_TEMPLATE_PATH}")
else:
    print(f"Template NOT FOUND: {DEFAULT_TEMPLATE_PATH}")

if os.path.exists(DEFAULT_MANUAL_PATH):
    print(f"Manual found: {DEFAULT_MANUAL_PATH}")
else:
    print(f"Manual NOT FOUND: {DEFAULT_MANUAL_PATH}")
    
print("\n--- Done ---")
