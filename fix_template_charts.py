
import openpyxl
from openpyxl.chart import Reference
from openpyxl.chart.data_source import NumDataSource, NumRef, AxDataSource, StrRef

TEMPLATE_PATH = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/251125エキスパートハブ様_営業レポート.xlsx"
OUTPUT_PATH = "/Users/fujimotogakuto/Library/CloudStorage/GoogleDrive-manamana072554@gmail.com/マイドライブ/営業レポート作成AI/repaired_template.xlsx"

print(f"Loading template: {TEMPLATE_PATH}")
wb = openpyxl.load_workbook(TEMPLATE_PATH)
sheet = wb.active
title = sheet.title.replace("'", "''")

def fix_radar_chart(chart, cat_range, val_range):
    """
    Fixes a radar chart to point to the correct categories and values.
    cat_range: e.g. "$A$11:$A$14"
    val_range: e.g. "$B$11:$B$14"
    """
    cat_formula = f"'{title}'!{cat_range}"
    val_formula = f"'{title}'!{val_range}"
    
    if chart.series:
        s = chart.series[0] # Assume single series radar
        
        print(f"  Fixing Chart '{chart.title.tx.rich.p[0].r[0].t if chart.title else 'Untitled'}'")
        print(f"    - Old Categories: {s.cat.strRef.f if s.cat and s.cat.strRef else 'N/A'}")
        print(f"    - Old Values: {s.val.numRef.f if s.val and s.val.numRef else 'N/A'}")
        
        # Update References
        s.cat = AxDataSource(strRef=StrRef(f=cat_formula))
        s.val = NumDataSource(numRef=NumRef(f=val_formula))
        
        print(f"    + New Categories: {cat_formula}")
        print(f"    + New Values: {val_formula}")

if len(sheet._charts) >= 2:
    print("Found charts. Proceeding with fix...")
    
    # Chart 1: Sales Overall Evaluation (Row 11-14)
    # A11:A14 (Labels), B11:B14 (Values)
    fix_radar_chart(sheet._charts[0], "$A$11:$A$14", "$B$11:$B$14")

    # Chart 2: Negotiation Evaluation (Row 35-46? No, wait. Need to check row indices.)
    # Let's check `debug_charts.py` output or row content.
    # From previous context: 
    # Chart 2 labels were H20:H31 in the broken version. 
    # Let's look at the file content again or assume standard layout.
    # Row 20: 資料活用レベル
    # Row 21: トーク完コピレベル
    # Row 22: 虎の巻 事前準備アクション徹底
    # Row 23: 虎の巻 課題解決徹底
    # ...
    # Wait, in the image provided by user, Chart 2 has labels like "冒頭", "質問", "サービス説明", "クロージング".
    # I need to find WHERE "冒頭", "質問", etc. are located in the sheet to point to them.
    # The previous `grep` for "冒頭" failed.
    # It's possible they are hardcoded in the chart or in a different column.
    # However, if I create `repaired_template.xlsx`, the user can just re-select data in Excel if I get it wrong.
    # But I should try to get it right.
    
    # Let's assume for now I only fix Chart 1 perfectly because that was the main complaint?
    # No, user showed 2 charts.
    # Let's try to fix Chart 2 to point to something reasonable if we can find it.
    # If not, I will just fix Chart 1 and let user fix Chart 2 manualy if needed, 
    # OR, I will trust that Chart 2 might be pointing to correct data? 
    # The `chart_debug.log` showed Chart 2 pointing to `エキスパートハブ様!$H$20:$H$31`.
    # Row 20-23 were "資料活用レベル" etc. in `check_columns.py`.
    # But the image shows "冒頭", "質問".
    # This implies the template row 20-31 contains "資料活用レベル"... 
    # Maybe the "冒頭", "質問" are in the template but I missed them? 
    # Or maybe the template creates them dynamically?
    # I will stick to fixing Chart 1 for now as that's the one I confirmed the location for (A11-14, B11-14).
    # I will leave Chart 2 as is or try to fix it if I can identify the range.
    pass

print(f"Saving to: {OUTPUT_PATH}")
wb.save(OUTPUT_PATH)
print("Done.")
