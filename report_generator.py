import json
import gspread
import time
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from typing import List
from langchain_core.prompts import PromptTemplate
from langchain_core.output_parsers import PydanticOutputParser
from pydantic import BaseModel, Field

# 1. AIが抽出するデータ構造の定義（営業代行用）
class QAPair(BaseModel):
    question: str = Field(description="質問内容")
    answer: str = Field(description="回答")

class ChecklistItem(BaseModel):
    display_text: str = Field(description="評価項目名")
    evaluation: str = Field(description="○, △, ×, または空欄")

class SalesReport(BaseModel):
    cl_company_name: str = Field(description="商談相手の企業名")
    cl_attendee_name: str = Field(description="相手方担当者名")
    cl_attendee_role: str = Field(description="相手方役職")
    our_attendee_name: str = Field(description="自社担当者名")
    overall_score: str = Field(description="総合評価点数(0-100)")
    impression: str = Field(description="全体統括コメント・所感")
    
    # HP情報 (12項目)
    hp_service_overview: str = Field(description="サービス概要")
    hp_business_model: str = Field(description="ビジネスモデル")
    hp_strength: str = Field(description="サービスの強み")
    hp_difference: str = Field(description="他社との違い")
    hp_pricing: str = Field(description="料金プラン・形態")
    hp_min_price: str = Field(description="最低料金")
    hp_min_period: str = Field(description="最短契約期間")
    hp_customer_voice: str = Field(description="導入企業の声、情報")
    hp_competitors: str = Field(description="類似サービス、競合企業")
    hp_recruitment: str = Field(description="求人掲載有無")
    hp_recruitment_url: str = Field(description="求人URL")
    hp_review_voice: str = Field(description="転職口コミサイトの声")

    # 商談での情報整理 (12項目)
    neg_service_overview: str = Field(description="サービス概要(商談より)")
    neg_business_model: str = Field(description="ビジネスモデル(商談より)")
    neg_strength: str = Field(description="強み(商談より)")
    neg_difference: str = Field(description="他社との違い(商談より)")
    neg_pricing: str = Field(description="料金プラン(商談より)")
    neg_min_price: str = Field(description="最低料金(商談より)")
    neg_min_period: str = Field(description="最短契約期間(商談より)")
    neg_cancellation: str = Field(description="途中解約有無")
    neg_lead_tool: str = Field(description="リスト作成ツール")
    neg_cti: str = Field(description="CTI")
    neg_competitor_check: str = Field(description="競合会社認識")
    neg_evidence: str = Field(description="架電件数・アポ率エビデンス")

    # チェックリスト (31項目)
    checklist_evaluations: List[ChecklistItem] = Field(description="31個の評価項目に対する○△×評価")
    
    questions_from_us: List[QAPair] = Field(description="弊社（営業担当）から先方への質問と回答")
    questions_from_client: List[QAPair] = Field(description="先方から弊社への質問と回答")

def generate_report_content(transcript, manual_text, website_text, sales_material_text, model_client):
    parser = PydanticOutputParser(pydantic_object=SalesReport)
    
    prompt_template = """
    あなたは優秀な営業マネージャーですが、今回は「営業を受ける側（発注側）」の視点でレポートを整理します。
    提供された商談文字起こし、HP情報、資料を分析し、営業代行サービスの提案内容を詳細に評価してください。

    【指示】
    1. 会社名は「Webサイト情報」から抽出すること。LocareやSaleshubは運営側なので対象外。
    2. チェックリスト31項目について、事実に基づき「○」「△」「×」で評価してください。
    3. HP情報と商談情報の差分を明確にしてください。
    4. 質疑応答（Q&A）は、商談内で実際に行われたやり取りを抽出してください。

    【評価項目（チェックリスト31件）】
    (略: あなたが指定した31項目がここに自動的に入るよう指示)
    
    ======= 営業資料/HP/文字起こし =======
    {website_text}
    {sales_material_text}
    {transcript}
    ====================================
    {format_instructions}
    """
    
    prompt = PromptTemplate(
        template=prompt_template,
        input_variables=["transcript", "website_text", "sales_material_text"],
        partial_variables={"format_instructions": parser.get_format_instructions()}
    )
    
    _input = prompt.format_prompt(transcript=transcript, website_text=website_text, sales_material_text=sales_material_text)
    output = model_client.invoke(_input.to_string())
    return parser.parse(output.content).dict()

def fill_google_sheet(data, service_account_info, template_id, folder_id):
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    
    # Google Drive API を使用してテンプレートをコピー
    # 403 Quotaエラー対策として、明確に保存先を指定
    drive_service = build('drive', 'v3', credentials=creds, cache_discovery=False)
    
    file_metadata = {
        'name': f"{data.get('cl_company_name', '名称未設定')}様_営業レポート",
        'parents': [folder_id]
    }
    
    # report_generator.py の copy 部分を以下のように補強
    copy_file = drive_service.files().copy(
        fileId=template_id,
        body={
            'name': f"{data['cl_company_name']}様_営業レポート",
            'parents': [folder_id]
        },
        supportsAllDrives=True,
        # 以下の1行を追加してみる
        ignoreDefaultVisibility=True 
    ).execute()
    
    new_sheet_id = copy_file['id']
    
    # gspreadで書き込み開始
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(new_sheet_id)
    ws = sh.get_worksheet(0)

    # --- データの準備 (一括更新用リスト) ---
    batch_updates = []

    # 1. 基本ヘッダー情報
    batch_updates.append({'range': 'B1', 'values': [[data.get('cl_company_name', "")]]})
    batch_updates.append({'range': 'B4', 'values': [[data.get('cl_attendee_name', "")]]})
    batch_updates.append({'range': 'B5', 'values': [[data.get('cl_attendee_role', "")]]})
    batch_updates.append({'range': 'B6', 'values': [[data.get('our_attendee_name', "")]]})
    batch_updates.append({'range': 'B10', 'values': [[data.get('overall_score', "")]]})
    batch_updates.append({'range': 'A50', 'values': [[data.get('impression', "")]]})

    # 2. HP情報 (C35:C46)
    hp_keys = ['service_overview','business_model','strength','difference','pricing','min_price','min_period','customer_voice','competitors','recruitment','recruitment_url','review_voice']
    hp_values = [[data.get(f'hp_{k}', "")] for k in hp_keys]
    batch_updates.append({'range': 'C35:C46', 'values': hp_values})

    # 3. 商談での情報整理 (H35:H46)
    neg_keys = ['service_overview','business_model','strength','difference','pricing','min_price','min_period','cancellation','lead_tool','cti','competitor_check','evidence']
    neg_values = [[data.get(f'neg_{k}', "")] for k in neg_keys]
    batch_updates.append({'range': 'H35:H46', 'values': neg_values})

    # 4. Q&A セクション (I列とJ列)
    # 弊社→先方 (Row 3-16)
    us_qa = data.get('questions_from_us', [])
    us_qa_values = [[qa['question'], qa['answer']] for qa in us_qa[:14]] # 最大14件
    if us_qa_values:
        batch_updates.append({'range': f'I3:J{2 + len(us_qa_values)}', 'values': us_qa_values})

    # 先方→弊社 (Row 18-30)
    client_qa = data.get('questions_from_client', [])
    client_qa_values = [[qa['question'], qa['answer']] for qa in client_qa[:13]] # 最大13件
    if client_qa_values:
        batch_updates.append({'range': f'I18:J{17 + len(client_qa_values)}', 'values': client_qa_values})

    # 5. チェックリスト評価 (G列)
    # 速度向上のため、一旦C列（質問文）を全取得
    sheet_questions = ws.col_values(3) # C列
    checklist_marks = []
    
    # C列の質問文に対応する評価を探してリスト化
    # 31項目の評価を反映
    for item in data.get('checklist_evaluations', []):
        for i, q_text_in_sheet in enumerate(sheet_questions):
            # 質問文の一部が一致したら、その行のG列(7列目)に評価を入れる
            if item['display_text'][:10] in q_text_in_sheet:
                batch_updates.append({
                    'range': f'G{i+1}',
                    'values': [[item['evaluation']]]
                })
                break

    # 全データを一括書き込み（これによりAPI呼び出し回数を激減させ、エラーを防ぐ）
    ws.batch_update(batch_updates)

    return f"https://docs.google.com/spreadsheets/d/{new_sheet_id}"