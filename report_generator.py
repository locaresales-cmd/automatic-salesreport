import json
import gspread
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
    
    questions_from_us: List[QAPair] = Field(description="自社から相手への質問")
    questions_from_client: List[QAPair] = Field(description="相手から自社への質問")

def generate_report_content(transcript, manual_text, website_text, sales_material_text, model_client):
    parser = PydanticOutputParser(pydantic_object=SalesReport)
    
    prompt_template = """
    あなたは営業代行会社の商談を評価するバイヤー側の責任者です。
    提供された情報を元に、営業代行サービスの提案内容を詳細に分析し、レポートを作成してください。

    【指示】
    1. 会社名は「Webサイト情報」から抽出すること。LocareやSaleshubは運営側なので対象外。
    2. チェックリスト31項目について、商談内容から厳しく「○」「△」「×」で評価してください。
    3. HP情報と商談情報の差分を明確にしてください。

    【評価項目（チェックリスト31件）】
    - 冒頭大きな声でさわやかに挨拶できているか？
    - 音声の聞こえ方やカメラ画面の見え方に違和感などがないか確認しているか？
    - いきなり質問から始めるのではなく商談の方向性や流れを明確にしてからスタートしているか？
    - 商談担当者の自己紹介は行っているか？
    - 打ち合わせのタイムスケジュール感を冒頭ですり合わせているか？
    - 商談をされる側が不快感を示すような一方的な質問責めをしていないか？
    - HPや事前情報を調べれば把握できるような質問をしていないか？
    - 質問に回答した後の相槌や反応はしているか？
    - 質問に回答した後のメモに時間を要して沈黙の時間を生んでいないか？
    - 質問回答後に「ありがとうございます」と御礼を伝えているか？
    - 自社の会社概要について話しているか？
    - 提供するサービスのビジネスモデルを分かりやすく簡潔に話しているか？
    - ビジネスモデルは企業HP記載の内容と一致しているか？(虚偽、誇大記載はしていないか？)
    - 他社との違いや優位性について触れているか？
    - 事例共有は営業を受ける側に近い事例を準備し、説明しているか？
    - 具体的な数値効果について仮説を交えながらわかりやすく説明しているか？
    - 一方的なサービス説明に終執するのではなく、中間で質疑応答の時間を設けているか？
    - サービスフローはスケジュール感を交えて説明できているか？
    - 料金プランは一部を切り出して説明するのではなく、全体でかかる費用を網羅的に説明できているか？
    - リスクやデメリットについても説明しているか？
    - 営業を受けている側のニーズに合わせた提案を行っているか？
    - 質問に対して結論ファーストで簡潔に回答できているか？
    - 自社・他社情報を共有する際には情報セキュリティに十分に配慮できているか？
    - BANT情報を聞けているか？
    - 導入に向けた懸念点を聞き出しているか？（テストクロージング）
    - 金額懸念に対しての交渉に応じる姿勢は見せているか？
    - ネクストアクションの提案やすり合わせを行っているか？
    - 決裁者を交えた次回打ち合わせ日程調整の打診を行っているか？
    - 検討期限を区切っているか？
    - 金額に対しての価値が伝わっており、価格に対し安いと思える価値訴求ができているか？
    - 強引なクロージングで終えるのではなく気分の良い終話だったか

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
    
    # テンプレートをコピー
    drive_service = build('drive', 'v3', credentials=creds)
    copy_file = drive_service.files().copy(
        fileId=template_id,
        body={'name': f"{data['cl_company_name']}様_営業レポート", 'parents': [folder_id]},
        supportsAllDrives=True
    ).execute()
    new_sheet_id = copy_file['id']
    
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(new_sheet_id)
    ws = sh.get_worksheet(0)

    # 書き込み (セル番地はスプレッドシートのレイアウトに合わせて調整してください)
    ws.update_acell('B1', data['cl_company_name'])
    ws.update_acell('B4', data['cl_attendee_name'])
    ws.update_acell('B5', data['cl_attendee_role'])
    ws.update_acell('B6', data['our_attendee_name'])
    ws.update_acell('B10', data['overall_score'])
    ws.update_acell('A50', data['impression'])

    # HP情報 (C35-C46)
    hp_list = [[data[f'hp_{k}']] for k in ['service_overview','business_model','strength','difference','pricing','min_price','min_period','customer_voice','competitors','recruitment','recruitment_url','review_voice']]
    ws.update('C35:C46', hp_list)

    # 商談での情報整理 (H35-H46)
    neg_list = [[data[f'neg_{k}']] for k in ['service_overview','business_model','strength','difference','pricing','min_price','min_period','cancellation','lead_tool','cti','competitor_check','evidence']]
    ws.update('H35:H46', neg_list)

    # チェックリストの書き込みロジック (G列に評価を流し込む例)
    # テンプレートの質問文(C列)と一致する行を探してG列に評価を書く
    sheet_questions = ws.col_values(3) # C列
    for item in data['checklist_evaluations']:
        for i, q_text in enumerate(sheet_questions):
            if item['display_text'][:10] in q_text:
                ws.update_cell(i+1, 7, item['evaluation']) # 7列目=G列
                break

    return f"https://docs.google.com/spreadsheets/d/{new_sheet_id}"