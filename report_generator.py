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
    evaluation: str = Field(description="〇, △, ✕, または空欄")
    comment: str = Field(description="△または✕の場合、その理由を具体的に1文で記載。〇の場合は空欄")

class SalesReport(BaseModel):
    cl_company_name: str = Field(description="商談相手の企業名")
    cl_attendee_name: str = Field(description="相手方担当者名")
    cl_attendee_role: str = Field(description="相手方役職")
    our_attendee_name: str = Field(description="自社担当者名")
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
    checklist_evaluations: List[ChecklistItem] = Field(description="以下の31項目のみを評価する。display_textは各項目の文章を15文字以上含む形で記載すること")    
    questions_from_us: List[QAPair] = Field(description="弊社（株式会社Locareの営業担当）から先方（営業をしている会社）への質問と先方の回答")
    questions_from_client: List[QAPair] = Field(description="先方（営業をしている会社）から弊社（株式会社Locare）への質問とLocareの回答")

def generate_report_content(transcript, manual_text, website_text, sales_material_text, model_client):
    parser = PydanticOutputParser(pydantic_object=SalesReport)
    
    prompt_template = """
    あなたは優秀な営業マネージャーですが、今回は「営業を受ける側（発注側）」の視点でレポートを整理します。
    提供された商談文字起こし、HP情報、資料を分析し、営業代行サービスの提案内容を詳細に評価してください。

    【指示】
    1. 会社名は「Webサイト情報」から抽出すること。LocareやSaleshubは運営側なので対象外。
       HP情報の抽出については以下の優先順位で情報を収集すること：
       - まず提供されたWebサイト情報から各項目を抽出する。
       - 情報が不明・不足している項目については、提供されたURLから遷移できる関連ページ
         （サービス詳細ページ、料金ページ、会社概要ページ、導入事例ページ等）も参照して補完する。
       - それでも情報が取得できない場合のみ「※情報なし」と記載する。
       - PDFファイルは参照しないこと。
       - Locare・Saleshub等、運営側の情報は含めないこと。

    2. チェックリスト31項目について、事実に基づき以下の基準で評価してください。
       - 「〇」：基準を満たしている／加点要素あり
       - 「△」：一部不足／普通
       - 「✕」：基準未達／改善が必要
       - 判断できない場合は空欄
       必ず全31項目に対していずれかを入力すること。
       △または✕の場合はcommentフィールドに理由を必ず記載すること。
       display_textは評価項目名を正確に記載すること。
    
    3. HP情報と商談情報の差分を明確にしてください。
    4. 質疑応答（Q&A）は、商談内で実際に行われたやり取りを抽出してください。その際、内容をそのまま記載するのではなく、質問と回答の要点を簡潔にまとめてください。

    【評価項目（チェックリスト31件）】
    冒頭大きな声でさわやかに挨拶できているか？		
    音声の聞こえ方やカメラ画面の見え方に違和感などがないか確認しているか？		
    いきなり質問から始めるのではなく商談の方向性や流れを明確にしてからスタートしているか？		
    商談担当者の自己紹介は行っているか？		
    打ち合わせのタイムスケジュール感を冒頭ですり合わせているか？		
    商談をされる側が不快感を示すような一方的な質問責めをしていないか？		
    HPや事前情報を調べれば把握できるような質問をしていないか？		
    質問に回答した後の相槌や反応はしているか？		
    質問に回答した後のメモに時間を要して沈黙の時間を生んでいないか？		
    質問回答後に「ありがとうございます」と御礼を伝えているか？		
    自社の会社概要について話しているか？		
    提供するサービスのビジネスモデルを分かりやすく簡潔に話しているか？		
    ビジネスモデルは企業HP記載の内容と一致しているか？(虚偽、誇大記載はしていないか？)		
    他社との違いや優位性について触れているか？		
    事例共有は営業を受ける側に近い事例を準備し、説明しているか？		
    具体的な数値効果について仮説を交えながらわかりやすく説明しているか？		
    一方的なサービス説明に終始するのではなく、中間で質疑応答の時間を設けているか？		
    サービスフローはスケジュール感を交えて説明できているか？		
    料金プランは一部を切り出して説明するのではなく、全体でかかる費用を網羅的に説明できているか？		
    リスクやデメリットについても説明しているか？		
    営業を受けている側のニーズに合わせた提案を行っているか？		
    質問に対して結論ファーストで簡潔に回答できているか？		
    自社・他社情報を共有する際には情報セキュリティに十分に配慮できているか？		
    BANT情報を聞けているか？		
    導入に向けた懸念点を聞き出しているか？（テストクロージング）		
    金額懸念に対しての交渉に応じる姿勢は見せているか？		
    ネクストアクションの提案やすり合わせを行っているか？		
    決裁者を交えた次回打ち合わせ日程調整の打診を行っているか？		
    検討期限を区切っているか？		
    金額に対しての価値が伝わっており、価格に対し安いと思える価値訴求ができているか？		
    強引なクロージングで終えるのではなく気分の良い終話だったか		
    
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

    # output.contentがリストの場合と文字列の場合の両方に対応
    if isinstance(output.content, list):
        output_text = "".join(
            block["text"] for block in output.content
            if isinstance(block, dict) and block.get("type") == "text"
        )
    else:
        output_text = output.content

    return parser.parse(output_text).dict()

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
    # まず共有ドライブIDを取得する
    folder_info = drive_service.files().get(
        fileId=folder_id,
        supportsAllDrives=True,
        fields='driveId'
    ).execute()
    shared_drive_id = folder_info.get('driveId', folder_id)

    copy_file = drive_service.files().copy(
        fileId=template_id,
        body={
            'name': f"{data['cl_company_name']}様_営業レポート",
            'parents': [folder_id],
            'driveId': shared_drive_id,          # ← 追加
            'teamDriveId': shared_drive_id,      # ← 追加（後方互換）
        },
        supportsAllDrives=True,
        ignoreDefaultVisibility=True,
        fields='id'
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

    # 5. チェックリスト評価（Row130〜160固定・31項目）
    # D列テキストとの照合でG列に評価、J列に備考を書き込む
    TARGET_ROWS = list(range(130, 161))  # Row130〜160
    sheet_d_col = ws.col_values(4)  # D列を取得（0始まりなのでrow-1がindex）

    for item in data.get('checklist_evaluations', []):
        for row_num in TARGET_ROWS:
            d_text = sheet_d_col[row_num - 1] if row_num - 1 < len(sheet_d_col) else ''
            if item['display_text'][:15] in str(d_text):
                # G列に評価（〇△✕）を書き込む
                # 〇の文字コードをスプレッドシートの形式に統一
                # AIが出力する可能性のある全パターンを正規化
                evaluation = item['evaluation']
                evaluation = evaluation.replace('○', '〇')  # U+25CB → U+3007
                evaluation = evaluation.replace('◯', '〇')  # U+25EF → U+3007
                evaluation = evaluation.replace('O', '〇')  # 半角英字O → U+3007
                evaluation = evaluation.replace('o', '〇')  # 半角英字o → U+3007
                evaluation = evaluation.replace('×', '✕')  # U+00D7 → U+2715
                evaluation = evaluation.replace('✗', '✕')  # U+2717 → U+2715
                evaluation = evaluation.replace('X', '✕')  # 半角英字X → U+2715
                evaluation = evaluation.replace('x', '✕')  # 半角英字x → U+2715
                evaluation = evaluation.strip()             # 前後の空白除去

                # G列に評価（〇△✕）を書き込む
                batch_updates.append({
                    'range': f'G{row_num}',
                    'values': [[evaluation]]
                })
                # △または✕の場合、J列に備考を書き込む
                comment = item.get('comment', '')
                if comment and evaluation in ['△', '✕']:
                    batch_updates.append({
                        'range': f'J{row_num}',
                        'values': [[comment]]
                    })
                break

    # 全データを一括書き込み（これによりAPI呼び出し回数を激減させ、エラーを防ぐ）
    ws.batch_update(batch_updates)

    return f"https://docs.google.com/spreadsheets/d/{new_sheet_id}"