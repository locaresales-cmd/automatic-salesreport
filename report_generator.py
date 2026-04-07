import json
import re
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from typing import List, Dict
from pydantic import BaseModel, Field

# ==============================
# 1. データ構造の定義
# ==============================

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

    # HP情報・商談情報は動的項目に対応したdict形式
    hp_info: Dict[str, str] = Field(
        description="HP情報整理。キーが項目名（hp_itemsで指定した名称と完全一致）、値がその内容"
    )
    neg_info: Dict[str, str] = Field(
        description="商談での情報整理。キーが項目名（neg_itemsで指定した名称と完全一致）、値がその内容"
    )

    # チェックリスト（31項目固定）
    checklist_evaluations: List[ChecklistItem] = Field(
        description="以下の31項目のみを評価する。display_textは各項目の文章を15文字以上含む形で記載すること"
    )
    questions_from_us: List[QAPair] = Field(
        description="弊社（株式会社Locareの営業担当）から先方（営業をしている会社）への質問と先方の回答"
    )
    questions_from_client: List[QAPair] = Field(
        description="先方（営業をしている会社）から弊社（株式会社Locare）への質問とLocareの回答"
    )


# ==============================
# 2. AIによるレポート内容生成
# ==============================

def generate_report_content(
    transcript: str,
    manual_text: str,
    website_text: str,
    sales_material_text: str,
    model_client,
    hp_items: list,
    neg_items: list,
) -> dict:
    """
    文字起こし・HP情報・資料をもとにAIがレポートを生成する。
    hp_items / neg_items は app.py のセッションステートから受け取る動的リスト。
    """

    # 項目リストをプロンプト用の文字列に変換
    hp_items_str  = "\n".join(f"  - {item}" for item in hp_items)
    neg_items_str = "\n".join(f"  - {item}" for item in neg_items)

    # JSON出力フォーマット指示（Pydanticパーサーを使わず直接指定）
    hp_json_example  = ", ".join(f'"{item}": "内容"' for item in hp_items[:3]) + ", ..."
    neg_json_example = ", ".join(f'"{item}": "内容"' for item in neg_items[:3]) + ", ..."

    prompt = f"""
あなたは超一流のBtoBセールス・イネーブルメント（営業組織強化）の専門家です。
今回は「営業を受ける側（発注側）」の視点でレポートを整理します。
提供された商談文字起こし、HP情報、資料を分析し、営業サービスの提案内容を詳細に評価してください。

【指示】
1. 会社名は「Webサイト情報」から抽出すること。LocareやSaleshubは運営側なので対象外とする。
   HP情報の抽出については以下の優先順位で情報を収集すること：
   - まず提供されたWebサイト情報から各項目を抽出する。
   - 情報が不明・不足している項目は「※情報なし」と記載する。
   - PDFファイルは参照しないこと。
   - Locare・Saleshub等、運営側の情報は含めないこと。

2. チェックリスト31項目について、事実に基づき以下の基準で評価してください。
   - 「〇」：基準を満たしている／加点要素あり
   - 「△」：一部不足／普通
   - 「✕」：基準未達／改善が必要
   - 判断できない場合は空欄
   必ず全31項目に対していずれかを入力すること。
   △または✕の場合はcommentフィールドに理由を必ず記載すること。

3. HP情報と商談情報の差分を明確にしてください。

4. 商談の文字起こしを深く分析し、「全体統括コメント」を【600〜800字程度】で作成してください。
   - 単なる要約や浅い感想は禁止。以下の観点から戦略的なフィードバックを行うこと。
   - 営業スタイルと優位性（Good/More）：受身的になっていないか。顕在・潜在ニーズへのアプローチと他社比較の強みは伝わったか。
   - 顧客の「納得度」と価格提示の戦略（More）：料金への納得度、価格提示タイミングと費用対効果のロジックは適切だったか。
   - 前置きや冗長な背景説明は一切省き、結論ファーストで記述すること。
   - 出力テキスト内にアスタリスク記号（*）は絶対に使用しないこと。見出しや強調には【】や「」を使用すること。
   - 実際の文字起こし内の発言を根拠として端的に交えること。

5. 質疑応答（Q&A）は商談内の実際のやり取りを抽出し、質問と回答の要点を簡潔にまとめること。

【HP情報として抽出する項目（hp_infoのキーと完全一致させること）】
{hp_items_str}

【商談情報として抽出する項目（neg_infoのキーと完全一致させること）】
{neg_items_str}

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

【出力形式】
必ず以下のJSON形式のみで出力すること。前置き・説明文・マークダウンのコードブロック（```）は一切含めないこと。

{{
  "cl_company_name": "商談相手の企業名",
  "cl_attendee_name": "相手方担当者名",
  "cl_attendee_role": "相手方役職",
  "our_attendee_name": "自社担当者名",
  "impression": "全体統括コメント（600〜800字）",
  "hp_info": {{
    {hp_json_example}
  }},
  "neg_info": {{
    {neg_json_example}
  }},
  "checklist_evaluations": [
    {{"display_text": "評価項目名（15文字以上）", "evaluation": "〇", "comment": ""}},
    ...
  ],
  "questions_from_us": [
    {{"question": "質問内容", "answer": "回答内容"}},
    ...
  ],
  "questions_from_client": [
    {{"question": "質問内容", "answer": "回答内容"}},
    ...
  ]
}}
"""

    output = model_client.invoke(prompt)

    # output.content がリスト形式の場合と文字列の場合の両方に対応
    if isinstance(output.content, list):
        output_text = "".join(
            block["text"] for block in output.content
            if isinstance(block, dict) and block.get("type") == "text"
        )
    else:
        output_text = output.content

    # コードブロック（```json ... ```）が含まれていれば除去
    output_text = re.sub(r"```json\s*", "", output_text)
    output_text = re.sub(r"```\s*", "", output_text)
    output_text = output_text.strip()

    # JSONとしてパース
    parsed = json.loads(output_text)

    # hp_info / neg_info のキーが存在しない項目を空文字で補完
    if "hp_info" not in parsed:
        parsed["hp_info"] = {}
    if "neg_info" not in parsed:
        parsed["neg_info"] = {}
    for item in hp_items:
        parsed["hp_info"].setdefault(item, "※情報なし")
    for item in neg_items:
        parsed["neg_info"].setdefault(item, "※情報なし")

    return parsed


# ==============================
# 3. Googleスプレッドシートへの書き込み
# ==============================

def _normalize_evaluation(evaluation: str) -> str:
    """AIが出力する可能性のある〇△✕の表記ゆれをすべて正規化する"""
    evaluation = evaluation.replace("○", "〇")   # U+25CB → U+3007
    evaluation = evaluation.replace("◯", "〇")   # U+25EF → U+3007
    evaluation = evaluation.replace("O", "〇")   # 半角英字O
    evaluation = evaluation.replace("o", "〇")   # 半角英字o
    evaluation = evaluation.replace("×", "✕")   # U+00D7 → U+2715
    evaluation = evaluation.replace("✗", "✕")   # U+2717 → U+2715
    evaluation = evaluation.replace("X", "✕")   # 半角英字X
    evaluation = evaluation.replace("x", "✕")   # 半角英字x
    return evaluation.strip()


def fill_google_sheet(
    data: dict,
    service_account_info: dict,
    template_id: str,
    folder_id: str,
    hp_items: list,
    neg_items: list,
) -> str:
    """
    テンプレートスプレッドシートをコピーし、dataの内容を書き込んで返す。
    hp_items / neg_items は動的項目リスト（app.pyのセッションステートから渡す）。
    """

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)

    # --- テンプレートをコピー ---
    try:
        folder_info = drive_service.files().get(
            fileId=folder_id,
            supportsAllDrives=True,
            fields="driveId",
        ).execute()
        shared_drive_id = folder_info.get("driveId")
    except Exception:
        shared_drive_id = None

    copy_body = {
        "name": f"{data.get('cl_company_name', '名称未設定')}様_営業レポート",
        "parents": [folder_id],
    }
    if shared_drive_id:
        copy_body["driveId"] = shared_drive_id
        copy_body["teamDriveId"] = shared_drive_id

    copy_file = drive_service.files().copy(
        fileId=template_id,
        body=copy_body,
        supportsAllDrives=True,
        fields="id",
    ).execute()

    new_sheet_id = copy_file["id"]

    # --- gspreadで書き込み開始 ---
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(new_sheet_id)
    ws = sh.get_worksheet(0)

    batch_updates = []

    # 1. 基本ヘッダー情報
    batch_updates.append({"range": "B1", "values": [[data.get("cl_company_name", "")]]})
    batch_updates.append({"range": "B3", "values": [[data.get("website_url", "")]]})
    batch_updates.append({"range": "B4", "values": [[data.get("cl_attendee_name", "")]]})
    batch_updates.append({"range": "B5", "values": [[data.get("cl_attendee_role", "")]]})
    batch_updates.append({"range": "B6", "values": [[data.get("our_attendee_name", "")]]})
    batch_updates.append({"range": "A50", "values": [[data.get("impression", "")]]})

    # 2. HP情報
    #    A列（Row35〜）: 項目名
    #    C列（Row35〜）: 内容
    hp_info = data.get("hp_info", {})
    hp_label_values   = [[item] for item in hp_items]
    hp_content_values = [[hp_info.get(item, "")] for item in hp_items]
    hp_end_row = 34 + len(hp_items)
    batch_updates.append({"range": f"A35:A{hp_end_row}", "values": hp_label_values})
    batch_updates.append({"range": f"C35:C{hp_end_row}", "values": hp_content_values})

    # 3. 商談での情報整理
    #    F列（Row35〜）: 項目名
    #    H列（Row35〜）: 内容
    neg_info = data.get("neg_info", {})
    neg_label_values   = [[item] for item in neg_items]
    neg_content_values = [[neg_info.get(item, "")] for item in neg_items]
    neg_end_row = 34 + len(neg_items)
    batch_updates.append({"range": f"F35:F{neg_end_row}", "values": neg_label_values})
    batch_updates.append({"range": f"H35:H{neg_end_row}", "values": neg_content_values})

    # 4. Q&A セクション
    #    弊社→先方（Row 3〜16、最大14件）
    us_qa = data.get("questions_from_us", [])
    us_qa_values = [[qa["question"], qa["answer"]] for qa in us_qa[:14]]
    if us_qa_values:
        batch_updates.append({
            "range": f"I3:J{2 + len(us_qa_values)}",
            "values": us_qa_values,
        })

    #    先方→弊社（Row 18〜30、最大13件）
    client_qa = data.get("questions_from_client", [])
    client_qa_values = [[qa["question"], qa["answer"]] for qa in client_qa[:13]]
    if client_qa_values:
        batch_updates.append({
            "range": f"I18:J{17 + len(client_qa_values)}",
            "values": client_qa_values,
        })

    # 5. チェックリスト評価（Row 130〜160固定・31項目）
    #    D列のテキストと照合し、G列に評価・J列に備考を書き込む
    TARGET_ROWS = list(range(130, 161))
    sheet_d_col = ws.col_values(4)  # D列を全取得

    for item in data.get("checklist_evaluations", []):
        for row_num in TARGET_ROWS:
            d_text = sheet_d_col[row_num - 1] if row_num - 1 < len(sheet_d_col) else ""
            if item["display_text"][:15] in str(d_text):
                evaluation = _normalize_evaluation(item["evaluation"])

                # G列：評価（〇△✕）
                batch_updates.append({
                    "range": f"G{row_num}",
                    "values": [[evaluation]],
                })
                # J列：備考（△または✕のときのみ）
                comment = item.get("comment", "")
                if comment and evaluation in ["△", "✕"]:
                    batch_updates.append({
                        "range": f"J{row_num}",
                        "values": [[comment]],
                    })
                break

    # 全データを一括書き込み
    ws.batch_update(batch_updates)

    return f"https://docs.google.com/spreadsheets/d/{new_sheet_id}"