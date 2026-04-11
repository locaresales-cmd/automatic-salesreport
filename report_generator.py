import json
import re
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from typing import List, Dict, Optional

# ==============================
# チェックリスト項目の定義
# カテゴリ別にD列テキストの先頭15文字で照合する
# ==============================
CHECKLIST_ITEMS_BY_CATEGORY = {
    "商談前IS": [
        "HP問い合わせに対してのレスポンス速度",
        "商談前段階における連絡の基本的な返信速度（平均）",
        "メールでのラリー回数は少なく、スムーズに日程調整ができた。",
        "問い合わせに対しての返信はテンプレートではなく自社に合わせた形での連絡",
        "メールの文面は言葉遣いが丁寧で、誤字脱字がなく、不快感や違和感を感じる",
        "メールで記載されていた候補日時はスケジュー",
        "メールで事前に資料が共有された。",
        "メールでの連絡を希望していたにも関わらず電話をかけていないか",
        "メールに事前質問は記載されていたか？",
        "商談前のリマインド連絡があった。",
    ],
    "営業人間力": [
        "不快感を持たないか？(説明、対応、返答など)",
        "言動や対応に違和感・不審感は無いか？",
        "相手の業界の知識を持ち合わせているか？この人はわかっている！と思えるか",
        "相手の商材について、イメージを持っているか？近しい業界の事例などを伝え",
        "ビジネスモデルの理解ができているか？キャッシュポイントなどを認識して、",
        "○○社長/代表/様と呼んでいるか",
        "貢献できます。など断言ができているか？※嘘はNG",
        "リレーションは築けているか？(最後に雑談等)",
    ],
    "商談対応力": [
        "冒頭大きな声でさわやかに挨拶できているか？",
        "音声の聞こえ方やカメラ画面の見え方に違和感などがないか確認しているか？",
        "いきなり質問から始めるのではなく商談の方向性や流れを明確にしてからスタ",
        "商談担当者の自己紹介は行っているか？",
        "打ち合わせのタイムスケジュール感を冒頭ですり合わせているか？",
        "商談をされる側が不快感を示すような一方的な質問責めをしていないか？",
        "HPや事前情報を調べれば把握できるような質問をしていないか？",
        "質問に回答した後の相槌や反応はしているか？",
        "質問に回答した後のメモに時間を要して沈黙の時間を生んでいないか？",
        "質問回答後に「ありがとうございます」と御礼を伝えているか？",
        "自社の会社概要について話しているか？",
        "提供するサービスのビジネスモデルを分かりやすく簡潔に話しているか？",
        "ビジネスモデルは企業HP記載の内容と一致しているか？(虚偽、誇大記載は",
        "他社との違いや優位性について触れているか？",
        "事例共有は営業を受ける側に近い事例を準備し、説明しているか？",
        "具体的な数値効果について仮説を交えながらわかりやすく説明しているか？",
        "一方的なサービス説明に終始するのではなく、中間で質疑応答の時間を設けて",
        "サービスフローはスケジュール感を交えて説明できているか？",
        "料金プランは一部を切り出して説明するのではなく、全体でかかる費用を網羅",
        "リスクやデメリットについても説明しているか？",
        "営業を受けている側のニーズに合わせた提案を行っているか？",
        "質問に対して結論ファーストで簡潔に回答できているか？",
        "自社・他社情報を共有する際には情報セキュリティに十分に配慮できているか",
        "BANT情報を聞けているか？",
        "導入に向けた懸念点を聞き出しているか？（テストクロージング）",
        "金額懸念に対しての交渉に応じる姿勢は見せているか？",
        "ネクストアクションの提案やすり合わせを行っているか？",
        "決裁者を交えた次回打ち合わせ日程調整の打診を行っているか？",
        "検討期限を区切っているか？",
        "金額に対しての価値が伝わっており、価格に対し安いと思える価値訴求ができ",
        "強引なクロージングで終えるのではなく気分の良い終話だったか",
    ],
    "商談後（メール）": [
        "商談後のサンクスメールは最速で送られてきたか？",
        "サンクスメールは商談企業ごとに内容を合わせて送っているか？",
        "商談後に追加質問を送った後の回答返信速度",
        "商談で区切った検討期限日時にフォローの連絡があったか",
        "営業資料だけではなく事例や提案資料など検討材料になりえる資料の共有があったか",
    ],
    "商談後（全体評価）": [
        "会社自体に信頼性を感じることができたか？",
        "サービス自体に信頼性を持てると感じることはできたか？",
        "営業マンの対応は正直かつ誠実で好印象を持つことはできたか？",
        "相手の立場にたった商談だったと評価することはできるか？",
        "サービス内容はわかりやすく、商談の時間で理解しきることはできたか？",
        "営業代行会社への発注をしている企業に勧めたいと思えたか？",
    ],
}


# ==============================
# 共通ユーティリティ
# ==============================

def _normalize_evaluation(evaluation: str) -> str:
    """AIが出力する〇△✕の表記ゆれをすべて正規化する"""
    evaluation = evaluation.replace("○", "〇").replace("◯", "〇")
    evaluation = evaluation.replace("O", "〇").replace("o", "〇")
    evaluation = evaluation.replace("×", "✕").replace("✗", "✕")
    evaluation = evaluation.replace("X", "✕").replace("x", "✕")
    return evaluation.strip()


def _parse_json_output(raw_text: str) -> dict:
    """AIの出力からJSONを安全にパースする"""
    text = re.sub(r"```json\s*", "", raw_text)
    text = re.sub(r"```\s*", "", text)
    return json.loads(text.strip())


def _get_gspread_client(service_account_info: dict):
    """gspreadクライアントとCredentialsを返す"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc, creds


def _copy_template(creds, template_id: str, folder_id: str, name: str) -> str:
    """テンプレートをコピーして新しいシートIDを返す"""
    drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)

    try:
        folder_info = drive_service.files().get(
            fileId=folder_id, supportsAllDrives=True, fields="driveId"
        ).execute()
        shared_drive_id = folder_info.get("driveId")
    except Exception:
        shared_drive_id = None

    copy_body = {"name": name, "parents": [folder_id]}
    if shared_drive_id:
        copy_body["driveId"] = shared_drive_id
        copy_body["teamDriveId"] = shared_drive_id

    copy_file = drive_service.files().copy(
        fileId=template_id,
        body=copy_body,
        supportsAllDrives=True,
        fields="id",
    ).execute()

    return copy_file["id"]


# ==============================
# 1. データ構造の定義（新規作成モード用）
# ==============================

class QAPair:
    def __init__(self, question: str, answer: str):
        self.question = question
        self.answer = answer

class ChecklistItem:
    def __init__(self, display_text: str, evaluation: str, comment: str):
        self.display_text = display_text
        self.evaluation = evaluation
        self.comment = comment


# ==============================
# 2. AIによるレポート内容生成（新規作成モード）
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
    """文字起こし・HP情報・資料をもとにAIがフルレポートを生成する"""

    hp_items_str  = "\n".join(f"  - {item}" for item in hp_items)
    neg_items_str = "\n".join(f"  - {item}" for item in neg_items)
    hp_json_example  = ", ".join(f'"{item}": "内容"' for item in hp_items[:3]) + ", ..."
    neg_json_example = ", ".join(f'"{item}": "内容"' for item in neg_items[:3]) + ", ..."

    # 商談対応力31項目をプロンプトに埋め込む
    checklist_str = "\n".join(CHECKLIST_ITEMS_BY_CATEGORY["商談対応力"])

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

2. チェックリスト項目について、事実に基づき以下の基準で評価してください。
   - 「〇」：基準を満たしている／加点要素あり
   - 「△」：一部不足／普通
   - 「✕」：基準未達／改善が必要
   - 判断できない場合は空欄
   必ず全項目に対していずれかを入力すること。
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

【評価項目（チェックリスト）】
{checklist_str}

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
  "hp_info": {{ {hp_json_example} }},
  "neg_info": {{ {neg_json_example} }},
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
    output_text = (
        "".join(b["text"] for b in output.content if isinstance(b, dict) and b.get("type") == "text")
        if isinstance(output.content, list)
        else output.content
    )

    parsed = _parse_json_output(output_text)

    # キーが欠けている項目を補完
    parsed.setdefault("hp_info", {})
    parsed.setdefault("neg_info", {})
    for item in hp_items:
        parsed["hp_info"].setdefault(item, "※情報なし")
    for item in neg_items:
        parsed["neg_info"].setdefault(item, "※情報なし")

    return parsed


# ==============================
# 3. 新規スプレッドシートへの書き込み（新規作成モード）
# ==============================

def fill_google_sheet(
    data: dict,
    service_account_info: dict,
    template_id: str,
    folder_id: str,
    hp_items: list,
    neg_items: list,
) -> str:
    """テンプレートをコピーしてフルレポートを書き込む"""

    gc, creds = _get_gspread_client(service_account_info)
    new_sheet_id = _copy_template(
        creds, template_id, folder_id,
        f"{data.get('cl_company_name', '名称未設定')}様_営業レポート"
    )

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

    # 2. HP情報（A列=項目名、C列=内容）
    hp_info = data.get("hp_info", {})
    hp_end = 34 + len(hp_items)
    batch_updates.append({"range": f"A35:A{hp_end}", "values": [[k] for k in hp_items]})
    batch_updates.append({"range": f"C35:C{hp_end}", "values": [[hp_info.get(k, "")] for k in hp_items]})

    # 3. 商談情報（F列=項目名、H列=内容）
    neg_info = data.get("neg_info", {})
    neg_end = 34 + len(neg_items)
    batch_updates.append({"range": f"F35:F{neg_end}", "values": [[k] for k in neg_items]})
    batch_updates.append({"range": f"H35:H{neg_end}", "values": [[neg_info.get(k, "")] for k in neg_items]})

    # 4. Q&A
    us_qa = data.get("questions_from_us", [])
    us_qa_values = [[qa["question"], qa["answer"]] for qa in us_qa[:14]]
    if us_qa_values:
        batch_updates.append({"range": f"I3:J{2 + len(us_qa_values)}", "values": us_qa_values})

    client_qa = data.get("questions_from_client", [])
    client_qa_values = [[qa["question"], qa["answer"]] for qa in client_qa[:13]]
    if client_qa_values:
        batch_updates.append({"range": f"I18:J{17 + len(client_qa_values)}", "values": client_qa_values})

    # 5. チェックリスト評価（商談対応力 Row130〜160、H列=評価、J列=備考）
    TARGET_ROWS = list(range(130, 161))
    sheet_d_col = ws.col_values(4)

    for item in data.get("checklist_evaluations", []):
        for row_num in TARGET_ROWS:
            d_text = sheet_d_col[row_num - 1] if row_num - 1 < len(sheet_d_col) else ""
            if item["display_text"][:15] in str(d_text):
                evaluation = _normalize_evaluation(item["evaluation"])
                batch_updates.append({"range": f"G{row_num}", "values": [[evaluation]]})
                comment = item.get("comment", "")
                if comment and evaluation in ["△", "✕"]:
                    batch_updates.append({"range": f"J{row_num}", "values": [[comment]]})
                break

    ws.batch_update(batch_updates)
    return f"https://docs.google.com/spreadsheets/d/{new_sheet_id}"


# ==============================
# 4. チェックリストのみAI評価（評価追記モード）
# ==============================

def evaluate_checklist_only(
    text: str,
    model_client,
    target_categories: List[str],
) -> List[dict]:
    """
    文字起こしまたはメール文章から、指定カテゴリのチェックリストのみを評価する。
    文字起こし（商談内容）でも、メール文章（商談前後の連絡）でも対応。
    """

    # 対象カテゴリの評価項目を収集
    all_items = []
    # all_items組み立ての直後に追加
    criteria_section = ""
    if "商談前IS" in target_categories:
        criteria_section += """
    【商談前ISの評価基準（必ずこの基準に従うこと）】
    No.1 HP問い合わせに対してのレスポンス速度
    〇：6時間以内　△：24時間以内　✕：24時間以上
    No.2 商談前段階における連絡の基本的な返信速度（平均）
    〇：6時間以内　△：24時間以内　✕：24時間以上
    No.3 メールでのラリー回数は少なく、スムーズに日程調整ができた。
    〇：ラリー回数1〜2　△：ラリー回数2〜3　✕：3往復以上のラリー（✕は備考に具体的内容を記載）
    No.4 問い合わせへの返信はテンプレートではなく自社に合わせた形での連絡だったか。
    〇：自社に合わせた返信　✕：テンプレートと判断できる
    No.5 メールの文面は言葉遣いが丁寧で、誤字脱字がなく、不快感や違和感がなかったか。
    〇：一切懸念なし　△：若干の懸念を感じる　✕：違和感・懸念を強く感じる（△✕は備考に記載）
    No.6 メールで記載されていた候補日時はスケジューラーURLではなくテキストで送ってきた。
    〇：テキストのみ候補日程多　△：テキストのみ候補日程少　✕：スケジューラー添付のみ
    No.7 メールで事前に資料が共有された。
    〇：2つ以上の共有あり　△：1つのみ共有有　✕：共有なし（備考に資料内容を記載）
    No.8 メールでの連絡を希望していたにも関わらず電話をかけていないか
    〇：希望通りメールで連携　△：電話をかけてきた　✕：電話を繰り返しかけてきた
    No.9 メールに事前質問は記載されていたか？
    〇：3つ以上の質問記載　△：〜3つ以内の質問記載　✕：質問記載なし（備考に質問内容を記載）
    No.10 商談前のリマインド連絡があった。
    〇：前日と当日リマインド　△：当日リマインド　✕：リマインドなし（備考にタイミングを記載）
    """

    if "商談後（メール）" in target_categories:
        criteria_section += """
    【商談後フォローの評価基準（必ずこの基準に従うこと）】
    No.1 商談後のサンクスメールは最速で送られてきたか？
    〇：6時間以内　△：24時間以内　✕：24時間以上
    No.2 サンクスメールは商談企業ごとに内容を合わせて送っているか？
    〇：合わせている　△：一部合わせている　✕：定型文のみ
    No.3 商談後に追加質問を送った後の回答返信速度
    〇：6時間以内　△：24時間以内　✕：24時間以上
    No.4〜5（検討期限フォロー・資料共有）
    〇：明確に確認できる　△：判断が難しい・一部確認できる　✕：確認できない・問題あり
    """

    if "商談後（全体評価）" in target_categories:
        criteria_section += """
    【商談後 全体評価の評価基準（文字起こしをもとに評価すること）】
    以下6項目は商談の内容・印象・理解度・推薦意向を文字起こしから判断すること。
    〇：明確に確認できる・好印象　△：判断が難しい・一部確認できる　✕：確認できない・問題あり・悪印象
    """

    if not criteria_section:
        criteria_section = """
    【汎用評価基準】
    〇：基準を満たしている／確認できる
    △：一部不足／判断が難しい
    ✕：基準未達／確認できない／問題あり
    """
    for cat in target_categories:
        items = CHECKLIST_ITEMS_BY_CATEGORY.get(cat, [])
        for item in items:
            all_items.append({"category": cat, "text": item})

    items_str = "\n".join(
    f"{i['text']}" for i in all_items
    )

    prompt = f"""
あなたはプロの営業評価者です。
以下のテキスト（商談文字起こし、またはメール文章）をもとに、
チェックリストの各項目を〇△✕の三段階で評価してください。

{criteria_section}

テキストから判断できない場合は evaluation を空文字("")にし、commentに「判断不可」と記載すること。判断できないのに無理に✕をつけないこと。

【評価項目】
{items_str}

【評価対象テキスト】
{text}

【出力形式】
以下のJSON配列のみを出力すること。前置き・説明文・コードブロック（```）は含めないこと。

[
  {{
    "display_text": "評価項目のテキスト（15文字以上含める）",
    "evaluation": "〇",
    "comment": ""
  }},
  {{
    "display_text": "評価項目のテキスト",
    "evaluation": "✕",
    "comment": "理由を1文で記載"
  }},
  ...
]
"""

    output = model_client.invoke(prompt)
    output_text = (
        "".join(b["text"] for b in output.content if isinstance(b, dict) and b.get("type") == "text")
        if isinstance(output.content, list)
        else output.content
    )

    # JSONパース
    text_clean = re.sub(r"```json\s*", "", output_text)
    text_clean = re.sub(r"```\s*", "", text_clean).strip()
    result = json.loads(text_clean)

    # 評価の正規化
    for item in result:
        item["evaluation"] = _normalize_evaluation(item.get("evaluation", ""))

    return result


# ==============================
# 5. 既存シートへの評価書き込み（評価追記モード）
# ==============================

def write_evaluation_to_existing_sheet(
    checklist_result: List[dict],
    target_rows: List[int],
    service_account_info: dict,
    template_id: str,
    folder_id: str,
    existing_sheet_id: Optional[str],
    write_evaluation: bool,
    write_comment: bool,
) -> str:
    """
    既存シート（またはテンプレートの新規コピー）の指定行範囲に
    チェックリスト評価（G列）と備考（J列）を書き込む。

    existing_sheet_id が None の場合はテンプレートをコピーして新規作成する。
    書き込む列：
        G列 = 評価（〇△✕）
        J列 = 備考
    照合方法：D列テキストの先頭15文字 vs checklist_resultのdisplay_text先頭15文字
    """

    gc, creds = _get_gspread_client(service_account_info)

    # シートの解決（既存 or 新規コピー）
    if existing_sheet_id:
        sh = gc.open_by_key(existing_sheet_id)
        result_url = f"https://docs.google.com/spreadsheets/d/{existing_sheet_id}"
    else:
        new_id = _copy_template(creds, template_id, folder_id, "営業レポート（評価追記）")
        sh = gc.open_by_key(new_id)
        result_url = f"https://docs.google.com/spreadsheets/d/{new_id}"

    ws = sh.get_worksheet(0)

    # D列を全取得（照合用）
    sheet_d_col = ws.col_values(4)

    batch_updates = []

    for item in checklist_result:
        raw_text = item.get("display_text", "")
        clean_text = re.sub(r"^\[.+?\]\s*", "", raw_text)  # [xxx]プレフィックスを除去
        item_text_key = clean_text[:15]
        evaluation    = _normalize_evaluation(item.get("evaluation", ""))
        comment       = item.get("comment", "")
        ALL_CHECKLIST_ROWS = list(range(72, 173))  # シート全体のチェックリスト行
        for row_num in ALL_CHECKLIST_ROWS:
            d_text = sheet_d_col[row_num - 1] if row_num - 1 < len(sheet_d_col) else ""
            if item_text_key in str(d_text):
                # G列：評価（〇△✕）
                if write_evaluation:
                    batch_updates.append({
                        "range": f"G{row_num}",
                        "values": [[evaluation]],
                    })
                # J列：備考
                if write_comment and comment:
                    batch_updates.append({
                        "range": f"J{row_num}",
                        "values": [[comment]],
                    })
                break

    if batch_updates:
        ws.batch_update(batch_updates)

    return result_url