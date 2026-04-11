import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
import os
import re
from utils import extract_text_from_pdf, fetch_website_content
try:
    from report_generator import (
        generate_report_content,
        fill_google_sheet,
        evaluate_checklist_only,
        write_evaluation_to_existing_sheet,
    )
except ImportError as _e:
    import streamlit as st
    st.error(
        f"report_generator.py のインポートに失敗しました。\n\n"
        f"GitHubに最新の report_generator.py がpushされているか確認してください。\n\n"
        f"詳細: {_e}"
    )
    st.stop()
from langchain_google_genai import ChatGoogleGenerativeAI

st.set_page_config(page_title="営業レポート作成AI", layout="wide")
st.title("🗒️ 営業レポート作成AI")

# ==============================
# カテゴリ別デフォルト項目定義
# ==============================
CATEGORY_PRESETS = {
    "営業代行業界": {
        "hp": [
            "サービス概要", "サービスのビジネスモデル", "サービスの強み",
            "他社との違い", "料金プラン・料金形態", "最低料金", "最短契約期間",
            "導入企業の声、情報", "類似サービス、料金の競合企業",
            "コールスタッフの求人掲載有無", "コールスタッフ求人URL",
            "転職口コミサイトの声（退職率など）",
        ],
        "neg": [
            "サービス概要", "サービスのビジネスモデル", "サービスの強み",
            "他社との違い", "料金プラン・料金形態", "最低料金", "最短契約期間",
            "途中解約有無", "リスト作成ツール", "CTI", "競合会社認識",
            "1時間当たり架電件数・アポ率エビデンス",
        ],
    },
    "顧問サービス": {
        "hp": [
            "サービス概要", "サービスのビジネスモデル", "サービスの強み",
            "他社との違い", "料金プラン・料金形態", "初期費用有無",
            "最低料金", "最短契約期間", "競合企業、類似サービス",
            "顧問、会員登録人数", "依頼業務内容", "導入実績",
        ],
        "neg": [
            "サービス概要", "サービスのビジネスモデル", "サービスの強み",
            "他社との違い", "料金プラン・料金形態", "初期費用有無",
            "最低料金", "最短契約期間", "競合企業、類似サービス",
            "顧問、会員登録人数", "依頼業務内容", "導入実績",
        ],
    },
    "ISOコンサル": {
        "hp": [
            "サービス概要", "サービスのビジネスモデル", "サービスの強み",
            "他社との違い", "料金プラン・料金形態", "初期費用有無",
            "最低料金", "最短契約期間", "競合企業、類似サービス",
            "支援体制", "支援範囲（認証取得、内部監査、審査同席、育成等の可否）", "導入実績",
        ],
        "neg": [
            "サービス概要", "サービスのビジネスモデル", "サービスの強み",
            "他社との違い", "料金プラン・料金形態", "初期費用有無",
            "最低料金", "最短契約期間", "競合企業、類似サービス",
            "支援体制", "支援範囲（認証取得、内部監査、審査同席、育成等の可否）", "導入実績",
        ],
    },
    "カスタム（自由入力）": {"hp": [], "neg": []},
}

# チェックリストのカテゴリ別行範囲定義（スプレッドシートの実構造に基づく）
CHECKLIST_CATEGORIES = {
    "商談前IS":   {"rows": list(range(72, 82)),   "label": "商談前IS（No.1〜10）"},
    "商談態勢":   {"rows": list(range(83, 105)),  "label": "商談態勢（No.1〜22）"},
    "営業人間力": {"rows": list(range(105, 130)), "label": "営業人間力（No.1〜25）"},
    "商談対応力": {"rows": list(range(130, 161)), "label": "商談対応力（No.1〜31）"},
    "商談後（メール）":   {"rows": list(range(162, 167)), "label": "商談後フォロー（メール評価・No.1〜5）"},
    "商談後（全体評価）": {"rows": list(range(167, 173)), "label": "商談後 全体評価（文字起こし評価・No.6〜11）"},
}

# ==============================
# サイドバー（両モード共通）
# ==============================
with st.sidebar:
    st.header("設定")

    GEMINI_MODELS = {
        "Gemini 2.5 Pro（高精度）":        "gemini-2.5-pro",
        "Gemini 2.5 Flash（高速・推奨）":  "gemini-2.5-flash",
        "Gemini 2.5 Flash Lite（最軽量）": "gemini-2.5-flash-lite",
        "Gemini 2.0 Flash（安定版）":      "gemini-2.0-flash",
        "Gemini 3 Pro（最新・最高精度）":  "gemini-3-pro-preview",
        "Gemini 3 Flash（最新・高速）":    "gemini-3-flash-preview",
    }
    selected_label = st.selectbox("使用するGeminiモデル", list(GEMINI_MODELS.keys()), index=1)
    selected_model_name = GEMINI_MODELS[selected_label]
    st.caption(f"モデルID: `{selected_model_name}`")

    api_key = st.text_input("Gemini API Key", type="password")

    st.markdown("---")
    st.markdown("### 📂 調査カテゴリ・項目設定")
    st.caption("新規レポート作成タブで使用します")

    selected_category = st.selectbox(
        "調査カテゴリを選択",
        list(CATEGORY_PRESETS.keys()),
        index=0,
        key="selected_category",
    )

    if st.session_state.get("last_category") != selected_category:
        st.session_state.hp_items  = list(CATEGORY_PRESETS[selected_category]["hp"])
        st.session_state.neg_items = list(CATEGORY_PRESETS[selected_category]["neg"])
        st.session_state.last_category = selected_category

    st.markdown("#### 📄 HP情報の項目")
    for i, item in enumerate(st.session_state.hp_items):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.text(f"・{item}")
        with c2:
            if st.button("✕", key=f"hp_del_{i}"):
                st.session_state.hp_items.pop(i)
                st.rerun()
    new_hp = st.text_input("項目を追加", key="new_hp_input", placeholder="例：資本金")
    if st.button("➕ HP項目を追加", key="hp_add_btn", use_container_width=True):
        if new_hp.strip():
            st.session_state.hp_items.append(new_hp.strip())
            st.rerun()

    st.markdown("---")
    st.markdown("#### 💬 商談情報の項目")
    for i, item in enumerate(st.session_state.neg_items):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.text(f"・{item}")
        with c2:
            if st.button("✕", key=f"neg_del_{i}"):
                st.session_state.neg_items.pop(i)
                st.rerun()
    new_neg = st.text_input("項目を追加", key="new_neg_input", placeholder="例：担当者の決裁権限")
    if st.button("➕ 商談項目を追加", key="neg_add_btn", use_container_width=True):
        if new_neg.strip():
            st.session_state.neg_items.append(new_neg.strip())
            st.rerun()

    st.markdown("---")
    st.markdown("### マニュアルファイル")
    base_dir = os.path.dirname(os.path.abspath(__file__))
    DEFAULT_MANUAL_PATH = os.path.join(
        base_dir, "8ba0d12e-f2ee-4002-9533-54a0940f4eaa_営業レポートマニュアル.pdf"
    )
    if os.path.exists(DEFAULT_MANUAL_PATH):
        st.success("✅ マニュアル読み込み済み")
        manual_file = open(DEFAULT_MANUAL_PATH, "rb")
    else:
        st.error("デフォルトマニュアルが見つかりません。")
        manual_file = None

# ==============================
# タブ切り替え
# ==============================
tab_new, tab_eval = st.tabs(["🆕 新規レポート作成", "✏️ 評価を既存シートに追記"])

# ==================================================
# タブ1：新規レポート作成
# ==================================================
with tab_new:
    st.markdown("商談の文字起こしとHP URLから、新しいスプレッドシートを作成します。")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. 商談情報の入力")
        transcript_text_input = st.text_area(
            "商談文字起こしをここに貼り付けてください",
            height=350,
            placeholder="文字起こしテキストをここに貼り付けてください...",
            key="new_transcript",
        )
        website_url = st.text_input(
            "商談相手の企業HP URL（任意）",
            placeholder="https://example.com",
            key="new_website_url",
        )

    with col2:
        st.subheader("2. 生成")

        with st.expander(f"📋 現在の項目設定（カテゴリ：{selected_category}）"):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**HP情報の項目**")
                for item in st.session_state.get("hp_items", []):
                    st.markdown(f"- {item}")
            with c2:
                st.markdown("**商談情報の項目**")
                for item in st.session_state.get("neg_items", []):
                    st.markdown(f"- {item}")

        if st.button("🚀 レポート生成を開始", type="primary", use_container_width=True, key="btn_new"):
            if not api_key:
                st.error("APIキーを入力してください。")
            elif not transcript_text_input:
                st.error("商談テキストを入力してください。")
            elif not manual_file:
                st.error("マニュアルファイルが必要です。")
            elif not st.session_state.get("hp_items"):
                st.error("HP情報の項目が1件もありません。サイドバーで追加してください。")
            elif not st.session_state.get("neg_items"):
                st.error("商談情報の項目が1件もありません。サイドバーで追加してください。")
            else:
                manual_text = ""
                with st.spinner("マニュアルを読み込み中..."):
                    try:
                        manual_text = extract_text_from_pdf(manual_file)
                    except Exception as e:
                        st.error(f"マニュアルPDFの読み込みに失敗しました: {e}")

                website_text = ""
                if website_url:
                    try:
                        with st.spinner("Webサイト情報を取得中..."):
                            website_text = fetch_website_content(website_url)
                            if not website_text:
                                st.warning("指定されたURLから情報を取得できませんでした。")
                            else:
                                st.success(f"Webサイト情報を取得しました ({len(website_text)}文字)")
                    except Exception as e:
                        st.warning(f"Webサイト情報の取得中にエラーが発生しました: {e}")

                with st.spinner("AIが営業代行レポートを分析中..."):
                    try:
                        llm = ChatGoogleGenerativeAI(model=selected_model_name, google_api_key=api_key)

                        data = generate_report_content(
                            transcript_text_input,
                            manual_text,
                            website_text,
                            "",
                            llm,
                            hp_items=st.session_state.hp_items,
                            neg_items=st.session_state.neg_items,
                        )
                        st.success("分析完了！スプレッドシートを作成します...")

                        gcp_info   = st.secrets["gcp_service_account"]
                        TEMPLATE_ID = st.secrets["google_drive"]["template_id"]
                        FOLDER_ID   = st.secrets["google_drive"]["folder_id"]

                        data["website_url"] = website_url
                        sheet_url = fill_google_sheet(
                            data, gcp_info, TEMPLATE_ID, FOLDER_ID,
                            hp_items=st.session_state.hp_items,
                            neg_items=st.session_state.neg_items,
                        )

                        st.balloons()
                        st.success("レポートが共有ドライブに作成されました！")
                        st.link_button("🔥 完成したスプレッドシートを開く", sheet_url)

                        with st.expander("抽出データ（JSON）の確認"):
                            st.json(data)

                    except Exception as e:
                        st.error(f"エラーが発生しました: {e}")

# ==================================================
# タブ2：評価を既存シートに追記
# ==================================================
with tab_eval:
    st.markdown("既存のスプレッドシートに、文字起こしまたはメール文章から評価を追記します。")
    st.info("💡 追記先のスプレッドシートはサービスアカウントと共有済みである必要があります。")

    col_e1, col_e2 = st.columns(2)

    with col_e1:
        st.subheader("1. 追記先シートの指定")

        sheet_mode = st.radio(
            "追記先を選択",
            ["既存のスプレッドシートURLを入力", "テンプレートから新規作成して追記"],
            key="sheet_mode",
        )

        if sheet_mode == "既存のスプレッドシートURLを入力":
            existing_sheet_url = st.text_input(
                "スプレッドシートのURL",
                placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit",
                key="existing_sheet_url",
            )
        else:
            st.caption("テンプレートをコピーして新規作成し、そこに評価を書き込みます。")
            existing_sheet_url = None

        st.subheader("2. 評価対象テキストの入力")
        transcript_eval_input = st.text_area(
            "📹 商談文字起こし（営業人間力・商談対応力の評価に使用）",
            height=200,
            placeholder="商談文字起こしをここに貼り付けてください...",
            key="eval_transcript",
        )
        email_eval_input = st.text_area(
            "📧 メール文章（商談前IS・商談後フォローの評価に使用）",
            height=150,
            placeholder="メール本文をここに貼り付けてください...",
            key="eval_email",
        )

    with col_e2:
        st.subheader("3. 評価の範囲と種類を選択")

        st.markdown("**📹 文字起こしから評価できるカテゴリ**")
        TRANSCRIPT_CATS = ["営業人間力", "商談対応力", "商談後（全体評価）"]
        EMAIL_CATS      = ["商談前IS", "商談後（メール）"]

        selected_cats = {}
        for cat_name in TRANSCRIPT_CATS:
            cat_info = CHECKLIST_CATEGORIES[cat_name]
            default_on = (cat_name == "商談対応力")
            selected_cats[cat_name] = st.checkbox(
                cat_info["label"], value=default_on, key=f"cat_{cat_name}"
            )

        st.markdown("**📧 メール文章から評価できるカテゴリ**")
        email_available = bool(email_eval_input.strip())  # メール入力があるか
        if not email_available:
            st.caption("⚠️ メール文章を入力すると選択できます")
        for cat_name in EMAIL_CATS:
            cat_info = CHECKLIST_CATEGORIES[cat_name]
            selected_cats[cat_name] = st.checkbox(
                cat_info["label"],
                value=False,
                key=f"cat_{cat_name}",
                disabled=not email_available,  # メールなしは選択不可
            )

        st.markdown("---")
        st.markdown("**書き込む内容**")
        write_evaluation = st.checkbox("〇△✕ の評価を書き込む", value=True, key="write_eval")
        write_comment    = st.checkbox("備考欄のコメントを書き込む", value=True, key="write_comment")

        st.markdown("---")

        if st.button("✏️ 評価を追記する", type="primary", use_container_width=True, key="btn_eval"):
            # バリデーション
            errors = []
            if not api_key:
                errors.append("APIキーを入力してください。")
            if not transcript_eval_input.strip() and not email_eval_input.strip():
                errors.append("文字起こしまたはメール文章を少なくとも1つ入力してください。")
            if not any(selected_cats.values()):
                errors.append("評価するカテゴリを1つ以上選択してください。")
            if not write_evaluation and not write_comment:
                errors.append("「評価を書き込む」または「備考を書き込む」を選択してください。")
            if sheet_mode == "既存のスプレッドシートURLを入力" and not existing_sheet_url:
                errors.append("スプレッドシートのURLを入力してください。")
            mail_cats_selected = any(selected_cats.get(c) for c in EMAIL_CATS)
            if mail_cats_selected and not email_eval_input.strip():
                errors.append("商談前IS・商談後フォローの評価にはメール文章の入力が必要です。")

            for err in errors:
                st.error(err)

            if not errors:
                # 対象カテゴリのrow範囲を収集
                target_rows = []
                for cat_name, is_selected in selected_cats.items():
                    if is_selected:
                        target_rows.extend(CHECKLIST_CATEGORIES[cat_name]["rows"])

                checklist_result = []
                with st.spinner("AIがチェックリストを評価中..."):
                    try:
                        llm = ChatGoogleGenerativeAI(model=selected_model_name, google_api_key=api_key)
                        checklist_result = []

                        # 文字起こし系カテゴリ
                        transcript_cats_selected = [c for c in TRANSCRIPT_CATS if selected_cats.get(c)]
                        if transcript_cats_selected and transcript_eval_input.strip():
                            result_t = evaluate_checklist_only(
                                text=transcript_eval_input,
                                model_client=llm,
                                target_categories=transcript_cats_selected,
                            )
                            checklist_result.extend(result_t)

                        # メール系カテゴリ
                        email_cats_selected = [c for c in EMAIL_CATS if selected_cats.get(c)]
                        if email_cats_selected and email_eval_input.strip():
                            result_e = evaluate_checklist_only(
                                text=email_eval_input,
                                model_client=llm,
                                target_categories=email_cats_selected,
                            )
                            checklist_result.extend(result_e)

                        # display_textの先頭15文字で重複除去（メール系を優先して残す）
                        seen_keys = {}
                        for item in checklist_result:
                            key = item.get("display_text", "")[:15]
                            seen_keys[key] = item  # 後から来たもの（メール系）で上書き
                        checklist_result = list(seen_keys.values())

                        st.success(f"評価完了！{len(checklist_result)}件の項目を評価しました。")

                        # 評価結果プレビュー
                        with st.expander("📊 評価結果プレビュー（書き込み前に確認）", expanded=True):
                            for item in checklist_result:
                                mark = item.get("evaluation", "")
                                if not mark:  # 空欄は表示・書き込みスキップ
                                    continue
                                if mark == "〇":
                                    color = "green"
                                elif mark == "△":
                                    color = "orange"
                                else:
                                    color = "red"
                                comment_text = f"　└ {item['comment']}" if item.get("comment") else ""
                                st.markdown(
                                    f":{color}[**{mark}**]　{item.get('display_text', '')}"
                                    + (f"  \n{comment_text}" if comment_text else "")
                                )

                    except Exception as e:
                        st.error(f"AI評価中にエラーが発生しました: {e}")

                if checklist_result:
                    with st.spinner("スプレッドシートに書き込み中..."):
                        try:
                            gcp_info    = st.secrets["gcp_service_account"]
                            TEMPLATE_ID = st.secrets["google_drive"]["template_id"]
                            FOLDER_ID   = st.secrets["google_drive"]["folder_id"]

                            # 既存シートのIDを解決
                            if sheet_mode == "既存のスプレッドシートURLを入力":
                                match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", existing_sheet_url)
                                if not match:
                                    st.error("有効なスプレッドシートURLを入力してください。")
                                    st.stop()
                                target_sheet_id = match.group(1)
                            else:
                                target_sheet_id = None  # Noneのときは新規コピー作成

                            final_url = write_evaluation_to_existing_sheet(
                                checklist_result=checklist_result,
                                target_rows=target_rows,
                                service_account_info=gcp_info,
                                template_id=TEMPLATE_ID,
                                folder_id=FOLDER_ID,
                                existing_sheet_id=target_sheet_id,
                                write_evaluation=write_evaluation,
                                write_comment=write_comment,
                            )

                            st.balloons()
                            st.success("書き込みが完了しました！")
                            st.link_button("🔥 スプレッドシートを開く", final_url)

                        except Exception as e:
                            st.error(f"書き込み中にエラーが発生しました: {e}")

st.markdown("---")
st.caption("Powered by Streamlit & LangChain")