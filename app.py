import sys
# Force UTF-8 encoding for stdout/stderr to prevent UnicodeEncodeError in restricted environments
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
import os
from utils import extract_text_from_pdf, fetch_website_content
from report_generator import generate_report_content, fill_google_sheet
from langchain_google_genai import ChatGoogleGenerativeAI

st.set_page_config(page_title="営業レポート作成AI", layout="wide")

st.title("🗒️ 営業レポート作成AI")
st.markdown("商談の文字起こしとマニュアルから、指定フォーマットの営業レポートを自動生成します。")

# ==============================
# カテゴリ別デフォルト項目定義
# ==============================
CATEGORY_PRESETS = {
    "営業代行業界": {
        "hp": [
            "サービス概要",
            "サービスのビジネスモデル",
            "サービスの強み",
            "他社との違い",
            "料金プラン・料金形態",
            "最低料金",
            "最短契約期間",
            "導入企業の声、情報",
            "類似サービス、料金の競合企業",
            "コールスタッフの求人掲載有無",
            "コールスタッフ求人URL",
            "転職口コミサイトの声（退職率など）",
        ],
        "neg": [
            "サービス概要",
            "サービスのビジネスモデル",
            "サービスの強み",
            "他社との違い",
            "料金プラン・料金形態",
            "最低料金",
            "最短契約期間",
            "途中解約有無",
            "リスト作成ツール",
            "CTI",
            "競合会社認識",
            "1時間当たり架電件数・アポ率エビデンス",
        ],
    },
    "顧問サービス": {
        "hp": [
            "サービス概要",
            "サービスのビジネスモデル",
            "サービスの強み",
            "他社との違い",
            "料金プラン・料金形態",
            "初期費用有無",
            "最低料金",
            "最短契約期間",
            "競合企業、類似サービス",
            "顧問、会員登録人数",
            "依頼業務内容",
            "導入実績",
        ],
        "neg": [
            "サービス概要",
            "サービスのビジネスモデル",
            "サービスの強み",
            "他社との違い",
            "料金プラン・料金形態",
            "初期費用有無",
            "最低料金",
            "最短契約期間",
            "競合企業、類似サービス",
            "顧問、会員登録人数",
            "依頼業務内容",
            "導入実績",
        ],
    },
    "ISOコンサル": {
        "hp": [
            "サービス概要",
            "サービスのビジネスモデル",
            "サービスの強み",
            "他社との違い",
            "料金プラン・料金形態",
            "初期費用有無",
            "最低料金",
            "最短契約期間",
            "競合企業、類似サービス",
            "支援体制",
            "支援範囲（認証取得、内部監査、審査同席、育成等の可否）",
            "導入実績",
        ],
        "neg": [
            "サービス概要",
            "サービスのビジネスモデル",
            "サービスの強み",
            "他社との違い",
            "料金プラン・料金形態",
            "初期費用有無",
            "最低料金",
            "最短契約期間",
            "競合企業、類似サービス",
            "支援体制",
            "支援範囲（認証取得、内部監査、審査同席、育成等の可否）",
            "導入実績",
        ],
    },
    "カスタム（自由入力）": {
        "hp": [],
        "neg": [],
    },
}

# ==============================
# サイドバー
# ==============================
with st.sidebar:
    st.header("設定")

    # --- モデル選択 ---
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

    # --- カテゴリ選択 ---
    st.markdown("### 調査カテゴリ・項目設定")
    selected_category = st.selectbox(
        "調査カテゴリを選択",
        list(CATEGORY_PRESETS.keys()),
        index=0,
        key="selected_category",
    )

    # カテゴリが切り替わったときにセッションステートをリセット
    if st.session_state.get("last_category") != selected_category:
        st.session_state.hp_items = list(CATEGORY_PRESETS[selected_category]["hp"])
        st.session_state.neg_items = list(CATEGORY_PRESETS[selected_category]["neg"])
        st.session_state.last_category = selected_category

    # --- HP情報の項目編集 ---
    st.markdown("#### 📄 HP情報の項目")
    for i, item in enumerate(st.session_state.hp_items):
        col_item, col_del = st.columns([5, 1])
        with col_item:
            st.text(f"・{item}")
        with col_del:
            if st.button("✕", key=f"hp_del_{i}", help="この項目を削除"):
                st.session_state.hp_items.pop(i)
                st.rerun()

    new_hp_item = st.text_input(
        "項目を追加", key="new_hp_input", placeholder="例：資本金"
    )
    if st.button("➕ HP項目を追加", key="hp_add_btn", use_container_width=True):
        if new_hp_item.strip():
            st.session_state.hp_items.append(new_hp_item.strip())
            st.rerun()

    st.markdown("---")

    # --- 商談情報の項目編集 ---
    st.markdown("#### 💬 商談情報の項目")
    for i, item in enumerate(st.session_state.neg_items):
        col_item, col_del = st.columns([5, 1])
        with col_item:
            st.text(f"・{item}")
        with col_del:
            if st.button("✕", key=f"neg_del_{i}", help="この項目を削除"):
                st.session_state.neg_items.pop(i)
                st.rerun()

    new_neg_item = st.text_input(
        "項目を追加", key="new_neg_input", placeholder="例：担当者の決裁権限"
    )
    if st.button("➕ 商談項目を追加", key="neg_add_btn", use_container_width=True):
        if new_neg_item.strip():
            st.session_state.neg_items.append(new_neg_item.strip())
            st.rerun()

    st.markdown("---")

    # --- マニュアルファイル ---
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
# メインコンテンツ
# ==============================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 商談情報の入力")

    transcript_text_input = st.text_area(
        "商談文字起こしをここに貼り付けてください",
        height=350,
        placeholder="文字起こしテキストをここに貼り付けてください...",
    )

    website_url = st.text_input(
        "商談相手の企業HP URL（任意）", placeholder="https://example.com"
    )

with col2:
    st.subheader("2. 生成")

    # 現在の項目設定をプレビュー表示
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

    if st.button("🚀 レポート生成を開始", type="primary", use_container_width=True):
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
            sales_material_text = ""

            with st.spinner("マニュアルを読み込み中..."):
                try:
                    manual_text = extract_text_from_pdf(manual_file)
                except Exception as e:
                    st.error(f"マニュアルPDFの読み込みに失敗しました: {e}")

            transcript_text = transcript_text_input

            if len(transcript_text) > 0:
                # Webサイト情報の取得
                website_text = ""
                if website_url:
                    try:
                        with st.spinner("Webサイト情報を取得中..."):
                            website_text = fetch_website_content(website_url)
                            if not website_text:
                                st.warning("指定されたURLから情報を取得できませんでした。")
                            else:
                                st.success(
                                    f"Webサイト情報を取得しました ({len(website_text)}文字)"
                                )
                    except Exception as e:
                        st.warning(f"Webサイト情報の取得中にエラーが発生しました: {e}")

                with st.spinner("AIが営業代行レポートを分析中..."):
                    try:
                        llm = ChatGoogleGenerativeAI(
                            model=selected_model_name, google_api_key=api_key
                        )

                        # hp_items / neg_items をセッションから取得して渡す
                        data = generate_report_content(
                            transcript_text,
                            manual_text,
                            website_text,
                            sales_material_text,
                            llm,
                            hp_items=st.session_state.hp_items,
                            neg_items=st.session_state.neg_items,
                        )

                        st.success("分析完了！スプレッドシートを作成します...")

                        gcp_info = st.secrets["gcp_service_account"]
                        TEMPLATE_ID = st.secrets["google_drive"]["template_id"]
                        FOLDER_ID = st.secrets["google_drive"]["folder_id"]

                        data["website_url"] = website_url
                        sheet_url = fill_google_sheet(
                            data,
                            gcp_info,
                            TEMPLATE_ID,
                            FOLDER_ID,
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

st.markdown("---")
st.caption("Powered by Streamlit & LangChain")