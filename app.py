import sys
# Force UTF-8 encoding for stdout/stderr to prevent UnicodeEncodeError in restricted environments
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

import streamlit as st
import os
import tempfile
from utils import extract_text_from_pdf, fetch_website_content
from report_generator import generate_report_content, fill_google_sheet
from langchain_openai import ChatOpenAI
from langchain_google_genai import ChatGoogleGenerativeAI
import google.generativeai as genai
import pandas as pd

st.set_page_config(page_title="営業レポート作成AI", layout="wide")

st.title("🗒️ 営業レポート作成AI")
st.markdown("商談の文字起こしとマニュアルから、指定フォーマットの営業レポートを自動生成します。")

# Sidebar for configuration
with st.sidebar:
    st.header("設定")

    GEMINI_MODELS = {
        "Gemini 3 Pro（最新・最高精度）":  "gemini-3-pro-preview",
        "Gemini 3 Flash（最新・高速）":    "gemini-3-flash-preview",
        "Gemini 2.5 Pro（高精度）":        "gemini-2.5-pro",
        "Gemini 2.5 Flash（高速・推奨）":  "gemini-2.5-flash",
        "Gemini 2.5 Flash Lite（最軽量）": "gemini-2.5-flash-lite",
        "Gemini 2.0 Flash（安定版）":      "gemini-2.0-flash",
    }

    selected_label = st.selectbox("使用するGeminiモデル", list(GEMINI_MODELS.keys()), index=3)
    selected_model_name = GEMINI_MODELS[selected_label]
    st.caption(f"モデルID: `{selected_model_name}`")

    api_key = st.text_input("Gemini API Key", type="password")

    st.markdown("---")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    DEFAULT_MANUAL_PATH = os.path.join(base_dir, "8ba0d12e-f2ee-4002-9533-54a0940f4eaa_営業レポートマニュアル.pdf")

    st.markdown("### マニュアルファイル")
    if os.path.exists(DEFAULT_MANUAL_PATH):
        st.success("✅ マニュアル読み込み済み")
        manual_file = open(DEFAULT_MANUAL_PATH, "rb")
    else:
        st.error("デフォルトマニュアルが見つかりません。")
        manual_file = None

# Main Content
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 商談情報の入力")

    transcript_text_input = st.text_area(
        "商談文字起こしをここに貼り付けてください",
        height=350,
        placeholder="文字起こしテキストをここに貼り付けてください..."
    )

    website_url = st.text_input("商談相手の企業HP URL（任意）", placeholder="https://example.com")

    

with col2:
    st.subheader("2. 生成")
    if st.button("🚀 レポート生成を開始", type="primary", use_container_width=True):
        if not api_key:
            st.error("APIキーを入力してください。")
        elif not transcript_text_input:
            st.error("商談テキストを入力してください。")
        elif not manual_file:
            st.error("マニュアルファイルが必要です。")
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
                # Fetch website content if URL is provided
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
                        # LLMの初期化 (SecretsからAPIキー取得)
                        llm = ChatGoogleGenerativeAI(model=selected_model_name, google_api_key=api_key)
                        
                        data = generate_report_content(transcript_text, manual_text, website_text, sales_material_text, llm)
                        
                        st.success("分析完了！スプレッドシートを作成します...")
                        
                        # SecretsからIDと認証情報を取得
                        gcp_info = st.secrets["gcp_service_account"]
                        TEMPLATE_ID = st.secrets["google_drive"]["template_id"]
                        FOLDER_ID = st.secrets["google_drive"]["folder_id"]
                        
                        sheet_url = fill_google_sheet(data, gcp_info, TEMPLATE_ID, FOLDER_ID)
                        
                        st.balloons()
                        st.success("レポートが共有ドライブに作成されました！")
                        st.link_button("🔥 完成したスプレッドシートを開く", sheet_url)
                        
                        with st.expander("抽出データ（JSON）の確認"):
                            st.json(data)

                    except Exception as e:
                        st.error(f"エラーが発生しました: {e}")

st.markdown("---")
st.caption("Powered by Streamlit & LangChain")
