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

st.title("🗒️ 営業レポート作成AIプラットフォーム")
st.markdown("商談の文字起こしとマニュアルから、指定フォーマットの営業レポートを自動生成します。")

# Sidebar for configuration
with st.sidebar:
    st.header("設定")
    model_provider = st.selectbox("使用するモデルプロバイダー", ["Google Gemini", "OpenAI GPT-4"])
    
    selected_model_name = None
    if model_provider == "Google Gemini":
        base_models = ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash", "gemini-2.5-flash", "Custom"]
        selected_option = st.selectbox("Geminiモデルを選択", base_models, index=2)
        
        if selected_option == "Custom":
            selected_model_name = st.text_input("モデル名を入力 (例: gemini-1.5-flash-8b)", value="gemini-2.0-flash")
        else:
            selected_model_name = selected_option
    
    api_key = st.text_input("API Key", type="password")

    if api_key and model_provider == "Google Gemini":
        if st.button("接続テスト & 対応モデル確認"):
            try:
                genai.configure(api_key=api_key)
                models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                st.success(f"接続成功！利用可能なモデル: {len(models)}個")
                st.json(models)
            except Exception as e:
                st.error(f"接続エラー: {e}")
    
    st.markdown("---")
    st.markdown("### テンプレートファイル")
    # Default path handling
    # Default path handling - Use relative paths for deployment
    base_dir = os.path.dirname(os.path.abspath(__file__))
    DEFAULT_TEMPLATE_PATH = os.path.join(base_dir, "repaired_template.xlsx")
    DEFAULT_MANUAL_PATH = os.path.join(base_dir, "8ba0d12e-f2ee-4002-9533-54a0940f4eaa_営業レポートマニュアル.pdf")
    
    use_default_files = st.checkbox("フォルダ内のデフォルトファイルを使用", value=True)

# Main Content
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 素材のアップロード")
    input_method = st.radio("商談文字起こしの入力方法", ["ファイルアップロード", "テキスト直接入力"])
    
    transcript_file = None
    transcript_text_input = ""

    if input_method == "ファイルアップロード":
        transcript_file = st.file_uploader("商談文字起こし (Text/Transcript)", type=["txt", "pdf"])
    else:
        transcript_text_input = st.text_area("商談テキストをここに貼り付けてください", height=300)
    
    if use_default_files:
        if os.path.exists(DEFAULT_MANUAL_PATH):
            st.info(f"デフォルトマニュアルを使用中: {os.path.basename(DEFAULT_MANUAL_PATH)}")
            try:
                manual_file = open(DEFAULT_MANUAL_PATH, "rb")
            except Exception as e:
                st.error(f"デフォルトマニュアルの読み込みエラー: {e}")
                manual_file = None
        else:
            st.error("デフォルトマニュアルが見つかりません。")
            manual_file = st.file_uploader("営業レポートマニュアル (PDF)", type=["pdf"])
            
        if os.path.exists(DEFAULT_TEMPLATE_PATH):
            st.info(f"デフォルトテンプレートを使用中: {os.path.basename(DEFAULT_TEMPLATE_PATH)}")
            template_path = DEFAULT_TEMPLATE_PATH
        else:
            st.error("デフォルトテンプレートが見つかりません。")
            template_file = st.file_uploader("レポートフォーマット (Excel)", type=["xlsx"])
            if template_file:
                # Save uploaded template strictly for processing
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(template_file.read())
                    template_path = tmp.name
            else:
                template_path = None
    else:
        manual_file = st.file_uploader("営業レポートマニュアル (PDF)", type=["pdf"])
        template_file = st.file_uploader("レポートフォーマット (Excel)", type=["xlsx"])
        if template_file:
             with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(template_file.read())
                template_path = tmp.name
        else:
            template_path = None
    
    website_url = st.text_input("商談相手の企業HP URL (任意)", placeholder="https://example.com")
    sales_material_file = st.file_uploader("営業資料 (PDF) - 任意", type=["pdf"])

with col2:
    st.subheader("2. 生成")
    if st.button("レポート生成を開始", type="primary"):
        if not api_key:
            st.error("APIキーを入力してください。")
        elif input_method == "ファイルアップロード" and not transcript_file:
            st.error("商談文字起こしファイルをアップロードしてください。")
        elif input_method == "テキスト直接入力" and not transcript_text_input:
            st.error("商談テキストを入力してください。")
        elif not manual_file:
             st.error("マニュアルファイルが必要です。")
        #elif not template_path:
         #    st.error("テンプレートファイルが必要です。")
        else:
            transcript_text = ""
            manual_text = ""
            
            with st.spinner("マニュアルを読み込み中..."):
                try:
                    manual_text = extract_text_from_pdf(manual_file)
                except Exception as e:
                    st.error(f"マニュアルPDFの読み込みに失敗しました: {e}")
            
            # Extract text from sales material
            sales_material_text = ""
            if sales_material_file:
                with st.spinner("営業資料を読み込み中..."):
                    try:
                        sales_material_text = extract_text_from_pdf(sales_material_file)
                        st.info(f"営業資料読み込み完了: {len(sales_material_text)}文字")
                    except Exception as e:
                         st.error(f"営業資料の読み込みに失敗しました: {e}")
            
            with st.spinner("文字起こしを読み込み中..."):
                try:
                    if input_method == "テキスト直接入力":
                        transcript_text = transcript_text_input
                    elif transcript_file.type == "application/pdf":
                        transcript_text = extract_text_from_pdf(transcript_file)
                    else:
                        transcript_text = transcript_file.read().decode("utf-8")
                    
                    st.info(f"読み込み完了: 文字数 {len(transcript_text)}")
                    # Debug: show first 100 characters to verify content
                    if len(transcript_text) > 0:
                         with st.expander("読み込んだテキストの先頭を確認"):
                             st.text(transcript_text[:500])
                    
                except Exception as e:
                    st.error(f"商談文字起こしの読み込みに失敗しました: {e}") 
                    # If it's the specific Unicode error, suggest solution
                    if "codec" in str(e):
                         st.warning("ヒント: ファイル名や内容に特殊な文字が含まれている可能性があります。")
                
                if len(transcript_text) == 0:
                     st.error("文字起こしファイルからテキストを抽出できませんでした。PDFが画像（スキャンデータ）の場合、テキストとして読み込めない可能性があります。テキスト形式のファイルを使用してください。")
            
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
