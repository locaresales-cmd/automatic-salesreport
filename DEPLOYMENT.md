# Streamlit Community Cloud デプロイ手順

このアプリケーションをWeb上で公開（デプロイ）するための手順書です。

## 1. 準備

以下のファイルが揃っていることを確認してください（今回の作業で自動的に準備されています）。
- `app.py`: メインアプリケーション
- `requirements.txt`: 必要なライブラリ一覧
- `repaired_template.xlsx`: Excelテンプレート
- `8ba0d12e-f2ee-4002-9533-54a0940f4eaa_営業レポートマニュアル.pdf`: マニュアルPDF
- `report_generator.py`, `utils.py`: 補助スクリプト

## 2. GitHubへのアップロード

Streamlit Community Cloudは、GitHub上のコードを読み込んで動きます。

1.  GitHubに新しいリポジトリ（例: `sales-report-ai`）を作成します。
2.  このフォルダ内のファイルをすべてアップロード（Push）します。

## 3. Streamlit Community Cloudでの設定

1.  [Streamlit Community Cloud](https://streamlit.io/cloud) にアクセスし、サインインします。
2.  「New app」ボタンをクリックします。
3.  **Repository**: 作成したGitHubリポジトリを選択します。
4.  **Branch**: `main` (または `master`) を選択します。
5.  **Main file path**: `app.py` を入力します。
6.  **Advanced settings** を開き、APIキーを設定します。

### APIキーの設定 (Secrets)

ローカルで `.env` ファイルや直接入力していたAPIキーは、クラウド上では「Secrets」として安全に管理します。

Advanced settings の「Secrets」欄に以下のように入力してください：

```toml
# Google Gemini APIキーを使用する場合
GOOGLE_API_KEY = "ここにGOOGEL_API_KEYを入力"

# OpenAI APIキーを使用する場合
OPENAI_API_KEY = "ここにOPENAI_API_KEYを入力"
```

## 4. デプロイ

「Deploy!」ボタンをクリックします。
数分待つと、アプリケーションが起動し、URLが発行されます。

## 注意点

- **ファイルアップロード**: クラウド上でアップロードしたファイルは一時的なものです。永続的に保存したい場合は、Google Drive API連携やS3などのクラウドストレージ設定が別途必要になります（現状のコードでは、生成されたExcelファイルはその場でダウンロードする仕様なので問題ありません）。
- **パス**: 今回の修正で、ファイルパスを「相対パス」に変更しました。これにより、クラウド上のどのディレクトリに配置されても動作するようになっています。
