# 営業レポート作成AI

商談の文字起こし・メール文章・企業HPから、Googleスプレッドシート形式の営業評価レポートを自動生成するStreamlitアプリです。

---

## ファイル構成

```
.
├── app.py                        # メインUI（Streamlit）
├── report_generator.py           # AI処理・スプレッドシート書き込み
├── utils.py                      # PDF読み込み・Webスクレイピング
├── requirements.txt              # 依存ライブラリ
├── README.md                     # このファイル
└── 8ba0d12e-...営業レポートマニュアル.pdf  # 評価マニュアル（必須）
```

---

## 機能概要

### タブ1：新規レポート作成
- 商談文字起こし＋企業HP URLを入力
- AIがフルレポートを生成してGoogleスプレッドシートに書き込む
- 出力内容：基本ヘッダー・HP情報整理・商談情報整理・Q&A・チェックリスト評価・全体統括コメント

### タブ2：評価を既存シートに追記
- 既存スプレッドシートのURLを指定、またはテンプレートから新規作成
- **文字起こし**から評価できるカテゴリ：営業人間力・商談対応力・商談後（全体評価）
- **メール文章**から評価できるカテゴリ：商談前IS・商談後（メール）
- 評価（〇△✕）と備考をG列・J列に書き込む

---

## セットアップ

### 1. 依存ライブラリのインストール

```bash
pip install -r requirements.txt
```

### 2. Google Cloud の設定

#### サービスアカウントの作成
1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. 「APIとサービス」→「認証情報」→「サービスアカウントを作成」
3. 作成したサービスアカウントのJSONキーをダウンロード
4. Google Drive API・Google Sheets APIを有効化

#### スプレッドシートの共有設定
- テンプレートスプレッドシートをサービスアカウントのメールアドレス（`xxx@xxx.iam.gserviceaccount.com`）と**編集者**として共有する
- 既存シートに追記する場合も同様に共有が必要

### 3. Streamlit Secrets の設定

`Streamlit Community Cloud` の「Advanced settings → Secrets」または ローカルの `.streamlit/secrets.toml` に以下を設定：

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "xxx"
private_key = "-----BEGIN RSA PRIVATE KEY-----\n..."
client_email = "xxx@xxx.iam.gserviceaccount.com"
client_id = "xxx"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"

[google_drive]
template_id = "スプレッドシートテンプレートのID"
folder_id   = "保存先フォルダのID"
```

> **スプレッドシートIDの取得方法**  
> `https://docs.google.com/spreadsheets/d/【ここがID】/edit` のURLから取得

### 4. Gemini APIキーの取得

[Google AI Studio](https://aistudio.google.com/) でAPIキーを発行し、アプリのサイドバーに入力して使用します。

---

## Streamlit Community Cloud へのデプロイ

1. GitHubリポジトリに全ファイルをpush
2. [Streamlit Cloud](https://streamlit.io/cloud) で「New app」→ リポジトリ・ブランチ・`app.py` を指定
3. Advanced settings の Secrets に上記のTOML設定を貼り付け
4. 「Deploy!」をクリック

---

## コードの主要な設計ポイント（改修者向け）

### カテゴリ・項目の追加・変更

#### 調査カテゴリ（HP情報・商談情報の項目）
`app.py` の `CATEGORY_PRESETS` 辞書に新しいカテゴリを追加する。  
`hp`（HP情報項目リスト）と `neg`（商談情報項目リスト）をそれぞれ定義する。

#### チェックリスト項目の追加・変更
2箇所の修正が必要：

1. **`report_generator.py`** の `CHECKLIST_ITEMS_BY_CATEGORY`  
   各カテゴリのリストに項目文字列を追加する。  
   ※照合はD列テキストの**先頭15文字**で行うため、先頭15文字が他の項目と被らないよう注意。

2. **`app.py`** の `CHECKLIST_CATEGORIES`  
   カテゴリ名・行範囲（`rows`）・表示ラベル（`label`）を定義する。  
   行範囲はスプレッドシートテンプレートの実際の行番号と一致させること。

#### 文字起こし系 / メール系の振り分け変更
`app.py` の以下2行で制御している：
```python
TRANSCRIPT_CATS = ["営業人間力", "商談対応力", "商談後（全体評価）"]
EMAIL_CATS      = ["商談前IS", "商談後（メール）"]
```

### スプレッドシートへの書き込み列
| 内容 | 列 |
|------|-----|
| 評価（〇△✕） | G列 |
| 備考 | J列 |
| HP情報・項目名 | A列（Row35〜） |
| HP情報・内容 | C列（Row35〜） |
| 商談情報・項目名 | F列（Row35〜） |
| 商談情報・内容 | H列（Row35〜） |
| Q&A（弊社→先方） | I・J列（Row3〜16） |
| Q&A（先方→弊社） | I・J列（Row18〜30） |
| 全体統括コメント | A50 |

### 照合ロジック（書き込み時）
`write_evaluation_to_existing_sheet()` でAIが返した `display_text` の**先頭15文字**をD列の各行と照合して書き込み先を特定している。  
項目を追加する際は先頭15文字が既存項目と重複しないか確認すること。

### 生成ファイル名
`report_generator.py` の以下2箇所で制御：
- 新規レポート作成：`f"AI作成_{data.get('cl_company_name', '名称未設定')}様_営業レポート"`
- 評価追記（新規コピー）：`"AI作成_営業レポート（評価追記）"`

---

## よくあるエラーと対処法

| エラー | 原因 | 対処 |
|--------|------|------|
| `ImportError: cannot import name '...'` | GitHubのファイルが古い | 最新の`report_generator.py`をpushし直す。Streamlit CloudでReboot appを実行 |
| 書き込みが0件で完了する | D列の照合が一致しない | `CHECKLIST_ITEMS_BY_CATEGORY`の先頭15文字とシートのD列テキストが一致しているか確認 |
| `gspread.exceptions.APIError` | サービスアカウントの権限不足 | スプレッドシートをサービスアカウントのメールと共有しているか確認 |
| `JSONDecodeError` | AIのJSON出力が壊れている | モデルを`gemini-2.5-flash`に変更するか、入力テキストを短くして再試行 |
| マニュアルが見つからない | PDFファイルのパスが違う | `app.py`の`DEFAULT_MANUAL_PATH`に設定されているファイル名と実際のファイル名を一致させる |

---

## 使用技術

- **Streamlit** — WebアプリUI
- **Google Gemini API (LangChain経由)** — AI評価・レポート生成
- **gspread** — Googleスプレッドシートへの書き込み
- **Google Drive API** — テンプレートのコピー
- **pypdf** — マニュアルPDFのテキスト抽出
- **BeautifulSoup / requests** — 企業HPのスクレイピング