import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ここにご自身の情報を入れてください
service_account_info = {"type": "service_account",
  "project_id": "gen-lang-client-0647999793",
  "private_key_id": "f389a8a1c017be35360edb798eb00beaafe97fbc",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDWzmCLbupCNH+5\n9/q/kXKEXo4vMeDwCkJAr/2/jmtcsCrCxiCzOvJlsFjGVBlOVkVWkP/EGxF6wZvc\nklVas1r70HVWMkgeeEQrOOONrTyp9e7GqatgDG51nJFm3us3QFgz+JXTJhuFa+So\njPIFzHKd2EQTFaKVZeUH3sBZJetE5IpRhDRtyxXk/tJDnlleGCyTGV9agjQkI3ap\naDYHybRVqO3e1XTCoKKDqR/zZpyyzuO7H+gIX3MORIfklHGAGHAdFa8r7VCsQpYV\nOHTeJaIJ83uwfl1j6mVXTAAkPvraoADcE8RtxQSG8HJJDDo8HcAUIaCax5ULCmnB\nbQI8/yYjAgMBAAECggEAI0T+OSaxtBkmolmt2ZJ/eRiFbrOujqXrPEUpUYXdlE8g\nq5vDIWqbrkgG64RCJh39gznfpPVvffZtfwlP68qTNSbZ0pitmsMofj2jBKTd1xJD\nGGoiHMKDHds39z9sOPA0YKQqWuG2QcaguW7Ftl9cBRV4vUwfcTmqd2EDDXWFFxRv\nia+p4yqSVEx0XHIzfCU/IufeYdCAIQpDnuCfIlZFNO59lWuonHyEXkEan7y1q1aF\nKiYasL5a0F7vp+4lFUsd/SrfK/Gt49Rqhq3SfYqtkLN33eQ/YXn/YhdkgdmhX0JW\nxfQGZErY48GO2dO9Avtc7LYKqGe2axlOxrkH7JhMMQKBgQD4+VYvpJhZIHAbFUQu\n5lLODH6mSGd19TBpzRn9jCV8muR8oCh0FV87AY+LKPXjZPjvlfeyVBy+QhImvMsv\nKc/vI/4rZQBIKnAKKVCRKei+EgghSkSduipmFt/ALYBnXhtm1UkwrCsyM9DVmntt\ndKFNUfO3PeJv+h7r7FoFBRM0DwKBgQDc3jOs0jpEYkDDPsERlzxy90yUpbXjl76V\nsPxJUdD8cne+8jtMxU3x/0wbXFat+oyWkdlJAHGx3LKHkGu9LArrc6vEpFCDNqtd\nAzpZyGJfpTsyJTpGBYxFs2AOEdinKeSeZaHjMU6FUyUBhR/rJnrhXbKY8Ny8Md48\nP+GuHdOIrQKBgFrwp/xrAIK9iHU8BVWkJ2a/xZrzI2dAkdhzZCTqhd7HrOGglmYg\nUFJ7NXU9FuNiRFMu0fS/KGiONZcUqpqliR/uY65yC/JQHfB4OsdrKWoTqAiQ2hNK\npqX3gO7vL9GR3Cxph3xRxs1lg8ghzyehzDEz1/N8lTMVhynhgNgIjIUdAoGBANpj\nN1M9t3FgeUrU5RBgqtu+XNFqHLRCmabnjj1tEahcAr0iRLI/MTgESBuRrP9wCszi\nv6doMgM9BqX2jiFJyC5RfFj+Y8GqL7zTcUHPWj3aYfLOTpVn7PAKUgL3cHKxgKWC\nNpUvbsVzldav7ASWUtA91ldVad0HrgeC3sJMKZotAoGAGxyYNpnhaGOU0vXLqI7B\nTSWmFCTfnB5Raer4U7Bxx8xVrewefv5sq+geWxvFGP2OFxL0V7wq2g5wBUcHwUqD\n1Y2El6Lij0OKOOIOgGabfsL4fmQzb6umgQVAXCVqtMHKrVsrgEc/5qKGjQDLA9gp\n8BJsvC8KTDa89GPPiTt2G1I=\n-----END PRIVATE KEY-----\n",
  "client_email": "report-bot@gen-lang-client-0647999793.iam.gserviceaccount.com",
  "client_id": "104242993315022279581",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/report-bot%40gen-lang-client-0647999793.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"}
folder_id = "1NOIx0avBHLI1LxQFjseGhtFg71btvnwj"

creds = Credentials.from_service_account_info(
    service_account_info, 
    scopes=["https://www.googleapis.com/auth/drive"]
)
drive_service = build('drive', 'v3', credentials=creds)

# テスト1: 非常に小さなテキストファイルを作成してみる
try:
    print("テスト1: 新規ファイル作成を試行中...")
    file_metadata = {'name': 'test.txt', 'parents': [folder_id]}
    drive_service.files().create(body=file_metadata).execute()
    print("成功: 新規作成は可能です。")
except Exception as e:
    print(f"失敗: 新規作成でエラーが発生しました。\n{e}")

# テスト2: 容量情報を取得してみる
try:
    print("\nテスト2: アカウントの容量情報を取得中...")
    about = drive_service.about().get(fields="storageQuota").execute()
    print(f"容量情報: {about['storageQuota']}")
except Exception as e:
    print(f"容量情報の取得に失敗しました。\n{e}")