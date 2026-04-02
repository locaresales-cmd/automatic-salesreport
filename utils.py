import pypdf
import requests
from bs4 import BeautifulSoup

def extract_text_from_pdf(pdf_file):
    """
    Simples text extraction from a PDF file.
    Args:
        pdf_file: A file-like object (e.g., from st.file_uploader)
    Returns:
        str: Extracted text.
    """
    text = ""
    try:
        reader = pypdf.PdfReader(pdf_file)
        for page in reader.pages:
            text += page.extract_text() + "\n"
    except Exception as e:
        raise e  # Propagate error to caller
    return text

def fetch_page_text(url, headers):
    """
    1ページ分のテキストを取得して返す内部関数
    """
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        for script in soup(["script", "style"]):
            script.decompose()
        text = soup.get_text()
        lines = (line.strip() for line in text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        return '\n'.join(chunk for chunk in chunks if chunk)
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return ""

def extract_internal_links(url, soup):
    """
    同一ドメイン内のリンクを最大10件抽出する内部関数
    PDFは除外する
    """
    from urllib.parse import urljoin, urlparse
    base_domain = urlparse(url).netloc
    links = set()
    for a_tag in soup.find_all('a', href=True):
        href = a_tag['href']
        full_url = urljoin(url, href)
        parsed = urlparse(full_url)
        # 同一ドメインかつPDF以外のみ収集
        if parsed.netloc == base_domain and not full_url.lower().endswith('.pdf'):
            links.add(full_url)
        if len(links) >= 10:
            break
    return list(links)

def fetch_website_content(url):
    """
    指定URLとその内部リンク先（最大10ページ）からテキストを取得する。
    情報が不足している場合に関連ページも参照できるようにする。
    """
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    all_text = []
    visited = set()

    try:
        # まずトップページを取得
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        # トップページのテキストを取得
        top_text = fetch_page_text(url, headers)
        if top_text:
            all_text.append(f"=== {url} ===\n{top_text}")
        visited.add(url)

        # 内部リンクを抽出して各ページも取得
        internal_links = extract_internal_links(url, soup)
        for link in internal_links:
            if link in visited:
                continue
            page_text = fetch_page_text(link, headers)
            if page_text:
                all_text.append(f"=== {link} ===\n{page_text}")
            visited.add(link)

        combined = '\n\n'.join(all_text)
        # トークン制限対策で全体を30000文字に制限
        return combined[:30000]

    except Exception as e:
        print(f"Error fetching website content: {e}")
        return ""
