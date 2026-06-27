#!/usr/bin/env python3
"""
Digital Zoo Linkbuilding Article Generator
==========================================
讀取 Excel keyword 配對 → AI 生成文章 → 產出格式化 Google Doc

Usage:
    python generate.py --excel data.xlsx --batch 1 --creds service_account.json
    python generate.py --excel data.xlsx --batch 1 --creds sa.json --folder FOLDER_ID
    python generate.py --excel data.xlsx --batch 1 --creds sa.json --start 5 --end 10
"""

import os
import sys
import json
import time
import argparse
import re
from pathlib import Path

import requests as http_req
from openpyxl import load_workbook
from tqdm import tqdm

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ================================================================
# Configuration
# ================================================================
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
MODEL = os.environ.get("LB_MODEL", "z-ai/glm-5.2")
DELAY_BETWEEN_CALLS = int(os.environ.get("LB_DELAY", "3"))

# 藍色 highlight RGB（Google Docs 0-1 scale）
HIGHLIGHT_COLOR = {"red": 0.60, "green": 0.84, "blue": 0.92}

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]

# ================================================================
# Writing Guidelines (injected into AI prompt)
# ================================================================
WRITING_GUIDELINES = """
## 反向連結文章寫作指引

### 字數要求
下限 750 至 1000 字

### 內容書寫原則
- 寫作以網絡媒體文章為標準，成品愈接近這類文章愈理想。
- 如有兩個關鍵字，應構思一個能觸及兩個關鍵字主題的主題，內容盡量能關聯兩個關鍵字。
- 不應該直接以關鍵字為主題，因有可能最終令文章排名比客戶的目標連結頁面更高。

### 目標關鍵字與連結位置規則
- 在正文中首次出現「完全符合關鍵字」時標記。
- 連結必須放置在與目標頁面內容高度相關的上下文中。
- 不可將兩個關鍵字放在同一段落，也不可安排在前段尾部與後段開首緊接出現。
- 關鍵字不應加到全文的第一段或最後一段。
- H1 與 H2 標題中不可出現關鍵字。

### 標題與段落格式
- H1 標題需在 30 個全形字（或 60 字元）以內，不可有標點符號
- H1 中絕對不可出現與客戶目標頁面完全相同的「目標關鍵字」
- 不使用 H3
- 不應全篇都使用 Bullet Point
- 避免以密集的短句作分段，建議結合成段落
"""


# ================================================================
# Excel Parser
# ================================================================
def parse_excel(filepath, batch_number, sheet_name="202605"):
    """Parse Excel and return list of article dicts for the specified batch."""
    wb = load_workbook(filepath)
    ws = wb[sheet_name]

    articles = []
    current_batch = None
    target_batch = f"Batch {batch_number}"

    row = 2  # skip header
    while row <= ws.max_row:
        # Check batch label
        batch_val = ws[f"A{row}"].value
        if batch_val:
            current_batch = str(batch_val).strip()

        if current_batch != target_batch:
            row += 1
            continue

        art_num = ws[f"B{row}"].value
        kw_num = ws[f"C{row}"].value

        # Article starts at kw_num == 1
        if art_num is not None and kw_num == 1.0:
            kw1 = ws[f"D{row}"].value or ""
            url1 = ws[f"E{row}"].value or ""
            category = ws[f"K{row}"].value or "General"

            # Read keyword 2 from next row
            kw2 = ""
            url2 = ""
            if row + 1 <= ws.max_row and ws[f"C{row+1}"].value == 2.0:
                kw2_val = ws[f"D{row+1}"].value or "--"
                url2_val = ws[f"E{row+1}"].value or "--"
                if kw2_val != "--":
                    kw2 = kw2_val
                if url2_val != "--":
                    url2 = url2_val

            articles.append({
                "number": int(art_num),
                "keyword1": kw1.strip(),
                "url1": url1.strip(),
                "keyword2": kw2.strip(),
                "url2": url2.strip(),
                "category": category.strip(),
            })

        row += 1

    wb.close()
    return articles


# ================================================================
# Language Detection
# ================================================================
def detect_language(text):
    """Detect if text is primarily Chinese or English."""
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    return "zh-HK" if chinese_chars > len(text) * 0.3 else "en"


# ================================================================
# Article Generator (OpenRouter API)
# ================================================================
def build_prompt(article):
    """Build the AI prompt for article generation."""
    kw1 = article["keyword1"]
    url1 = article["url1"]
    kw2 = article["keyword2"]
    url2 = article["url2"]
    category = article["category"]

    lang = detect_language(kw1 + kw2)

    if lang == "zh-HK":
        lang_instruction = """語言要求：
- 使用繁體中文書面語撰寫，語體風格對標香港主流網絡媒體（如《香港01》、《經濟日報》副刊）
- 嚴禁使用廣東話口語、粵語語助詞或任何口語化表達（例如「嘅」「咗」「啲」「點解」「攞」「揀」「搞掂」「嚟」等）
- 用詞正式但自然，語氣平實而有深度，行文要流暢，段落之間要有邏輯銜接
- 用「的」不用「嘅」，用「了」不用「咗」，用「一些」不用「啲」，用「為什麼」不用「點解」"""
    else:
        lang_instruction = """Language requirements:
- Write in fluent, professional English suitable for online media
- Maintain a formal but accessible tone, similar to quality editorial content
- Ensure logical flow between paragraphs with smooth transitions"""

    # Build keyword section
    if kw2:
        kw_section = f"""目標關鍵字：
- 關鍵字 1: 「{kw1}」（目標連結: {url1}）
- 關鍵字 2: 「{kw2}」（目標連結: {url2}）

文章需同時觸及兩個關鍵字的主題。在正文中，用 {{{{KW1}}}} 標記關鍵字 1 首次出現的位置，用 {{{{KW2}}}} 標記關鍵字 2 首次出現的位置。"""
    else:
        kw_section = f"""目標關鍵字：
- 關鍵字 1: 「{kw1}」（目標連結: {url1}）

在正文中，用 {{{{KW1}}}} 標記關鍵字 1 首次出現的位置。"""

    prompt = f"""你是一位專業的 SEO 內容寫手。請根據以下指引撰寫一篇反向連結文章。

{WRITING_GUIDELINES}

---

### 本篇任務

{kw_section}

網站類別 / 口吻: {category}
{lang_instruction}

### 輸出格式

請嚴格以下列 JSON 格式輸出（不要加 markdown code fence）：

{{
  "h1": "文章標題（30全形字內，無標點，不含關鍵字）",
  "sections": [
    {{
      "h2": null,
      "body": "開篇段落（不含關鍵字）"
    }},
    {{
      "h2": "第一個 H2 標題（不含關鍵字）",
      "body": "段落正文...在適當上下文中嵌入 {{{{KW1}}}}...後續內容"
    }},
    {{
      "h2": "第二個 H2 標題（不含關鍵字）",
      "body": "段落正文...在適當上下文中嵌入 {{{{KW2}}}}...後續內容"
    }},
    {{
      "h2": "結尾 H2 標題",
      "body": "總結段落（不含關鍵字）"
    }}
  ]
}}

### 重要規則
1. {{{{KW1}}}} 和 {{{{KW2}}}} 各只能出現一次，且必須使用標記形式
2. 關鍵字的原文（包括部分匹配）不可在標記以外的任何位置出現。例如關鍵字是「迷你倉 推介」，則正文其他段落不可出現「迷你倉」或「推介迷你倉」等字眼。如需提及相關概念，必須用同義詞或改寫方式表達（例如用「小型倉儲」代替「迷你倉」）
3. 兩個標記不可在同一 section
4. 標記不可在第一個 section（開篇段落）或最後一個 section
5. H1 和 H2 標題中不可包含關鍵字原文或其中任何部分
6. {{{{KW1}}}} 和 {{{{KW2}}}} 所在位置必須與目標連結頁面主題高度相關
7. 文章主題不可直接等於關鍵字，應找到一個能自然串聯兩個關鍵字的上層主題
8. 中文文章至少 800 字，建議 900-1100 字；英文文章至少 800 words，建議 900-1200 words
9. 段落長度適中，避免密集短句，每個 section 的 body 至少 120 字（中文）或 120 words（英文）
10. 只輸出 JSON，不要任何其他文字
"""
    return prompt


def _clean_keyword_duplicates(result, article):
    """Remove keyword text that appears outside of {{KW1}}/{{KW2}} markers."""
    kw1 = article.get("keyword1", "").strip()
    kw2 = article.get("keyword2", "").strip()

    for section in result["sections"]:
        body = section.get("body", "")
        if not body:
            continue

        # Temporarily replace markers with placeholders
        body = body.replace("{{KW1}}", "\x00KW1\x00")
        body = body.replace("{{KW2}}", "\x00KW2\x00")

        # Remove bare keyword occurrences (case-insensitive for English)
        if kw1:
            body = re.sub(re.escape(kw1), "", body, flags=re.IGNORECASE)
        if kw2:
            body = re.sub(re.escape(kw2), "", body, flags=re.IGNORECASE)

        # Clean up double spaces left by removal
        body = re.sub(r'  +', ' ', body)

        # Restore markers
        body = body.replace("\x00KW1\x00", "{{KW1}}")
        body = body.replace("\x00KW2\x00", "{{KW2}}")

        section["body"] = body

    return result


def generate_article_content(article, api_key, model, max_retries=3):
    """Call OpenRouter API to generate article content."""
    prompt = build_prompt(article)

    for attempt in range(max_retries):
        try:
            resp = http_req.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": model,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.7,
                    "max_tokens": 6000,
                },
                timeout=120,
            )
            resp.raise_for_status()
            data = resp.json()

            raw = data["choices"][0]["message"]["content"].strip()
            # Strip markdown fences if present
            raw = re.sub(r'^```(?:json)?\s*', '', raw)
            raw = re.sub(r'\s*```$', '', raw)

            result = json.loads(raw)

            # Validate structure
            assert "h1" in result, "Missing h1"
            assert "sections" in result, "Missing sections"
            assert len(result["sections"]) >= 3, "Need at least 3 sections"

            # Post-process: remove keyword duplicates outside markers
            result = _clean_keyword_duplicates(result, article)

            return result

        except (json.JSONDecodeError, AssertionError, KeyError) as e:
            print(f"  ⚠ Attempt {attempt+1} parse error: {e}")
            if attempt < max_retries - 1:
                time.sleep(5)
            continue
        except http_req.exceptions.RequestException as e:
            print(f"  ⚠ Attempt {attempt+1} API error: {e}")
            if attempt < max_retries - 1:
                time.sleep(10)
            continue

    print(f"  ✗ Failed after {max_retries} attempts")
    return None


# ================================================================
# Document Builder
# ================================================================
class DocBuilder:
    """Builds document text and tracks formatting ranges."""

    def __init__(self):
        self.text = ""
        self.bold_ranges = []       # (start, end)
        self.keyword_ranges = []    # (start, end, url)
        self.bullet_ranges = []     # (start, end)

    @property
    def pos(self):
        """Current position (0-indexed in the string)."""
        return len(self.text)

    def add(self, content, bold=False):
        """Add text, optionally tracking as bold."""
        start = self.pos
        self.text += content
        if bold:
            self.bold_ranges.append((start, self.pos))
        return start

    def add_line(self, content, bold=False):
        """Add text followed by newline."""
        start = self.add(content, bold=bold)
        self.text += "\n"
        return start

    def add_blank(self):
        """Add empty line."""
        self.text += "\n"

    def add_keyword(self, keyword_text, url):
        """Add keyword text and track for highlight + link."""
        start = self.pos
        self.text += keyword_text
        self.keyword_ranges.append((start, self.pos, url))

    def build_article(self, article, content):
        """Build one article's formatted text."""
        num = article["number"]
        kw1 = article["keyword1"]
        kw2 = article["keyword2"]
        url1 = article["url1"]
        url2 = article["url2"]

        # Article header: #N
        self.add_line(f"#{num}", bold=True)
        self.add_blank()

        # Keyword listing
        self.add_line("Keyword：", bold=True)
        self.add_blank()
        self.add_line(f"● {kw1}")
        if kw2:
            self.add_line(f"● {kw2}")
        self.add_blank()

        # H1
        h1 = content["h1"]
        self.add_line(f"H1：{h1}", bold=True)
        self.add_blank()

        # Sections
        for section in content["sections"]:
            # H2 (if present)
            h2 = section.get("h2")
            if h2:
                self.add_line(f"H2：{h2}", bold=True)
                self.add_blank()

            # Body with keyword markers
            body = section.get("body", "")
            self._add_body_with_keywords(body, kw1, url1, kw2, url2)
            self.add_blank()

        # Add page separator
        self.add_blank()

    def _add_body_with_keywords(self, body, kw1, url1, kw2, url2):
        """Parse body text and replace {{KW1}}/{{KW2}} markers."""
        # Split by markers and rebuild with tracking
        parts = re.split(r'(\{\{KW[12]\}\})', body)

        for part in parts:
            if part == "{{KW1}}" and kw1:
                self.add_keyword(kw1, url1)
            elif part == "{{KW2}}" and kw2:
                self.add_keyword(kw2, url2)
            else:
                self.text += part

        self.text += "\n"

    def get_gdocs_requests(self):
        """Convert tracked ranges to Google Docs API batchUpdate requests.

        Google Docs indices are 1-based (our string is 0-based),
        so every position gets +1 offset.
        """
        requests = []

        # 1. Insert all text
        requests.append({
            "insertText": {
                "location": {"index": 1},
                "text": self.text,
            }
        })

        # 2. Bold ranges
        for start, end in self.bold_ranges:
            requests.append({
                "updateTextStyle": {
                    "range": {
                        "startIndex": start + 1,
                        "endIndex": end + 1,
                    },
                    "textStyle": {"bold": True},
                    "fields": "bold",
                }
            })

        # 3. Keyword highlight + hyperlink
        for start, end, url in self.keyword_ranges:
            requests.append({
                "updateTextStyle": {
                    "range": {
                        "startIndex": start + 1,
                        "endIndex": end + 1,
                    },
                    "textStyle": {
                        "backgroundColor": {
                            "color": {"rgbColor": HIGHLIGHT_COLOR}
                        },
                        "link": {"url": url},
                    },
                    "fields": "backgroundColor,link",
                }
            })

        return requests


# ================================================================
# DOCX File Builder (no Google API needed)
# ================================================================
def _add_hyperlink_run(paragraph, text, url):
    """Add a highlighted hyperlink run to a paragraph."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run_el = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    # Blue text color
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0070C0")
    rPr.append(color)

    # Underline
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    # Highlight (cyan/turquoise background)
    hl = OxmlElement("w:highlight")
    hl.set(qn("w:val"), "cyan")
    rPr.append(hl)

    run_el.append(rPr)

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run_el.append(t)

    hyperlink.append(run_el)
    paragraph._element.append(hyperlink)


def _add_body_paragraph(doc, body, article):
    """Add a body paragraph, replacing {{KW1}}/{{KW2}} with highlighted hyperlinks."""
    p = doc.add_paragraph()

    parts = re.split(r'(\{\{KW[12]\}\})', body)
    for part in parts:
        if part == "{{KW1}}" and article["keyword1"]:
            _add_hyperlink_run(p, article["keyword1"], article["url1"])
        elif part == "{{KW2}}" and article["keyword2"]:
            _add_hyperlink_run(p, article["keyword2"], article["url2"])
        elif part:
            p.add_run(part)


def build_docx_file(articles_with_content, output_path):
    """Build a formatted .docx file from articles and their generated content.

    Args:
        articles_with_content: list of (article_dict, content_dict) tuples
        output_path: path to save the .docx file
    """
    from docx import Document as DocxDocument
    from docx.shared import Pt, Twips
    from docx.enum.text import WD_LINE_SPACING

    doc = DocxDocument()

    # Set Normal style: Arial 12pt, single spacing, no paragraph spacing
    style = doc.styles["Normal"]
    style.font.size = Pt(12)
    style.font.name = "Arial"
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    for idx, (article, content) in enumerate(articles_with_content):
        # Article number header
        p = doc.add_paragraph()
        run = p.add_run(f"#{article['number']}")
        run.bold = True

        # Blank line
        doc.add_paragraph()

        # Keyword listing
        p = doc.add_paragraph()
        run = p.add_run("Keyword：")
        run.bold = True

        doc.add_paragraph()
        doc.add_paragraph(f"● {article['keyword1']}")
        if article["keyword2"]:
            doc.add_paragraph(f"● {article['keyword2']}")

        doc.add_paragraph()

        # H1
        p = doc.add_paragraph()
        run = p.add_run(f"H1：{content['h1']}")
        run.bold = True

        doc.add_paragraph()

        # Sections
        for section in content["sections"]:
            h2 = section.get("h2")
            if h2:
                p = doc.add_paragraph()
                run = p.add_run(f"H2：{h2}")
                run.bold = True
                doc.add_paragraph()

            body = section.get("body", "")
            if body:
                _add_body_paragraph(doc, body, article)

            doc.add_paragraph()

        # Page break between articles (except last)
        if idx < len(articles_with_content) - 1:
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement

            p = doc.add_paragraph()
            run = p.add_run()
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run._element.append(br)

    doc.save(output_path)
    return output_path
def get_google_services(credentials_path):
    """Authenticate from JSON file and return Docs + Drive services."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(
        credentials_path, scopes=SCOPES
    )
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return docs_service, drive_service


def get_google_services_from_info(credentials_info):
    """Authenticate from dict (for Streamlit Cloud secrets) and return services."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    info = dict(credentials_info)
    if "private_key" in info:
        info["private_key"] = info["private_key"].replace("\\n", "\n")
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=SCOPES
    )
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return docs_service, drive_service


def create_formatted_doc(
    docs_service, drive_service, builder, title,
    folder_id=None, share_email=None
):
    """Create Google Doc, insert formatted content, and share."""

    # 1. Create empty doc
    doc = docs_service.documents().create(body={"title": title}).execute()
    doc_id = doc["documentId"]
    print(f"  📄 Created doc: {doc_id}")

    # 2. Apply all formatting
    requests = builder.get_gdocs_requests()

    # Split into chunks if too many requests (API limit ~100 per call)
    chunk_size = 80
    for i in range(0, len(requests), chunk_size):
        chunk = requests[i:i + chunk_size]
        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": chunk},
        ).execute()

    # 3. Move to folder (if specified)
    if folder_id:
        # Get current parent
        file_meta = drive_service.files().get(
            fileId=doc_id, fields="parents"
        ).execute()
        prev_parents = ",".join(file_meta.get("parents", []))

        drive_service.files().update(
            fileId=doc_id,
            addParents=folder_id,
            removeParents=prev_parents,
            fields="id, parents",
        ).execute()
        print(f"  📁 Moved to folder: {folder_id}")

    # 4. Share with user (if specified)
    if share_email:
        drive_service.permissions().create(
            fileId=doc_id,
            body={
                "type": "user",
                "role": "writer",
                "emailAddress": share_email,
            },
            sendNotificationEmail=False,
        ).execute()
        print(f"  👤 Shared with: {share_email}")

    doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"
    return doc_id, doc_url


# ================================================================
# Main Pipeline
# ================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Digital Zoo Linkbuilding Article Generator"
    )
    parser.add_argument("--excel", required=True, help="Excel 檔案路徑")
    parser.add_argument("--batch", required=True, type=int, help="Batch 編號 (1-4)")
    parser.add_argument("--creds", required=True, help="Google Service Account JSON")
    parser.add_argument("--folder", default=None, help="Google Drive 資料夾 ID")
    parser.add_argument("--share", default=None, help="分享給指定 email")
    parser.add_argument("--start", type=int, default=None, help="從第 N 篇開始")
    parser.add_argument("--end", type=int, default=None, help="到第 N 篇結束")
    parser.add_argument("--sheet", default="202605", help="Excel sheet 名稱")
    parser.add_argument("--dry-run", action="store_true", help="只生成文字不建 Google Doc")
    args = parser.parse_args()

    # Validate API key
    api_key = OPENROUTER_API_KEY
    if not api_key:
        print("✗ 請設定環境變數 OPENROUTER_API_KEY")
        sys.exit(1)

    # Parse Excel
    print(f"\n📊 讀取 Excel: {args.excel} (Batch {args.batch})")
    articles = parse_excel(args.excel, args.batch, args.sheet)

    if not articles:
        print("✗ 找不到指定 Batch 的文章")
        sys.exit(1)

    # Filter by start/end
    if args.start:
        articles = [a for a in articles if a["number"] >= args.start]
    if args.end:
        articles = [a for a in articles if a["number"] <= args.end]

    print(f"  找到 {len(articles)} 篇文章 (#{articles[0]['number']}-#{articles[-1]['number']})")

    # Generate articles
    print(f"\n🤖 開始生成文章 (Model: {MODEL})")
    builder = DocBuilder()
    failed = []

    for article in tqdm(articles, desc="生成進度"):
        num = article["number"]
        kw1 = article["keyword1"]
        kw2 = article["keyword2"]
        tqdm.write(f"  #{num}: {kw1}" + (f" + {kw2}" if kw2 else ""))

        content = generate_article_content(article, api_key, MODEL)

        if content:
            builder.build_article(article, content)
            tqdm.write(f"  ✓ #{num} 完成")
        else:
            failed.append(num)
            tqdm.write(f"  ✗ #{num} 失敗")

        time.sleep(DELAY_BETWEEN_CALLS)

    # Summary
    print(f"\n📝 生成完成: {len(articles) - len(failed)}/{len(articles)} 篇")
    if failed:
        print(f"  ⚠ 失敗文章: {failed}")

    # Dry run - save to local file
    if args.dry_run:
        out_path = f"batch_{args.batch}_output.txt"
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(builder.text)
        print(f"\n💾 Dry run 輸出: {out_path}")
        return

    # Create Google Doc
    if not Path(args.creds).exists():
        print(f"✗ 找不到 credentials 檔案: {args.creds}")
        sys.exit(1)

    print(f"\n📄 建立 Google Doc...")
    docs_svc, drive_svc = get_google_services(args.creds)

    title = f"Combined_2026_May_Internal_Batch_{args.batch}"
    if args.start or args.end:
        title += f"_#{args.start or articles[0]['number']}-{args.end or articles[-1]['number']}"

    doc_id, doc_url = create_formatted_doc(
        docs_svc, drive_svc, builder, title,
        folder_id=args.folder,
        share_email=args.share,
    )

    print(f"\n✅ 完成！")
    print(f"   文件名: {title}")
    print(f"   連結: {doc_url}")


if __name__ == "__main__":
    main()
