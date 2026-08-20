"""
Digital Zoo Linkbuild Generator — Web UI
=========================================
streamlit run app.py
"""

import io
import os
import time
import zipfile
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st

st.set_page_config(
    page_title="DZ Linkbuild Generator",
    page_icon="🔗",
    layout="wide",
)

from generate import (
    PLACEHOLDER_URL,
    build_docx_file,
    build_single_docx,
    detect_language,
    generate_article_content,
    parse_excel,
)

# ================================================================
# Secrets
# ================================================================
try:
    API_KEY = st.secrets["OPENROUTER_API_KEY"]
except Exception:
    API_KEY = os.environ.get("OPENROUTER_API_KEY", "")

if not API_KEY:
    st.error("⚠️ 請設定 OPENROUTER_API_KEY（Streamlit Secrets 或環境變數）")
    st.stop()

DEFAULT_MODEL = "~deepseek/deepseek-v4-flash-latest"
try:
    MODEL = st.secrets.get("LB_MODEL", DEFAULT_MODEL)
except Exception:
    MODEL = os.environ.get("LB_MODEL", DEFAULT_MODEL)


# ================================================================
# Cached Excel parsing (only runs ONCE per file, not on every click)
# ================================================================
@st.cache_data(show_spinner="📊 讀取 Excel 中...")
def load_all_batches(file_bytes, sheet_name=None):
    """Parse Excel once and cache results. Re-runs only when file changes."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    try:
        batch_counts = {}
        warnings = []
        for b in range(1, 10):
            report = []
            arts = parse_excel(tmp_path, b, sheet_name, report=report)
            if arts:
                batch_counts[b] = arts
                warnings.extend(f"Batch {b}：{w}" for w in report)
        return batch_counts, warnings
    finally:
        os.unlink(tmp_path)


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        from openpyxl import load_workbook
        wb = load_workbook(tmp_path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    finally:
        os.unlink(tmp_path)


# ================================================================
# UI
# ================================================================
st.title("🔗 DZ Linkbuild Generator")
st.caption("上傳 Excel → 選擇 Batch → 合規檢查 → 生成 .docx → 下載後拖入 Google Drive")
st.caption(f"🤖 Model：`{MODEL}`")

# ── Upload ──
excel_file = st.file_uploader("📊 上傳 Linkbuilding Excel", type=["xlsx"])

if not excel_file:
    st.info("👆 請先上傳 Excel 檔案")
    st.stop()

file_bytes = excel_file.getvalue()
sheet_names = list_sheets(file_bytes)
sheet_name = st.selectbox("📄 選擇 Sheet", sheet_names, index=0)

batch_counts, parse_warnings = load_all_batches(file_bytes, sheet_name)

if parse_warnings:
    with st.expander(f"⚠️ Excel 讀取警告（{len(parse_warnings)} 項）", expanded=True):
        for w in parse_warnings:
            st.warning(w)

if not batch_counts:
    st.error(f"Sheet「{sheet_name}」搵唔到任何 Batch 資料。請確認 A 欄有「Batch N」標籤。")
    st.stop()

# ── Select (instant, no re-parsing) ──
st.divider()

batch_options = list(batch_counts.keys())
batch_labels = [
    f"Batch {b} — {len(arts)} 篇（#{arts[0]['number']}-#{arts[-1]['number']}）"
    for b, arts in batch_counts.items()
]
selected_idx = st.selectbox(
    "📁 選擇 Batch",
    range(len(batch_options)),
    format_func=lambda i: batch_labels[i],
)
selected_batch = batch_options[selected_idx]
articles = batch_counts[selected_batch]

first_num = articles[0]["number"]
last_num = articles[-1]["number"]

select_mode = st.radio(
    "選擇方式",
    ["📏 連續範圍", "🔢 指定文章編號"],
    horizontal=True,
)

if select_mode == "📏 連續範圍":
    col_s, col_e = st.columns(2)
    with col_s:
        start_num = st.number_input(
            "起始 #", min_value=first_num, max_value=last_num, value=first_num,
        )
    with col_e:
        end_num = st.number_input(
            "結束 #", min_value=first_num, max_value=last_num, value=last_num,
        )
    filtered = [a for a in articles if start_num <= a["number"] <= end_num]
else:
    article_nums_input = st.text_input(
        "輸入文章編號（用逗號分隔）",
        placeholder=f"例如：{first_num}, {first_num+2}, {first_num+5}",
    )
    if article_nums_input:
        try:
            selected_nums = set(
                int(n.strip()) for n in article_nums_input.split(",") if n.strip()
            )
            filtered = [a for a in articles if a["number"] in selected_nums]
            not_found = selected_nums - {a["number"] for a in filtered}
            if not_found:
                st.warning(f"以下編號在此 Batch 中找不到：{sorted(not_found)}")
        except ValueError:
            st.error("格式錯誤，請輸入數字並用逗號分隔")
            filtered = []
    else:
        filtered = []

# ── Preview ──
with st.expander(f"📋 預覽（{len(filtered)} 篇）", expanded=False):
    preview_data = []
    for a in filtered:
        lang = detect_language(a["keyword1"] + a.get("keyword2", ""))
        preview_data.append({
            "#": a["number"],
            "Keyword 1": a["keyword1"],
            "Keyword 2": a["keyword2"] or "—",
            "Category": a["category"],
            "語言": "中文" if lang == "zh-HK" else "EN",
            "備註": "；".join(a.get("warnings", [])) or "—",
        })
    st.dataframe(preview_data, width="stretch", hide_index=True)

# ── Excel-level 警告 ──
placeholder_articles = [
    a["number"] for a in filtered
    if PLACEHOLDER_URL in (a["url1"], a["url2"])
]
if placeholder_articles:
    st.warning(
        f"⚠️ 以下文章喺 Excel 冇 target URL，會用 placeholder "
        f"`{PLACEHOLDER_URL}`，交稿前必須換返真實連結：{placeholder_articles}"
    )

# ── Settings ──
with st.expander("⚙️ 進階設定", expanded=False):
    parallel = st.slider(
        "同時生成篇數", 1, 5, 3,
        help="同時呼叫 API 的數量。越大越快，但太高可能觸發 rate limit。",
    )
    stagger = st.slider(
        "每次呼叫間隔（秒）", 0.0, 5.0, 1.5, step=0.5,
        help="避免同時打爆 OpenRouter rate limit。",
    )

# ── Buttons ──
st.divider()
no_articles = len(filtered) == 0
col_a, col_b, col_spacer = st.columns([1, 1, 2])
with col_a:
    btn_generate = st.button("🚀 生成文章", type="primary", width="stretch", disabled=no_articles)
with col_b:
    btn_dry = st.button("📝 Dry Run（預覽文字）", width="stretch", disabled=no_articles)

if no_articles and select_mode == "🔢 指定文章編號":
    st.info("👆 請輸入要生成的文章編號")


# ================================================================
# Generation
# ================================================================
def _summarise(log):
    """由 log 抽出一句人睇得明嘅失敗原因。"""
    text = "\n".join(log)
    if "429" in text or "rate" in text.lower():
        return "API rate limit（調低同時生成篇數 / 加大間隔）"
    if "OpenRouter 4" in text or "OpenRouter 5" in text:
        line = next((l for l in log if "OpenRouter" in l), "")
        return f"API 出錯 — {line.split('OpenRouter', 1)[-1].strip()[:120]}"
    if "API 出錯" in text or "Timeout" in text or "ConnectionError" in text:
        return "API 連線失敗 / timeout"
    if "JSONDecodeError" in text or "Expecting" in text:
        return "Model 回覆唔係合法 JSON"
    bullets = [l.strip("• ").strip() for l in log if l.strip().startswith("•")]
    if bullets:
        top = {}
        for b in bullets:
            key = ("字數" if "字數" in b else
                   "關鍵字字面重複" if "字面出現" in b else
                   "marker 缺失" if "搵唔到 marker" in b else
                   "關鍵字入咗標題" if "H1 或 H2" in b else
                   "H1 有標點" if "H1 唔准有標點" in b else
                   "H1 過長" if "H1 長度" in b else
                   "簡體字" if "簡體" in b else
                   "廣東話口語" if "廣東話" in b else
                   "英文稿夾中文" if "中文字" in b else
                   "破折號" if "破折號" in b else
                   "Keyword buffer" if "buffer" in b.lower() else b[:40])
            top[key] = top.get(key, 0) + 1
        worst = max(top.items(), key=lambda kv: kv[1])[0]
        return f"合規過唔到 — 主要係「{worst}」"
    return "生成失敗（詳情見下面 log）"


_rate_lock = threading.Lock()
_last_call = [0.0]


def _throttle(min_gap):
    """全域節流，避免 parallel worker 同一刻打爆 rate limit。"""
    if min_gap <= 0:
        return
    with _rate_lock:
        wait = _last_call[0] + min_gap - time.monotonic()
        if wait > 0:
            time.sleep(wait)
        _last_call[0] = time.monotonic()


def _generate_one(article, min_gap):
    """Worker function for parallel generation。log 會一齊帶返去畀 UI 顯示。"""
    _throttle(min_gap)
    log = []
    content = generate_article_content(article, API_KEY, MODEL, log=log)
    return article, content, log


if btn_generate or btn_dry:
    is_dry_run = btn_dry
    articles_with_content = []
    failed = []
    total = len(filtered)

    progress_bar = st.progress(0, text="準備中...")
    status_area = st.container()

    def _report(article, content, log):
        num = article["number"]
        if not content:
            failed.append((num, log))
            reason = _summarise(log)
            with status_area:
                st.error(f"❌ #{num} — {reason}")
                with st.expander(f"🔍 #{num} 失敗詳情（{len(log)} 行）", expanded=False):
                    st.code("\n".join(log) or "（冇任何訊息）", language="text")
            return
        articles_with_content.append((article, content))
        unit = "words" if content["_lang"] == "en" else "字"
        with status_area:
            st.success(
                f"✅ #{num} — {content['h1']}　"
                f"（{content['_word_count']} {unit}，合規通過）"
            )
            for w in content.get("_warnings", []):
                st.caption(f"　　⚠️ {w}")
            if content.get("_log"):
                with st.expander(f"🔍 #{num} 生成過程（{len(content['_log'])} 行）", expanded=False):
                    st.code("\n".join(content["_log"]), language="text")

    if parallel <= 1:
        for i, article in enumerate(filtered):
            label = f"#{article['number']}: {article['keyword1']}"
            progress_bar.progress(i / total, text=f"⏳（{i+1}/{total}）{label}")
            _report(*_generate_one(article, stagger))
    else:
        done_count = 0
        with ThreadPoolExecutor(max_workers=parallel) as executor:
            futures = [
                executor.submit(_generate_one, art, stagger)
                for art in filtered
            ]
            for future in as_completed(futures):
                article, content, log = future.result()
                done_count += 1
                progress_bar.progress(
                    done_count / total,
                    text=f"⏳（{done_count}/{total}）已完成 #{article['number']}",
                )
                _report(article, content, log)

        articles_with_content.sort(key=lambda x: x[0]["number"])

    progress_bar.progress(1.0, text=f"✅ 完成 {len(articles_with_content)}/{total} 篇")

    # ── Output ──
    st.divider()

    if not articles_with_content:
        st.error("所有文章都生成失敗")
    elif is_dry_run:
        st.subheader("📝 文字預覽")
        for article, content in articles_with_content:
            with st.expander(f"#{article['number']} — {content['h1']}", expanded=False):
                text = f"[H1: {content['h1']}]\n\n"
                for sec in content["sections"]:
                    if sec.get("h2"):
                        text += f"[H2: {sec['h2']}]\n\n"
                    text += sec.get("body", "") + "\n\n"
                st.text(text)
    else:
        with st.spinner("📄 建立 Word 文件中..."):
            nums = [a["number"] for a, _ in articles_with_content]
            filename = f"Combined_Batch_{selected_batch}"
            if len(articles_with_content) < len(articles):
                filename += f"_#{nums[0]}-{nums[-1]}" if len(nums) > 2 \
                    else "_#" + "_#".join(str(n) for n in nums)

            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
                build_docx_file(articles_with_content, tmp_docx.name)
                tmp_docx_path = tmp_docx.name
            with open(tmp_docx_path, "rb") as f:
                docx_bytes = f.read()
            os.unlink(tmp_docx_path)

            # 每篇一個檔，可以直接跑 dz-linkbuild 嘅 validate.py
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for article, content in articles_with_content:
                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
                        build_single_docx(article, content, t.name)
                        single_path = t.name
                    with open(single_path, "rb") as f:
                        zf.writestr(
                            f"{article['number']:03d}_{article['keyword1'][:20]}.docx",
                            f.read(),
                        )
                    os.unlink(single_path)
            zip_bytes = zip_buf.getvalue()

        st.balloons()
        st.success(f"🎉 已生成 {len(articles_with_content)} 篇文章，全部通過合規檢查")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label=f"⬇️ 下載 {filename}.docx（合併稿）",
                data=docx_bytes,
                file_name=f"{filename}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                width="stretch",
            )
            st.caption("💡 下載後拖入 Google Drive → 自動轉為 Google Doc")
        with col_d2:
            st.download_button(
                label=f"⬇️ 下載 {filename}_single.zip（每篇一個檔）",
                data=zip_bytes,
                file_name=f"{filename}_single.zip",
                mime="application/zip",
                width="stretch",
            )
            st.caption("💡 呢個 zip 入面嘅檔可以直接跑 `validate.py` 做最終覆核")

        used_placeholder = [
            a["number"] for a, _ in articles_with_content
            if PLACEHOLDER_URL in (a["url1"], a["url2"])
        ]
        if used_placeholder:
            st.warning(f"⚠️ 仲用緊 placeholder URL 嘅文章：{used_placeholder}")

    if failed:
        st.divider()
        st.subheader(f"⚠️ 失敗文章（{len(failed)} 篇）")
        buckets = {}
        for num, log in failed:
            buckets.setdefault(_summarise(log), []).append(num)
        for reason, nums in sorted(buckets.items(), key=lambda kv: -len(kv[1])):
            st.warning(f"**{reason}** — {len(nums)} 篇：{nums}")
        st.download_button(
            "⬇️ 下載完整失敗 log（.txt）",
            data="\n\n".join(
                f"===== #{num} =====\n" + "\n".join(log) for num, log in failed
            ),
            file_name="linkbuild_failures.log",
            mime="text/plain",
        )
