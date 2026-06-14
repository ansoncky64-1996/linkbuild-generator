"""
Digital Zoo Linkbuild Generator — Web UI
=========================================
streamlit run app.py
"""

import os
import time
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st

st.set_page_config(
    page_title="DZ Linkbuild Generator",
    page_icon="🔗",
    layout="wide",
)

from generate import (
    parse_excel,
    generate_article_content,
    build_docx_file,
    detect_language,
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

try:
    MODEL = st.secrets.get("LB_MODEL", "deepseek/deepseek-chat-v3-0324")
except Exception:
    MODEL = os.environ.get("LB_MODEL", "deepseek/deepseek-chat-v3-0324")

# ================================================================
# UI
# ================================================================
st.title("🔗 DZ Linkbuild Generator")
st.caption("上傳 Excel → 選擇 Batch → 生成 .docx → 下載後拖入 Google Drive")

# ── Upload ──
excel_file = st.file_uploader("📊 上傳 Linkbuilding Excel", type=["xlsx"])

if not excel_file:
    st.info("👆 請先上傳 Excel 檔案")
    st.stop()

with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
    tmp.write(excel_file.read())
    tmp_excel_path = tmp.name

try:
    from openpyxl import load_workbook
    wb = load_workbook(tmp_excel_path)
    sheet_name = wb.sheetnames[0]
    batch_counts = {}
    for b in range(1, 10):
        arts = parse_excel(tmp_excel_path, b, sheet_name)
        if arts:
            batch_counts[b] = arts
    wb.close()
except Exception as e:
    st.error(f"Excel 讀取失敗：{e}")
    st.stop()

if not batch_counts:
    st.error("找不到任何 Batch 資料")
    st.stop()

# ── Select ──
st.divider()
col1, col2, col3 = st.columns([2, 1, 1])

with col1:
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

with col2:
    first_num = articles[0]["number"]
    last_num = articles[-1]["number"]
    start_num = st.number_input(
        "起始 #", min_value=first_num, max_value=last_num, value=first_num,
    )

with col3:
    end_num = st.number_input(
        "結束 #", min_value=first_num, max_value=last_num, value=last_num,
    )

filtered = [a for a in articles if start_num <= a["number"] <= end_num]

# ── Preview table ──
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
        })
    st.dataframe(preview_data, use_container_width=True, hide_index=True)

# ── Settings ──
with st.expander("⚙️ 進階設定", expanded=False):
    parallel = st.slider(
        "同時生成篇數", 1, 5, 3,
        help="同時呼叫 API 的數量。數字越大越快，但太高可能觸發 rate limit。",
    )

# ── Buttons ──
st.divider()
col_a, col_b, col_spacer = st.columns([1, 1, 2])
with col_a:
    btn_generate = st.button("🚀 生成文章", type="primary", use_container_width=True)
with col_b:
    btn_dry = st.button("📝 Dry Run（預覽文字）", use_container_width=True)


# ================================================================
# Generation
# ================================================================
def _generate_one(article):
    """Worker function for parallel generation."""
    content = generate_article_content(article, API_KEY, MODEL)
    return article, content


if btn_generate or btn_dry:
    is_dry_run = btn_dry
    articles_with_content = []
    failed = []
    total = len(filtered)

    progress_bar = st.progress(0, text="準備中...")
    status_area = st.container()

    if parallel <= 1:
        # Sequential
        for i, article in enumerate(filtered):
            num = article["number"]
            label = f"#{num}: {article['keyword1']}"
            progress_bar.progress(i / total, text=f"⏳（{i+1}/{total}）{label}")

            content = generate_article_content(article, API_KEY, MODEL)
            if content:
                articles_with_content.append((article, content))
                with status_area:
                    st.success(f"✅ #{num} — {content['h1']}")
            else:
                failed.append(num)
                with status_area:
                    st.error(f"❌ #{num} — 生成失敗")
    else:
        # Parallel
        done_count = 0
        with ThreadPoolExecutor(max_workers=parallel) as executor:
            futures = {
                executor.submit(_generate_one, art): art
                for art in filtered
            }
            for future in as_completed(futures):
                article, content = future.result()
                done_count += 1
                num = article["number"]
                progress_bar.progress(
                    done_count / total,
                    text=f"⏳（{done_count}/{total}）已完成 #{num}",
                )
                if content:
                    articles_with_content.append((article, content))
                    with status_area:
                        st.success(f"✅ #{num} — {content['h1']}")
                else:
                    failed.append(num)
                    with status_area:
                        st.error(f"❌ #{num} — 生成失敗")

        # Sort by article number (parallel results come in random order)
        articles_with_content.sort(key=lambda x: x[0]["number"])

    progress_bar.progress(1.0, text=f"✅ 完成 {len(articles_with_content)}/{total} 篇")

    # ── Output ──
    st.divider()

    if not articles_with_content:
        st.error("所有文章都生成失敗")
    elif is_dry_run:
        # Text preview
        st.subheader("📝 文字預覽")
        for article, content in articles_with_content:
            with st.expander(f"#{article['number']} — {content['h1']}", expanded=False):
                text = f"H1：{content['h1']}\n\n"
                for sec in content["sections"]:
                    if sec.get("h2"):
                        text += f"H2：{sec['h2']}\n\n"
                    text += sec.get("body", "") + "\n\n"
                st.text(text)
    else:
        # Build .docx
        with st.spinner("📄 建立 Word 文件中..."):
            filename = f"Combined_Batch_{selected_batch}"
            if start_num != first_num or end_num != last_num:
                filename += f"_#{start_num}-{end_num}"
            filename += ".docx"

            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
                build_docx_file(articles_with_content, tmp_docx.name)
                tmp_docx_path = tmp_docx.name

            with open(tmp_docx_path, "rb") as f:
                docx_bytes = f.read()
            os.unlink(tmp_docx_path)

        st.balloons()
        st.success(f"🎉 已生成 {len(articles_with_content)} 篇文章")
        st.download_button(
            label=f"⬇️ 下載 {filename}",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
        st.caption("💡 下載後拖入 Google Drive → 自動轉為 Google Doc")

    if failed:
        st.warning(f"⚠️ 失敗文章：{failed}")

try:
    os.unlink(tmp_excel_path)
except Exception:
    pass
