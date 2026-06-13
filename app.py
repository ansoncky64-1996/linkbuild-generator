"""
Digital Zoo Linkbuild Generator — Web UI (Zero-Config Version)
==============================================================
所有認證由 Streamlit Cloud Secrets 管理，用戶只需上傳 Excel 即可。

Local testing:  streamlit run app.py
Production:     Deploy to Streamlit Cloud
"""

import os
import time
import tempfile

import streamlit as st

st.set_page_config(
    page_title="DZ Linkbuild Generator",
    page_icon="🔗",
    layout="wide",
)

from generate import (
    parse_excel,
    generate_article_content,
    DocBuilder,
    get_google_services_from_info,
    create_formatted_doc,
    detect_language,
)

# ================================================================
# Read secrets (all auth is server-side, users never see these)
# ================================================================
try:
    API_KEY = st.secrets["OPENROUTER_API_KEY"]
    FOLDER_ID = st.secrets.get("GOOGLE_FOLDER_ID", "")
    SHARE_EMAIL = st.secrets.get("SHARE_EMAIL", "")
    MODEL = st.secrets.get("LB_MODEL", "deepseek/deepseek-v4-0324")
    GOOGLE_CREDS = st.secrets["GOOGLE_CREDENTIALS"]
except Exception:
    st.error("⚠️ 伺服器設定未完成。請聯絡 Anson 設定 Streamlit Secrets。")
    st.stop()

# ================================================================
# Session state
# ================================================================
for key, default in {
    "generated": [],
    "failed": [],
    "doc_url": None,
    "running": False,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ================================================================
# UI
# ================================================================
st.title("🔗 DZ Linkbuild Generator")
st.caption("上傳 Excel → 選擇 Batch → 自動生成格式化 Google Doc")

# ── Step 1: Upload Excel ──
excel_file = st.file_uploader(
    "📊 上傳 Linkbuilding Excel",
    type=["xlsx"],
    help="每月的 Internal Linkbuilding List",
)

if not excel_file:
    st.info("👆 請先上傳 Excel 檔案")
    st.stop()

# Save upload to temp file
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
    tmp.write(excel_file.read())
    tmp_excel_path = tmp.name

# Parse all batches
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
    st.error("在 Excel 中找不到任何 Batch 資料")
    st.stop()

# ── Step 2: Select Batch & Range ──
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
        })
    st.dataframe(preview_data, use_container_width=True, hide_index=True)

# ── Step 3: Generate ──
st.divider()

col_a, col_b, col_spacer = st.columns([1, 1, 2])
with col_a:
    btn_generate = st.button(
        "🚀 生成 Google Doc",
        type="primary",
        use_container_width=True,
    )
with col_b:
    btn_dry = st.button(
        "📝 Dry Run（預覽文字）",
        use_container_width=True,
    )

# ================================================================
# Generation
# ================================================================
if btn_generate or btn_dry:
    is_dry_run = btn_dry

    st.session_state.generated = []
    st.session_state.failed = []
    st.session_state.doc_url = None

    builder = DocBuilder()
    progress_bar = st.progress(0, text="準備中...")
    status_area = st.container()

    total = len(filtered)
    success_count = 0

    for i, article in enumerate(filtered):
        num = article["number"]
        kw1 = article["keyword1"]
        kw2 = article["keyword2"]
        label = f"#{num}: {kw1}" + (f" + {kw2}" if kw2 else "")

        progress_bar.progress(i / total, text=f"⏳ 生成中（{i+1}/{total}）：{label}")

        content = generate_article_content(article, API_KEY, MODEL)

        if content:
            builder.build_article(article, content)
            success_count += 1
            st.session_state.generated.append({
                "number": num,
                "h1": content["h1"],
            })
            with status_area:
                st.success(f"✅ #{num} — {content['h1']}")
        else:
            st.session_state.failed.append(num)
            with status_area:
                st.error(f"❌ #{num} — 生成失敗")

        if i < total - 1:
            time.sleep(3)

    progress_bar.progress(1.0, text=f"✅ 完成 {success_count}/{total} 篇")

    # ── Output ──
    st.divider()

    if is_dry_run:
        st.subheader("📝 文字預覽")
        st.download_button(
            "⬇️ 下載 TXT",
            data=builder.text.encode("utf-8"),
            file_name=f"batch_{selected_batch}_preview.txt",
            mime="text/plain",
        )
        with st.expander("預覽內容", expanded=True):
            preview = builder.text[:5000]
            if len(builder.text) > 5000:
                preview += "\n\n... (更多內容請下載 TXT 檔案) ..."
            st.text(preview)

    else:
        if success_count == 0:
            st.error("所有文章都生成失敗，無法建立 Google Doc")
        else:
            with st.spinner("📄 建立 Google Doc 中..."):
                try:
                    docs_svc, drive_svc = get_google_services_from_info(GOOGLE_CREDS)

                    title = f"Combined_2026_May_Internal_Batch_{selected_batch}"
                    if start_num != first_num or end_num != last_num:
                        title += f"_#{start_num}-{end_num}"

                    doc_id, doc_url = create_formatted_doc(
                        docs_svc, drive_svc, builder, title,
                        folder_id=FOLDER_ID or None,
                        share_email=SHARE_EMAIL or None,
                    )
                    st.session_state.doc_url = doc_url

                except Exception as e:
                    st.error(f"Google Doc 建立失敗：{e}")

            if st.session_state.doc_url:
                st.balloons()
                st.success("🎉 完成！Google Doc 已建立並分享到指定資料夾。")
                st.markdown(f"### 📄 [打開 Google Doc]({st.session_state.doc_url})")

    if st.session_state.failed:
        st.warning(
            f"⚠️ 以下文章生成失敗：{st.session_state.failed}。"
            f"可調整範圍重新跑這些文章。"
        )

# Cleanup
try:
    os.unlink(tmp_excel_path)
except Exception:
    pass
