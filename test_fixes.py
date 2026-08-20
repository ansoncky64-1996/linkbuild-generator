# -*- coding: utf-8 -*-
"""Regression tests：逐個對返之前 confirm 到嘅 bug。"""
import json, os, sys, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "linkbuild-generator"))
import generate as g

PASS, FAIL = [], []
def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  ✅ " if cond else "  ❌ ") + name + (f"  → {detail}" if detail else ""))

def prep(result, article, lang):
    r = json.loads(json.dumps(result))
    r = g._normalize_output(r, lang)
    r = g._place_markers(r, article)
    r, dropped = g._strip_unmapped_markers(r, article)
    return r, dropped

print("\n【Bug 1】kw1 係 kw2 子字串 —— 唔可以再剷爛文字")
art = {"keyword1": "紅酒", "keyword2": "紅酒櫃", "url1": "u1", "url2": "u2"}
res = {"h1": "紅酒櫃入門指南", "sections": [
    {"h2": None, "body": "香港人近年開始鑽研飲食文化，家居收藏亦成為話題。"},
    {"h2": "家居收藏的興起", "body": "不少家庭會購置紅酒櫃來存放藏品，溫度控制十分重要。"},
    {"h2": "緩衝段落", "body": "除了溫度，濕度同樣影響存放質素。"},
    {"h2": "選購考量", "body": "選購時要留意容量與壓縮機類型。"},
    {"h2": "總結", "body": "長遠而言，妥善存放能保障品質。"}]}
out, _ = prep(res, art, "zh-HK")
check("H1 冇被剷爛", out["h1"] == "紅酒櫃入門指南", repr(out["h1"]))
check("正文冇被剷爛", "購置" in out["sections"][1]["body"] and "購置櫃" not in out["sections"][1]["body"],
      out["sections"][1]["body"])
check("長 keyword 優先認領（{{KW2}} 貼咗落「紅酒櫃」）",
      "{{KW2}}" in out["sections"][1]["body"])
fails, _ = g.validate_content(out, art, "zh-HK")
check("H1 含關鍵字 → 報 FAIL 而唔係靜靜地剷", any("H1" in f or "標題" in f for f in fails), str(fails[:2]))

print("\n【Bug 1b】英文 keyword 唔可以食詞中詞")
art_en = {"keyword1": "wine cellar", "keyword2": "wine fridge", "url1": "u1", "url2": "u2"}
res_en = {"h1": "How Collectors Store Their Bottles", "sections": [
    {"h2": None, "body": "Collecting has become a serious hobby in Hong Kong."},
    {"h2": "Storage Basics", "body": "A {{KW1}} keeps humidity stable, and many wine cellars are custom built."},
    {"h2": "Buffer", "body": "Temperature swings damage corks over time."},
    {"h2": "Compact Options", "body": "A {{KW2}} suits smaller flats."},
    {"h2": "Closing", "body": "Good storage protects value."}]}
out_en, _ = prep(res_en, art_en, "en")
check("\"many wine cellars\" 冇變 \"many s\"",
      "many wine cellars" in out_en["sections"][1]["body"], out_en["sections"][1]["body"])
fails_en, _ = g.validate_content(out_en, art_en, "en")
check("重複字面 → 報 FAIL", any("字面出現" in f for f in fails_en), str([f for f in fails_en if "字面" in f]))

print("\n【Bug 2】簡體 keyword：normalize 要行喺 keyword 處理之前")
art2 = {"keyword1": "紅酒投資", "keyword2": "威士忌收藏", "url1": "u1", "url2": "u2"}
res2 = {"h1": "另類資產的長線思維", "sections": [
    {"h2": None, "body": "近年另類資產受注目。"},
    {"h2": "市場背景", "body": "谈到红酒投资，回报要看年份。"},
    {"h2": "緩衝", "body": "流動性是另一個考慮。"},
    {"h2": "實物資產", "body": "{{KW2}}需要倉儲配套。"},
    {"h2": "總結", "body": "長線持有為主。"}]}
out2, _ = prep(res2, art2, "zh-HK")
body_all = "".join(s["body"] for s in out2["sections"])
check("簡體「红酒投资」轉繁後被認出做 marker", "{{KW1}}" in body_all, body_all)
check("冇留低無 link 嘅重複關鍵字",
      g._kw_count(g._mask_markers(body_all), "紅酒投資") == 0)
fails2, _ = g.validate_content(out2, art2, "zh-HK")
check("Buffer 規則:KW1(idx1) vs KW2(idx3) 合格", not any("buffer" in f.lower() for f in fails2))

print("\n【Bug 3】單 keyword 文章唔可以漏 {{KW2}} 落 .docx")
art3 = {"number": 7, "keyword1": "脫髮治療", "keyword2": "", "url1": "https://x.com/a", "url2": "", "category": "Health"}
res3 = {"h1": "頭皮健康的日常管理", "sections": [
    {"h2": None, "body": "頭皮健康愈受關注。"},
    {"h2": "成因", "body": "壓力與飲食都有影響。{{KW1}}要對症下藥。"},
    {"h2": "緩衝", "body": "作息同樣關鍵。"},
    {"h2": "選擇", "body": "坊間方案眾多 {{KW2}} 值得比較。"},
    {"h2": "總結", "body": "耐心最重要。"}]}
out3, dropped = prep(res3, art3, "zh-HK")
check("{{KW2}} 被清走", dropped == 1 and "{{KW2}}" not in json.dumps(out3, ensure_ascii=False))
g.build_docx_file([(art3, out3)], "/tmp/n1.docx")
from docx import Document
txt = "\n".join(p.text for p in Document("/tmp/n1.docx").paragraphs)
check(".docx 成品冇 marker 字面", "{{KW" not in txt)
check(".docx 用 [H1: ] 格式", "[H1: 頭皮健康的日常管理]" in txt, [l for l in txt.split("\n") if "H1" in l])

print("\n【Bug 4】Excel 文字格式序號唔可以靜靜地丟資料")
import openpyxl
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "202608"
rows = [
    ("Batch 1", 1, 1, "紅酒櫃", "https://a.com/1", "Lifestyle"),
    (None, 1, "2", "酒櫃推介", "https://a.com/2", None),
    (None, 2, "1", "wine cellar", "https://c.com/1", "Lifestyle"),
    (None, 2, "2", "wine fridge", "https://c.com/2", None),
    (None, 3, 1, "脫髮治療", None, "Health"),
    (None, 3, 2, "--", "--", None),
]
ws.append(["Batch", "Art", "KW#", "Keyword", "URL", "F", "G", "H", "I", "J", "Category"])
for b, a, c, d, e, k in rows:
    ws.append([b, a, c, d, e, None, None, None, None, None, k])
wb.save("/tmp/t.xlsx")
rep = []
arts = g.parse_excel("/tmp/t.xlsx", 1, "202608", report=rep)
check("3 篇全部讀到", len(arts) == 3, f"實際 {len(arts)} 篇")
check("文字格式 '2' 嘅 keyword 2 冇丟", arts[0]["keyword2"] == "酒櫃推介", arts[0])
check("文字格式 '1' 嘅文章冇丟", any(a["keyword1"] == "wine cellar" for a in arts))
check("冇 URL → 用 placeholder 兼有警告",
      arts[2]["url1"] == g.PLACEHOLDER_URL and any("placeholder" in w for w in arts[2]["warnings"]),
      arts[2]["warnings"])
check("sheet_name 唔指定會用第一個 sheet", len(g.parse_excel("/tmp/t.xlsx", 1)) == 3)

print("\n【Bug 5】字數 gate")
short = {"h1": "測試標題", "sections": [
    {"h2": None, "body": "短" * 50}, {"h2": "A", "body": "短" * 50},
    {"h2": "B", "body": "短" * 50}, {"h2": "C", "body": "短" * 50},
    {"h2": "D", "body": "短" * 50}]}
f5, _ = g.validate_content(short, {"keyword1": "", "keyword2": ""}, "zh-HK")
check("250 字會被擋", any("唔夠 750" in f for f in f5), str([f for f in f5 if "字數" in f]))
long_ = json.loads(json.dumps(short))
for s in long_["sections"]: s["body"] = "長" * 300
f5b, _ = g.validate_content(long_, {"keyword1": "", "keyword2": ""}, "zh-HK")
check("1500 字會被擋", any("超出 1000" in f for f in f5b))
check("字數計算包含 keyword anchor 文字",
      g._count_words({"sections": [{"body": "abc {{KW1}} def"}]}, "zh-HK",
                     {"keyword1": "紅酒櫃", "keyword2": ""}) == 3)

print("\n【Bug 6】中英文 prompt 唔可以撈埋")
import re as _re
p_en = g.build_prompt({"keyword1": "wine cellar", "keyword2": "wine fridge",
                       "url1": "u", "url2": "u", "category": "Lifestyle"})
check("英文文章 prompt 零中文字", len(g.CJK_RE.findall(p_en)) == 0,
      f"{len(g.CJK_RE.findall(p_en))} 個中文字")
p_zh = g.build_prompt({"keyword1": "紅酒櫃", "keyword2": "酒櫃推介",
                       "url1": "u", "url2": "u", "category": "生活"})
check("中文文章 prompt 係中文", len(g.CJK_RE.findall(p_zh)) > 200)
for kw, exp in [("wine cellar", "en"), ("紅酒櫃", "zh-HK"), ("wine 櫃", "zh-HK"),
                ("Château Margaux 2015", "en"), ("日本威士忌 Yamazaki", "zh-HK")]:
    check(f"detect_language({kw!r}) = {exp}", g.detect_language(kw) == exp)
check("英文稿混入中文 → 報 FAIL",
      any("中文字" in f for f in g.validate_content(
          {"h1": "A Clean Title", "sections": [
              {"h2": None, "body": "word " * 200 + "這是中文"},
              {"h2": "A", "body": "word " * 200}, {"h2": "B", "body": "word " * 200},
              {"h2": "C", "body": "word " * 200}, {"h2": "D", "body": "word " * 150}]},
          {"keyword1": "", "keyword2": ""}, "en")[0]))
check("中文稿有廣東話口語 → 報 FAIL",
      any("廣東話" in f for f in g.validate_content(
          {"h1": "測試標題", "sections": [
              {"h2": None, "body": "字" * 200 + "呢個係嘅"},
              {"h2": "A", "body": "字" * 200}, {"h2": "B", "body": "字" * 200},
              {"h2": "C", "body": "字" * 200}, {"h2": "D", "body": "字" * 150}]},
          {"keyword1": "", "keyword2": ""}, "zh-HK")[0]))

print("\n【Bug 8】空 URL 唔可以整壞 .docx")
art8 = {"number": 8, "keyword1": "紅酒櫃", "keyword2": "酒櫃推介",
        "url1": "https://x.com/a", "url2": "", "category": "L"}
c8 = {"h1": "測試標題", "sections": [
    {"h2": None, "body": "開篇"}, {"h2": "A", "body": "x{{KW1}}y"},
    {"h2": "B", "body": "b"}, {"h2": "C", "body": "p{{KW2}}q"}, {"h2": "D", "body": "z"}]}
g.build_docx_file([(art8, c8)], "/tmp/n2.docx")
import zipfile
rels = zipfile.ZipFile("/tmp/n2.docx").read("word/_rels/document.xml.rels").decode()
check("冇 Target=\"\" 嘅 relationship", 'Target=""' not in rels)

print("\n【Bug 9】marker 搵唔到唔可以硬塞落段尾")
art9 = {"keyword1": "威士忌投資", "keyword2": "橡木桶", "url1": "u", "url2": "u"}
res9 = {"h1": "烈酒市場觀察", "sections": [
    {"h2": None, "body": "烈酒市場近年轉變不少。"}, {"h2": "一", "body": "產區與年份決定價格。"},
    {"h2": "二", "body": "倉儲成本亦要計算。"}, {"h2": "三", "body": "轉手渠道有限。"},
    {"h2": "四", "body": "投資者需保持耐性。"}]}
out9, _ = prep(res9, art9, "zh-HK")
check("冇硬貼 keyword 落段尾",
      not any(s["body"].endswith("威士忌投資") or s["body"].endswith("橡木桶")
              for s in out9["sections"]))
f9, _ = g.validate_content(out9, art9, "zh-HK")
check("marker 缺失 → 報 FAIL 交去 repair", sum("搵唔到 marker" in f for f in f9) == 2)

print("\n【Bug 10 / 7】合規 gate + 輸出格式")
good_art = {"number": 1, "keyword1": "紅酒櫃", "keyword2": "威士忌收藏",
            "url1": "https://a.com/1", "url2": "https://a.com/2", "category": "L"}
good = {"h1": "家居藏酒空間的規劃思路", "sections": [
    {"h2": None, "body": "近年香港家居收藏文化" + "興" * 175 + "。"},
    {"h2": "空間規劃", "body": "在規劃時要考慮{{KW1}}的擺放位置" + "與" * 175 + "。"},
    {"h2": "環境條件", "body": "濕度與震動同樣重要" + "而" * 175 + "。"},
    {"h2": "藏品管理", "body": "談到{{KW2}}的紀錄方式" + "則" * 175 + "。"},
    {"h2": "長線維護", "body": "定期檢視可以保障價值" + "並" * 175 + "。"}]}
gout, _ = prep(good, good_art, "zh-HK")
gf, gw = g.validate_content(gout, good_art, "zh-HK")
check("合格稿 0 FAIL", not gf, str(gf))
g.build_single_docx(good_art, gout, "/tmp/single.docx")
check("build_single_docx 產出成功", os.path.exists("/tmp/single.docx"))

print("\n【新發現】OpenCC s2hk 唔可以改爛正確嘅繁體字")
for t in ["客戶服務", "說明書", "稅務安排", "脫髮治療", "溫度控制", "群組管理",
          "閱讀體驗", "兌換率", "臥室設計", "蘊含", "醞釀"]:
    check(f"{t} 保持原樣", g._to_traditional(t) == t, g._to_traditional(t))
check("簡體照樣轉到繁體", g._to_traditional("客户体验说明书") == "客戶體驗說明書",
      g._to_traditional("客户体验说明书"))

print("\n【End-to-end】mock API,行齊 normalize → marker → strip → validate")
CANNED_BAD = json.dumps({
    "h1": "家居藏酒空間的規劃思路",
    "sections": [
        {"h2": None, "body": "近年香港家居收藏文化" + "興" * 60 + "。"},
        {"h2": "空间规划", "body": "谈到红酒柜的摆放位置" + "与" * 60 + "。红酒柜很重要。"},
        {"h2": "環境條件", "body": "濕度與震動同樣重要" + "而" * 60 + "。"},
        {"h2": "藏品管理", "body": "至於{{KW2}}的紀錄方式" + "則" * 60 + "。"},
        {"h2": "長線維護", "body": "定期檢視可以保障價值" + "並" * 60 + "。{{KW3}}"}],
}, ensure_ascii=False)
CANNED_GOOD = json.dumps({
    "h1": "家居藏酒空間的規劃思路",
    "sections": [
        {"h2": None, "body": "近年香港家居收藏文化" + "興" * 175 + "。"},
        {"h2": "空間規劃", "body": "談到{{KW1}}的擺放位置" + "與" * 175 + "。"},
        {"h2": "環境條件", "body": "濕度與震動同樣重要" + "而" * 175 + "。"},
        {"h2": "藏品管理", "body": "至於{{KW2}}的紀錄方式" + "則" * 175 + "。"},
        {"h2": "長線維護", "body": "定期檢視可以保障價值" + "並" * 175 + "。"}],
}, ensure_ascii=False)

calls = []
def fake_chat(api_key, model, prompt, **kw):
    calls.append(prompt)
    return CANNED_BAD if len(calls) == 1 else CANNED_GOOD
g._chat = fake_chat

e2e_art = {"number": 1, "keyword1": "紅酒櫃", "keyword2": "威士忌收藏",
           "url1": "https://a.com/1", "url2": "https://a.com/2", "category": "Lifestyle"}
out = g.generate_article_content(e2e_art, "k", "m")
check("E2E 回到合規稿", out is not None)
if out:
    check("E2E 觸發咗 repair", len(calls) == 2, f"{len(calls)} 次 API call")
    check("E2E 簡體已清乾淨", not (set("".join(s["body"] for s in out["sections"])) & g.SIMPLIFIED))
    check("E2E 無效 marker {{KW3}} 已清走", "{{KW3}}" not in json.dumps(out, ensure_ascii=False))
    check("E2E 字數達標", g.MIN_BODY <= out["_word_count"] <= g.MAX_BODY, out["_word_count"])
    check("E2E 0 FAIL", not g.validate_content(out, e2e_art, "zh-HK")[0])

print("\n【End-to-end】改極都改唔好 → 一定要回 None,唔可以出稿")
calls2 = []
def always_bad(api_key, model, prompt, **kw):
    calls2.append(1)
    return CANNED_BAD
g._chat = always_bad
import time as _t
_orig_sleep = _t.sleep; _t.sleep = lambda *a: None
bad_out = g.generate_article_content(e2e_art, "k", "m", max_retries=1, max_repairs=1)
_t.sleep = _orig_sleep
check("屢改不成 → 明確標示未合規(唔會扮合格)",
      bad_out is not None and bad_out.get("_compliant") is False and bad_out.get("_fails"))
g._chat = always_bad
_t.sleep = lambda *a: None
none_out = g.generate_article_content(e2e_art, "k", "m", max_retries=1, max_repairs=1,
                                      return_best_effort=False)
_t.sleep = _orig_sleep
check("return_best_effort=False → 回 None", none_out is None)

print("\n【新修】reasoning model / 大細階 / JSON 容錯")
check("中英混合 keyword 唔分大細階",
      g._kw_count("業界普遍採用 MPLS 虛擬專用網絡 方案", "mpls 虛擬專用網絡") == 1)
check("純英文 keyword 一樣唔分大細階", g._kw_count("SD-WAN and SD WAN", "sd wan") == 1)
check("JSON 字串入面有原始換行都解析到",
      g._extract_json('{"h1": "a\nb", "sections": []}')["h1"] == "a\nb")
check("ARTICLE_SCHEMA 鎖死 5 個 section",
      g.ARTICLE_SCHEMA["schema"]["properties"]["sections"]["minItems"] == 5
      and g.ARTICLE_SCHEMA["schema"]["properties"]["sections"]["maxItems"] == 5)
check("h2 schema 容許 null",
      "null" in g.ARTICLE_SCHEMA["schema"]["properties"]["sections"]["items"]["properties"]["h2"]["type"])
check("max_tokens 預設由 6000 升到 16000", g.MAX_TOKENS == 16000)
check("預設關掉 reasoning", g.DISABLE_REASONING is True)

print("\n" + "=" * 60)
print(f"PASS {len(PASS)} / FAIL {len(FAIL)}")
if FAIL:
    for f in FAIL: print("   ❌ " + f)
sys.exit(1 if FAIL else 0)
