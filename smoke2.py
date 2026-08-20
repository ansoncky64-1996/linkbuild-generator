import sys, warnings; warnings.filterwarnings("ignore")
from streamlit.testing.v1 import AppTest
P, F = [], []
def check(n, c, d=""):
    (P if c else F).append(n); print(("  ✅ " if c else "  ❌ ") + n + (f"  → {d}" if d else ""))
def dls(at):
    try: return [e.label for e in at.get("download_button")]
    except Exception: return []

at = AppTest.from_file("smoke_wrapper.py", default_timeout=180); at.run()
check("上傳後 app 冇 exception", not at.exception, str(at.exception)[:300] if at.exception else "")
check("有生成掣", any("生成文章" in b.label for b in at.button))
check("未生成時冇下載掣", not dls(at), dls(at))

next(b for b in at.button if "生成文章" in b.label).click().run()
check("生成後冇 exception", not at.exception, str(at.exception)[:400] if at.exception else "")
check("有合併稿下載掣", any(".docx" in l for l in dls(at)), dls(at))
check("有單篇 zip 下載掣", any(".zip" in l for l in dls(at)))
subs = [str(s.value) for s in at.subheader]
check("有「已生成」區", any("已生成" in x for x in subs), subs)
check("未合規草稿獨立一區", any("未合規" in x for x in subs), subs)
check("合規稿只有 #1", list(at.session_state["done"]) == [(1, 1)], list(at.session_state["done"]))
check("未合規 #2 分開存", list(at.session_state["drafts"]) == [(1, 2)], list(at.session_state["drafts"]))
check("交付稿唔包未合規", any("#1-1.docx" in l for l in dls(at)), dls(at))

before = len(at.session_state["done"])
next(b for b in at.button if "生成文章" in b.label).click().run()
check("重跑跳過已完成", any("跳過已完成" in i.value for i in at.info), [i.value for i in at.info])
check("重跑後結果仲喺度", len(at.session_state["done"]) >= before)
check("重跑後下載掣仲喺度", any(".docx" in l for l in dls(at)))

print(); print("=" * 60); print(f"PASS {len(P)} / FAIL {len(F)}")
for f in F: print("  ❌", f)
sys.exit(1 if F else 0)
