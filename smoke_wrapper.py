# -*- coding: utf-8 -*-
"""把 file_uploader 同 API 都 patch 咗,再跑成個 app.py。"""
import io, json, os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "linkbuild-generator"))
os.environ["OPENROUTER_API_KEY"] = "test-key"

import streamlit as st
import openpyxl

# ── 造 Excel ──
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "202608"
ws.append(["Batch","Art","KW#","Keyword","URL","F","G","H","I","J","Category"])
for row in [("Batch 1",1,1,"紅酒櫃","https://a.com/1","Lifestyle"),
            (None,1,"2","酒櫃推介","https://a.com/2",None),
            (None,2,1,"威士忌收藏","https://a.com/3","Finance"),
            (None,2,2,"橡木桶","https://a.com/4",None)]:
    b,a,c,d,e,k = row
    ws.append([b,a,c,d,e,None,None,None,None,None,k])
buf = io.BytesIO(); wb.save(buf); XLSX = buf.getvalue()

class FakeUpload:
    name = "batch.xlsx"
    def getvalue(self): return XLSX

st.file_uploader = lambda *a, **k: FakeUpload()

# ── Patch API:第一篇成功,第二篇一直失敗（試 circuit breaker + best effort）──
import generate as g
GOOD = json.dumps({"h1":"家居藏酒空間的規劃思路","sections":[
    {"h2":None,"body":"近年香港家居收藏文化"+"興"*175+"。"},
    {"h2":"空間規劃","body":"談到{{KW1}}的擺放位置"+"與"*175+"。"},
    {"h2":"環境條件","body":"濕度與震動同樣重要"+"而"*175+"。"},
    {"h2":"藏品管理","body":"至於{{KW2}}的紀錄方式"+"則"*175+"。"},
    {"h2":"長線維護","body":"定期檢視可以保障價值"+"並"*175+"。"}]}, ensure_ascii=False)
SHORT = json.dumps({"h1":"太短的標題","sections":[
    {"h2":None,"body":"短"*30},{"h2":"a","body":"短"*30},{"h2":"b","body":"短"*30},
    {"h2":"c","body":"短"*30},{"h2":"d","body":"短"*30}]}, ensure_ascii=False)

def fake_chat(api_key, model, prompt, **kw):
    return GOOD if "紅酒櫃" in prompt or "酒櫃推介" in prompt else SHORT
g._chat = fake_chat

import time; time.sleep = lambda *a: None

exec(compile(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
     "linkbuild-generator", "app.py"), encoding="utf-8").read(), "app.py", "exec"))
