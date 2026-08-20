# linkbuild-generator

DZ Linkbuild Article Generator —— 由 keyword Excel 生成合規嘅外連文章（.docx / Google Doc）。

## 合規 gate

每篇文章喺回稿之前都要通過 `validate_content()`，規則對齊
`.claude/agents/dz-linkbuild/scripts/validate.py`：

- 結構：1 個 H1 + 剛好 4 個 H2，冇 H3
- H1：中文 < 30 全形字 / 英文 < 60 字元，完全冇標點，唔含關鍵字
- 關鍵字：每個喺正文剛好出現一次，只可以係 hyperlink；唔准入標題、開篇、結尾
- Buffer：兩個關鍵字之間要隔至少一整段
- 字數：正文 750–1,000（不含標題），生成目標 820–960
- 語言：中文稿唔准有簡體字、廣東話口語虛詞；英文稿唔准夾中文字；兩者都唔准破折號

過唔到會自動叫 model 修正（最多 2 次），仍然唔過就重新生成（最多 3 次）。
再唔得就保留最接近嘅草稿，**明確標示做未合規**並且獨立一區顯示／下載，
唔會混入交付稿 —— 唔會靜靜地交唔合規嘅嘢，亦唔會白白掉咗成篇。

## Model 選擇

幾乎所有現代 model 都係 reasoning model，推理 token 同 completion 共用 `max_tokens`。
本工具預設送 `reasoning={"enabled": false}` 並用較大嘅 `max_tokens`；provider 唔收
就自動除返個參數重試。所以揀 model 唔使避開 reasoning model。

## 輸出格式

- 標題寫成 `[H1: 標題文字]` / `[H2: 小標題]`，H1 置中
- 關鍵字係真 hyperlink，藍色 `0563C1` + 單底線 + cyan highlight
- 字體 Arial / Microsoft JhengHei
- Excel 冇 target URL 會用 `https://example.com` placeholder，並且喺 UI 同 CLI 標示出嚟

## 用法

Web UI：

```bash
streamlit run app.py
```

CLI：

```bash
export OPENROUTER_API_KEY=sk-...
python generate.py --excel batch.xlsx --batch 1 --docx out.docx
python generate.py --excel batch.xlsx --batch 1 --single-docx-dir ./out   # 每篇一個檔
python generate.py --excel batch.xlsx --batch 1 --creds sa.json --folder FOLDER_ID
```

`--sheet` 唔指定就會用第一個 sheet。

## 最終覆核

合併稿入面嘅 `#N` 同 `● keyword` 頭段會被 `validate.py` 當成正文，令關鍵字次數多咗。
要跑 validator 請用單篇檔（`--single-docx-dir`，或者 Web UI 嗰個 `_single.zip`）：

```bash
python3 .claude/agents/dz-linkbuild/scripts/validate.py out/001_關鍵字.docx
```

## 環境變數

| 變數 | 預設 | 說明 |
|---|---|---|
| `OPENROUTER_API_KEY` | — | 必填 |
| `LB_MODEL` | `~deepseek/deepseek-v4-flash-latest` | OpenRouter model id（UI 有 picker 可即場切換） |
| `LB_MAX_TOKENS` | `16000` | 單次 completion 上限。Reasoning model 嘅推理 token 同 completion 共用呢個 budget，太細會回空 content |
| `LB_DISABLE_REASONING` | `1` | 預設送 `reasoning={"enabled": false}`。設 `0` 就唔送 |
| `LB_DELAY` | `3` | CLI 每篇之間嘅間隔（秒） |
