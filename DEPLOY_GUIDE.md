# 部署指南：從零開始到上線

這份指南假設你完全沒有部署經驗。跟住做，大約 30-40 分鐘可以完成。

完成後你會得到一個固定網址，任何人打開就能用。

---

## 全程總覽（4 個大步驟）

```
Step 1  建立 Google Service Account（讓程式能建立 Google Doc）
Step 2  建立 GitHub Repo（存放程式碼）
Step 3  部署到 Streamlit Cloud（變成一個網站）
Step 4  設定 Secrets（把 API Key 等認證放到伺服器）
```

---

## Step 1：建立 Google Service Account

這一步讓程式有權限自動建立 Google Doc。

### 1.1 建立 Google Cloud Project

1. 打開 https://console.cloud.google.com/
2. 用你的 Google 帳號登入
3. 畫面頂部有一個 project 選擇器（可能顯示「Select a project」）→ 點擊它
4. 點右上角的「NEW PROJECT」
5. Project name 輸入：`linkbuild-generator`
6. 點「CREATE」
7. 等幾秒鐘，建立完成後確認頂部已切換到這個 project

### 1.2 啟用 Google Docs API

1. 在左側選單點「APIs & Services」→「Library」
2. 搜尋欄輸入：`Google Docs API`
3. 點擊搜尋結果中的「Google Docs API」
4. 點藍色的「ENABLE」按鈕
5. 等它啟用完成

### 1.3 啟用 Google Drive API

1. 再次回到 Library（左側選單 → APIs & Services → Library）
2. 搜尋欄輸入：`Google Drive API`
3. 點擊搜尋結果中的「Google Drive API」
4. 點藍色的「ENABLE」按鈕

### 1.4 建立 Service Account

1. 在左側選單點「APIs & Services」→「Credentials」
2. 點頂部的「+ CREATE CREDENTIALS」
3. 選「Service account」
4. Service account name 輸入：`linkbuild-bot`
5. 點「CREATE AND CONTINUE」
6. Role 欄位選：「Editor」
7. 點「CONTINUE」→ 點「DONE」

### 1.5 下載 JSON 金鑰

1. 你現在在 Credentials 頁面，會看到 Service Accounts 區域有剛建立的帳號
2. 點擊「linkbuild-bot@...」這個 email 連結
3. 點上方的「KEYS」tab
4. 點「ADD KEY」→「Create new key」
5. 選「JSON」→ 點「CREATE」
6. 瀏覽器會自動下載一個 `.json` 檔案
7. **用記事本打開這個 JSON 檔案，暫時不要關閉，Step 4 會用到裡面的內容**

### 1.6 分享 Google Drive 資料夾

1. 打開 Google Drive（drive.google.com）
2. 建立一個新資料夾（或選擇現有的），例如叫「Linkbuild Output」
3. 右鍵這個資料夾 →「Share」→「Share」
4. 在「Add people」欄位貼入 Service Account 的 email
   - 這個 email 在你剛下載的 JSON 檔案裡的 `client_email` 欄位
   - 格式類似：`linkbuild-bot@linkbuild-generator.iam.gserviceaccount.com`
5. 權限選「Editor」
6. 取消勾選「Notify people」
7. 點「Share」

8. **記下資料夾 ID**：
   - 打開這個資料夾
   - 看瀏覽器網址列，格式是 `https://drive.google.com/drive/folders/XXXXXXXXXXXXX`
   - `folders/` 後面那串就是資料夾 ID
   - 複製它，Step 4 會用到

---

## Step 2：建立 GitHub Repo

GitHub 是存放程式碼的地方。Streamlit Cloud 會從這裡讀取你的程式碼。

### 2.1 建立帳號（如已有可跳過）

1. 打開 https://github.com
2. 註冊帳號（或直接登入）

### 2.2 建立新 Repo

1. 登入後，點右上角的「+」→「New repository」
2. 填寫以下資料：
   - Repository name：`linkbuild-generator`
   - Description（可選）：`DZ Linkbuild Article Generator`
   - **選擇「Private」**（重要！不要選 Public）
   - 勾選「Add a README file」
3. 點「Create repository」

### 2.3 上傳程式碼

1. 在 repo 頁面，點「Add file」→「Upload files」
2. 將以下 4 個檔案拖拉進去：
   - `app.py`
   - `generate.py`
   - `requirements.txt`
   - `DEPLOY_GUIDE.md`（可選）
3. 在下方 Commit message 輸入：`initial upload`
4. 點「Commit changes」
5. 等幾秒鐘上傳完成

完成後你的 repo 頁面應該能看到這些檔案。

---

## Step 3：部署到 Streamlit Cloud

### 3.1 連接 Streamlit Cloud

1. 打開 https://share.streamlit.io
2. 點「Continue to sign-in」→ 用 GitHub 帳號登入
3. 授權 Streamlit 存取你的 GitHub

### 3.2 建立 App

1. 登入後點「Create app」(或「New app」)
2. 選擇：
   - Repository：選 `你的帳號/linkbuild-generator`
   - Branch：`main`
   - Main file path：`app.py`
3. 點「Deploy!」

### 3.3 等待部署

- Streamlit 會開始安裝套件和啟動應用
- 第一次部署需要 2-3 分鐘
- 你會看到 build log 在跑
- **這時候畫面會顯示錯誤「伺服器設定未完成」— 這是正常的，因為你還沒設定 Secrets**
- 先不用理會，繼續 Step 4

部署完成後你會得到一個固定網址，格式如：
```
https://你的帳號-linkbuild-generator-app-xxxxx.streamlit.app
```

---

## Step 4：設定 Secrets

這是最後一步。把所有認證資料放到 Streamlit 伺服器，用戶永遠看不到這些。

### 4.1 打開 Secrets 設定頁

1. 在 Streamlit Cloud 的 dashboard（https://share.streamlit.io）
2. 找到你剛部署的 app
3. 點右邊的「⋮」（三個點）→「Settings」
4. 點左邊的「Secrets」tab

### 4.2 貼入以下內容

在 Secrets 文字框中，貼入以下內容（替換 `<...>` 的部分）：

```toml
OPENROUTER_API_KEY = "<你的 OpenRouter API Key>"
GOOGLE_FOLDER_ID = "<Step 1.6 記下的資料夾 ID>"
SHARE_EMAIL = "growthwithdz@digitalzoo.com.hk"
LB_MODEL = "deepseek/deepseek-v4-0324"

[GOOGLE_CREDENTIALS]
type = "service_account"
project_id = "<從 JSON 檔案複製>"
private_key_id = "<從 JSON 檔案複製>"
private_key = "<從 JSON 檔案複製，整段包括 BEGIN 和 END>"
client_email = "<從 JSON 檔案複製>"
client_id = "<從 JSON 檔案複製>"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "<從 JSON 檔案複製>"
universe_domain = "googleapis.com"
```

### 4.3 如何從 JSON 檔案取得這些值

用記事本打開你在 Step 1.5 下載的 JSON 檔案，內容大概是：

```json
{
  "type": "service_account",
  "project_id": "linkbuild-generator",
  "private_key_id": "abc123def456...",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQ...(很長一串)...\n-----END PRIVATE KEY-----\n",
  "client_email": "linkbuild-bot@linkbuild-generator.iam.gserviceaccount.com",
  "client_id": "123456789",
  ...
}
```

逐個欄位複製值到 Secrets 的對應位置。注意：

- **private_key**：這是最長的一段，從 `-----BEGIN PRIVATE KEY-----` 一直到 `-----END PRIVATE KEY-----\n`，全部複製，包括中間的 `\n`
- 每個值都要用引號 `"..."` 包住
- 不要漏掉任何一個欄位

### 4.4 儲存並重啟

1. 貼好後點「Save」
2. 回到 app 頁面，點右上角的「⋮」→「Reboot app」
3. 等 30 秒左右 app 重新啟動
4. 刷新頁面 — 如果看到上傳 Excel 的畫面而不是錯誤訊息，就代表設定成功！

---

## 完成！🎉

現在你可以：

1. 把網址發給同事
2. 他們打開網頁、上傳 Excel、選 Batch、按按鈕
3. 等 15-20 分鐘
4. Google Doc 自動出現在你指定的 Drive 資料夾

---

## 常見問題

### 畫面顯示「伺服器設定未完成」
Secrets 設定有誤。回到 Streamlit Cloud → Settings → Secrets 檢查格式。

### Google Doc 建立失敗
1. 確認 Google Docs API 和 Google Drive API 都已啟用
2. 確認 Service Account email 已被加為 Drive 資料夾的 Editor
3. 確認 Secrets 中的 private_key 完整複製（包括 BEGIN/END 標記）

### 文章品質不夠好
在 Secrets 中把 `LB_MODEL` 改為 `anthropic/claude-sonnet-4-6`（品質更高但成本較高）。

### App 變慢或停止回應
Streamlit Cloud 免費版有資源限制。如果 34 篇一次跑完太久，可以分成兩次：
第一次跑 #1-#17，第二次跑 #18-#34。

### 想更新程式碼
在 GitHub repo 頁面，點擊要更新的檔案 → 鉛筆圖示 → 修改 → Commit。
Streamlit Cloud 會自動偵測到改動並重新部署。
