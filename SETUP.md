# 設定與部署說明

採購・資產整合平台（Streamlit + Google Sheets）。資料存在一份 Google Sheet，
所有同工用同一網址 + 共用密碼存取同一份資料。

檔案：
- `app.py`：主程式
- `requirements.txt`：套件
- `secrets.toml.example`：密鑰範本

---

## 一、建立 Google Sheet（資料庫）

1. 到 Google 雲端硬碟新建一份試算表，命名隨意（例如「採購資產資料庫」）。
2. 複製它的網址，等一下要貼進 secrets。工作表內容不用自己建，程式會自動建立。

## 二、建立服務帳戶（讓程式能讀寫試算表）

1. 進 https://console.cloud.google.com → 建立一個專案。
2. 「API 和服務 → 程式庫」搜尋並啟用兩個 API：**Google Sheets API** 與 **Google Drive API**。
3. 「API 和服務 → 憑證 → 建立憑證 → 服務帳戶」，建立後進入該帳戶 → 「金鑰 → 新增金鑰 → JSON」，下載一個 JSON 檔。
4. 打開那份 JSON，裡面的 `client_email`（長得像 `xxx@xxx.iam.gserviceaccount.com`）就是服務帳戶信箱。
5. 回到第一步的 Google Sheet，按右上「共用」，把**服務帳戶信箱**加為**編輯者**。
   （這一步最常被忘記，忘了會連不上。）

## 三、填寫 secrets

把 `secrets.toml.example` 複製成 `.streamlit/secrets.toml`，依下載的 JSON 對應填入：
- `app_password`：你們的共用密碼
- `[sheet] url`：第一步的試算表網址
- `[gcp_service_account]`：把 JSON 裡每個欄位對應貼上。
  `private_key` 要整段貼，並保留裡面的 `\n`（換行符號）。

**`.streamlit/secrets.toml` 不要上傳到 GitHub。** 在專案根目錄建一個 `.gitignore`，加入：
```
.streamlit/secrets.toml
```

## 四、本機測試

```bash
pip install -r requirements.txt
streamlit run app.py
```
打開後輸入密碼登入 → 左側「⚙️ 首次設定」→ 按「初始化試算表（含範例）」，
程式會在你的 Google Sheet 裡自動建立 suppliers / purchase_orders / po_items / assets
四個工作表並填入範例資料。之後回試算表就能看到資料同步進去。

## 五、部署到 Streamlit Community Cloud

1. 把 `app.py`、`requirements.txt` push 到一個 GitHub repo（**不要** push secrets.toml）。
2. 到 https://share.streamlit.io 用 GitHub 登入 → New app → 選你的 repo、分支、`app.py`。
3. 部署頁的 **Advanced settings → Secrets**，把 `.streamlit/secrets.toml` 的內容整段貼進去。
4. 部署完成會得到一個 `https://你的app.streamlit.app` 網址。
5. 把網址 + 共用密碼給同工即可。

---

## 常見問題

- **連不上 / 權限錯誤**：99% 是忘了把服務帳戶信箱加進試算表共用（步驟二之 5），
  或 `private_key` 的換行貼壞了。
- **App 休眠**：太久沒人用會休眠，下次有人開啟需等它喚醒，醒來後正常，資料不會掉。
- **改了程式**：push 到 GitHub，Streamlit Cloud 會自動重新部署。
- **多人同時改**：採最後寫入為準。內部小團隊使用通常足夠；若需要嚴格防衝突，
  之後可改用資料庫（例如 Supabase / PostgreSQL）取代 Google Sheets。

## 之後可以再加的東西

- 固定資產折舊（取得日、耐用年限、年折舊、帳面淨值）
- 採購請購人 / 簽核流程
- 報廢 / 移轉的歷史紀錄
- 匯出報表（Excel / PDF）
