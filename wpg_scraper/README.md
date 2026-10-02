# WPG Holdings 據點爬蟲

抓取 <https://www.wpgholdings.com/branches/search/en/0_0_0/0/> 上所有公司據點，輸出 Excel 資料庫。

| Company | Region/City | Zip Code | Tel | Fax | Email | Address |
|---|---|---|---|---|---|---|

## 方法一：在本機執行

```bash
cd wpg_scraper
pip install -r requirements.txt
python scraper.py                 # 產生 wpg_branches.xlsx
```

常用參數：

| 參數 | 說明 |
|---|---|
| `-o out.xlsx` | 輸出檔名 |
| `--delay 1.0` | 每頁間隔秒數（避免對網站造成負擔） |
| `--max-pages 300` | 最多抓取頁數 |
| `--browser` | 用 Playwright 渲染（資料由 JavaScript 載入時；需 `pip install playwright && playwright install chromium`） |
| `--html a.html b.html` | 離線解析已另存的網頁 |

## 方法二：用 GitHub Actions 一鍵產生

GitHub → **Actions** → **WPG branches scraper** → **Run workflow**，
跑完後在該次執行頁面下方 **Artifacts** 下載 `wpg_branches.xlsx`。

## 運作方式

1. 從起始頁開始，跟隨 `/branches/search/en/` 底下的所有分頁 / 篩選連結（BFS，自動去重）。
2. 每頁找出含有 `Tel` 標籤的據點區塊，依 `Zip Code / Tel / Fax / Email / Address` 標籤擷取值；
   公司名稱取區塊標題，Region/City 取區塊內標籤前文字或最近的區域標題。
3. 以「公司 + 地址 + 電話」去重後寫入 Excel（欄位皆存為文字，保留前導 0）。
