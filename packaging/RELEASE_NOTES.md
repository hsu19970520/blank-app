把 Microsoft [MarkItDown](https://github.com/microsoft/markitdown) 打包成 Windows 執行檔和 Claude skill，可以把 PDF、Word、PowerPoint、Excel、HTML、CSV、JSON、XML、EPUB、ZIP、Outlook .msg、Jupyter Notebook 和網頁（包含 YouTube、Wikipedia）轉成 Markdown。

## 這個版本的更新

- **Excel 公式**：每張工作表最後會列出所有公式的寫法和使用範圍。同一條公式往下或往右複製只算一種寫法，例如 61 萬個公式整理成 195 種。開頭新增「工作表之間的資料流向」，列出每張工作表的公式和樞紐分析表讀取了哪些工作表。選項 `--xlsx-formulas cells`（圖形介面：「選項 › Excel：在每個儲存格的值下方附上公式」）會在每個值下方附上該格的公式。
- **用 Excel 重新計算**：`--xlsx-recalc`（圖形介面：「選項 › Excel：轉換前先用 Excel 重新計算」）會呼叫你電腦上的 Microsoft Excel，在檔案副本中依序重新計算、重新整理所有樞紐分析表、再算一次，然後才轉換，原始檔不會被修改。需要 Windows 與 Excel 2021 或 Microsoft 365（檔案用到 XLOOKUP）；如果重新計算後出現新的 #NAME? 錯誤，結果開頭會警告。

### 1.2.1

- 修正：資料驗證的下拉清單如果引用其他工作表（Excel 會用 x14 擴充格式儲存），之前不會出現在輸出中，現在會列出套用範圍和選項來源。

### 1.2.0

- Excel 轉換改寫：用串流方式讀取（大型活頁簿也能在幾十秒內轉完）。會自動找出真正的表頭列（例如第 14 列，並把週別標籤和日期合併成欄名），數字依 Excel 格式顯示、不再出現科學記號。註解、自動篩選條件、合併儲存格、資料驗證、樞紐分析表結構，以及隱藏的工作表、列和欄，都會清楚標示。新增選項：`--xlsx-header-row`、`--xlsx-visible-only`、`--xlsx-raw-values`、`--xlsx-classic`；圖形介面的「選項」也可以設定只輸出看得到的內容。

### 1.1.0

- 重新設計圖形介面：採用 Windows 11 風格並支援深色模式。檔案清單直接顯示每個項目的狀態，右邊可以預覽 Markdown。轉換途中可以停止，失敗時會說明原因和解決方法。另外加入了選單列、快速鍵和右鍵選單。
- 新的 App 圖示。
- 預設關閉 ONNX Runtime 傳送給 Microsoft 的使用統計。

## 下載哪一個？

| 檔案 | 用途 |
|---|---|
| `markitdown.exe` | **Windows 單一執行檔**，不需安裝 Python。雙擊開啟圖形介面，或在命令列使用。 |
| `markitdown-skill-windows.zip` | Claude skill，內附 `markitdown.exe`（Windows 上的 Claude Code 用這個）。 |
| `markitdown-skill.zip` | Claude skill，僅含腳本（macOS / Linux，或上傳到 Claude App；需要 `pip install "markitdown[all]"`）。 |
| `SHA256SUMS.txt` | 檔案雜湊值，可用來驗證下載內容。 |

## markitdown.exe 用法

- **雙擊**：開啟圖形介面。把檔案或資料夾拖進視窗，按「轉換」。左邊清單會顯示每個項目的狀態，右邊可以預覽、複製 Markdown。外觀會跟著 Windows 的淺色／深色設定切換。
- **拖曳**：把檔案或資料夾拖到 `markitdown.exe` 圖示上，會在原位置產生 `.md`。
- **命令列**：
  ```
  markitdown.exe report.pdf                 # 輸出到螢幕
  markitdown.exe report.pdf -o report.md    # 存成檔案
  markitdown.exe C:\docs -r --out-dir C:\docs_md
  markitdown.exe --help
  ```

> 這個執行檔沒有數位簽章。第一次執行時，Windows SmartScreen 可能會跳出「Windows 已保護您的電腦」，請按「其他資訊」再按「仍要執行」。每次啟動需要幾秒鐘解壓縮。
