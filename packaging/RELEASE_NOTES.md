把 Microsoft [MarkItDown](https://github.com/microsoft/markitdown) 打包成 Windows 執行檔和 Claude skill，可以把 PDF、Word、PowerPoint、Excel、HTML、CSV、JSON、XML、EPUB、ZIP、Outlook .msg、Jupyter Notebook 和網頁（包含 YouTube、Wikipedia）轉成 Markdown。

## 這個版本的更新

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
