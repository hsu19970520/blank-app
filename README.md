# MarkItDown Skill + Windows 執行檔

這個專案把 Microsoft 的 [MarkItDown](https://github.com/microsoft/markitdown) 文件轉 Markdown 功能打包成兩種形式：

1. **Claude skill**（`skills/markitdown/`）：讓 Claude Code 或 Claude App 能直接幫你讀取、轉換各種文件。
2. **Windows 單一執行檔 `markitdown.exe`**：不需安裝 Python，雙擊就能用圖形介面轉檔，也可以在命令列使用。

支援格式：PDF、Word (.docx)、PowerPoint (.pptx)、Excel (.xlsx/.xls)、HTML、CSV、JSON、XML/RSS、EPUB、ZIP、Outlook .msg、Jupyter Notebook、圖片（EXIF）、音訊（語音轉文字），以及網址（一般網頁、YouTube 字幕、Wikipedia）。

## 下載

到本 repo 的 **[Releases](../../releases/latest)** 頁面下載：

| 檔案 | 用途 |
|---|---|
| `markitdown.exe` | Windows 單一執行檔 |
| `markitdown-skill-windows.zip` | Claude skill，內附 `markitdown.exe`（Windows 上的 Claude Code 用這個） |
| `markitdown-skill.zip` | Claude skill，只有腳本（macOS/Linux，或上傳到 Claude App） |

每次推送都會由 GitHub Actions 自動建置，也可以到 **Actions → Build Windows exe → 最新一次執行 → Artifacts → `markitdown-windows`** 下載。

## markitdown.exe 用法

**圖形介面**：雙擊 `markitdown.exe`，把檔案或資料夾拖進視窗（或按「加入檔案…」、貼上網址），再按右上角的「轉換 N 個項目」。

- 左邊清單會直接顯示每個項目的狀態：✓ 已完成、✕ 失敗、⚠ 沒有文字。選取項目後，右邊會預覽 Markdown，可以「複製」、「開啟」或「開啟檔案位置」。
- 轉換失敗時，預覽區會用白話說明原因和解決方法。轉換途中可以按「停止」。
- 外觀跟著 Windows 的淺色／深色設定自動切換。
- 常用快速鍵：`Ctrl+O` 加入檔案、`Ctrl+Shift+O` 加入資料夾、`Ctrl+L` 貼上網址、`Ctrl+Enter` 轉換、`Esc` 停止、`Ctrl+Shift+C` 複製 Markdown、`Delete` 移除。也可以在項目上按右鍵操作。

**拖曳轉檔**：把檔案或資料夾拖到 `markitdown.exe` 圖示上，會在原檔旁產生同名的 `.md`。

**命令列**：

```bat
markitdown.exe report.pdf                        :: Markdown 輸出到螢幕
markitdown.exe report.pdf -o report.md           :: 存成檔案
markitdown.exe a.docx b.pptx c.xlsx              :: 每個檔案旁各產生一個 .md
markitdown.exe C:\docs -r --out-dir C:\docs_md   :: 整個資料夾（含子資料夾）
markitdown.exe https://www.youtube.com/watch?v=ID
markitdown.exe --help                            :: 全部選項
```

> **注意**
> - 執行檔沒有數位簽章，第一次執行時 Windows SmartScreen 可能會阻擋。請按「其他資訊」再按「仍要執行」。
> - 單一執行檔每次啟動都要先解壓縮，大約需要幾秒鐘。
> - 掃描版 PDF（只有圖片、沒有文字層）轉不出文字，需要搭配 OCR 或 Azure Document Intelligence（`--docintel-endpoint`）。
> - Excel（.xlsx/.xlsm）使用本專案改寫的轉換器：自動找出真正的表頭列（依自動篩選、樞紐分析表或凍結窗格），數字依 Excel 格式顯示（不會變成科學記號）；註解、自動篩選、合併儲存格、資料驗證、樞紐分析表和隱藏的工作表／列／欄都會標示出來。公式只保留上次存檔時的計算結果，條件式格式（紅黃綠燈號）目前還不會顯示。相關選項：`--xlsx-header-row "工作表=列"`、`--xlsx-visible-only`、`--xlsx-raw-values`、`--xlsx-classic`。
> - MarkItDown 用來判斷檔案類型的 ONNX Runtime 預設會傳送使用統計給 Microsoft。本工具預設關閉這項功能；如果想開啟，請設定環境變數 `ORT_DISABLE_TELEMETRY=0`。

## 安裝 Claude skill

**Claude Code（Windows）**：把 `markitdown-skill-windows.zip` 解壓縮到 `%USERPROFILE%\.claude\skills\`，結果應該是：

```
%USERPROFILE%\.claude\skills\markitdown\SKILL.md
%USERPROFILE%\.claude\skills\markitdown\scripts\convert.py
%USERPROFILE%\.claude\skills\markitdown\bin\markitdown.exe
```

**Claude Code（macOS / Linux）**：把 `markitdown-skill.zip` 解壓縮到 `~/.claude/skills/`，再執行 `pip install "markitdown[all]"`。

**Claude App（claude.ai / 桌面版）**：在設定的 Skills 區塊上傳 `markitdown-skill.zip`。Claude 會在雲端環境用 Python 執行，需要開啟程式碼執行功能，並允許安裝 `markitdown[all]`。

裝好之後，直接跟 Claude 說「幫我把這份 PDF 轉成 Markdown」或「讀一下這個 Excel」，Claude 就會使用這個 skill。

## 自己建置

**在 Windows 本機**：安裝 Python 3.10 以上版本，然後執行：

```bat
packaging\build_windows.bat
```

產出的檔案在 `dist\`。

**用 GitHub Actions**：推送到 `main` 或 `claude/**` 分支，會自動建置並上傳 Artifact。要發佈 Release，可以推送 `v*` 標籤（例如 `v1.0.0`），或到 Actions → Build Windows exe → Run workflow，填入 `release_tag` 後執行。

**只用 Python（任何作業系統）**：

```bash
pip install "markitdown[all]"
python skills/markitdown/scripts/convert.py report.pdf
python app/markitdown_app.py            # 圖形介面
```

## 專案結構

```
skills/markitdown/SKILL.md            Claude skill 說明（Claude 讀這個檔案）
skills/markitdown/scripts/convert.py  轉換核心：CLI 與批次轉換（skill 和 exe 共用）
skills/markitdown/scripts/xlsx_converter.py  Excel 轉換器（保留表頭、格式與結構資訊）
app/markitdown_app.py                 exe 進入點：圖形介面 / 拖曳 / 命令列
app/gui.py                            圖形介面（Windows 11 風格、淺色／深色）
app/assets/                           App 圖示（由 packaging/make_icon.py 產生）
packaging/markitdown.spec             PyInstaller 設定
packaging/requirements-build.txt      建置用相依套件（固定版本）
packaging/build_windows.bat           Windows 本機建置腳本
packaging/package_skill.py            打包 skill zip
tests/smoke_test.py                   各格式端對端測試（可測 .py 或 .exe）
tests/gui_smoke_test.py               圖形介面測試
tests/xlsx_test.py                    Excel 轉換測試（表頭、數字格式、隱藏、註解、樞紐）
.github/workflows/build-windows-exe.yml  Windows 自動建置、測試與發佈
```

## 授權

MarkItDown 由 Microsoft 以 MIT 授權釋出，`markitdown.exe` 內含 MarkItDown 及其相依套件，這些套件各自依原授權條款使用。本 repo 其餘內容依 `LICENSE`（Apache 2.0）授權。

<sub>這個 repo 原本是 Streamlit 範本，`streamlit_app.py` 仍保留，執行方式是 `pip install -r requirements.txt && streamlit run streamlit_app.py`。</sub>
