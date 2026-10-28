---
name: file-toolkit
description: |
  老師的教學檔案批次處理工具包（本機 Windows／Python）。當使用者要對 Word、Excel、PowerPoint、PDF、圖片做「一次很多份」或「格式轉換」的雜事時，請一定要使用此技能。

  涵蓋能力：
  1. Word：名單套印獎狀／通知單／個人成績單、題目卷與答案卷兩版產出、Word 批次轉 PDF 或 JPG
  2. Excel：成績計算與排名、不及格標紅、依班級拆檔、隨機座位表與分組、答對率與分數分布圖
  3. PowerPoint：教材大綱轉整份簡報、圖片資料夾轉圖卡簡報、全簡報統一字型與加校徽
  4. PDF：合併／拆分／抽頁／重排、加浮水印、掃描檔繁中 OCR 轉可搜尋 PDF、指定頁轉圖並去白邊
  5. 其他：連結轉 QR Code、PDF／Word／PPT／Excel 轉 Markdown 餵 AI、圖片批次縮放與轉檔

  觸發語句包含：「把這些 PDF 合併」「Word 全部轉成 PDF」「這份掃描 PDF 做 OCR」「用名單套印獎狀」「算總分和排名、不及格標紅」「把成績表依班級拆開」「排一張座位表」「把這資料夾的圖片做成簡報」「這份 PPT 字型全部統一」「PDF 加浮水印」「課本第 12 頁轉成圖片」「幫我生 QR Code」「把這份 PDF 轉成 Markdown」，或任何「一個資料夾的檔案要批次處理」的請求。

  本技能負責在本機建立並使用固定版本的 Python 環境實際執行檔案處理；若使用者只是要從零撰寫一份新文件的內容（而非批次處理既有檔案），改用 docx／xlsx／pptx 等文件撰寫技能。
---

# 教學檔案處理工具包（file-toolkit）

判斷任務 → 準備環境 → 讀對應 recipe → 寫一次性腳本執行 → 驗收回報

處理的是老師手上的真實教材與學生資料，**首要原則是不動到原始檔**。

---

## 0. 先確認任務與輸入

開工前一定要弄清楚三件事，缺哪個就問哪個（不要自己猜路徑）：

| 要確認 | 說明 |
|--------|------|
| 輸入在哪 | 單一檔案還是整個資料夾？資料夾要不要含子目錄？ |
| 要什麼結果 | 檔案格式、份數、命名規則（例：`302_15_王小明.pdf`） |
| 有無模板 | 套印、簡報、考卷這類任務通常有既有 Word／PPT 模板要沿用 |

輸入若是 Excel 名單或成績表，**先讀前 5 列印出欄位名**再寫處理邏輯，不要假設欄位叫「姓名」「班級」。

---

## 1. 準備 Python 環境（每次都要先做）

本技能的所有處理都跑在固定版本的核心環境上（12 個核心套件）。執行：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File "<本技能目錄>/scripts/ensure_env.ps1"
```

一定要用 PowerShell 7 的 `pwsh`，**不要用 `powershell.exe`**（5.1）：5.1 會把腳本裡的中文讀成亂碼，直接解析失敗。

Agent 的權限機制若擋下 `-ExecutionPolicy Bypass`，不要想辦法繞過，改在 pwsh 裡直接執行 `& "<本技能目錄>/scripts/ensure_env.ps1"`；這樣也因執行原則失敗時，請使用者自己執行上面那行。

**環境還沒好之前，不要改用別的工具硬做**：要讀 PDF 時，不要用系統 Python 臨時 `pip install`，也不要用 Git Bash 內建的 `pdftotext`（它缺中文字碼對照表 `Adobe-CNS1`，中文會全部抽成空白）。先把本節的環境建好，再照 recipe 做。

腳本會依序尋找可用環境，找不到就用 `uv` 自動建立共用環境並安裝核心套件，**最後一行印出 Python 直譯器完整路徑**。後續所有腳本都用這個直譯器執行，不要用系統 `python`：

```powershell
& "<上一步印出的路徑>" "./_work/do_task.py"
```

尋找順序：`FILE_TOOLKIT_PYTHON` 環境變數 → 目前專案的 `.\.venv` → 共用環境 `%LOCALAPPDATA%\file-toolkit\.venv`。

新建時架構由腳本決定：Windows on ARM 先試原生 ARM64，被 Smart App Control 擋或裝不起來會自動改建 x64。**照它印出的路徑用就好，不要自己另建 ARM64 或 x64 環境**。使用者要求「重試 ARM64」時，加 `-RetryNative` 再跑一次：沒通過會保留原本的 x64 環境，不要叫使用者手動刪環境。

**OCR 任務（掃描 PDF 轉可搜尋文字）額外執行一次**，它會裝好 Tesseract 與繁體中文模型：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File "<本技能目錄>/scripts/ensure_ocr.ps1"
```

---

## 2. 讀對應的 recipe

依任務類別讀取（**只讀需要的那一份**，不要一次全讀）：

| 任務 | 檔案 |
|------|------|
| 套印、考卷兩版、Word 轉 PDF／JPG | `references/word.md` |
| 成績計算、排名標紅、拆檔、座位表、統計圖 | `references/excel.md` |
| 大綱轉簡報、圖卡簡報、統一字型加校徽 | `references/pptx.md` |
| 合併拆分、浮水印、OCR、頁面轉圖 | `references/pdf.md` |
| QR Code、轉 Markdown、圖片批次處理 | `references/misc.md` |

recipe 裡的寫法都是實測過的版本，**遇到同名任務直接沿用，不要自己另發明**（例如套印一定要用 recipe 的跨 run 取代寫法，直接 `run.text = ...` 會掉格式）。

---

## 3. 寫腳本 → 執行 → 驗收

1. 腳本寫成獨立 `.py` 檔（放 `_work\` 或系統暫存區），不要用 `python -c` 塞多行中文
2. 檔頭固定加 `# -*- coding: utf-8 -*-`；Windows 中文路徑用 `pathlib.Path`
3. 執行後**一定要驗收**再回報：檢查輸出檔數量、抽一份檢查內容（PDF 用 pypdfium2 抽文字、Excel 用 openpyxl 讀回關鍵儲存格、圖片檢查尺寸）
4. 回報時列出：輸出位置、產出份數、跳過或失敗的項目及原因

處理數量大時（> 20 份），先只跑前 2 份給使用者確認格式，再跑全部。

---

## 安全規則

- **絕不覆蓋原始檔**。輸出一律寫到輸入資料夾旁的 `output\`（或使用者指定的資料夾）；同名時加 `_1`、`_2`，不要靜默覆蓋
- 刪除、移動使用者既有檔案前一定要先問
- 學生資料只用班級代碼與座號命名；需要真實姓名時只寫在檔案內容，不外傳、不上雲
- 一切在本機處理，不要把教材或名單傳到任何外部服務
- 失敗的檔案要列出來，不要吞掉錯誤假裝全部成功

---

## 選裝套件

以下三個套件**不必另外安裝**，核心套件會順便裝進共用環境，可以直接 `import`：

| 套件 | 由誰帶進來 | 適合用在 |
|------|-----------|---------|
| `pandas` | `markitdown[xlsx]` | 大量資料的樞紐分析；一般成績計算仍以 `openpyxl` 為主（recipe 都是 openpyxl 寫法） |
| `pdfplumber` | `markitdown[pdf]` | 精準擷取 PDF 表格 |
| `xlsxwriter` | `python-pptx` | 只寫不讀的大型 Excel 輸出 |

其他套件（`docxcompose`、`yt-dlp`、`youtube-transcript-api`）**預設不裝**。真的需要時先跟使用者說明用途，同意後單獨安裝一項：

```powershell
uv pip install --python "<Python 路徑>" docxcompose
```

**不要裝 `PyMuPDF`**：已於 2026-10 移出核心，抽文字與轉圖一律用 `pypdfium2`（理由見 `references/pdf.md`）。

---

## 疑難排解

| 症狀 | 原因與處理 |
|------|-----------|
| `docx2pdf` 失敗或卡住 | 需要桌面版 Microsoft Word，且不支援 `python -m docx2pdf`；確認 Word 沒有開著跳對話框 |
| OCR 找不到 `chi_tra` | 跑 `scripts\ensure_ocr.ps1`；模型不能放 `Program Files`，要在 `%LOCALAPPDATA%\Tesseract-OCR\tessdata` |
| 圖表中文變空格 | matplotlib 未設中文字型，見 `references/misc.md` |
| 裝完系統工具仍找不到指令 | PATH 尚未更新，需重開終端機或重啟 Agent |
| PDF 文字抽不出來 | 是掃描圖片型 PDF，要先做 OCR 才有文字層 |
