# 🤖 Agent 安裝指南：教學檔案處理「核心工具包」

> **這份檔案是寫給 AI Agent 讀的。**
> 適用於 Windows 上的 Claude Code、ChatGPT 的 Codex App、Google Antigravity、OpenCode。
> 預設只安裝已篩選過的核心工具；進階工具一律不自動安裝。

---

## 🎯 給 Agent 的必要規則

使用者要求「依這份檔案安裝」時，請遵守：

1. **只在本 repo 內工作**：以本檔所在資料夾為專案根目錄，不要搜尋其他磁碟或其他 Agent 的資料夾。
2. **只執行核心安裝**：核心包含 Tesseract OCR 與繁中模型；不要安裝其餘選用套件或系統工具，除非使用者明確點名用途。
3. **不要使用全域 `pip install`**：核心套件安裝到共用環境 `%LOCALAPPDATA%\file-toolkit\.venv`，避免污染使用者原有 Python。也不要改建在本 repo 內——repo 位於雲端同步資料夾，venv 綁死 base Python 的絕對路徑，換電腦會壞。
4. **不要逐項上網研究**：套件與版本交給 `uv` 解析；不要為每個套件另開網頁、產生長篇計畫或重複說明。
5. **最多重試一次**：失敗時先回報原始錯誤與建議，不要反覆改指令、重裝或自動改用系統管理員權限。
6. **不自動切換執行環境**：Windows、WSL、沙盒是不同環境；不要為了安裝而自行改用 WSL、Docker 或另一個 Agent。

---

## 🚀 核心安裝（預設只做這段）

在本 repo 根目錄執行。所有腳本都要用 PowerShell 7 的 `pwsh`，**不要用 Windows 內建的 `powershell.exe`（5.1）**——
5.1 會把腳本裡的中文讀成亂碼而解析失敗。電腦沒有 `pwsh` 時先裝：`winget install Microsoft.PowerShell`。

**兩個平台的入口不同**：

```powershell
# Windows
pwsh -NoProfile -ExecutionPolicy Bypass -File "./install_windows.ps1"
```

```powershell
# macOS：install_windows.ps1 是 Windows 專屬（winget、Program Files 偵測），
# mac 改跑這兩支跨平台腳本，再手動驗證
pwsh -File skill/scripts/ensure_env.ps1      # 最後一行印出直譯器路徑
pwsh -File skill/scripts/ensure_ocr.ps1      # brew 安裝 tesseract / tesseract-lang / ghostscript
<上一步印出的直譯器路徑> verify_core.py
```

`install_windows.ps1` 本身不重做安裝邏輯，而是依序呼叫技能的兩支跨平台腳本，再驗證：

1. `skill/scripts/ensure_env.ps1`：找到 `uv`（沒有就透過 WinGet 安裝官方 `astral-sh.uv`），以 Python 3.12 建立共用環境 `%LOCALAPPDATA%\file-toolkit\.venv`（本機沒有 3.12 時由 `uv` 下載），安裝 [`skill/scripts/requirements-core.txt`](skill/scripts/requirements-core.txt) 的 12 項核心套件。架構依平台決定：Windows x64 建 x64；**Windows on ARM 先建原生 ARM64，裝不起來或被 Smart App Control 擋就自動改建 x64**；macOS Apple Silicon 建原生 arm64。
2. 環境已存在時（不論架構），`install_windows.ps1` 以 `ensure_env.ps1 -Sync` 依同一份清單同步一次，確保版本與清單一致。
3. 確認 Microsoft Word；若未安裝，清楚回報 `docx2pdf` 的必要前置，不自動安裝 Office。
4. `skill/scripts/ensure_ocr.ps1`：找到 Tesseract（沒有就透過 WinGet 安裝官方套件），並設定英文、方向與繁中 `chi_tra` 模型。
5. 執行 [`verify_core.py`](verify_core.py)，驗證套件匯入、Word 轉 PDF、英文 OCR 與繁中顛倒頁 OCR；通過標準是 `CORE_OK: 12/12`。

> 安裝期間若 Antigravity、Claude Code、Codex 或 OpenCode 顯示執行指令／安裝程式的權限確認，
> 請讓使用者看清楚目標是 `%LOCALAPPDATA%\file-toolkit\.venv`、官方 `astral-sh.uv`、`UB-Mannheim.TesseractOCR`
> 與使用者層級的 Tesseract 語言資料後自行核准。

---

## ✅ 核心必裝套件

這 12 項涵蓋研習最常見的 Word、Excel、PowerPoint、PDF、圖片、圖表、QR Code、繁中 OCR 與教材轉 Markdown：

| 套件 | 用途 |
|------|------|
| `python-docx` | 生成／讀寫 Word |
| `openpyxl` | 讀寫與格式化 Excel |
| `python-pptx` | 生成／改寫 PowerPoint |
| `pypdf` | PDF 合併、拆分、浮水印 |
| `reportlab` | 生成 PDF 與浮水印圖層 |
| `pillow` | 圖片裁切、去白邊、合成 |
| `matplotlib` | 產生統計圖表 |
| `qrcode[pil]` | 產生 QR Code |
| `markitdown[pdf,docx,pptx,xlsx]` | 將 PDF／Word／PPT／Excel 轉成 Markdown |
| `docx2pdf` | 使用本機 Microsoft Word 將 Word 轉成 PDF |
| `ocrmypdf` | 將掃描 PDF 轉成可搜尋、可複製的 PDF |
| `pypdfium2` | PDF 抽文字、轉圖片；同時是 OCRmyPDF 的核心轉圖引擎 |

> 重要修正：只裝裸的 `markitdown` 不會啟用所有文件格式。
> 本工具包改裝 `pdf,docx,pptx,xlsx` 四組官方 extras，才符合影片示範用途。

> 這 12 項還會順便帶進 `pandas`（`markitdown[xlsx]`）、`pdfplumber`（`markitdown[pdf]`）、`xlsxwriter`（`python-pptx`），
> 三者都**不必另外安裝**。

> ⚠️ **PDF 抽文字與轉圖一律用 `pypdfium2`，不要另外裝 `PyMuPDF`（`import fitz`）。**
> 部分 Windows 11 預設開啟 Smart App Control，會永久封鎖它的原生 DLL `_mupdf.pyd`
> （`import fitz` → `No module named 'mupdf'`），而 `pypdfium2` 不受影響，所以 PyMuPDF 已移出核心清單。
> 也不要建議使用者關閉 Smart App Control（它一旦關閉就無法再開啟，必須重設 Windows）。

---

## 🟡 非必要／按需求選裝（預設不要裝）

| 套件 | 為什麼不是核心 |
|------|----------------|
| `docxcompose` | 只有「合併多份 Word」才需要。 |
| `PyMuPDF` | 抽文字與轉圖已由 `pypdfium2` 負責；部分 Windows 11 的 Smart App Control 會封鎖它的原生 DLL。 |
| `pdf2image` | 需另裝 Poppler；PDF 轉圖可先用 `pypdfium2`。 |
| `pywin32` | `docx2pdf` 會自動帶入；只有直接撰寫 Office COM 自動化時才需另外操作。 |
| `yt-dlp` | 只有下載影音時需要，合併影音通常還要 ffmpeg。 |
| `youtube-transcript-api` | 只有抓 YouTube 已存在字幕時需要；無字幕影片無法處理。 |

使用者之後若明確點名某個任務，再由 Agent 把相應套件裝進同一個 `.venv`。不要一次把這張表全部安裝。

---

## ⚙️ 核心前置與選用系統工具

| 系統工具 | 狀態 | 重要限制 |
|----------|------|----------|
| Tesseract OCR | **核心；腳本自動安裝** | 同時設定英文、方向與繁中 `chi_tra` 語言資料。 |
| Microsoft Word | **核心前置；不自動安裝** | `docx2pdf` 需要有授權的桌面版 Word；沒有時核心驗證會停止並回報。 |
| Ghostscript | 選用 | OCRmyPDF 17 的基本 OCR 可用 `pypdfium2`；部分 PDF/A 流程才需要。 |
| Poppler | 使用 `pdf2image` | 核心包已用 `pypdfium2` 轉圖，可先不裝。 |
| ffmpeg | 影音下載、轉檔、合併 | 文件處理不需要。 |

核心腳本只會自動安裝 Tesseract；不要因此安裝 Ghostscript、Poppler、ffmpeg 或 Microsoft Office。這些工具仍需使用者明確點名並核准。

---

## 🧭 四套 Agent 相容性提醒

| Agent | 核心安裝 | 需注意 |
|-------|----------|--------|
| ChatGPT 的 Codex App | ✅ | 必須開啟本機 repo／工作區；一般 ChatGPT 對話本身不會替主機安裝套件。 |
| Claude Code | ✅ | Windows 原生可用 PowerShell；若目前使用 Git Bash，仍請執行上方完整的 `pwsh` 指令。 |
| Google Antigravity | ✅ | 預設權限模式會要求使用者核准安裝；若在沙盒模式，套件只會留在沙盒，不會安裝到 Windows 的共用環境。 |
| OpenCode | ✅（Windows 原生） | 若 OpenCode 跑在 WSL，WSL 與 Windows 的 Python 環境分開，不能共用 Windows 的 venv，需在 WSL 另建環境。 |

> 「一個 Agent 裝一次，其他 Agent 都能用」只在它們使用**同一個作業系統、同一個共用環境**時成立。
> Windows、WSL、Docker、雲端與沙盒彼此不能共用 Python 套件。

---

## 📌 給 Agent 的最終回報格式

```text
核心安裝完成回報：
✅ uv：已存在／已安裝
✅ Python：3.12.x
✅ 環境：<共用 venv 路徑>
✅ 核心套件：CORE_OK 12/12 匯入與煙霧測試成功
✅ Word 轉檔：Microsoft Word＋docx2pdf 可用
✅ OCR：Tesseract（eng／osd／chi_tra）＋OCRmyPDF 可用
🟡 選用套件：未安裝（正確）
下一步：請使用 <venv 直譯器完整路徑> 執行本 repo 的 Python 程式
```

共用 venv 的位置與直譯器檔名兩個平台不同，照 `ensure_env.ps1` 實際印出的那一行填：

| | 共用 venv | 直譯器 |
|---|---|---|
| Windows | `%LOCALAPPDATA%\file-toolkit\.venv` | `Scripts\python.exe` |
| macOS | `~/.local/share/file-toolkit/.venv` | `bin/python` |

若失敗，列出失敗步驟與原始錯誤摘要即可；不要自動安裝選用工具補救。

---

> 來源：file-toolkit
