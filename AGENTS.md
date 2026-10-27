# teach-file-toolkit（專案藍圖）

> 本檔為跨 Agent 通用的專案藍圖（AGENTS.md 開放標準）。任何 Agent 的每個 session 都應先讀本檔＋`handoff.md`。

## 專案簡介

提供老師在本機處理 Word、Excel、PowerPoint、PDF、圖片、圖表、QR Code 與 Markdown 教材的工具清單、安裝腳本及驗證方式。核心清單 12 項全部固定版本，驗證通過為 `CORE_OK: 12/12`（`PyMuPDF` 已於 2026-10-18 移出核心：開啟 Smart App Control 的 Windows 會封鎖它的原生 DLL，且沒有任何 recipe 用它）。本專案不使用 Firebase，也沒有部署目標。

## 關鍵時程

<!-- 目前無固定時程 -->

## 目標與路線圖

- [x] 階段一：核心工具清單、安裝腳本 `install_windows.ps1` 與煙霧測試 `verify_core.py` 成形
- [x] 階段二：納入 `docx2pdf`、`ocrmypdf`、`pypdfium2`，核心清單由 10 項擴充為 13 項
- [x] 階段三：安裝腳本加上 Word 檢查、自動裝 Tesseract＋繁中模型（固定版本 SHA-256）
- [x] 階段四：`verify_core.py` 補上 Word 轉 PDF 與圖片型 PDF OCR 實測
- [x] 階段五：把檔案處理能力包裝成 `skill/`（SKILL.md＋環境腳本＋五份 recipe），16 段 recipe 程式碼實測通過
- [x] 階段六：PDF 抽文字與轉圖改用 `pypdfium2`，讓技能在開啟 Smart App Control 的電腦也能完整運作（`CORE_OK: 12/12`）
- [x] 階段七：把 `Python範例.md`、`教學檔案處理_工具列表.md`、`AGENT_SETUP_教學檔案處理工具包.md`（連同 `make_pdf.py` 與工具列表 PDF）裡「PDF 轉圖用 PyMuPDF」的敘述同步改為 `pypdfium2`
- [x] 階段八：`install_windows.ps1` 改建共用環境 `%LOCALAPPDATA%\file-toolkit\.venv`（與技能同一份，不再建在雲端同步的 repo 內）；`Python範例.md` 與 `教學檔案處理_工具列表.md` 合併為單一清單，移除 `make_pdf.py` 與工具列表 PDF
- [x] 階段九：以實際教材測試 Word 批次轉 PDF／JPG（通過），以及繁中掃描 PDF OCR（D3 補上 `rotate_pages`＋門檻 5、稀疏頁與框線標題兩條陷阱）
- [x] 階段十：全專案檢視整理——移除 PyMuPDF、選裝文件改寫間接相依、`install_windows.ps1` 改串接技能兩支腳本、統一 pwsh、修 `SKILL.md` 與 Excel E1／E3、`verify_core.py` 加繁中顛倒頁 OCR 測試
- [ ] 階段十一：只有需要特定 PDF/A 流程時才另外評估 Ghostscript

## 資料夾結構

```
teach-file-toolkit/
├─ README.md                              # 專案用途與初學者使用方式
├─ AGENT_SETUP_教學檔案處理工具包.md        # 交給 AI Agent 的安裝與安全指引
├─ install_windows.ps1                    # Windows 入口：串接 skill/scripts 兩支腳本＋Word 檢查＋驗證
├─ requirements-core.txt                  # 核心 Python 相依套件（固定版本）
├─ verify_core.py                         # 核心套件匯入與檔案處理煙霧測試
├─ 教學檔案處理_工具列表.md                 # 工具清單唯一來源（痛點／套件／一句話＋安裝與選裝）
├─ skill/                                 # 打包成 Agent 技能的原始檔（安裝名：file-toolkit）
│  ├─ SKILL.md                            # 技能主檔：判斷任務→備環境→讀 recipe→執行驗收
│  ├─ scripts/ensure_env.ps1              # 尋找或建立含 12 核心套件的 Python 環境
│  ├─ scripts/ensure_ocr.ps1              # 安裝設定 Tesseract＋繁中模型（OCR 任務前執行）
│  ├─ scripts/requirements-core.txt       # 技能自帶的核心套件清單（與根目錄同版本）
│  └─ references/                         # word／excel／pptx／pdf／misc 五份實測 recipe
├─ AGENTS.md                              # 本檔：專案藍圖
├─ handoff.md                             # 交接檔（每次收工必更新）
├─ .agents/  .gitignore                    # 環境不在 repo 內，見 %LOCALAPPDATA%\file-toolkit\.venv
```

## 同步層級（本專案初始化至第 3 層級）

| 層級 | 平台 | 位置 | 讀取時機 |
|------|------|------|---------|
| L1 | 本地（GDrive） | `AGENTS.md`＋`handoff.md` | 每個 session |
| L2 | GitHub | https://github.com/changyiwu/teach-file-toolkit （公開，預設分支 `main`） | 指定時 |
| L3 | Obsidian | `teach-file-toolkit/專案工作流程.md` | 有需要時 |

## 三個檔案的職責（依「時效性」分家，不是依「詳細程度」）

| 檔案 | 時效 | 寫入方式 | 放什麼 |
|------|------|---------|--------|
| `handoff.md` | **只對下一個 session 有效**，過期即丟 | 每次收工整份重寫 | 做到哪、下一步、**這次**的暫時 workaround |
| `AGENTS.md`（本檔） | **長期有效**，每個 session 都適用 | 只有規則本身變了才改 | 目標、路線圖、常設規則、結構 |
| Obsidian／`git log` | **歷史**：發生過什麼、為什麼 | 只增不刪 | 決策紀錄、踩坑完整版、逐次進度 |

驗收標準：**`handoff.md` 整份刪掉，不應損失任何長期資訊**——會的話代表該升級進本檔卻沒升級。

**本檔不要出現的東西**：❌ `## 最近進度`／逐次工作紀錄、❌ 決策理由與踩坑完整版。2026-08-03 移除了 `## 最近進度`，內容逐條比對後已在 L3 筆記的〈🗓️ 最近更動紀錄〉——**是主動移除，不是遺漏，不要補回來**。踩過的坑只把**結論**收斂成一條祈使句寫進〈工作約定〉，原因留 L3。

## 工作約定

- 任何 Agent、任何電腦：**開工先讀 `handoff.md`，收工必更新 `handoff.md`**
- 修改共用檔案前先讀最新內容，避免覆蓋其他 Agent 的變更
- 所有回應與文件使用繁體中文
- 推送前先抓取並確認 `origin/main`、提交範圍及工作區狀態；**禁止 force push**
- 保留既有 Git 歷史與 `main` 分支
- 本專案位於雲端同步資料夾，Git 本機設定應維持 `windows.appendAtomically=false`（**Windows 專屬**，macOS 不需要）
- 更新 Obsidian 專案筆記時，不要修改 `02-知識庫/log.md`
- 工具清單只有 `教學檔案處理_工具列表.md` 一份，不要再另外複製一份工具地圖或產生 PDF 版
- PDF 抽文字與轉圖一律用 `pypdfium2`，程式碼與說明文件都不要再出現「用 PyMuPDF／`import fitz`」的敘述；也不要把 `PyMuPDF` 加回 `requirements-core.txt`
- `pandas`、`pdfplumber`、`xlsxwriter` 是核心套件的間接相依（`markitdown[xlsx]`、`markitdown[pdf]`、`python-pptx`），文件一律寫「已隨核心安裝」，不要再列為選裝
- 環境建置邏輯只寫在 `skill/scripts/ensure_env.ps1`＋`ensure_ocr.ps1`；`install_windows.ps1` 只負責串接、Word 檢查與驗證，**不要再把安裝邏輯複製回去**
- 根目錄與 `skill/scripts/` 各有一份 `requirements-core.txt`（技能拷出 repo 也要能用），兩份必須逐位元組相同，`install_windows.ps1` 會檢查
- Python 環境一律建在 repo 外的共用位置（Windows `%LOCALAPPDATA%\file-toolkit\.venv`、macOS `~/.local/share/file-toolkit/.venv`），`install_windows.ps1` 透過 `skill/scripts/ensure_env.ps1` 用同一份；**不要在 repo 內建 `.venv`**（repo 在雲端同步資料夾，而 venv 會寫死 base Python 的絕對路徑，換電腦就壞）
- **Windows on ARM（RTX Spark、Snapdragon）的共用環境一律建成 x64 Python**（`cpython-3.12-windows-x86_64-none`，靠 Prism 模擬）。原生 ARM64 裝不起核心套件：`cryptography`（`pdfminer.six` 拉進來的）、`pi-heif`、`uharfbuzz` 都沒有 win_arm64 wheel。這段邏輯只在 `ensure_env.ps1`，`install_windows.ps1` 呼叫它
- **`install_windows.ps1` 是 Windows 專屬**（winget、Program Files 偵測），已列入 `.platform-ok`。macOS 的環境建置走 `skill/scripts/ensure_env.ps1` ＋ `ensure_ocr.ps1`，那兩支是跨平台的；跨平台約定見 `agent-project-flow/platform.md`
- **PowerShell 一律 pwsh 7**，三支腳本開頭都有 `$IsWindows` 的 guard；`.ps1` 一律 UTF-8 無 BOM。所有文件裡的執行指令都寫 `pwsh -NoProfile -ExecutionPolicy Bypass -File ...`，**不要寫 `powershell.exe`**；Agent 的權限機制擋下 `-ExecutionPolicy Bypass` 時，改在 pwsh 裡直接 `& "<腳本路徑>"` 執行（執行原則允許本機腳本時可跑，2026-10-27 實測），`SKILL.md` 第 1 節有寫這個備案
- 共用環境是既有環境就重用（3.12 以上即可），不要因為版本不是 3.12 就叫使用者刪掉——那是技能正在用的同一份
- **產出一律放 `output/`**：本專案產生的檔案或資料夾（成品、匯出、報表、生成圖、渲染等）預設放專案根目錄的 `output/`，不另開 `outputs/`、`exports/`、`generated/`、`out/` 之類的資料夾，方便找；`output/` 不進 git。部署工具指定的建置目錄（`dist/`、`build/`）不在此限

## 安全邊界

- 保留使用者既有修改；不要覆蓋原始教材或輸入檔，處理結果使用新檔名輸出
- `README.md` 的頻道推廣署名已於 2026-08-22 依使用者決定移除（頻道名、YouTube 連結、「製作：」行）。**不要加回來**；原作者歸屬另以「原作者三師爸」一句保留於出處說明，那句不可刪
- 不提交 `.env`、憑證、權杖、密碼、`.codex/` 或 `.claude/`
- 學生資料只使用班級代碼與座號，不儲存真實姓名；預設在本機處理
- 初始化工作區不代表授權部署、變更公開狀態、啟用 GitHub Pages、commit 或 push

## 已知環境陷阱

- `chi_tra.traineddata` 不能寫入 `Program Files`，改放 `%LOCALAPPDATA%\Tesseract-OCR\tessdata`（並複製 `configs`／`tessconfigs`）
- `docx2pdf` 需桌面版 Word，且不支援 `python -m docx2pdf`
- `powershell.exe`（5.1）會用 Big5 解讀無 BOM 的 UTF-8，中文註解的結尾會吃掉下一行程式碼，腳本直接解析失敗（2026-10-18 實測 `ensure_env.ps1` 在第 114 行出錯）——所以一律用 `pwsh` 執行，不是把腳本改成 ASCII
- **Smart App Control（部分 Windows 11 預設開啟；NB-YI 與 PC-YI-SL 於 2026-08-20 皆為 `1`——PC-YI-SL 在 2026-08-05 曾記為未開啟，狀態會變，每台都要當場查）**會封鎖**尚未取得信譽**的執行檔與原生 DLL，典型症狀是 `uv.exe` 無法執行、PyMuPDF 的 `_mupdf.pyd` 載入失敗（`import fitz` → `No module named 'mupdf'`）。**但信譽會隨版本累積，不要假設某台電腦一定不行——當場實測**（2026-08-05 於 NB-YI／SAC=1 實測：`uv 0.11.31` 可執行、`PyMuPDF 1.28.0` 可 import）。真的被擋時的備案是改用系統 Python 的 `venv`＋`pip` 建環境。判斷指令：`(Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy').VerifiedAndReputablePolicyState`，`1` 為開啟。**SAC 一旦關閉就無法再開啟（需重設 Windows），不要建議使用者關掉它**
- SAC 也會擋 **uv 寫進 `Scripts\` 的 45 KB 啟動器**，症狀是 `.venv\Scripts\python.exe` 一執行就回「應用程式控制原則已封鎖此檔案」，但**套件與 base Python 都是好的，壞的只有入口**。`ensure_env.ps1`（`install_windows.ps1` 也是透過它）已會自動偵測並用 base Python 的 `-m venv` 就地重建入口（保留 `site-packages`，不重裝套件）。同樣**信譽逐版累積**：2026-08-20 於 PC-YI-SL 實測，uv 0.11.25 於 2026-08-06 建的入口被擋、當下版本新建的沒被擋
- **Windows on ARM 被 SAC 擋時，不要退回 ARM64 的系統 Python**（上面「系統 Python 的 `venv`＋`pip`」備案只適用 x64 電腦）：2026-10-27 於 LT-YI-RS（ARM64、SAC=`1`）實測，原生 ARM64 版 `pypdfium2` 的 `pdfium.dll` 被擋（`WinError 4551`），`pdfminer.six` 要編譯 `cryptography` 而裝不起來；同一天建的 x64 共用環境裡 `pypdfium2` 正常載入。ARM 電腦一律修 x64 共用環境
- `.venv` 綁定建立當下的 base Python 絕對路徑，放在雲端同步資料夾**不能跨電腦共用**；共用環境一律建在 `%LOCALAPPDATA%\file-toolkit\.venv`
- `ocrmypdf` 執行時會探測選用外部程式，印出數行 `[WinError 2] 系統找不到指定的檔案。` 屬正常，不影響結果
- 跑 OCR 前必須有 `TESSDATA_PREFIX`。`ensure_ocr.ps1` 會設在使用者層級，但**同一個 session 已開的 shell 讀不到**——要重開 shell 或當場再設一次
- `pypdfium2` 的 `scale` 以 72 dpi 為 1.0，換算是 `scale = dpi / 72`；抽文字用 `page.get_textpage().get_text_range()`
