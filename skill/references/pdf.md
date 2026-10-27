# PDF 處理 recipe

套件：`pypdf`（合併拆分）、`pypdfium2`（抽文字與轉圖）、`reportlab`（生成頁面）、`ocrmypdf`＋Tesseract（OCR）

> 抽文字與轉圖一律用 `pypdfium2`，**不要另外裝 PyMuPDF（`import fitz`）**：它已移出核心，
> 有些 Windows 11 開了 Smart App Control，PyMuPDF 的原生 DLL `_mupdf.pyd` 會被永久封鎖
> （`應用程式控制原則已封鎖此檔案`），而共用環境裡的 `pypdfium2` 不受影響。`pypdfium2` 的 `scale` 以 72 dpi 為 1.0，所以 `scale = dpi / 72`。
>
> **一定要用 `ensure_env.ps1` 印出的直譯器**：Windows on ARM 的共用環境是 x64 Python，裡面的 `pypdfium2` 可以載入；
> 改用 ARM64 系統 Python 另裝的 `pypdfium2` 會被 Smart App Control 擋（`WinError 4551`，2026-10-27 實測）。
> 也不要用 Git Bash 內建的 `pdftotext` 抽中文：它缺 `Adobe-CNS1` 字碼對照表，中文會全部變空白。

---

## D1. 合併／拆分／抽頁／重排

```python
# -*- coding: utf-8 -*-
from pathlib import Path
from pypdf import PdfReader, PdfWriter

Path("output").mkdir(exist_ok=True)

# 合併（依檔名排序，數字檔名要注意 10 會排在 2 前面）
writer = PdfWriter()
for pdf_path in sorted(Path("考古題").glob("*.pdf"), key=lambda p: p.stem):
    writer.append(str(pdf_path))
with open("output/合併.pdf", "wb") as file:
    writer.write(file)

# 抽出第 5～8 頁（人看到的頁碼；程式從 0 開始）
reader = PdfReader("合併.pdf")
part = PdfWriter()
for page_number in range(5, 9):
    part.add_page(reader.pages[page_number - 1])
with open("output/第5-8頁.pdf", "wb") as file:
    part.write(file)

# 每頁拆成一個檔
for index, page in enumerate(reader.pages, start=1):
    single = PdfWriter()
    single.add_page(page)
    with open(f"output/p{index:03d}.pdf", "wb") as file:
        single.write(file)
```

**驗收**：`len(PdfReader("output/合併.pdf").pages)` 要等於各來源頁數總和。

有密碼的 PDF：`reader.decrypt(password)`；密碼要跟使用者拿，不要嘗試破解。

---

## D2. 加浮水印（班級／防外流）

用 `reportlab` 依「每一頁自己的尺寸」畫一張透明浮水印，再疊上去。
頁面尺寸不一致時共用同一張浮水印會歪掉，所以用快取依尺寸各做一張。

```python
# -*- coding: utf-8 -*-
import io
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

TEXT = "302 班 期中複習"
FONT_PATH = Path(r"C:\Windows\Fonts\msjh.ttc")      # 微軟正黑體，中文必備
pdfmetrics.registerFont(TTFont("ZhFont", str(FONT_PATH), subfontIndex=0))


def make_watermark(width, height):
    buffer = io.BytesIO()
    pdf_canvas = canvas.Canvas(buffer, pagesize=(width, height))
    pdf_canvas.setFont("ZhFont", 46)
    pdf_canvas.setFillColor(colors.grey, alpha=0.18)
    pdf_canvas.saveState()
    pdf_canvas.translate(width / 2, height / 2)
    pdf_canvas.rotate(45)
    pdf_canvas.drawCentredString(0, 0, TEXT)
    pdf_canvas.restoreState()
    pdf_canvas.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]


reader = PdfReader("講義.pdf")
writer = PdfWriter()
cache = {}
for page in reader.pages:
    size = (round(float(page.mediabox.width), 1), round(float(page.mediabox.height), 1))
    if size not in cache:
        cache[size] = make_watermark(*size)
    page.merge_page(cache[size])                     # 浮水印疊在原內容之上
    writer.add_page(page)

Path("output").mkdir(exist_ok=True)
with open("output/講義_浮水印.pdf", "wb") as file:
    writer.write(file)
```

浮水印是視覺標示，**擋不住複製文字**，不要跟使用者說這樣就防得住外流。

---

## D3. 掃描 PDF OCR → 可搜尋、可複製

前置：先跑一次本技能的 `scripts\ensure_ocr.ps1`（安裝 Tesseract 與 `chi_tra` 繁中模型）。

```python
# -*- coding: utf-8 -*-
from pathlib import Path
import ocrmypdf
import pypdfium2 as pdfium

source = Path("掃描講義.pdf")
target = Path("output/掃描講義_OCR.pdf")
target.parent.mkdir(exist_ok=True)

ocrmypdf.ocr(
    source,
    target,
    language=["chi_tra", "eng"],     # 繁中講義固定用這組；純英文用 ["eng"]
    output_type="pdf",               # 不要用 pdfa，會需要 Ghostscript
    optimize=0,
    skip_text=True,                  # 已經有文字層的頁面跳過，避免重複疊字
    progress_bar=False,
    deskew=True,                     # 掃歪的紙本先校正
    rotate_pages=True,               # 紙本放反、掃成上下顛倒時自動轉正，不加的話整頁辨識不到
    rotate_pages_threshold=5.0,      # 預設 14 太嚴：純中文頁的方向信心常只有 5～13，會判斷對卻不轉
)

document = pdfium.PdfDocument(target)
try:
    text = "\n".join(page.get_textpage().get_text_range() for page in document)
finally:
    document.close()
print(f"抽到 {len(text.strip())} 字")
print(text[:200])
```

**陷阱**

- OCR 很慢：A4 每頁約 2～5 秒，整本 100 頁先跟使用者說要等；先只做前 3 頁確認辨識品質再全做
- `output_type="pdfa"` 需要 Ghostscript，核心環境沒裝，會直接失敗
- 找不到 `chi_tra`：模型不能放在 `Program Files`，要在 `%LOCALAPPDATA%\Tesseract-OCR\tessdata`
- 辨識率差通常是掃描 dpi 太低（< 200）或紙張太歪，這是來源問題，要如實告訴使用者
- 字很少、印章很大的頁面（簽章頁、核章頁）用預設設定常常一個字都抽不到：方向信心只有 3 左右，低於門檻 5 就不會轉正，就算轉正了預設的版面分析也找不到字。這種頁面先用 pypdf 的 `page.rotate(180)` 手動轉正，再加 `tesseract_pagesegmode=11` 重跑
- 不要為了這種頁面把 `rotate_pages_threshold` 再往下調：信心越低越容易轉錯方向，整批文件都會受影響
- 框線裡的大字標題（例如封面的「學生名冊」）可能整個漏抓，調整設定也救不回來；要靠這個標題搜尋的話，先跟使用者說明

---

## D4. 指定頁轉圖片、去白邊

```python
# -*- coding: utf-8 -*-
from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image, ImageChops

PAGES = [12]              # 人看到的頁碼
DPI = 300                 # 要貼進學習單列印用 300

Path("output").mkdir(exist_ok=True)
document = pdfium.PdfDocument("課本.pdf")
try:
    for page_number in PAGES:
        page = document[page_number - 1]
        image = page.render(scale=DPI / 72).to_pil().convert("RGB")   # scale 以 72 dpi 為 1.0

        background = Image.new("RGB", image.size, (255, 255, 255))
        bbox = ImageChops.difference(image, background).getbbox()     # 去白邊
        if bbox:
            image = image.crop(bbox)

        image.save(f"output/課本_p{page_number}.png")
finally:
    document.close()
```

- 抽文字：`page.get_textpage().get_text_range()`；抽不到就是圖片型 PDF，要先做 D3 的 OCR
- 只要頁面上的某一張圖（不是整頁）：核心包沒有直接抽內嵌圖的簡單做法，先整頁轉圖再裁切
- 轉 JPG 用 `image.save(..., quality=90)`；有透明背景要先 `convert("RGB")`
