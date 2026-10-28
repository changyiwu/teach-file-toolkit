"""驗證核心教學檔案處理套件。"""

from __future__ import annotations

import importlib
import os
import re
import shutil
import subprocess
import sys
import sysconfig
import tempfile
from pathlib import Path


MODULES = {
    "docx": "python-docx",
    "openpyxl": "openpyxl",
    "pptx": "python-pptx",
    "pypdf": "pypdf",
    "reportlab": "reportlab",
    "PIL": "Pillow",
    "matplotlib": "matplotlib",
    "qrcode": "qrcode",
    "markitdown": "markitdown",
    "docx2pdf": "docx2pdf",
    "ocrmypdf": "ocrmypdf",
    "pypdfium2": "pypdfium2",
}

# 繁中 OCR 測試頁的內容：中英混排，方向信心才穩定高於 recipe 的門檻 5（純中文約 6，太貼近門檻）
CJK_OCR_LINES = [
    "教學檔案處理工具包", "掃描講義轉成可搜尋文字", "這一頁是上下顛倒的測試", "老師批次處理文件更輕鬆",
    "Grade 7 Math Quiz 2026", "第一次評量 1-1 至 1-4", "Name ______ No. 12", "考試時間 45 分鐘 Score 100",
]
CJK_FONT_CANDIDATES = [
    Path(r"C:\Windows\Fonts\msjh.ttc"),              # 微軟正黑體
    Path(r"C:\Windows\Fonts\mingliu.ttc"),           # 細明體
    Path("/System/Library/Fonts/PingFang.ttc"),
    Path("/System/Library/Fonts/STHeiti Medium.ttc"),
]


def find_word() -> Path | None:
    if sys.platform == "darwin":
        # docx2pdf 在 macOS 走 AppleScript 驅動 Word for Mac，跟 Windows 走 COM 是同一個
        # 套件的兩條路——兩邊都要求裝桌面版 Word。.app 是資料夾，所以用 exists()。
        mac_word = Path("/Applications/Microsoft Word.app")
        return mac_word if mac_word.exists() else None

    if sys.platform != "win32":
        return None

    candidates = [
        Path(r"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\Root\Office16\WINWORD.EXE"),
    ]
    try:
        import winreg

        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(
                    hive,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
                ) as key:
                    candidates.insert(0, Path(winreg.QueryValue(key, None)))
            except OSError:
                pass
    except ImportError:
        pass

    return next((path for path in candidates if path.is_file()), None)


def find_tesseract() -> Path | None:
    command = shutil.which("tesseract")
    candidates = [Path(command) if command else None]
    if sys.platform == "win32":
        # macOS 靠 Homebrew 裝，一定在 PATH 上，不需要候選清單
        candidates += [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
    return next((path for path in candidates if path and path.is_file()), None)


def verify_external_tools() -> Path:
    word = find_word()
    if word is None:
        hint = ("安裝 Word for Mac" if sys.platform == "darwin" else "安裝桌面版 Microsoft Word")
        raise RuntimeError(f"core docx2pdf workflow 需要桌面版 Microsoft Word：{hint}")

    tesseract = find_tesseract()
    if tesseract is None:
        hint = ("brew install tesseract tesseract-lang ghostscript"
                if sys.platform == "darwin" else "跑 skill/scripts/ensure_ocr.ps1")
        raise RuntimeError(f"core OCRmyPDF workflow 需要 Tesseract OCR：{hint}")

    result = subprocess.run(
        [str(tesseract), "--list-langs"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=os.environ.copy(),
    )
    languages = f"{result.stdout}\n{result.stderr}"
    missing = [language for language in ("eng", "osd", "chi_tra") if language not in languages]
    if result.returncode != 0 or missing:
        raise RuntimeError(f"Tesseract languages missing: {', '.join(missing) or 'unknown error'}")

    os.environ["PATH"] = f"{tesseract.parent}{os.pathsep}{os.environ.get('PATH', '')}"
    return tesseract


def run_smoke_tests() -> None:
    import matplotlib
    import ocrmypdf
    import pypdfium2
    import qrcode
    from docx import Document
    from docx2pdf import convert as convert_docx_to_pdf
    from markitdown import MarkItDown
    from openpyxl import Workbook
    from PIL import Image, ImageDraw, ImageFont
    from pptx import Presentation
    from pptx.util import Inches
    from pypdf import PdfReader
    from reportlab.pdfgen import canvas

    matplotlib.use("Agg")
    from matplotlib import pyplot as plt

    with tempfile.TemporaryDirectory(prefix="teacher-file-core-") as temp_dir:
        root = Path(temp_dir)

        docx_path = root / "sample.docx"
        document = Document()
        document.add_paragraph("TEACHER FILE DOCX")
        document.save(docx_path)

        docx_pdf_path = root / "sample-docx.pdf"
        convert_docx_to_pdf(str(docx_path), str(docx_pdf_path))
        assert docx_pdf_path.is_file(), "docx2pdf did not create a PDF"
        assert len(PdfReader(docx_pdf_path).pages) == 1

        xlsx_path = root / "sample.xlsx"
        workbook = Workbook()
        workbook.active["A1"] = "TEACHER FILE XLSX"
        workbook.save(xlsx_path)

        pptx_path = root / "sample.pptx"
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text = "TEACHER FILE PPTX"
        presentation.save(pptx_path)

        pdf_path = root / "sample.pdf"
        pdf_canvas = canvas.Canvas(str(pdf_path))
        pdf_canvas.drawString(72, 720, "TEACHER FILE PDF")
        pdf_canvas.save()
        assert len(PdfReader(pdf_path).pages) == 1
        pdf_document = pypdfium2.PdfDocument(pdf_path)
        try:
            assert len(pdf_document) == 1
            pdf_document[0].render(scale=0.5).to_pil()
        finally:
            pdf_document.close()

        image_path = root / "sample.png"
        Image.new("RGB", (32, 32), "white").save(image_path)
        qrcode.make("https://example.com/file-toolkit").save(root / "qr.png")

        figure = plt.figure(figsize=(2, 1))
        plt.plot([0, 1], [0, 1])
        figure.savefig(root / "chart.png")
        plt.close(figure)

        converter = MarkItDown()
        expected_text = {
            docx_path: "TEACHER FILE DOCX",
            xlsx_path: "TEACHER FILE XLSX",
            pptx_path: "TEACHER FILE PPTX",
            pdf_path: "TEACHER FILE PDF",
        }
        for source_path, expected in expected_text.items():
            result = converter.convert(str(source_path))
            assert expected in result.text_content, f"MarkItDown failed for {source_path.suffix}"

        scan_pdf_path = root / "sample-scan.pdf"
        scan_image = Image.new("RGB", (1200, 500), "white")
        try:
            ocr_font = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 72)
        except OSError:
            ocr_font = ImageFont.load_default()
        ImageDraw.Draw(scan_image).text((80, 180), "OCR TEST 123", fill="black", font=ocr_font)
        scan_image.save(scan_pdf_path, "PDF", resolution=200.0)

        ocr_pdf_path = root / "sample-ocr.pdf"
        ocrmypdf.ocr(
            scan_pdf_path,
            ocr_pdf_path,
            language=["eng"],
            output_type="pdf",
            pdf_renderer="sandwich",   # 原生 Windows ARM64 環境沒有 uharfbuzz，預設的 fpdf2 渲染器會失敗
            optimize=0,
            progress_bar=False,
        )
        ocr_document = pypdfium2.PdfDocument(ocr_pdf_path)
        try:
            ocr_text = " ".join(page.get_textpage().get_text_range() for page in ocr_document)
        finally:
            ocr_document.close()
        assert "OCR" in ocr_text and "123" in ocr_text, "OCRmyPDF text smoke test failed"

        verify_cjk_upside_down_ocr(root)


def verify_cjk_upside_down_ocr(root: Path) -> None:
    """繁中掃描頁上下顛倒：用 references/pdf.md D3 的參數，確認會自動轉正且辨識得到中文。"""
    import ocrmypdf
    import pypdfium2
    from PIL import Image, ImageDraw, ImageFont

    font_path = next((path for path in CJK_FONT_CANDIDATES if path.is_file()), None)
    if font_path is None:
        raise RuntimeError("找不到中文字型，無法產生繁中 OCR 測試頁：" + "、".join(map(str, CJK_FONT_CANDIDATES)))

    page = Image.new("RGB", (1654, 2339), "white")      # A4 @ 200 dpi，跟一般影印機掃描相同
    draw = ImageDraw.Draw(page)
    font = ImageFont.truetype(str(font_path), 48)
    for index, line in enumerate(CJK_OCR_LINES * 2):
        draw.text((150, 200 + index * 86), line, fill="black", font=font)
    source = root / "sample-cjk-upside-down.pdf"
    page.rotate(180).save(source, "PDF", resolution=200.0)

    target = root / "sample-cjk-ocr.pdf"
    ocrmypdf.ocr(
        source,
        target,
        language=["chi_tra", "eng"],
        output_type="pdf",
        pdf_renderer="sandwich",
        optimize=0,
        skip_text=True,
        progress_bar=False,
        deskew=True,
        rotate_pages=True,
        rotate_pages_threshold=5.0,
    )
    document = pypdfium2.PdfDocument(target)
    try:
        text = "".join("".join(p.get_textpage().get_text_range() for p in document).split())
    finally:
        document.close()

    expected = set(re.findall(r"[一-鿿]", "".join(CJK_OCR_LINES)))
    coverage = len(expected & set(text)) / len(expected)
    assert coverage >= 0.8 and "教學檔案處理" in text, (
        f"繁中顛倒頁 OCR 失敗（中文涵蓋 {coverage:.0%}）：確認 chi_tra／osd 模型，以及 rotate_pages 參數"
    )


def main() -> int:
    failures: list[tuple[str, str]] = []
    for module_name, package_name in MODULES.items():
        try:
            importlib.import_module(module_name)
        except Exception as exc:  # pragma: no cover - intended diagnostic path
            failures.append((package_name, f"{type(exc).__name__}: {exc}"))

    if failures:
        print(f"CORE_FAIL: {len(MODULES) - len(failures)}/{len(MODULES)}")
        for package_name, error in failures:
            print(f"- {package_name}: {error}")
        return 1

    try:
        verify_external_tools()
        run_smoke_tests()
    except Exception as exc:  # pragma: no cover - intended diagnostic path
        print(f"CORE_SMOKE_FAIL: {type(exc).__name__}: {exc}")
        return 1

    version = '.'.join(str(part) for part in sys.version_info[:3])
    print(f"CORE_OK: {len(MODULES)}/{len(MODULES)} imports and file smoke tests (Python {version}, {sysconfig.get_platform()})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
