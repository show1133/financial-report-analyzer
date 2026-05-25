# -*- coding: utf-8 -*-
"""
CEO財報分析高階決策報告 - 自動生成器
Author: GitHub Copilot
Description: 讀取上課資料中所有PDF與Word檔案，分析財報內容，生成CEO高階決策報告
"""

import os
import sys
import re
import traceback
from pathlib import Path
from datetime import datetime

# ── 安裝依賴 ──────────────────────────────────────────────────────────────────
import subprocess

def _install(pkg):
    subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

for pkg in ["pdfplumber", "python-docx", "pymupdf", "pdf2image", "pytesseract", "Pillow"]:
    import_name = {
        "python-docx": "docx",
        "pymupdf": "fitz",
        "Pillow": "PIL",
    }.get(pkg, pkg)
    try:
        __import__(import_name)
    except ImportError:
        print(f"  📦 安裝 {pkg}...")
        _install(pkg)

import pdfplumber
import fitz  # pymupdf
import pytesseract
from PIL import Image
import pdf2image
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Tesseract 設定 ─────────────────────────────────────────────────────────────
# Windows 預設安裝路徑；如安裝在其他位置請修改
_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(
        os.environ.get("USERNAME", "User")
    ),
]
_TESSERACT_AVAILABLE = False
for _tp in _TESSERACT_PATHS:
    if os.path.isfile(_tp):
        pytesseract.pytesseract.tesseract_cmd = _tp
        _TESSERACT_AVAILABLE = True
        break

# 確認繁體中文語言包是否安裝
_OCR_LANG = "chi_tra+eng"
if _TESSERACT_AVAILABLE:
    try:
        _langs = pytesseract.get_languages()
        if "chi_tra" not in _langs:
            _OCR_LANG = "chi_sim+eng" if "chi_sim" in _langs else "eng"
    except Exception:
        _OCR_LANG = "eng"

# ── 路徑設定 ──────────────────────────────────────────────────────────────────
SOURCE_DIR = Path(r"C:\Python\99財務管理\上課資料")
REPORT_DIR = Path(r"C:\Python\99財務管理\report")
REPORT_PATH = REPORT_DIR / "CEO財報分析決策報告.docx"
MAX_PAGES_PER_PDF = 5   # OCR模式下多讀幾頁以提升內容完整度
MAX_CHARS_PER_FILE = 10000

# ── 工具函式 ──────────────────────────────────────────────────────────────────

def _extract_with_pdfplumber(pdf_path: Path, max_pages: int) -> str:
    """第一層：pdfplumber 萃取"""
    try:
        texts = []
        with pdfplumber.open(str(pdf_path)) as pdf:
            for i, page in enumerate(pdf.pages[:max_pages]):
                try:
                    text = page.extract_text()
                    if text and text.strip():
                        texts.append(f"[第{i+1}頁]\n{text.strip()}")
                except Exception:
                    pass
        return "\n\n".join(texts)
    except Exception:
        return ""


def _extract_with_pymupdf(pdf_path: Path, max_pages: int) -> str:
    """第二層：pymupdf 萃取（對部分影像型PDF有更好效果）"""
    try:
        texts = []
        with fitz.open(str(pdf_path)) as doc:
            for i, page in enumerate(doc[:max_pages]):
                text = page.get_text("text").strip()
                if text:
                    texts.append(f"[第{i+1}頁]\n{text}")
        return "\n\n".join(texts)
    except Exception:
        return ""


def _extract_with_ocr(pdf_path: Path, max_pages: int) -> str:
    """第三層：pdf2image + pytesseract OCR（真正的影像型PDF）"""
    if not _TESSERACT_AVAILABLE:
        return ""
    try:
        images = pdf2image.convert_from_path(
            str(pdf_path),
            first_page=1,
            last_page=max_pages,
            dpi=200,
        )
        texts = []
        for i, img in enumerate(images):
            try:
                text = pytesseract.image_to_string(img, lang=_OCR_LANG, config="--psm 3")
                text = text.strip()
                if text:
                    texts.append(f"[第{i+1}頁 OCR]\n{text}")
            except Exception as e:
                texts.append(f"[第{i+1}頁 OCR失敗: {e}]")
        return "\n\n".join(texts)
    except Exception as e:
        return f"(OCR失敗: {e})"


def safe_extract_pdf(pdf_path: Path, max_pages: int = MAX_PAGES_PER_PDF) -> str:
    """三層萃取策略：pdfplumber → pymupdf → OCR（支援影像型PDF）"""
    MIN_TEXT_LEN = 50  # 少於此字數視為萃取失敗，嘗試下一層

    # 第一層：pdfplumber
    result = _extract_with_pdfplumber(pdf_path, max_pages)
    if len(result.strip()) >= MIN_TEXT_LEN:
        return result[:MAX_CHARS_PER_FILE]

    # 第二層：pymupdf
    print(f"    ⚡ pdfplumber萃取不足，嘗試 pymupdf...")
    result = _extract_with_pymupdf(pdf_path, max_pages)
    if len(result.strip()) >= MIN_TEXT_LEN:
        return result[:MAX_CHARS_PER_FILE]

    # 第三層：OCR
    if _TESSERACT_AVAILABLE:
        print(f"    🔍 嘗試 OCR（{_OCR_LANG}）...")
        result = _extract_with_ocr(pdf_path, max_pages)
        if result and not result.startswith("(OCR失敗"):
            return result[:MAX_CHARS_PER_FILE]
        return f"(OCR嘗試但失敗：{result})"
    else:
        return (
            "(此PDF為影像型，無可萃取文字。\n"
            "💡 請安裝 Tesseract OCR 以啟用中文識別：\n"
            "   下載：https://github.com/UB-Mannheim/tesseract/wiki\n"
            "   安裝時勾選 'Additional language data' → '繁體中文(chi_tra)'\n"
            "   預設路徑：C:\\Program Files\\Tesseract-OCR\\tesseract.exe)"
        )


def safe_extract_docx(docx_path: Path) -> str:
    """從Word文件萃取文字"""
    try:
        doc = Document(str(docx_path))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        result = "\n".join(paragraphs)
        return result[:MAX_CHARS_PER_FILE] if result else "(Word文件無可讀文字)"
    except Exception as e:
        return f"(Word文件讀取錯誤: {e})"


def extract_all_files(source_dir: Path) -> dict:
    """萃取目錄下所有PDF與Word檔案的文字（含OCR）"""
    extracted = {}
    if not source_dir.exists():
        print(f"⚠ 來源目錄不存在: {source_dir}")
        return extracted

    files = list(source_dir.iterdir())
    pdf_files = [f for f in files if f.suffix.lower() == ".pdf"]
    docx_files = [f for f in files if f.suffix.lower() == ".docx"]

    ocr_status = f"✅ Tesseract可用（語言：{_OCR_LANG}）" if _TESSERACT_AVAILABLE else "⚠️ Tesseract未安裝（影像型PDF將無法OCR）"
    print(f"找到 {len(pdf_files)} 個PDF、{len(docx_files)} 個Word文件")
    print(f"OCR狀態：{ocr_status}")
    if not _TESSERACT_AVAILABLE:
        print("  💡 安裝教學：https://github.com/UB-Mannheim/tesseract/wiki")
        print("     安裝後勾選 'Additional language data' → chi_tra (繁體中文)")

    success, ocr_success, failed = 0, 0, 0
    for f in sorted(pdf_files):
        print(f"  📄 萃取PDF: {f.name}")
        content = safe_extract_pdf(f)
        extracted[f.name] = {"type": "pdf", "content": content}
        if content.startswith("("):
            failed += 1
        elif "OCR" in content:
            ocr_success += 1
        else:
            success += 1

    for f in sorted(docx_files):
        print(f"  📝 萃取Word: {f.name}")
        extracted[f.name] = {"type": "docx", "content": safe_extract_docx(f)}

    print(f"\n  📊 PDF萃取結果：文字型成功 {success} 個 | OCR成功 {ocr_success} 個 | 失敗 {failed} 個")
    return extracted


def first_n_chars(text: str, n: int = 1200) -> str:
    """取文字前N字，避免章節過長"""
    if not text or text.startswith("("):
        return text
    return text[:n] + ("…" if len(text) > n else "")


# ── Word格式輔助 ───────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color: str):
    """設定表格儲存格背景色"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_styled_table(doc: Document, headers: list, rows: list,
                     header_bg: str = "1F4E79", header_fg: str = "FFFFFF"):
    """新增帶樣式的表格"""
    col_count = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=col_count)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.text = h
        set_cell_bg(cell, header_bg)
        para = cell.paragraphs[0]
        run = para.runs[0] if para.runs else para.add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor.from_string(header_fg)
        run.font.size = Pt(10)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for r_idx, row_data in enumerate(rows):
        data_row = table.rows[r_idx + 1]
        for c_idx, cell_val in enumerate(row_data):
            cell = data_row.cells[c_idx]
            cell.text = str(cell_val)
            if r_idx % 2 == 0:
                set_cell_bg(cell, "D6E4F0")
            para = cell.paragraphs[0]
            run = para.runs[0] if para.runs else para.add_run(str(cell_val))
            run.font.size = Pt(9)
    return table


def add_heading(doc: Document, text: str, level: int):
    p = doc.add_heading(text, level=level)
    run = p.runs[0] if p.runs else p.add_run(text)
    if level == 1:
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    elif level == 2:
        run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    elif level == 3:
        run.font.color.rgb = RGBColor(0x4A, 0x86, 0xC8)
    return p


def add_body(doc: Document, text: str, bold: bool = False, italic: bool = False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(6)
    return p


def add_bullet(doc: Document, text: str, level: int = 0):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    run.font.size = Pt(10)
    return p


def add_page_break(doc: Document):
    doc.add_page_break()


def add_info_box(doc: Document, title: str, content: str):
    """新增資訊框（表格式）"""
    table = doc.add_table(rows=2, cols=1)
    table.style = "Table Grid"
    title_cell = table.rows[0].cells[0]
    title_cell.text = title
    set_cell_bg(title_cell, "2E75B6")
    tp = title_cell.paragraphs[0]
    tr = tp.runs[0] if tp.runs else tp.add_run(title)
    tr.bold = True
    tr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tr.font.size = Pt(11)
    content_cell = table.rows[1].cells[0]
    content_cell.text = content
    set_cell_bg(content_cell, "EBF3FB")
    cp = content_cell.paragraphs[0]
    if cp.runs:
        cp.runs[0].font.size = Pt(10)
    doc.add_paragraph()


def build_cover_page(doc: Document):
    for _ in range(4):
        doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("CEO 財報分析高階決策報告")
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    doc.add_paragraph()
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = subtitle.add_run("涵蓋企業：TSMC・Walmart・Amazon・Novartis・Nestle・Intel")
    run2.font.size = Pt(14)
    run2.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    doc.add_paragraph()
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run3 = date_p.add_run(f"報告生成日期：{datetime.now().strftime('%Y年%m月%d日')}")
    run3.font.size = Pt(12)
    run3.italic = True
    doc.add_paragraph()
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run4 = note.add_run("本報告依據課程上課資料自動生成，供高階主管決策參考之用")
    run4.font.size = Pt(10)
    run4.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
    add_page_break(doc)


def build_chapter_1(doc: Document, extracted: dict):
    add_heading(doc, "第一章：財報分析方法論（CEO視角）", 1)
    add_body(doc,
        "財務報表是企業最誠實的鏡子。對於CEO而言，閱讀財報不僅是了解企業過去的成績，"
        "更是預判未來走向、制定戰略決策的核心工具。本章從CEO視角出發，系統性地介紹如何"
        "透過財報分析獲取決策洞見，並將財務數字轉化為可執行的戰略行動。")
    add_heading(doc, "1.1 財報的戰略意義", 2)
    add_body(doc, "在競爭激烈的商業環境中，財報是企業與利害關係人溝通的主要語言。CEO需要從戰略高度理解財報的多重功能：")
    for b in [
        "🎯 戰略對標工具：透過比較競爭對手財報，識別相對優勢與差距，指引資源配置",
        "📊 績效評估基準：以財務指標衡量戰略執行成效，及時調整方向",
        "💰 資本市場溝通：財報是企業對投資人承諾的具體呈現，影響資本成本",
        "⚠️ 風險預警系統：財報異常訊號往往先於危機出現，是早期預警的關鍵",
        "🔄 決策回饋循環：將戰略決策的財務後果量化，形成持續改進的閉環",
    ]:
        add_bullet(doc, b)
    lecture_key = "0307_長庚大學 AI 精鍊班_財報基礎觀念複習_課前自習講義.pdf"
    if lecture_key in extracted:
        content = first_n_chars(extracted[lecture_key]["content"], 600)
        if not content.startswith("("):
            add_info_box(doc, "📚 課程資料萃取：財報基礎觀念（長庚大學AI精鍊班講義）", content)
    add_heading(doc, "1.2 三大財報解讀框架", 2)
    add_styled_table(doc,
        ["財報類型", "核心問題", "CEO關注重點", "決策應用"],
        [
            ["損益表 (P&L)", "賺了多少錢？", "營收成長、毛利率、EBITDA", "定價策略、成本控制、業務擴張"],
            ["資產負債表 (B/S)", "企業有多少資產？欠多少錢？", "負債比率、流動比率、股東權益", "資本結構、槓桿管理、M&A評估"],
            ["現金流量表 (C/F)", "現金從哪來、往哪去？", "自由現金流、資本支出、營運現金", "投資優先序、股利政策、流動性管理"],
        ])
    doc.add_paragraph()
    add_heading(doc, "1.3 CEO如何從財報中提取決策信號", 2)
    for title_l, desc_l in [
        ("第一層：數字層", "直接閱讀各項財務數字，了解規模大小與基本比例"),
        ("第二層：趨勢層", "比較歷年數字，識別成長/衰退趨勢與轉折點"),
        ("第三層：對比層", "與競爭對手、行業均值對比，評估相對競爭力"),
        ("第四層：關聯層", "分析三大財報數字間的邏輯關聯，發現異常或矛盾"),
        ("第五層：戰略層", "將財務表現與戰略意圖連結，評估戰略執行品質"),
    ]:
        p = doc.add_paragraph()
        r1 = p.add_run(f"► {title_l}：")
        r1.bold = True
        r1.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
        r1.font.size = Pt(10)
        r2 = p.add_run(desc_l)
        r2.font.size = Pt(10)
    add_heading(doc, "1.4 關鍵財務指標與其決策意涵", 2)
    add_styled_table(doc,
        ["指標類別", "指標名稱", "計算公式", "CEO決策意涵", "警示閾值"],
        [
            ["獲利能力", "毛利率", "毛利 ÷ 營收", "產品定價力與成本效率", "< 行業均值5%以上需關注"],
            ["獲利能力", "ROE（股東權益報酬率）", "淨利 ÷ 股東權益", "為股東創造價值的能力", "< 10% 長期不可持續"],
            ["獲利能力", "EBITDA Margin", "EBITDA ÷ 營收", "核心業務獲利能力", "下滑超過3個百分點需警覺"],
            ["流動性", "流動比率", "流動資產 ÷ 流動負債", "短期償債能力", "< 1.0 有流動性危機風險"],
            ["槓桿", "淨負債/EBITDA", "淨負債 ÷ EBITDA", "財務槓桿安全度", "> 4x 財務壓力顯著"],
            ["效率", "存貨周轉天數", "存貨 ÷ (COGS÷365)", "供應鏈效率", "較行業均值高50%以上需改善"],
            ["成長", "營收CAGR", "複合年均成長率", "業務擴張速度", "連續3年低於行業成長率需重視"],
            ["現金", "自由現金流FCF", "營業現金流 − 資本支出", "企業真實盈利質量", "FCF/淨利 < 0.8 需關注"],
        ])
    doc.add_paragraph()
    for story_key in ["財報就像一本故事書Part2_1.pdf", "財報就像一本故事書Part2_2.pdf"]:
        if story_key in extracted:
            content = first_n_chars(extracted[story_key]["content"], 500)
            if not content.startswith("("):
                add_info_box(doc, f"📖 課程資料：{story_key}", content)
    add_page_break(doc)


def build_chapter_2(doc: Document, extracted: dict):
    add_heading(doc, "第二章：本課程財報資料分析總覽", 1)
    add_body(doc, "本課程涵蓋六家全球頂尖企業的財務報表與經營策略分析，涵蓋半導體、零售、製藥、食品及科技電商等多個行業，提供跨產業的CEO決策視角。")
    add_heading(doc, "2.1 分析企業概覽", 2)
    add_styled_table(doc,
        ["企業", "行業", "主要分析面向", "資料來源", "分析重點"],
        [
            ["台積電 (TSMC)", "半導體製造", "Q4 2025合併財報", "TSMC 2025Q4財報(中文版)", "毛利率、資本支出、技術護城河"],
            ["Walmart", "全球零售", "1994/2024/2026年報", "三份年報對照", "30年財務演變、數位轉型財務影響"],
            ["Amazon", "電商/雲端", "2025年報", "Amazon 2025 Annual Report", "AWS獲利貢獻、電商vs雲端比例"],
            ["Novartis", "製藥", "ESG整合財報", "Novartis ESG簡報", "ESG KPI與財務表現連結"],
            ["Nestle", "食品飲料", "ESG整合財報", "Nestle ESG簡報", "永續投資與品牌價值"],
            ["Intel", "半導體設計", "轉型期財報", "陳立武轉型案例", "財務重組、轉型成本"],
        ])
    doc.add_paragraph()
    add_heading(doc, "2.2 課程資料清單與萃取狀態", 2)
    rows2 = []
    for fname, info in sorted(extracted.items()):
        ftype = "PDF" if info["type"] == "pdf" else "Word"
        content = info["content"]
        if content.startswith("(此PDF為影像型"):
            status, preview = "⚠️ 影像型PDF", "無可萃取文字"
        elif content.startswith("("):
            status, preview = "❌ 錯誤", content[:50]
        else:
            status = "✅ 成功"
            preview = content[:60].replace("\n", " ") + "…"
        rows2.append([fname[:35], ftype, status, preview])
    add_styled_table(doc, ["檔案名稱", "類型", "萃取狀態", "主要內容"], rows2, header_bg="2E75B6")
    doc.add_paragraph()
    add_page_break(doc)


def build_chapter_3(doc: Document, extracted: dict):
    add_heading(doc, "第三章：台積電（TSMC）財報深度分析", 1)
    tsmc_key = "TSMC 2025Q4 Consolidated Financial Statements_C.pdf"
    if tsmc_key in extracted:
        content = extracted[tsmc_key]["content"]
        if not content.startswith("("):
            add_info_box(doc, "📊 TSMC 2025Q4 合併財務報表（原文摘錄）", first_n_chars(content, 1000))
    add_heading(doc, "3.1 台積電財務表現亮點", 2)
    add_styled_table(doc,
        ["財務指標", "2025 Q4表現", "行業意義", "CEO決策啟示"],
        [
            ["毛利率", "~57-60%", "半導體代工業最高水準之一", "定價能力強，技術壁壘高"],
            ["資本支出", "年度約$350-400億美元", "全球最大單一資本支出計畫", "需評估報酬率與現金流衝擊"],
            ["先進製程比例", "N3/N5佔收入約60%+", "技術領先優勢持續", "保持研發投入不可動搖"],
            ["AI相關營收", "快速成長", "CoWoS等先進封裝需求爆增", "AI基礎設施投資週期長"],
            ["地理分散", "台、美、日、歐建廠", "地緣政治風險對沖", "資本效率 vs 政治保險的取捨"],
            ["研發費用率", "約8-10%", "持續創新的護城河", "R&D是不可削減的戰略投資"],
        ])
    doc.add_paragraph()
    add_heading(doc, "3.2 TSMC vs Intel：財務競爭力比較", 2)
    add_styled_table(doc,
        ["比較維度", "TSMC (台積電)", "Intel (英特爾)", "差距分析"],
        [
            ["商業模式", "純晶圓代工（Foundry）", "IDM整合製造（設計+製造）", "TSMC專注核心，效率更高"],
            ["毛利率", "~57-60%", "~40-45%（近年下滑）", "TSMC高15-20個百分點"],
            ["資本回報率ROE", "~25-30%", "~5-10%（轉型期低迷）", "TSMC資本效率遠優於Intel"],
            ["財務健康度", "高現金流、低負債", "負債增加、現金流承壓", "TSMC財務彈性更大"],
        ])
    doc.add_paragraph()
    add_page_break(doc)


def build_chapter_4(doc: Document, extracted: dict):
    add_heading(doc, "第四章：全球零售巨頭比較分析（Walmart vs Amazon）", 1)
    add_heading(doc, "4.1 Walmart 30年財務演變（1994 → 2024 → 2026）", 2)
    add_styled_table(doc,
        ["財務指標", "1994年", "2024年", "2026年（預測/實際）", "30年CAGR"],
        [
            ["年營收", "~$675億", "~$6,480億", "~$6,800億+", "~8-9%"],
            ["淨利潤", "~$21億", "~$157億", "~$200億+", "~8%"],
            ["電商佔比", "<1%", "~15%+", "~20%+", "數位轉型加速"],
            ["毛利率", "~20%", "~24%", "~25%", "持續小幅改善"],
        ])
    doc.add_paragraph()
    for wm_key, wm_label in [
        ("1994 Walmart Financial Report.pdf", "1994 Walmart Annual Report"),
        ("2024 Walmart Financial Report.pdf", "2024 Walmart Financial Report"),
        ("Walmart 2026 Annual Report.pdf", "Walmart 2026 Annual Report"),
    ]:
        if wm_key in extracted:
            content = first_n_chars(extracted[wm_key]["content"], 600)
            if not content.startswith("("):
                add_info_box(doc, f"📊 {wm_label} — 原文摘錄", content)
    add_heading(doc, "4.2 Amazon 2025年財報分析", 2)
    amz_key = "Amazon-2025-Annual-Report.pdf"
    if amz_key in extracted:
        content = first_n_chars(extracted[amz_key]["content"], 800)
        if not content.startswith("("):
            add_info_box(doc, "📋 Amazon 2025 Annual Report — 原文摘錄", content)
    add_page_break(doc)


def build_chapter_5(doc: Document, extracted: dict):
    add_heading(doc, "第五章：ESG與財務表現", 1)
    for key, label in [
        ("Novartis_ESG與財報_高階主管簡報.pdf", "💊 Novartis ESG與財報高階主管簡報"),
        ("Nestle_ESG與財報_高階主管簡報.pdf", "☕ Nestle ESG與財報高階主管簡報"),
    ]:
        if key in extracted:
            content = first_n_chars(extracted[key]["content"], 1000)
            if not content.startswith("("):
                add_info_box(doc, f"{label} — 原文摘錄", content)
    add_page_break(doc)


def build_chapter_6(doc: Document, extracted: dict):
    add_heading(doc, "第六章：企業轉型案例研究與策略哲學", 1)
    intel_key = "20260516陳立武turnaround Intel.pdf"
    if intel_key in extracted:
        content = first_n_chars(extracted[intel_key]["content"], 1000)
        if not content.startswith("("):
            add_info_box(doc, "🔄 陳立武 Intel Turnaround — 課程資料摘錄", content)
    add_styled_table(doc,
        ["轉型面向", "Intel面臨的挑戰", "轉型措施", "財務影響"],
        [
            ["技術競爭力", "製程技術落後TSMC/Samsung", "IDM 2.0策略，對外開放代工", "短期成本上升，長期商業化"],
            ["成本結構", "人力成本過高、效率低下", "裁員1.5萬人、工廠關閉/出售", "一次性重組費用，長期降本"],
            ["業務聚焦", "業務過於分散", "出售非核心業務（如Mobileye）", "優化資本配置，聚焦核心"],
        ])
    doc.add_paragraph()
    for key, label in [
        ("孫子兵法原文與白話翻譯.docx", "⚔️ 孫子兵法原文與白話翻譯"),
        ("如何立於不敗之地.docx", "🛡️ 如何立於不敗之地"),
        ("20260523Porter與創造共享價值.docx", "📐 Porter與創造共享價值"),
        ("Steve Jobs Stanford Address.docx", "🎓 Steve Jobs Stanford Address"),
    ]:
        if key in extracted:
            content = first_n_chars(extracted[key]["content"], 600)
            if not content.startswith("("):
                add_info_box(doc, f"{label} — 課程資料摘錄", content)
    add_page_break(doc)


def build_chapter_7(doc: Document, extracted: dict):
    add_heading(doc, "第七章：CEO高階決策建議與戰略指引", 1)
    add_heading(doc, "7.1 CEO關鍵決策指標儀表板", 2)
    add_styled_table(doc,
        ["類別", "指標", "監控頻率", "警示條件", "行動觸發"],
        [
            ["成長", "營收YoY成長率", "月度", "< 行業均值 -5%", "審查市場策略"],
            ["獲利", "毛利率趨勢", "月度", "環比下滑 >1%", "成本結構審查"],
            ["獲利", "EBITDA Margin", "季度", "< 年度目標 -2%", "全面成本優化"],
            ["現金", "自由現金流", "月度", "連續3個月負數", "資本支出優先序重排"],
            ["槓桿", "淨負債/EBITDA", "季度", "> 3.5x", "降槓桿計畫啟動"],
            ["效率", "存貨周轉天數", "月度", "較均值高 >20%", "庫存優化專案"],
            ["人才", "關鍵人才留任率", "半年", "< 90%", "薪酬&文化診斷"],
        ])
    doc.add_paragraph()
    add_heading(doc, "7.2 風險管理視角", 2)
    add_styled_table(doc,
        ["風險類別", "具體風險", "財務影響", "緩解策略"],
        [
            ["流動性風險", "現金流驟降、融資渠道收窄", "無法支付運營支出", "維持12-18個月現金儲備"],
            ["地緣政治風險", "貿易制裁、市場准入限制", "市場份額損失、成本上升", "供應鏈多元化、市場分散"],
            ["科技顛覆風險", "新技術使現有產品過時", "收入快速下滑", "持續R&D、開放式創新"],
            ["ESG風險", "碳稅、極端氣候供應中斷", "成本上升、資產擱置", "低碳轉型、供應鏈韌性"],
        ])
    doc.add_paragraph()
    add_page_break(doc)


def build_appendix(doc: Document, extracted: dict):
    add_heading(doc, "附錄：課程資料完整摘錄", 1)
    add_body(doc, "以下為本課程所有資料來源的完整文字萃取（影像型PDF除外），供參考查閱。")
    for fname, info in sorted(extracted.items()):
        content = info["content"]
        if content.startswith("(") or len(content.strip()) < 50:
            continue
        add_heading(doc, fname, 3)
        add_body(doc, first_n_chars(content, 400))


def main():
    print("=" * 65)
    print("  CEO財報分析高階決策報告 — 自動生成器（含OCR）")
    print("=" * 65)
    if _TESSERACT_AVAILABLE:
        print(f"  🔍 Tesseract OCR 已啟用 | 語言：{_OCR_LANG}")
    else:
        print("  ⚠️  Tesseract OCR 未安裝 — 影像型PDF將顯示安裝指引")
        print("     下載：https://github.com/UB-Mannheim/tesseract/wiki")

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"✅ 輸出目錄就緒：{REPORT_DIR}")

    print(f"\n📂 開始萃取來源目錄：{SOURCE_DIR}")
    extracted = extract_all_files(SOURCE_DIR)
    print(f"✅ 共萃取 {len(extracted)} 個檔案")

    print("\n📝 開始生成Word報告...")
    doc = Document()
    from docx.oxml import OxmlElement
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.5)
    doc.styles["Normal"].font.name = "微軟正黑體"
    doc.styles["Normal"].font.size = Pt(11)

    print("  📌 生成封面頁...")
    build_cover_page(doc)
    print("  📌 生成第一章：財報分析方法論...")
    build_chapter_1(doc, extracted)
    print("  📌 生成第二章：課程財報資料分析...")
    build_chapter_2(doc, extracted)
    print("  📌 生成第三章：台積電財報深度分析...")
    build_chapter_3(doc, extracted)
    print("  📌 生成第四章：Walmart vs Amazon...")
    build_chapter_4(doc, extracted)
    print("  📌 生成第五章：ESG與財務表現...")
    build_chapter_5(doc, extracted)
    print("  📌 生成第六章：企業轉型案例...")
    build_chapter_6(doc, extracted)
    print("  📌 生成第七章：CEO決策建議...")
    build_chapter_7(doc, extracted)
    print("  📌 生成附錄...")
    build_appendix(doc, extracted)

    save_path = REPORT_PATH
    if save_path.exists():
        try:
            with open(save_path, "a"):
                pass
        except PermissionError:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = REPORT_DIR / f"CEO財報分析決策報告_{ts}.docx"
            print(f"\n⚠️  原始檔案已被開啟（Word尚未關閉），改存為：{save_path.name}")

    doc.save(str(save_path))
    print(f"\n{'=' * 65}")
    print(f"✅ 報告生成完成！")
    print(f"📄 報告路徑：{save_path}")
    size_kb = save_path.stat().st_size // 1024
    print(f"📦 檔案大小：{size_kb} KB")
    print(f"{'=' * 65}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 發生錯誤：{e}")
        traceback.print_exc()
        sys.exit(1)
