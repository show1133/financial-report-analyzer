# 📊 財務報表分析報告自動生成器

> CEO財報分析高階決策報告 — 支援PDF文字萃取、OCR影像型PDF識別、多企業財務深度分析

[![Python](https://img.shields.io/badge/Python-3.8+-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 📋 專案說明

本工具讀取課程上課資料（PDF / Word），自動分析財務報表內容，並生成一份結構完整的 CEO 高階決策報告（Word格式）。

**涵蓋企業：**
- 🏭 台積電 (TSMC) — 半導體製造
- 🛒 Walmart — 全球零售（1994 / 2024 / 2026 年報）
- 📦 Amazon — 電商/雲端（2025年報）
- 💊 Novartis — 製藥 ESG
- ☕ Nestle — 食品飲料 ESG
- 💻 Intel — 半導體轉型

---

## 🚀 快速開始

### 1. 安裝 Python 套件

腳本啟動時會自動安裝以下套件，也可手動安裝：

```bash
pip install pdfplumber pymupdf python-docx pdf2image pytesseract Pillow
```

### 2. 安裝 Tesseract OCR（選用，用於影像型PDF）

1. 下載：https://github.com/UB-Mannheim/tesseract/wiki
2. 安裝時勾選 **Additional language data**：
   - ✅ `chi_tra` — 繁體中文
   - ✅ `chi_sim` — 簡體中文（選用）
3. 預設安裝路徑：`C:\Program Files\Tesseract-OCR\tesseract.exe`

> ⚠️ 若未安裝 Tesseract，文字型 PDF 仍可正常萃取；影像型 PDF 會顯示安裝提示。

### 3. 執行報告生成器

```bash
python report_generator.py
```

**輸出位置：** `C:\Python\99財務管理\report\CEO財報分析決策報告.docx`

---

## 📂 專案結構

```
financial-report-analyzer/
├── report_generator.py          # 主程式：CEO多企業財報分析報告
├── generate_walmart_report.py   # Walmart 跨時代財報分析（孫子兵法框架）
└── README.md
```

**來源資料目錄（本機）：**
```
C:\Python\99財務管理\上課資料\
├── TSMC 2025Q4 Consolidated Financial Statements_C.pdf
├── Walmart 2026 Annual Report.pdf
├── Amazon-2025-Annual-Report.pdf
├── Novartis_ESG與財報_高階主管簡報.pdf
├── Nestle_ESG與財報_高階主管簡報.pdf
├── 20260516陳立武turnaround Intel.pdf
├── 孫子兵法原文與白話翻譯.docx
└── ... (共27份資料)
```

---

## 📖 報告結構（report_generator.py）

| 章節 | 內容 |
|------|------|
| 第一章 | 財報分析方法論（CEO視角）— 三大財報框架、五層深讀法、8大決策指標 |
| 第二章 | 本課程財報資料分析總覽 |
| 第三章 | 台積電（TSMC）財報深度分析 vs Intel |
| 第四章 | 全球零售巨頭比較（Walmart vs Amazon）|
| 第五章 | ESG與財務表現（Novartis / Nestle）|
| 第六章 | 企業轉型案例研究（Intel、孫子兵法、Porter）|
| 第七章 | CEO高階決策建議 + 風險管理框架 |
| 附錄 | 課程資料完整文字摘錄 |

---

## 🔍 PDF 三層萃取策略

```
第一層：pdfplumber   → 一般文字型 PDF
   ↓ 若萃取不足
第二層：pymupdf      → 補強萃取嵌入文字
   ↓ 若仍不足
第三層：pytesseract  → OCR 識別影像型 PDF（需安裝 Tesseract）
```

---

## 📊 Walmart 跨時代分析（generate_walmart_report.py）

以 **孫子兵法「道・天・地・將・法」五構面框架** 分析 Walmart 30年財務演變：

| 構面 | 企業分析維度 |
|------|-------------|
| 道（The Way）| 企業使命、天天低價哲學、股利政策 |
| 天（Heaven）| 時代背景、稅務環境、外部競爭 |
| 地（Earth）| 資產結構、市場佈局、護城河 |
| 將（General）| ROE杜邦分析、庫藏股策略、併購 |
| 法（Method）| 現金流、運營效率、資本支出 |

```bash
python generate_walmart_report.py
# 輸出：C:\Python\Walmart_跨時代財報分析報告_王道智謀五構面.docx
```

---

## ⚙️ 環境需求

- **Python** 3.8+
- **Windows** 10/11（路徑設定為 Windows 格式，可修改）
- **Tesseract OCR**（選用，影像型PDF必須）
- **Poppler**（pdf2image 依賴，Windows需手動安裝）
  - 下載：https://github.com/oschwartz10612/poppler-windows/releases
  - 解壓後將 `bin/` 路徑加入系統 PATH

---

## 📝 License

MIT License — 供學術研究與教學使用，不構成投資建議。
