#!/usr/bin/env python3
# 提出用一式PDFの結合・圧縮テンプレート
# 構成: 事業計画書(本体) + 代表者プロフィール + FC実績資料 + エビデンス(見積/火災保険)
# 使い方: 各加盟店のPDFパスを下記に差し替えて実行。
import subprocess, sys
from pypdf import PdfWriter

# === 差し替え項目 ===
PARTS = [
    "事業計画書.pdf",        # build_docx.js -> soffice で生成した本体
    "代表者プロフィール.pdf",  # オーナーのプロフィール(pptx等をPDF化)
    "FC実績.pdf",            # monet FC本部の実績資料
    "エビデンス.pdf",        # 見積書・火災保険等の根拠資料
]
OUT = "提出用一式.pdf"
# ===================

w = PdfWriter()
for p in PARTS:
    try:
        w.append(p)
    except FileNotFoundError:
        print(f"skip (not found): {p}", file=sys.stderr)
with open("_merged_raw.pdf", "wb") as f:
    w.write(f)

# ghostscriptで圧縮 (/ebook)
subprocess.run([
    "gs","-sDEVICE=pdfwrite","-dCompatibilityLevel=1.4",
    "-dPDFSETTINGS=/ebook","-dNOPAUSE","-dQUIET","-dBATCH",
    f"-sOutputFile={OUT}","_merged_raw.pdf"
], check=True)
print("done ->", OUT)
