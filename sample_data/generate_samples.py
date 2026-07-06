# -*- coding: utf-8 -*-
"""
Generates fully synthetic demo PDFs used to exercise the app in this repo:
  - meisai_sample.pdf    : 注文明細票 with real scannable CODE128 barcodes
  - delivery_sample.pdf  : matching 納品書 (no barcodes) + one intentionally
                           unmatched part, to show the "unmatched" status too
  - po_sample.pdf        : 発注書 (PO) for the demo vendor in po_to_csv.py

All company names, part numbers, and figures here are made up for
demonstration only. Regenerating requires two extra dev-only packages
not needed to run the app itself:

    pip install reportlab python-barcode

Run: python sample_data/generate_samples.py
"""
import io
import os

import barcode
from barcode.writer import ImageWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image

# Built-in CJK font (no external font file needed) so the Japanese labels
# actually render as glyphs instead of blank boxes — otherwise Azure OCR
# can't read them back out of the PDF.
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
JP_FONT = "HeiseiKakuGo-W5"

OUT_DIR = os.path.dirname(__file__)
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="JP", fontName=JP_FONT, fontSize=11, leading=14))
styles.add(ParagraphStyle(name="JPTitle", fontName=JP_FONT, fontSize=16, leading=20))


def _barcode_image(value, width=1.3 * inch, height=0.35 * inch):
    buf = io.BytesIO()
    barcode.Code128(value, writer=ImageWriter()).write(
        buf, options={"write_text": False, "module_height": 8.0, "quiet_zone": 1.0}
    )
    buf.seek(0)
    return Image(buf, width=width, height=height)


TABLE_STYLE = TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
    ("FONTNAME", (0, 0), (-1, -1), JP_FONT),
    ("FONTSIZE", (0, 0), (-1, -1), 9),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("ALIGN", (2, 0), (-1, -1), "CENTER"),
])


def make_meisai():
    doc = SimpleDocTemplate(os.path.join(OUT_DIR, "meisai_sample.pdf"), pagesize=A4)
    rows = [
        ("10010-93652", "シャフトA", "5", "250", "10010001-001"),
        ("10020-84521", "ブラケットB", "3", "480", "10010002-002"),
        ("10030-77410", "カバーC", "10", "120", "10010003-003"),
        ("10040-65210", "プレートD", "2", "990", "10010004-004"),
        ("10050-51230", "パイプE", "5", "250", "10010005-005"),
    ]
    data = [["品番", "品名", "数量", "単価", "バーコード"]]
    data += [[p, n, q, t, _barcode_image(bc)] for p, n, q, t, bc in rows]
    table = Table(data, colWidths=[1.5 * inch, 1.3 * inch, 0.7 * inch, 0.8 * inch, 1.5 * inch])
    table.setStyle(TABLE_STYLE)

    story = [
        Paragraph("注文明細票 (サンプル / Demo data — not a real order)", styles["JPTitle"]),
        Paragraph("伝票番号: DEMO-MEISAI-0001", styles["JP"]),
        Spacer(1, 12),
        table,
    ]
    doc.build(story)
    print("wrote meisai_sample.pdf")


def make_delivery():
    doc = SimpleDocTemplate(os.path.join(OUT_DIR, "delivery_sample.pdf"), pagesize=A4)
    # same 5 parts as meisai (should all match), plus 1 part not present in
    # meisai at all (demonstrates the "unmatched" status in the output).
    rows = [
        ("10010-93652", "シャフトA", "5", "250"),
        ("10020-84521", "ブラケットB", "3", "480"),
        ("10030-77410", "カバーC", "10", "120"),
        ("10040-65210", "プレートD", "2", "990"),
        ("10050-51230", "パイプE", "5", "250"),
        ("10090-11111", "不明部品X", "1", "300"),
    ]
    data = [["品番", "品名", "数量", "単価"]] + [list(r) for r in rows]
    table = Table(data, colWidths=[1.5 * inch, 1.5 * inch, 0.8 * inch, 0.9 * inch])
    table.setStyle(TABLE_STYLE)

    story = [
        Paragraph("納品書 (サンプル / Demo data — not a real delivery)", styles["JPTitle"]),
        Paragraph("伝票番号: DEMO-DELIVERY-0001", styles["JP"]),
        Spacer(1, 12),
        table,
    ]
    doc.build(story)
    print("wrote delivery_sample.pdf")


def make_po():
    doc = SimpleDocTemplate(os.path.join(OUT_DIR, "po_sample.pdf"), pagesize=A4)
    rows = [
        ("132001", "AB1234-01", "NK(C) 20Y-60-61151-01 COVER", "10", "07/15", "1200", ""),
        ("132002", "CD5678-02", "XC(D) 30Z-70-72222-02 BRACKET", "5", "07/20", "800", ""),
        ("132003", "EF9012-03", "PLATE", "20", "07/25", "300", ""),
    ]
    total = sum(int(q) * int(t) for _, _, _, q, _, t, _ in rows)
    data = [["注文番号", "品番", "品名", "数量", "納期", "単価", "備考"]] + [list(r) for r in rows]
    table = Table(data, colWidths=[0.7 * inch, 0.9 * inch, 2.0 * inch, 0.6 * inch, 0.6 * inch, 0.6 * inch, 0.6 * inch])
    table.setStyle(TABLE_STYLE)

    story = [
        Paragraph("注文書 (サンプル / Demo data — fictional company)", styles["JPTitle"]),
        Paragraph("発行日: 2026年07月01日", styles["JP"]),
        Paragraph("有限会社 Sample Vendor 御中", styles["JP"]),  # 発注先 (excluded from 顧客)
        Paragraph("Demo Customer 株式会社　Osaka", styles["JP"]),  # 顧客 (should get normalized)
        Spacer(1, 12),
        table,
        Spacer(1, 12),
        Paragraph(f"税抜金額: {total:,}", styles["JP"]),
    ]
    doc.build(story)
    print(f"wrote po_sample.pdf (checksum total = {total:,})")


if __name__ == "__main__":
    make_meisai()
    make_delivery()
    make_po()
