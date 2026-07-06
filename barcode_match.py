# -*- coding: utf-8 -*-
"""
barcode_match.py
================
ตรรกะจับคู่ barcode ระหว่าง 注文明細票 (meisai, มี barcode) กับ 納品書 (delivery, ไม่มี barcode)

แนวคิด (พิสูจน์กับไฟล์จริงแล้ว):
  1. meisai อ่านด้วย Azure prebuilt-layout + ฟีเจอร์ BARCODES
     → ได้ทั้งตาราง (Part No.) และ barcode พร้อมพิกัด ในระบบพิกัดเดียวกัน
     → จับ barcode เข้าบรรทัดด้วยพิกัด Y (แม่นมาก, ห่าง < 0.02 นิ้ว)
  2. จับคู่เอกสาร meisai ↔ delivery ด้วย "ชุด Part No. ที่ทับกัน" (ลำดับสแกนสลับได้)
  3. แมตช์บรรทัดในคู่ด้วย Part No. (+ จำนวน/ราคา/ยอด ช่วยตัดสินตัวซ้ำ)
  4. สถานะหลายระดับ: ✅ จับคู่ได้ / ⚠ ต้องตรวจ / barcodeなし / ❌ ไม่จับคู่

โมดูลนี้ไม่พึ่ง GUI — รับ normalize_fn / pick_table_fn เข้ามาเพื่อเลี่ยง circular import
"""

import io
import re
import time

from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import (
    HttpResponseError, ServiceRequestError, ServiceResponseError,
)
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import (
    AnalyzeDocumentRequest,
    DocumentAnalysisFeature,
)

# HTTP status ที่ retry ได้ (ชั่วคราว): 408 timeout, 429 too many, 5xx
_RETRYABLE_STATUS = {408, 429, 500, 502, 503, 504}


def call_with_retry(fn, *args, attempts=4, base_delay=2.0, log_fn=None, **kwargs):
    """
    เรียก fn พร้อม retry แบบ exponential backoff เฉพาะ error ชั่วคราว (429/5xx/เน็ตหลุด)
    error ถาวร (401/403/400) จะโยนทันทีไม่ retry
    """
    last = None
    for i in range(attempts):
        try:
            return fn(*args, **kwargs)
        except HttpResponseError as e:
            code = getattr(getattr(e, "response", None), "status_code", None)
            if code not in _RETRYABLE_STATUS or i == attempts - 1:
                raise
            last = e
        except (ServiceRequestError, ServiceResponseError, TimeoutError, ConnectionError) as e:
            if i == attempts - 1:
                raise
            last = e
        wait = base_delay * (2 ** i)
        if log_fn:
            log_fn(f"      ⏳ ネットワーク再試行 {i + 1}/{attempts} — {wait:.0f}s 待機 ({type(last).__name__})\n")
        time.sleep(wait)
    if last:
        raise last

# barcode บนเอกสารชุดนี้เป็น CODE128 รูปแบบ 8 หลัก - 3 หลัก
BARCODE_RE = re.compile(r"\d{8}-\d{3}")

# ── keyword หาบทบาทคอลัมน์ ──
# Azure OCR เอกสารสแกนมักแทรกช่องว่างกลางคันจิ ("単 価") หรืออ่านได้บางตัว ("数" จาก 数量)
# จึง (1) ตัดช่องว่างก่อนเทียบ (2) ใช้ตัวอักษรเฉพาะตัว (価, 額, 位) เพื่อเลี่ยงชนกัน (単価/単位 ต่างกันที่ 価/位)
QTY_KEYS = ["数量", "員数", "個数", "数"]
PRICE_KEYS = ["単価", "価"]
AMOUNT_KEYS = ["金額", "額", "小計"]
DESC_KEYS = ["品名", "摘要", "品名寸法", "商品名"]
UNIT_KEYS = ["単位", "位"]
REMARK_KEYS = ["備考", "考"]
NOUKI_KEYS = ["納期", "納入", "出荷", "納品"]
HDR_KEYS = ["品名", "品番", "数量", "単価", "金額", "納期", "数", "価", "額"]


def _hk(h):
    """normalize หัวคอลัมน์ก่อนเทียบ keyword: ตัด whitespace ทั้งหมด"""
    return re.sub(r"\s", "", h or "")


def _find_col(headers, keys):
    for c, h in enumerate(headers):
        hk = _hk(h)
        if any(k in hk for k in keys):
            return c
    return -1


def _find_header_row(grid, n_rows):
    for r in range(min(6, n_rows)):
        rt = _hk(" ".join(grid[r]))
        if any(k in rt for k in HDR_KEYS):
            return r
    return 0

# สถานะการจับคู่
STATUS_OK = "ok"            # ✅ จับคู่ได้แน่นอน
STATUS_REVIEW = "review"    # ⚠ กำกวม ต้องให้คนตรวจ
STATUS_NO_BARCODE = "no_bc" # บรรทัดนี้ไม่มี barcode บน meisai (เช่น เขียนมือ)
STATUS_UNMATCHED = "unmatched"  # ❌ หา Part No. ในคู่ meisai ไม่เจอ


# ──────────────────────────────────────────────────────────────
# Part No. helpers (ย้ายมาไว้ที่นี่เพื่อให้โมดูลทดสอบได้อิสระ)
# ──────────────────────────────────────────────────────────────
def normalize_part_no(part_code: str) -> str:
    """
    Normalize Part No.:
      1. OCR correction: I/T → 1 ใน segment ที่มีตัวเลขปน
      2. ตัด trailing -0 / -00 ออก (เป็น revision suffix)
    """
    code = (part_code or "").strip().upper()

    def fix_segment(seg):
        if re.search(r"\d", seg):
            seg = seg.replace("I", "1").replace("T", "1")
        return seg

    parts = [fix_segment(p) for p in code.split("-")]
    code = "-".join(parts)
    code = re.sub(r"(-0+)+$", "", code)
    return code


def extract_part_no(text: str) -> str:
    """
    ดึง Part No. จากข้อความ เช่น "50010-93652-0 シャフト" → "50010-93652"
    เลือก candidate ที่มีตัวเลขและยาวที่สุด
    """
    pattern = re.compile(r"\b([A-Z0-9]{2,}(?:-[A-Z0-9]+){1,})\b")
    candidates = pattern.findall((text or "").upper())
    valid = [c for c in candidates if re.search(r"\d", c) and len(c) >= 6]
    if not valid:
        return ""
    return normalize_part_no(max(valid, key=len))


def _num_lead(text: str) -> str:
    """ดึงตัวเลขนำหน้า เช่น '5¥250-' → '5', '5 個' → '5', '1,250' → '1250'"""
    s = re.sub(r"[,\s]", "", str(text or ""))
    m = re.search(r"\d+", s)
    return m.group() if m else ""


# ──────────────────────────────────────────────────────────────
# Geometry helpers
# ──────────────────────────────────────────────────────────────
def _poly_y(poly):
    ys = poly[1::2]
    return sum(ys) / len(ys) if ys else 0.0


def _poly_x(poly):
    xs = poly[0::2]
    return sum(xs) / len(xs) if xs else 0.0


def _cell_yband(cell):
    """คืน (page_number, y_min, y_max) ของ cell จาก bounding region แรก"""
    brs = getattr(cell, "bounding_regions", None)
    if not brs:
        return (None, None, None)
    poly = brs[0].polygon
    ys = poly[1::2]
    if not ys:
        return (None, None, None)
    return (brs[0].page_number, min(ys), max(ys))


# ──────────────────────────────────────────────────────────────
# Azure call (meisai): layout + barcodes
# ──────────────────────────────────────────────────────────────
def analyze_layout_with_barcodes(endpoint: str, key: str, pdf_path: str, log_fn=None):
    """อ่าน meisai ด้วย prebuilt-layout + ฟีเจอร์ barcodes (ถูกกว่า prebuilt-invoice) + retry"""
    client = DocumentIntelligenceClient(endpoint, AzureKeyCredential(key))
    with open(pdf_path, "rb") as f:
        data = f.read()

    def _run():
        poller = client.begin_analyze_document(
            "prebuilt-layout",
            AnalyzeDocumentRequest(bytes_source=data),
            features=[DocumentAnalysisFeature.BARCODES],
        )
        return poller.result()

    return call_with_retry(_run, log_fn=log_fn)


# ──────────────────────────────────────────────────────────────
# Decode barcodes จากรูปจริงด้วย pyzbar (recall สูงกว่า Azure บนหน้าที่แน่น/จาง)
# ──────────────────────────────────────────────────────────────
def decode_barcodes_pyzbar(pdf_path, dpi=300):
    """
    คืน { page_number(1-based): [ {value, y(inch)} ] }
    ถ้าไม่มี pymupdf/pyzbar ติดตั้ง → คืน {} (fallback ไปใช้ Azure อย่างเดียว)
    """
    try:
        import fitz
        from pyzbar.pyzbar import decode as _zbar_decode
        from PIL import Image
    except Exception:
        return {}

    out = {}
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return {}
    try:
        for pno in range(len(doc)):
            page = doc[pno]
            ph_inch = page.rect.height / 72.0  # points → inch
            try:
                pix = page.get_pixmap(dpi=dpi)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
            except Exception:
                continue
            lst = []
            for c in _zbar_decode(img):
                try:
                    val = c.data.decode("ascii", "replace").strip()
                except Exception:
                    continue
                m = BARCODE_RE.search(val)
                if not m:
                    continue
                y_in = (c.rect.top + c.rect.height / 2.0) / pix.height * ph_inch
                lst.append({"value": m.group(), "y": y_in})
            out[pno + 1] = lst
    finally:
        doc.close()
    return out


# ──────────────────────────────────────────────────────────────
# Index meisai: associate each barcode to its table row via Y
# ──────────────────────────────────────────────────────────────
def index_meisai(result, normalize_fn, pick_table_fn, pyzbar_pages=None):
    """
    คืน dict:
      {
        'rows': [ {part_no, barcode, qty, price, amount} ... ]  (ตามลำดับบรรทัด)
        'part_index': { part_no: [idx, ...] },
        'bc_total': int, 'bc_matched': int,
      }
    """
    norm = normalize_fn

    # 1) เก็บ barcode ทุกตัวแยกตามหน้า พร้อมพิกัด (Azure + pyzbar union)
    page_barcodes = {}   # page -> [ {value, y, used} ]
    for p in getattr(result, "pages", None) or []:
        lst = []
        for b in getattr(p, "barcodes", None) or []:
            val = (b.value or "").strip()
            m = BARCODE_RE.search(val)
            if not m:
                continue
            lst.append({"value": m.group(), "y": _poly_y(b.polygon), "used": False})
        page_barcodes[p.page_number] = lst

    # merge pyzbar (เติมตัวที่ Azure อ่านไม่ได้) — dedupe ด้วย value + y ใกล้กัน
    if pyzbar_pages:
        for pg, items in pyzbar_pages.items():
            existing = page_barcodes.setdefault(pg, [])
            for it in items:
                dup = any(e["value"] == it["value"] and abs(e["y"] - it["y"]) < 0.15
                          for e in existing)
                if not dup:
                    existing.append({"value": it["value"], "y": it["y"], "used": False})

    bc_total = sum(len(v) for v in page_barcodes.values())

    # 2) เลือกตารางรายการหลัก
    _, table = pick_table_fn(result)
    rows = []
    part_index = {}
    if not table:
        return {"rows": rows, "part_index": part_index, "bc_total": bc_total, "bc_matched": 0}

    n_rows, n_cols = table.row_count, table.column_count
    grid = [["" for _ in range(n_cols)] for _ in range(n_rows)]
    row_band = {}  # row_index -> (page, ymin, ymax)

    for cell in table.cells:
        r, c = cell.row_index, cell.column_index
        rs = max(getattr(cell, "row_span", 1) or 1, 1)
        cs = max(getattr(cell, "column_span", 1) or 1, 1)
        val = norm(cell.content or "")
        for dr in range(rs):
            for dc in range(cs):
                rr, cc = r + dr, c + dc
                if rr < n_rows and cc < n_cols and dr == 0 and dc == 0:
                    grid[rr][cc] = val
        pg, ymin, ymax = _cell_yband(cell)
        if pg is not None:
            for dr in range(rs):
                rr = r + dr
                if rr in row_band:
                    pg0, a, b2 = row_band[rr]
                    row_band[rr] = (pg0, min(a, ymin), max(b2, ymax))
                else:
                    row_band[rr] = (pg, ymin, ymax)

    # หา header row + บทบาทคอลัมน์
    hdr = _find_header_row(grid, n_rows)
    headers = grid[hdr]
    qty_c = _find_col(headers, QTY_KEYS)
    price_c = _find_col(headers, PRICE_KEYS)
    amt_c = _find_col(headers, AMOUNT_KEYS)
    nouki_c = _find_col(headers, NOUKI_KEYS)

    # 3) ต่อบรรทัด: ดึง Part No. + จับ barcode ด้วย Y
    for r in range(hdr + 1, n_rows):
        part = extract_part_no(grid[r][0])
        if not part:
            continue
        qty = _num_lead(grid[r][qty_c]) if qty_c >= 0 else ""
        price = _num_lead(grid[r][price_c]) if price_c >= 0 else ""
        amount = _num_lead(grid[r][amt_c]) if amt_c >= 0 else ""
        nouki = grid[r][nouki_c] if nouki_c >= 0 else ""

        bc = ""
        band = row_band.get(r)
        if band:
            pg, ymin, ymax = band
            mid = (ymin + ymax) / 2.0
            best, bestd = None, None
            # (ก) barcode ที่ y อยู่ในแถบของบรรทัด
            for b in page_barcodes.get(pg, []):
                if b["used"]:
                    continue
                if ymin - 0.05 <= b["y"] <= ymax + 0.05:
                    d = abs(b["y"] - mid)
                    if bestd is None or d < bestd:
                        best, bestd = b, d
            # (ข) ถ้าไม่เจอในแถบ ลองตัวที่ใกล้สุดในระยะเผื่อ
            if best is None:
                for b in page_barcodes.get(pg, []):
                    if b["used"]:
                        continue
                    d = abs(b["y"] - mid)
                    if d <= 0.12 and (bestd is None or d < bestd):
                        best, bestd = b, d
            if best is not None:
                bc = best["value"]
                best["used"] = True

        rows.append({"part_no": part, "barcode": bc, "qty": qty, "price": price,
                     "amount": amount, "nouki": nouki})
        part_index.setdefault(part, []).append(len(rows) - 1)

    bc_matched = sum(1 for x in rows if x["barcode"])
    return {"rows": rows, "part_index": part_index, "bc_total": bc_total, "bc_matched": bc_matched}


# ──────────────────────────────────────────────────────────────
# Document pairing: meisai ↔ delivery ด้วยชุด Part No. ที่ทับกัน
# ──────────────────────────────────────────────────────────────
def pair_documents(meisai_indexes: dict, delivery_partsets: dict):
    """
    meisai_indexes:    { meisai_path: index_dict (จาก index_meisai) }
    delivery_partsets: { delivery_path: set(part_no) }

    คืน { delivery_path: {'meisai': meisai_path|None, 'confidence': float, 'overlap': int} }

    NON-exclusive: meisai 1 ใบ (เป็น batch ใบสั่ง) มักรองรับ 納品書 หลายใบ
    → แต่ละ delivery เลือก meisai ที่ part ทับกันมากสุดได้อิสระ (ใช้ meisai ซ้ำได้)
    confidence = สัดส่วน part ของ delivery ที่อยู่ใน meisai นั้น (coverage)
    """
    meisai_sets = {mp: set(idx["part_index"].keys()) for mp, idx in meisai_indexes.items()}

    result = {}
    for dpath, dset in delivery_partsets.items():
        best = None  # (overlap, coverage, mpath)
        for mpath, mset in meisai_sets.items():
            inter = dset & mset
            if not inter:
                continue
            coverage = len(inter) / len(dset) if dset else 0.0
            cand = (len(inter), coverage, mpath)
            if best is None or (cand[0], cand[1]) > (best[0], best[1]):
                best = cand
        if best:
            result[dpath] = {"meisai": best[2], "confidence": round(best[1], 3), "overlap": best[0]}
        else:
            result[dpath] = {"meisai": None, "confidence": 0.0, "overlap": 0}

    return result


# ──────────────────────────────────────────────────────────────
# Line matching ภายในคู่
# ──────────────────────────────────────────────────────────────
def _assign_group(idxs, avail, line_qty, mrows):
    """
    พยายามจับคู่ delivery lines (idxs) กับ barcode candidates (avail) ของ part เดียวกัน
    คืน dict {line_idx: cand_idx} ถ้าจับคู่ได้สะอาด, ไม่งั้น None

    Case A: แยกตาม 数量 ได้ลงตัว (จำนวนบรรทัด = จำนวน barcode ในแต่ละค่า 数量)
    Case B: 数量 แยกไม่ได้ แต่ "จำนวนบรรทัดรวม = จำนวน barcode รวม" → แปะ 1:1 ตามลำดับ
    """
    from collections import defaultdict

    line_by_q = defaultdict(list)
    for i in idxs:
        line_by_q[line_qty.get(i, "")].append(i)
    cand_by_q = defaultdict(list)
    for j in avail:
        cand_by_q[_num_lead(mrows[j]["qty"])].append(j)

    # Case A — ทุกค่า 数量 ของบรรทัดมี barcode จำนวนเท่ากัน และใช้ candidate ครบพอดี
    same_keys = set(line_by_q.keys()) <= set(cand_by_q.keys())
    counts_ok = all(len(cand_by_q.get(q, [])) == len(ls) for q, ls in line_by_q.items())
    if same_keys and counts_ok and len(idxs) == len(avail):
        out = {}
        for q, ls in line_by_q.items():
            for k, i in enumerate(ls):
                out[i] = cand_by_q[q][k]
        return out

    # Case B — จำนวนรวมตรงกัน → แปะ 1:1 ตามลำดับ (ตัวซ้ำที่เหมือนกัน แปะตัวไหนก็ครบ)
    if len(idxs) == len(avail):
        return {i: avail[k] for k, i in enumerate(idxs)}

    # จำนวนไม่ตรง → ให้คนตรวจ
    return None


def match_delivery(rows, roles, meisai_index):
    """
    จับคู่ทั้งใบ delivery กับ barcode ใน meisai_index (จับกลุ่มตาม Part No.)
    คืน list ขนาดเท่ากับ rows: แต่ละตัวเป็น dict
      {barcode, status, part_no, note, candidates}

    กติกาตามที่ตกลง:
      - จำนวนบรรทัด(ของ part) = จำนวน barcode → แปะ 1:1 อัตโนมัติ (OK)
      - จำนวนไม่ตรง                              → ติดธงทั้งกลุ่ม (REVIEW)
      - ไม่มี barcode สำหรับ part (เขียนมือ)      → NO_BARCODE
      - part ไม่อยู่ใน meisai / ไม่มีคู่ใบ          → UNMATCHED
    """
    n = len(rows)
    results = [None] * n
    pidx = roles.get("part", 0)

    def cell(row, role):
        c = roles.get(role, -1)
        return row[c] if (c is not None and 0 <= c < len(row)) else ""

    line_qty = {}
    groups = {}
    parts = []
    for i, row in enumerate(rows):
        part = extract_part_no(row[pidx] if pidx < len(row) else "")
        parts.append(part)
        line_qty[i] = _num_lead(cell(row, "qty"))
        groups.setdefault(part, []).append(i)

    if meisai_index is None:
        for i in range(n):
            results[i] = {"barcode": "", "status": STATUS_UNMATCHED, "part_no": parts[i],
                          "note": "no paired meisai", "candidates": [], "nouki": ""}
        return results

    mrows = meisai_index["rows"]
    pindex = meisai_index["part_index"]

    for part, idxs in groups.items():
        if not part:
            for i in idxs:
                results[i] = {"barcode": "", "status": STATUS_UNMATCHED, "part_no": "",
                              "note": "no part no", "candidates": [], "nouki": ""}
            continue

        in_meisai = part in pindex
        avail = [j for j in pindex.get(part, []) if mrows[j]["barcode"] and not mrows[j].get("_used")]

        if not avail:
            if in_meisai and not any(mrows[j]["barcode"] for j in pindex[part]):
                st, note = STATUS_NO_BARCODE, "no barcode on meisai"
            elif in_meisai:
                st, note = STATUS_UNMATCHED, "barcode exhausted"
            else:
                st, note = STATUS_UNMATCHED, "part not in meisai"
            for i in idxs:
                results[i] = {"barcode": "", "status": st, "part_no": part, "note": note,
                              "candidates": [], "nouki": ""}
            continue

        assigned = _assign_group(idxs, avail, line_qty, mrows)
        if assigned is not None:
            multi = len(idxs) > 1
            for i, j in assigned.items():
                mrows[j]["_used"] = True
                results[i] = {"barcode": mrows[j]["barcode"], "status": STATUS_OK, "part_no": part,
                              "note": ("複数1:1自動" if multi else ""), "candidates": [],
                              "nouki": mrows[j].get("nouki", "")}
        else:
            cands = [mrows[j]["barcode"] for j in avail]
            for i in idxs:
                results[i] = {"barcode": "", "status": STATUS_REVIEW, "part_no": part,
                              "note": f"件数不一致 {len(idxs)}行/{len(avail)}個", "candidates": cands,
                              "nouki": ""}

    return results


def reconcile(meisai_index):
    """คืน barcode ที่ยังไม่ถูกใช้ใน meisai (ไว้รายงาน)"""
    leftovers = []
    for rec in meisai_index["rows"]:
        if rec["barcode"] and not rec.get("_used"):
            leftovers.append({"part_no": rec["part_no"], "barcode": rec["barcode"]})
    return leftovers


# ──────────────────────────────────────────────────────────────
# อ่านตารางรายการ (ใช้ร่วมทั้ง delivery matching และ canonical export)
# ──────────────────────────────────────────────────────────────
def read_items_table(result, normalize_fn, pick_table_fn):
    """
    คืน dict:
      {
        'headers': [str, ...]      ชื่อหัวคอลัมน์ตามตำแหน่ง (จาก header row)
        'rows':    [ [str, ...] ]  ข้อมูลแต่ละบรรทัด (เฉพาะ data row, ตัด subtotal)
        'roles':   {part, qty, price, amount, desc}  index คอลัมน์ตามบทบาท (-1 ถ้าไม่มี)
      }
    วางข้อมูลตาม "ตำแหน่งคอลัมน์จริงของไฟล์นั้น" — ผู้เรียกค่อย map เข้า canonical ด้วยชื่อหัว
    """
    norm = normalize_fn
    _, table = pick_table_fn(result)
    if not table:
        return {"headers": [], "rows": [], "roles": {}}

    n_rows, n_cols = table.row_count, table.column_count
    grid = [["" for _ in range(n_cols)] for _ in range(n_rows)]
    for cell in table.cells:
        r, c = cell.row_index, cell.column_index
        rs = max(getattr(cell, "row_span", 1) or 1, 1)
        cs = max(getattr(cell, "column_span", 1) or 1, 1)
        val = norm(cell.content or "")
        for dr in range(rs):
            for dc in range(cs):
                rr, cc = r + dr, c + dc
                if rr < n_rows and cc < n_cols and dr == 0 and dc == 0:
                    grid[rr][cc] = val

    hdr = _find_header_row(grid, n_rows)
    headers = grid[hdr]
    roles = {
        "part": 0,  # คอลัมน์ 0 = 品番/品名 (ใช้ extract_part_no)
        "qty": _find_col(headers, QTY_KEYS),
        "unit": _find_col(headers, UNIT_KEYS),
        "price": _find_col(headers, PRICE_KEYS),
        "amount": _find_col(headers, AMOUNT_KEYS),
        "remark": _find_col(headers, REMARK_KEYS),
    }

    data = []
    for r in range(hdr + 1, n_rows):
        row = grid[r]
        txt = " ".join(row)
        if not txt.strip():
            continue
        if any(k in txt for k in ["小計", "合計", "総計", "頁計"]):
            continue
        data.append(list(row))

    return {"headers": list(headers), "rows": data, "roles": roles}
