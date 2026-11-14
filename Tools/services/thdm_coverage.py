"""Kiểm tra đủ BOM theo đợt để lập THĐM — logic thuần + truy vấn chỉ đọc.

PM chia sản phẩm của đơn hàng theo đợt (B30BizDocDetailFactory). THĐM của một đợt chỉ lập được khi
mỗi dòng sản phẩm của đợt đều có BOM. Danh sách BOM của đợt lấy từ SP của Bravo
(usp_B30BizDocDemand_BOMIdList_GetData) để khớp đúng cách Bravo chọn BOM/version.
"""
import datetime
import re

_PERIODS_SQL = """
SELECT f.PeriodId, ISNULL(MAX(p.Name), f.PeriodId) AS Name
FROM   B30BizDocDetailFactory f
LEFT JOIN B20Period p ON p.Id = TRY_CAST(f.PeriodId AS int)
WHERE  f.BizDocId = ? AND f.PeriodId IS NOT NULL AND f.IsActive = 1
GROUP  BY f.PeriodId
ORDER  BY TRY_CAST(f.PeriodId AS int), f.PeriodId
"""

_LINES_SQL = """
SELECT f.BuiltinOrder0, f.ItemId, i.Code, i.Name, f.Quantity9, f.Unit
FROM   B30BizDocDetailFactory f
LEFT JOIN B20Item i ON i.Id = f.ItemId
WHERE  f.BizDocId = ? AND f.PeriodId = ? AND f.IsActive = 1 AND ISNULL(f.IsGroup, 0) = 0
ORDER  BY f.BuiltinOrder
"""

_BOM_SQL = "SELECT Id, Code, Version, ItemId0, IsActive FROM B20BOM WHERE Id IN ({ids})"


def _clean_name(name):
    """Bỏ hậu tố '(File Costing)' khỏi tên sản phẩm khi hiển thị."""
    return re.sub(r"\s*\(\s*file\s+costing\s*\)", "", str(name or ""), flags=re.IGNORECASE).strip()


def muc_key(muc):
    """Khóa sắp xếp Mục số kiểu '1', '1.1', '2.10' (so từng phần là số)."""
    parts = re.findall(r"\d+", str(muc or ""))
    return tuple(int(p) for p in parts) or (10 ** 9,)


def version_key(version):
    try:
        return float(str(version).strip())
    except (TypeError, ValueError):
        return -1.0


def compute_coverage(lines, boms):
    """Gắn BOM cho từng dòng sản phẩm của đợt.

    lines: [{muc, item_id, code, name, qty, unit}]
    boms:  [{id, code, version, item_id0, is_active}] — danh sách BOM Bravo trả về cho đợt.
    BOM không còn dùng (is_active falsy) coi như không có. Một sản phẩm có nhiều BOM → lấy version cao nhất.
    Trả về (rows, summary): rows = dòng + {bom_id, bom_code, bom_version, status: 'ok'|'missing'}.
    """
    best = {}
    for b in boms:
        if not b.get("is_active") or b.get("item_id0") is None:
            continue
        cur = best.get(b["item_id0"])
        if cur is None or (version_key(b.get("version")), b["id"]) > (version_key(cur.get("version")), cur["id"]):
            best[b["item_id0"]] = b
    rows = []
    for ln in sorted(lines, key=lambda x: muc_key(x.get("muc"))):
        bom = best.get(ln.get("item_id"))
        rows.append(dict(ln,
                         bom_id=bom["id"] if bom else None,
                         bom_code=bom.get("code") if bom else "",
                         bom_version=bom.get("version") if bom else "",
                         status="ok" if bom else "missing"))
    n_ok = sum(1 for r in rows if r["status"] == "ok")
    return rows, dict(total=len(rows), ok=n_ok, missing=len(rows) - n_ok)


def _fetch_all(cur, sql, params=()):
    cur.execute(sql, params)
    return cur.fetchall()


def load_periods(conn, order_id):
    """[(period_id:str, name:str)] các đợt của đơn."""
    cur = conn.cursor()
    return [(str(r[0]), str(r[1] or r[0])) for r in _fetch_all(cur, _PERIODS_SQL, (order_id,))]


def load_lines(conn, order_id, period_id):
    cur = conn.cursor()
    return [dict(muc=str(r[0] or ""), item_id=r[1], code=r[2] or "", name=_clean_name(r[3]),
                 qty=r[4], unit=r[5] or "")
            for r in _fetch_all(cur, _LINES_SQL, (order_id, str(period_id)))]


def load_period_boms(conn, order_id, period_id, today=None):
    """Danh sách BOM của đợt theo SP Bravo → [{id, code, version, item_id0, is_active}]."""
    today = (today or datetime.date.today()).isoformat()
    cur = conn.cursor()
    cur.execute("EXEC dbo.usp_B30BizDocDemand_BOMIdList_GetData "
                "@_DocDate=?, @_PeriodId=?, @_BizDocId_SO=?", (today, int(period_id), order_id))
    while cur.description is None and cur.nextset():
        pass
    row = cur.fetchone() if cur.description else None
    try:
        conn.rollback()   # SP chỉ đọc; bỏ giao dịch cho chắc
    except Exception:
        pass
    ids = [int(x) for x in re.findall(r"\d+", (row[0] or "") if row else "")]
    if not ids:
        return []
    cur = conn.cursor()
    sql = _BOM_SQL.format(ids=",".join("?" * len(ids)))
    return [dict(id=r[0], code=r[1] or "", version=r[2], item_id0=r[3], is_active=bool(r[4]))
            for r in _fetch_all(cur, sql, ids)]


def load_order_coverage(conn, order_id):
    """Toàn bộ các đợt của đơn: [{period_id, name, rows, summary, bom_ids}] (đã sắp theo đợt)."""
    result = []
    for period_id, name in load_periods(conn, order_id):
        lines = load_lines(conn, order_id, period_id)
        boms = load_period_boms(conn, order_id, period_id)
        rows, summary = compute_coverage(lines, boms)
        result.append(dict(period_id=period_id, name=name, rows=rows, summary=summary,
                           bom_ids=[b["id"] for b in boms if b.get("is_active")]))
    return result


def period_message(name, summary, rows):
    """Câu kết luận cho 1 đợt → (text, level) với level 'ok' | 'error' | 'empty'."""
    if not summary["total"]:
        return f"{name} chưa có dòng sản phẩm nào.", "empty"
    if not summary["missing"]:
        return (f"{name} đã đủ {summary['ok']}/{summary['total']} BOM — "
                f"có thể lập THĐM trong Bravo.", "ok")
    miss = ", ".join(r["muc"] for r in rows if r["status"] == "missing")
    return (f"{name} chưa đủ BOM — chưa nên lập THĐM. "
            f"Còn thiếu {summary['missing']} dòng (Mục {miss}).", "error")
