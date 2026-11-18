"""Database payroll per bulan.

Menyimpan slip, profil petugas, dan pengajuan banding dari file tarikan.
Tanggal masuk dan tanggal resign di Officer Profile dipakai untuk menandai
petugas baru dan resign pada periode file itu.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from calculate_payroll import Payroll, sheet_rows, text
from db import connect
from openpyxl import load_workbook


def person_key(lokasi: str, username: str, nama: str = "") -> str:
    identity = text(username).lower() or text(nama).lower()
    return f"{text(lokasi).lower()}|{identity}"


def parse_day(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    raw = text(value)[:10]
    if len(raw) == 10 and raw[4] == "-" and raw[7] == "-":
        return raw
    return ""


def split_period(period: str) -> tuple[str, str]:
    if "s/d" in (period or ""):
        left, right = period.split("s/d", 1)
        return parse_day(left), parse_day(right)
    return "", ""


def file_status(joined: str, resigned: str, start: str, end: str) -> tuple[str, str, str]:
    """Status dari tanggal di Officer Profile terhadap periode payroll."""
    if resigned and (not end or resigned <= end):
        return "resign", resigned, f"Officer Profile mencatat resign {resigned}"
    if resigned and end and resigned > end:
        return "aktif", "", f"Officer Profile: resign {resigned} masih sesudah periode ini"
    if joined and start and end and start <= joined <= end:
        return "baru", joined, f"Officer Profile: masuk {joined}"
    return "aktif", "", ""


def _cell(row: tuple, index: dict[str, int], name: str):
    pos = index.get(name)
    if pos is None or pos >= len(row):
        return None
    return row[pos]


def load_profiles(path: Path) -> dict[str, dict]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = sheet_rows(wb, "Officer Profile")
    finally:
        wb.close()
    if not rows:
        return {}
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    if "Username" not in index:
        return {}
    found = {}
    for row in rows[1:]:
        if not row:
            continue
        lokasi = text(_cell(row, index, "Lokasi Primary"))
        username = text(_cell(row, index, "Username"))
        nama = text(_cell(row, index, "Nama"))
        if not lokasi and not username:
            continue
        record = {
            "nama": nama,
            "username": username,
            "lokasi": lokasi,
            "numeric_id": text(_cell(row, index, "Numeric ID")),
            "bank": text(_cell(row, index, "Bank")),
            "rekening": text(_cell(row, index, "No. Rekening")),
            "phone": text(_cell(row, index, "Phone")),
            "email": text(_cell(row, index, "Email")),
            "joined": parse_day(_cell(row, index, "Joined Date")),
            "resigned": parse_day(_cell(row, index, "Resigned Date")),
            "activated": text(_cell(row, index, "Activated At")),
        }
        found[person_key(lokasi, username, nama)] = record
    return found


def load_appeals(path: Path) -> list[dict]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = sheet_rows(wb, "Pengajuan Banding")
    finally:
        wb.close()
    if not rows:
        return []
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    if "Username" not in index:
        return []
    appeals = []
    for row in rows[1:]:
        if not row or not text(_cell(row, index, "Username")) and not text(_cell(row, index, "Nama")):
            continue
        lokasi = text(_cell(row, index, "Lokasi"))
        username = text(_cell(row, index, "Username"))
        nama = text(_cell(row, index, "Nama"))
        appeals.append(
            {
                "person_key": person_key(lokasi, username, nama),
                "lokasi": lokasi,
                "nama": nama,
                "username": username,
                "submitted_at": text(_cell(row, index, "Waktu Pengajuan")),
                "attendance_date": parse_day(_cell(row, index, "Tanggal Attendance")),
                "label": text(_cell(row, index, "Label")),
                "reason": text(_cell(row, index, "Reason")),
                "detail": text(_cell(row, index, "Detail")),
                "status": text(_cell(row, index, "Status")),
                "before_status": text(_cell(row, index, "Attendance Sebelum")),
                "after_status": text(_cell(row, index, "Attendance Sesudah")),
                "sequence": int(float(_cell(row, index, "Urutan Appeal") or 0) or 0),
            }
        )
    return appeals


def profile_for(officer, profiles: dict[str, dict]) -> dict:
    direct = profiles.get(person_key(officer.lokasi, officer.username, officer.nama))
    if direct:
        return direct
    username = text(officer.username).lower()
    matches = [row for row in profiles.values() if text(row.get("username")).lower() == username]
    if len(matches) == 1:
        return matches[0]
    return {}


def save_month(path: Path, payroll: Payroll) -> dict:
    start, end = split_period(payroll.period)
    mtime = str(int(path.stat().st_mtime))
    if start and end:
        with connect() as db:
            existing = db.execute(
                "SELECT * FROM months WHERE period_start = ? AND period_end = ?",
                (start, end),
            ).fetchone()
        if existing and existing["source_mtime"] == mtime:
            return {
                "id": existing["id"],
                "period_start": start,
                "period_end": end,
                "period": f"{start} s/d {end}",
                "source_name": existing["source_name"] or payroll.source_name,
                "officers": existing["officer_count"] or 0,
                "hris": existing["hris_total"] or 0,
                "hitung": existing["hitung_total"] or 0,
                "baru": existing["baru_count"] or 0,
                "resign": existing["resign_count"] or 0,
                "appeals": existing["appeal_count"] or 0,
                "approved": existing["approved_count"] or 0,
                "pending": existing["pending_count"] or 0,
            }
    if not start or not end:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            from calculate_payroll import attendance_bounds

            start, end = attendance_bounds(sheet_rows(wb, "Attendance"))
        finally:
            wb.close()
    profiles = load_profiles(path)
    appeals = load_appeals(path)
    lines = []
    for officer in payroll.officers:
        profile = profile_for(officer, profiles)
        status, _when, _note = file_status(profile.get("joined", ""), profile.get("resigned", ""), start, end)
        lines.append(
            {
                "person_key": person_key(officer.lokasi, officer.username, officer.nama),
                "lokasi": officer.lokasi,
                "nama": officer.nama,
                "username": officer.username,
                "numeric_id": officer.numeric_id or profile.get("numeric_id", ""),
                "basic": int(round(officer.basic)),
                "paid": int(round(officer.paid)),
                "hitung": int(round(officer.gaji_hitung)),
                "selisih": int(round(officer.selisih_gaji)),
                "alpha": int(officer.alpha or 0),
                "status": status,
                "joined_date": profile.get("joined", ""),
                "resigned_date": profile.get("resigned", ""),
                "activated": profile.get("activated", ""),
            }
        )
    approved = sum(1 for row in appeals if row["status"].endswith("APPROVED"))
    pending = sum(1 for row in appeals if "PENDING" in row["status"])
    summary = {
        "period_start": start,
        "period_end": end,
        "period": f"{start} s/d {end}" if start and end else payroll.period,
        "source_name": payroll.source_name,
        "officers": len(lines),
        "hris": sum(row["paid"] for row in lines),
        "hitung": sum(row["hitung"] for row in lines),
        "baru": sum(1 for row in lines if row["status"] == "baru"),
        "resign": sum(1 for row in lines if row["status"] == "resign"),
        "appeals": len(appeals),
        "approved": approved,
        "pending": pending,
    }
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    with connect() as db:
        existing = None
        if start and end:
            existing = db.execute(
                "SELECT id, source_mtime FROM months WHERE period_start = ? AND period_end = ?",
                (start, end),
            ).fetchone()
        if existing:
            month_id = existing["id"]
            db.execute("DELETE FROM payroll_lines WHERE month_id = ?", (month_id,))
            db.execute("DELETE FROM appeals WHERE month_id = ?", (month_id,))
            db.execute(
                """
                UPDATE months
                SET source_name = ?, source_mtime = ?, saved_at = ?, officer_count = ?,
                    hris_total = ?, hitung_total = ?, baru_count = ?, resign_count = ?,
                    appeal_count = ?, approved_count = ?, pending_count = ?
                WHERE id = ?
                """,
                (
                    payroll.source_name,
                    mtime,
                    now,
                    summary["officers"],
                    summary["hris"],
                    summary["hitung"],
                    summary["baru"],
                    summary["resign"],
                    summary["appeals"],
                    summary["approved"],
                    summary["pending"],
                    month_id,
                ),
            )
        else:
            cursor = db.execute(
                """
                INSERT INTO months (
                    period_start, period_end, source_name, source_mtime, saved_at,
                    officer_count, hris_total, hitung_total, baru_count, resign_count,
                    appeal_count, approved_count, pending_count
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                RETURNING id
                """,
                (
                    start,
                    end,
                    payroll.source_name,
                    mtime,
                    now,
                    summary["officers"],
                    summary["hris"],
                    summary["hitung"],
                    summary["baru"],
                    summary["resign"],
                    summary["appeals"],
                    summary["approved"],
                    summary["pending"],
                ),
            )
            month_id = cursor.fetchone()["id"]
        db.executemany(
            """
            INSERT INTO payroll_lines (
                month_id, person_key, lokasi, nama, username, numeric_id,
                basic, paid, hitung, selisih, alpha, status, joined_date, resigned_date, activated
            ) VALUES (
                :month_id, :person_key, :lokasi, :nama, :username, :numeric_id,
                :basic, :paid, :hitung, :selisih, :alpha, :status, :joined_date, :resigned_date, :activated
            )
            """,
            [{**row, "month_id": month_id} for row in lines],
        )
        if appeals:
            db.executemany(
                """
                INSERT INTO appeals (
                    month_id, person_key, lokasi, nama, username, submitted_at, attendance_date,
                    label, reason, detail, status, before_status, after_status, sequence
                ) VALUES (
                    :month_id, :person_key, :lokasi, :nama, :username, :submitted_at, :attendance_date,
                    :label, :reason, :detail, :status, :before_status, :after_status, :sequence
                )
                """,
                [{**row, "month_id": month_id} for row in appeals],
            )
        for record in profiles.values():
            key = person_key(record["lokasi"], record["username"], record["nama"])
            current = db.execute("SELECT joined_date, resigned_date FROM officers WHERE person_key = ?", (key,)).fetchone()
            joined = record["joined"] or (current["joined_date"] if current else "")
            resigned = record["resigned"] or (current["resigned_date"] if current else "")
            db.execute(
                """
                INSERT INTO officers (
                    person_key, numeric_id, nama, username, lokasi, bank, rekening, phone, email,
                    joined_date, resigned_date, activated, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(person_key) DO UPDATE SET
                    numeric_id = excluded.numeric_id,
                    nama = excluded.nama,
                    username = excluded.username,
                    lokasi = excluded.lokasi,
                    bank = CASE WHEN excluded.bank != '' THEN excluded.bank ELSE officers.bank END,
                    rekening = CASE WHEN excluded.rekening != '' THEN excluded.rekening ELSE officers.rekening END,
                    phone = CASE WHEN excluded.phone != '' THEN excluded.phone ELSE officers.phone END,
                    email = CASE WHEN excluded.email != '' THEN excluded.email ELSE officers.email END,
                    joined_date = excluded.joined_date,
                    resigned_date = excluded.resigned_date,
                    activated = excluded.activated,
                    updated_at = excluded.updated_at
                """,
                (
                    key,
                    record["numeric_id"],
                    record["nama"],
                    record["username"],
                    record["lokasi"],
                    record["bank"],
                    record["rekening"],
                    record["phone"],
                    record["email"],
                    joined,
                    resigned,
                    record["activated"],
                    now,
                ),
            )
        db.commit()
    summary["id"] = month_id
    return summary


def legend_terms(path: Path) -> list[tuple[str, str]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = sheet_rows(wb, "Info & Legenda")
    finally:
        wb.close()
    pairs = []
    for row in rows[1:]:
        if not row:
            continue
        label = text(row[0]).strip().lstrip("-").strip()
        detail = text(row[1]) if len(row) > 1 else ""
        if not label or not detail or label.lower().startswith("tab "):
            continue
        pairs.append((label, detail))
    return pairs


def person_history(lokasi: str, username: str) -> dict:
    from roster import person_key

    key = person_key(lokasi, username)
    try:
        with connect() as db:
            lines = db.execute(
                """
                SELECT m.period_start, m.period_end, m.source_name,
                       l.paid, l.hitung, l.selisih, l.alpha, l.status, l.basic
                FROM payroll_lines l
                JOIN months m ON m.id = l.month_id
                WHERE l.person_key = ? OR (lower(l.username) = lower(?) AND lower(l.lokasi) = lower(?))
                ORDER BY m.period_end DESC, m.id DESC
                LIMIT 8
                """,
                (key, username, lokasi),
            ).fetchall()
            appeals = db.execute(
                """
                SELECT m.period_start, a.attendance_date, a.label, a.status,
                       a.before_status, a.after_status, a.reason
                FROM appeals a
                JOIN months m ON m.id = a.month_id
                WHERE lower(a.username) = lower(?) AND lower(a.lokasi) = lower(?)
                ORDER BY m.period_end DESC, a.attendance_date DESC
                LIMIT 8
                """,
                (username, lokasi),
            ).fetchall()
    except Exception:
        return {"lines": [], "appeals": []}
    return {"lines": lines, "appeals": appeals}


def list_months() -> list[dict]:
    with connect() as db:
        rows = db.execute("SELECT * FROM months ORDER BY period_end DESC, id DESC").fetchall()
    return [dict(row) for row in rows]


def month_page(month_id: int) -> dict | None:
    with connect() as db:
        month = db.execute("SELECT * FROM months WHERE id = ?", (month_id,)).fetchone()
        if month is None:
            return None
        lines = db.execute(
            """
            SELECT * FROM payroll_lines
            WHERE month_id = ?
            ORDER BY CASE status WHEN 'resign' THEN 0 WHEN 'baru' THEN 1 ELSE 2 END, lokasi, nama
            """,
            (month_id,),
        ).fetchall()
        appeals = db.execute(
            """
            SELECT * FROM appeals
            WHERE month_id = ?
            ORDER BY status, attendance_date, nama
            """,
            (month_id,),
        ).fetchall()
    return {"month": dict(month), "lines": [dict(row) for row in lines], "appeals": [dict(row) for row in appeals]}
