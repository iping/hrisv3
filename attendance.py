"""Absensi harian per petugas dari tab Attendance, Shift, dan Overtime."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from calculate_payroll import overtime_fee, person_key, sheet_rows, text, tokens
from openpyxl import load_workbook

WEEKDAYS = ("Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu")
MONTHS = ("Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des")
STATUS = {
    "OT": ("Hadir", "hadir"),
    "L": ("Terlambat", "telat"),
    "A": ("Alpa", "alpa"),
    "P": ("Izin", "izin"),
    "S": ("Sakit", "sakit"),
    "LEAVE": ("Cuti", "cuti"),
    "OFF": ("Off", "off"),
    "O": ("Tanpa jadwal", "kosong"),
}

_cache: dict = {"path": "", "mtime": 0.0, "people": {}}


def as_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = text(value)[:10]
    if len(raw) < 10:
        return None
    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError:
        return None


def clock(value) -> str:
    raw = text(value)
    if len(raw) >= 16 and raw[10] == " ":
        return raw[11:16]
    parsed = value if isinstance(value, datetime) else None
    if parsed:
        return parsed.strftime("%H:%M")
    return ""


def badge(code: str) -> dict:
    label, kind = STATUS.get(code, (code, "kosong"))
    return {"code": code, "label": label, "kind": kind}


def blank_person(lokasi: str, nama: str, username: str) -> dict:
    return {
        "lokasi": lokasi,
        "nama": nama,
        "username": username,
        "days": {},
    }


def load_people(path: Path) -> dict:
    resolved = str(path.resolve())
    mtime = path.stat().st_mtime
    if _cache["path"] == resolved and _cache["mtime"] == mtime:
        return _cache["people"]

    workbook = load_workbook(path, read_only=True, data_only=True)
    people: dict[tuple[str, str], dict] = {}

    attendance = sheet_rows(workbook, "Attendance")
    if attendance:
        headers = [as_date(cell) for cell in attendance[0]]
        for row in attendance[1:]:
            if not row or not text(row[0]):
                continue
            key = person_key(row[0], row[3] if len(row) > 3 else "", row[1])
            person = people.setdefault(key, blank_person(text(row[0]), text(row[1]), text(row[3] if len(row) > 3 else "")))
            if not person["nama"]:
                person["nama"] = text(row[1])
            for pos, day in enumerate(headers):
                if day is None or pos >= len(row):
                    continue
                codes = tokens(row[pos])
                if not codes and not text(row[pos]):
                    person["days"].setdefault(day, {"shift": "", "codes": [], "overtime": []})
                    continue
                slot = person["days"].setdefault(day, {"shift": "", "codes": [], "overtime": []})
                slot["codes"].extend(codes)

    shifts = sheet_rows(workbook, "Shift")
    if shifts:
        headers = [as_date(cell) for cell in shifts[0]]
        for row in shifts[1:]:
            if not row or not text(row[0]):
                continue
            key = person_key(row[0], row[3] if len(row) > 3 else "", row[1])
            person = people.setdefault(key, blank_person(text(row[0]), text(row[1]), text(row[3] if len(row) > 3 else "")))
            for pos, day in enumerate(headers):
                if day is None or pos >= len(row) or not text(row[pos]):
                    continue
                slot = person["days"].setdefault(day, {"shift": "", "codes": [], "overtime": []})
                if not slot["shift"]:
                    slot["shift"] = text(row[pos])

    overtime = sheet_rows(workbook, "Overtime")
    if overtime:
        headers = [text(cell) for cell in overtime[0]]
        index = {name: pos for pos, name in enumerate(headers)}
        for row in overtime[1:]:
            if not row or not any(row):
                continue
            lokasi = text(row[index.get("Lokasi", 2)]) if index.get("Lokasi", 2) < len(row) else ""
            nama = text(row[index.get("Petugas", 3)]) if index.get("Petugas", 3) < len(row) else ""
            username = text(row[index.get("Username", 4)]) if index.get("Username", 4) < len(row) else ""
            mulai = row[index["Mulai Lembur"]] if "Mulai Lembur" in index and index["Mulai Lembur"] < len(row) else ""
            day = as_date(mulai)
            if day is None:
                continue
            jam_pos = index.get("Durasi Lembur (Jam)")
            menit_pos = index.get("Durasi Lembur (Menit)")
            jam = float(row[jam_pos] or 0) if jam_pos is not None and jam_pos < len(row) and row[jam_pos] not in (None, "") else 0.0
            if jam == 0 and menit_pos is not None and menit_pos < len(row) and row[menit_pos]:
                jam = float(row[menit_pos]) / 60.0
            fee_pos = index.get("Overtime Fee")
            fee = float(row[fee_pos] or 0) if fee_pos is not None and fee_pos < len(row) else 0.0
            masuk_pos = index.get("Absen Masuk")
            keluar_pos = index.get("Absen Keluar")
            key = person_key(lokasi, username, nama)
            person = people.setdefault(key, blank_person(lokasi, nama, username))
            slot = person["days"].setdefault(day, {"shift": "", "codes": [], "overtime": []})
            jam_text = f"{jam:.2f}".replace(".", ",").rstrip("0").rstrip(",")
            slot["overtime"].append(
                {
                    "jam": jam,
                    "jam_text": jam_text or "0",
                    "fee": fee,
                    "fee_hitung": overtime_fee(jam),
                    "masuk": clock(row[masuk_pos]) if masuk_pos is not None and masuk_pos < len(row) else "",
                    "keluar": clock(row[keluar_pos]) if keluar_pos is not None and keluar_pos < len(row) else "",
                    "mulai": clock(mulai),
                    "akhir": clock(row[index["Akhir Lembur"]]) if "Akhir Lembur" in index and index["Akhir Lembur"] < len(row) else "",
                }
            )

    workbook.close()
    _cache["path"] = resolved
    _cache["mtime"] = mtime
    _cache["people"] = people
    return people


def count_codes(person: dict | None) -> dict:
    counts = {kind: 0 for _, kind in STATUS.values()}
    counts["hari"] = 0
    if not person:
        return counts
    for slot in person["days"].values():
        if slot["codes"]:
            counts["hari"] += 1
        for code in slot["codes"]:
            kind = STATUS.get(code, ("", "kosong"))[1]
            counts[kind] = counts.get(kind, 0) + 1
    return counts


def attach_counts(rows: list[dict], path: Path | None) -> None:
    people = load_people(path) if path is not None else {}
    for row in rows:
        person = people.get(person_key(row.get("lokasi") or "", row.get("username") or "", row.get("nama") or ""))
        counts = count_codes(person)
        row["absensi"] = person is not None
        row["hadir"] = counts["hadir"]
        row["telat"] = counts["telat"]
        row["alpa"] = counts["alpa"]


def officer_attendance(path: Path, lokasi: str, username: str) -> dict | None:
    people = load_people(path)
    wanted = (text(lokasi).lower(), text(username).lower())
    person = people.get(wanted)
    if person is None:
        return None
    days = []
    for day in sorted(person["days"]):
        slot = person["days"][day]
        badges = [badge(code) for code in slot["codes"]]
        flags = sorted({item["kind"] for item in badges})
        days.append(
            {
                "iso": day.isoformat(),
                "day": day.day,
                "month": MONTHS[day.month - 1],
                "weekday": WEEKDAYS[day.weekday()],
                "shift": slot["shift"],
                "badges": badges,
                "flags": " ".join(flags),
                "warn": "alpa" in flags or "telat" in flags,
                "overtime": slot["overtime"],
            }
        )
    counts = count_codes(person)
    return {
        "lokasi": person["lokasi"],
        "nama": person["nama"],
        "username": person["username"],
        "days": days,
        "hadir": counts["hadir"],
        "telat": counts["telat"],
        "alpa": counts["alpa"],
        "izin": counts["izin"],
        "sakit": counts["sakit"],
        "cuti": counts["cuti"],
        "off": counts["off"],
        "lembur": sum(len(day["overtime"]) for day in days),
        "terisi": sum(1 for day in days if day["badges"]),
    }
