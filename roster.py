"""Daftar lokasi aktif, petugas resign, dan petugas baru.

Disimpan di output/roster.json dan dipakai ulang setiap bulan.
Status yang sudah disimpan manual tidak ditimpa oleh tarikan HRIS.
"""

from __future__ import annotations

import json
from pathlib import Path

from openpyxl import load_workbook

from calculate_payroll import Payroll, sheet_rows, text

ROOT = Path(__file__).resolve().parent
ROSTER_PATH = ROOT / "output" / "roster.json"
STATUSES = ("aktif", "resign", "baru")


def person_key(lokasi: str, username: str, nama: str = "") -> str:
    identity = text(username).lower() or text(nama).lower()
    return f"{text(lokasi).lower()}|{identity}"


def empty_roster() -> dict:
    return {"meta": {}, "locations": {}, "people": {}}


def load_roster() -> dict:
    if not ROSTER_PATH.exists():
        return empty_roster()
    data = json.loads(ROSTER_PATH.read_text(encoding="utf-8"))
    data.setdefault("meta", {})
    data.setdefault("locations", {})
    data.setdefault("people", {})
    return data


def save_roster(roster: dict) -> None:
    ROSTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    ROSTER_PATH.write_text(json.dumps(roster, ensure_ascii=False, indent=2), encoding="utf-8")


def date_text(value) -> str:
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    raw = text(value)
    if len(raw) >= 10 and raw[4] == "-" and raw[7] == "-":
        return raw[:10]
    return ""


def resign_dates(path: Path) -> dict[str, str]:
    found = {}
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = sheet_rows(wb, "Anomaly")
    finally:
        wb.close()
    if not rows:
        return found
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    if "Tipe Anomali" not in index:
        return found
    for row in rows[1:]:
        if not row or text(row[index["Tipe Anomali"]]) != "resigned_in_period":
            continue
        lokasi = text(row[index["Lokasi"]]) if "Lokasi" in index else ""
        username = text(row[index["Username"]]) if "Username" in index else ""
        nama = text(row[index["Nama"]]) if "Nama" in index else ""
        tanggal = date_text(row[index["Tanggal"]]) if "Tanggal" in index else ""
        found[person_key(lokasi, username, nama)] = tanggal
    return found


def looks_testing(lokasi: str = "", nama: str = "", username: str = "") -> bool:
    return "testing" in f"{lokasi} {nama} {username}".lower()


def sync_flags(roster: dict | None = None) -> dict:
    """Tandai akun testing dan usulkan mutasi. Pilihan manual tidak ditimpa."""
    data = roster if roster is not None else load_roster()
    people = data.setdefault("people", {})
    by_user: dict[str, list[tuple[str, dict]]] = {}
    for key, row in people.items():
        username = (row.get("username") or "").strip().lower()
        if username:
            by_user.setdefault(username, []).append((key, row))
    changed = False
    for key, row in people.items():
        if looks_testing(row.get("lokasi") or "", row.get("nama") or "", row.get("username") or "") and not row.get("testing_manual"):
            if not row.get("testing"):
                row["testing"] = True
                changed = True
            note = "Akun testing. Tidak masuk transfer payroll berikutnya."
            if not row.get("saran"):
                row["saran"] = note
                changed = True
        username = (row.get("username") or "").strip().lower()
        others = [item[1].get("lokasi") or "" for item in by_user.get(username, []) if item[0] != key and item[1].get("lokasi")]
        row["kantor_lain"] = others
        if others and not row.get("saran") and not row.get("mutasi"):
            row["saran"] = "Username yang sama ada di " + ", ".join(others) + ". Cek mutasi kantor."
            changed = True
    if changed and roster is None:
        save_roster(data)
    elif changed and roster is not None:
        save_roster(data)
    return data


def merge_payroll(payroll: Payroll, source: Path | None = None) -> dict:
    """Tambah lokasi dan petugas baru dari tarikan. Pilihan manual tetap."""
    roster = load_roster()
    locations = roster["locations"]
    people = roster["people"]
    resigned = resign_dates(source) if source else {}
    profiles = {}
    period_start, period_end = "", ""
    if source is not None:
        from archive import file_status, load_profiles, profile_for, split_period

        profiles = load_profiles(source)
        period_start, period_end = split_period(payroll.period)

    for officer in payroll.officers:
        location = locations.setdefault(
            officer.lokasi,
            {"active": "testing" not in officer.lokasi.lower(), "manual": False},
        )
        location["seen"] = True
        key = person_key(officer.lokasi, officer.username, officer.nama)
        current = people.get(key)
        if current is None:
            status = "resign" if key in resigned else "aktif"
            current = {
                "lokasi": officer.lokasi,
                "nama": officer.nama,
                "username": officer.username,
                "status": status,
                "tanggal": resigned.get(key, ""),
                "catatan": "",
                "manual": False,
                "saran": "",
            }
            people[key] = current
        else:
            current["nama"] = officer.nama or current.get("nama") or ""
            current["lokasi"] = officer.lokasi
            if not current.get("manual") and key in resigned:
                current["status"] = "resign"
                current["tanggal"] = resigned.get(key) or current.get("tanggal") or ""
        profile = profile_for(officer, profiles) if profiles else {}
        suggested, tanggal, note = file_status(
            profile.get("joined", ""),
            profile.get("resigned", "") or resigned.get(key, ""),
            period_start,
            period_end,
        ) if profile or key in resigned else ("", "", "")
        if not profile and key in resigned:
            suggested, tanggal, note = "resign", resigned[key], f"HRIS mencatat resign {resigned[key]}"
        if suggested in {"resign", "baru"} and (
            not current.get("manual") or current.get("status") in {"aktif", suggested}
        ):
            current["status"] = suggested
            if tanggal:
                current["tanggal"] = tanggal
        if note:
            current["saran"] = note
        elif profile.get("activated") == "Tidak":
            current["saran"] = "Officer Profile menandai petugas tidak aktif."
        elif officer.amounts["prorate"] > 0 and current.get("status") != "resign":
            current["saran"] = "Slip punya prorate. Cek apakah petugas baru."
        else:
            current["saran"] = ""
        current["joined_date"] = profile.get("joined", "") or current.get("joined_date", "")
        current["resigned_date"] = profile.get("resigned", "") or current.get("resigned_date", "")
        current["paid"] = officer.paid

    roster["meta"] = {
        "source_name": payroll.source_name,
        "period": payroll.period,
        "month": payroll.month,
    }
    save_roster(roster)
    return sync_flags(roster)


def apply_choices(
    location_names: list[str],
    active_names: set[str],
    people_rows: list[dict],
    testing_keys: set[str] | None = None,
    mutasi_keys: set[str] | None = None,
) -> dict:
    roster = load_roster()
    posted_locations = set(location_names)
    for name, row in roster["locations"].items():
        if name not in posted_locations:
            continue
        row["active"] = name in active_names
        row["manual"] = True
    for item in people_rows:
        key = item["key"]
        current = roster["people"].setdefault(
            key,
            {
                "lokasi": item.get("lokasi", ""),
                "nama": item.get("nama", ""),
                "username": item.get("username", ""),
                "paid": 0,
                "saran": "",
            },
        )
        status = item.get("status") or "aktif"
        if status not in STATUSES:
            status = "aktif"
        current["status"] = status
        current["tanggal"] = item.get("tanggal") or ""
        current["catatan"] = item.get("catatan") or ""
        current["manual"] = True
        if testing_keys is not None:
            current["testing"] = key in testing_keys
            current["testing_manual"] = True
        if mutasi_keys is not None:
            current["mutasi"] = key in mutasi_keys
            current["mutasi_manual"] = True
            current["mutasi_dari"] = item.get("mutasi_dari") or ""
    save_roster(roster)
    return sync_flags(roster)


def set_employment(lokasi: str, username: str, fields: dict) -> None:
    roster = load_roster()
    wanted_user = username.strip().lower()
    wanted_loc = lokasi.strip().lower()
    for row in roster.get("people", {}).values():
        if (row.get("username") or "").lower() != wanted_user or (row.get("lokasi") or "").lower() != wanted_loc:
            continue
        status = fields.get("status") or row.get("status") or "aktif"
        row["status"] = status if status in STATUSES else "aktif"
        row["tanggal"] = fields.get("tanggal") or ""
        row["testing"] = bool(fields.get("testing"))
        row["testing_manual"] = True
        row["mutasi"] = bool(fields.get("mutasi"))
        row["mutasi_manual"] = True
        row["mutasi_dari"] = fields.get("mutasi_dari") or ""
        row["manual"] = True
        break
    save_roster(roster)
    sync_flags(roster)


def roster_for_view(payroll: Payroll, roster: dict | None = None) -> dict:
    data = roster if roster is not None else load_roster()
    locations = data.get("locations") or {}
    people = data.get("people") or {}

    def active(name: str) -> bool:
        row = locations.get(name)
        if row is None:
            return True
        return bool(row.get("active", True))

    seen: dict[str, float] = {}
    active_pay = 0.0
    inactive_pay = 0.0
    for officer in payroll.officers:
        seen[officer.lokasi] = seen.get(officer.lokasi, 0) + officer.paid
        if active(officer.lokasi):
            active_pay += officer.paid
        else:
            inactive_pay += officer.paid

    inactive = [
        {"lokasi": name, "amount": amount}
        for name, amount in sorted(seen.items(), key=lambda item: item[0].lower())
        if not active(name)
    ]
    resign = []
    baru = []
    for officer in payroll.officers:
        row = people.get(person_key(officer.lokasi, officer.username, officer.nama)) or {}
        status = row.get("status") or "aktif"
        item = {
            "lokasi": officer.lokasi,
            "nama": officer.nama,
            "username": officer.username,
            "paid": officer.paid,
            "tanggal": row.get("tanggal") or "",
            "catatan": row.get("catatan") or "",
            "aktif_lokasi": active(officer.lokasi),
        }
        if status == "resign":
            resign.append(item)
        elif status == "baru":
            baru.append(item)
    resign.sort(key=lambda item: (item["tanggal"], item["lokasi"].lower(), item["nama"].lower()))
    baru.sort(key=lambda item: (item["lokasi"].lower(), item["nama"].lower()))
    return {
        "active_count": sum(1 for name in seen if active(name)),
        "location_count": len(seen),
        "active_pay": active_pay,
        "inactive_pay": inactive_pay,
        "inactive": inactive,
        "resign": resign,
        "baru": baru,
    }


def person_rows(data: dict) -> list[dict]:
    people = []
    for key, row in data.get("people", {}).items():
        people.append(
            {
                "key": key,
                "lokasi": row.get("lokasi") or "",
                "nama": row.get("nama") or "",
                "username": row.get("username") or "",
                "status": row.get("status") or "aktif",
                "tanggal": row.get("tanggal") or "",
                "catatan": row.get("catatan") or "",
                "saran": row.get("saran") or "",
                "paid": row.get("paid") or 0,
                "testing": bool(row.get("testing")),
                "mutasi": bool(row.get("mutasi")),
                "mutasi_dari": row.get("mutasi_dari") or "",
                "kantor_lain": row.get("kantor_lain") or [],
                "diabaikan": bool(row.get("testing")) or (row.get("status") or "") == "resign",
            }
        )
    return people


def page_rows(roster: dict | None = None) -> dict:
    data = roster if roster is not None else load_roster()
    people = person_rows(data)
    grouped: dict[str, list[dict]] = {}
    for person in people:
        grouped.setdefault(person["lokasi"], []).append(person)
    locations = []
    for name, row in sorted(data.get("locations", {}).items(), key=lambda item: item[0].lower()):
        group = grouped.get(name, [])
        locations.append(
            {
                "name": name,
                "active": bool(row.get("active", True)),
                "people_count": len(group),
                "resign_count": sum(1 for person in group if person["status"] == "resign"),
                "baru_count": sum(1 for person in group if person["status"] == "baru"),
                "paid": sum(float(person["paid"] or 0) for person in group),
            }
        )
    return {
        "meta": data.get("meta") or {},
        "locations": locations,
        "active_count": sum(1 for row in locations if row["active"]),
        "resign_count": sum(1 for row in people if row["status"] == "resign"),
        "baru_count": sum(1 for row in people if row["status"] == "baru"),
    }


def people_page(roster: dict | None = None) -> dict:
    data = roster if roster is not None else load_roster()
    people = person_rows(data)
    people.sort(
        key=lambda row: (
            row["status"] != "resign",
            row["status"] != "baru",
            row["lokasi"].lower(),
            row["nama"].lower(),
        )
    )
    locations = sorted({row["lokasi"] for row in people if row["lokasi"]}, key=str.lower)
    return {
        "meta": data.get("meta") or {},
        "people": people,
        "locations": locations,
        "aktif_count": sum(1 for row in people if row["status"] == "aktif" and not row["testing"]),
        "resign_count": sum(1 for row in people if row["status"] == "resign"),
        "baru_count": sum(1 for row in people if row["status"] == "baru"),
        "testing_count": sum(1 for row in people if row["testing"]),
        "mutasi_count": sum(1 for row in people if row["mutasi"]),
    }


def location_detail(name: str, roster: dict | None = None) -> dict | None:
    data = roster if roster is not None else load_roster()
    locations = data.get("locations") or {}
    match = name if name in locations else ""
    if not match:
        wanted = (name or "").strip().lower()
        match = next((key for key in locations if key.lower() == wanted), "")
    if not match:
        return None
    people = [person for person in person_rows(data) if person["lokasi"] == match]
    people.sort(key=lambda row: (row["status"] != "resign", row["status"] != "baru", row["nama"].lower()))
    return {
        "meta": data.get("meta") or {},
        "name": match,
        "active": bool(locations[match].get("active", True)),
        "people": people,
        "paid": sum(float(person["paid"] or 0) for person in people),
        "resign_count": sum(1 for person in people if person["status"] == "resign"),
        "baru_count": sum(1 for person in people if person["status"] == "baru"),
    }
