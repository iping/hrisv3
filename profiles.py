"""Profil petugas: lokasi, rekening, dan kode transfer.

Diisi dari payroll final (file transfer bank), lalu dipakai lagi setiap bulan.
"""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from calculate_payroll import Payroll, text
from roster import load_roster, person_key

ROOT = Path(__file__).resolve().parent
PROFILE_PATH = ROOT / "output" / "profiles.json"
MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def empty_profiles() -> dict:
    return {"locations": {}, "people": {}}


def load_profiles() -> dict:
    if not PROFILE_PATH.exists():
        return empty_profiles()
    data = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
    data.setdefault("locations", {})
    data.setdefault("people", {})
    return data


def save_profiles(data: dict) -> None:
    PROFILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    PROFILE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def norm(value) -> str:
    raw = text(value).lower().replace("-", " ")
    raw = re.sub(r"[^a-z0-9\s]", " ", raw)
    raw = re.sub(r"\b(s\s*h|s\s*pd|s\s*p)\b", " ", raw)
    return " ".join(raw.split())


def profile_key(lokasi: str, username: str, nama: str) -> str:
    if text(username):
        return person_key(lokasi, username, nama)
    return f"nama|{norm(lokasi)}|{norm(nama)}"


def find_profile(lokasi: str, username: str = "", nama: str = "", data: dict | None = None) -> dict | None:
    profiles = data if data is not None else load_profiles()
    people = profiles.get("people") or {}
    direct = people.get(profile_key(lokasi, username, nama))
    if direct:
        return direct
    wanted_loc = norm(lokasi)
    wanted_user = text(username).lower()
    wanted_name = norm(nama)
    by_user = []
    by_name = []
    for row in people.values():
        if wanted_user and text(row.get("username")).lower() == wanted_user and norm(row.get("lokasi")) == wanted_loc:
            return row
        if wanted_user and text(row.get("username")).lower() == wanted_user:
            by_user.append(row)
        if wanted_name and norm(row.get("nama")) == wanted_name and norm(row.get("lokasi")) == wanted_loc:
            return row
        if wanted_name and norm(row.get("nama")) == wanted_name:
            by_name.append(row)
    if len(by_user) == 1:
        return by_user[0]
    if len(by_name) == 1:
        return by_name[0]
    return None


def location_kode(name: str, data: dict | None = None) -> str:
    profiles = data if data is not None else load_profiles()
    row = (profiles.get("locations") or {}).get(name)
    if row and row.get("kode"):
        return text(row.get("kode"))
    wanted = norm(name)
    for loc, item in (profiles.get("locations") or {}).items():
        if norm(loc) == wanted and item.get("kode"):
            return text(item.get("kode"))
    return ""


def set_location_kode(name: str, kode: str) -> None:
    data = load_profiles()
    row = data["locations"].setdefault(name, {})
    row["kode"] = text(kode)
    row["manual"] = True
    save_profiles(data)


def set_profile(lokasi: str, username: str, nama: str, fields: dict) -> dict:
    data = load_profiles()
    key = profile_key(lokasi, username, nama)
    current = data["people"].get(key)
    if current is None:
        found = find_profile(lokasi, username, nama, data)
        if found and text(found.get("username")).lower() == text(username).lower():
            current = found
            key = profile_key(found.get("lokasi") or lokasi, found.get("username") or username, found.get("nama") or nama)
    if current is None:
        current = {
            "lokasi": lokasi,
            "nama": nama,
            "username": username,
            "nama_rekening": "",
            "bank": "BCA",
            "rekening": "",
            "catatan": "",
        }
    current["lokasi"] = lokasi or current.get("lokasi") or ""
    current["nama"] = nama or current.get("nama") or ""
    current["username"] = username or current.get("username") or ""
    current["nama_rekening"] = text(fields.get("nama_rekening"))
    current["bank"] = text(fields.get("bank")) or "BCA"
    current["rekening"] = re.sub(r"\s+", "", text(fields.get("rekening")))
    current["catatan"] = text(fields.get("catatan"))
    current["manual"] = True
    data["people"][key] = current
    save_profiles(data)
    return current


def attach_banks(rows: list[dict]) -> None:
    data = load_profiles()
    for row in rows:
        profile = find_profile(row.get("lokasi") or "", row.get("username") or "", row.get("nama") or "", data)
        row["nama_rekening"] = (profile or {}).get("nama_rekening") or ""
        row["bank"] = (profile or {}).get("bank") or ""
        row["rekening"] = (profile or {}).get("rekening") or ""
        row["punya_rekening"] = bool(row["rekening"])


def coverage(roster: dict | None = None) -> dict:
    data = load_profiles()
    people = (roster if roster is not None else load_roster()).get("people") or {}
    filled = 0
    for row in people.values():
        profile = find_profile(row.get("lokasi") or "", row.get("username") or "", row.get("nama") or "", data)
        if profile and profile.get("rekening"):
            filled += 1
    return {"filled": filled, "total": len(people), "stored": len(data.get("people") or {})}


def import_final(path: Path) -> dict:
    """Salin rekening dan kode lokasi dari file transfer final. Profil manual tidak ditimpa."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["Accounting"] if "Accounting" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    if not rows:
        return load_profiles()
    header = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(header)}

    def cell(row, name, fallback=0):
        pos = index.get(name, fallback)
        return row[pos] if pos < len(row) else ""

    roster_people = list(load_roster().get("people", {}).values())
    by_loc_name = {(norm(row.get("lokasi")), norm(row.get("nama"))): row for row in roster_people}
    by_name: dict[str, list] = defaultdict(list)
    for row in roster_people:
        by_name[norm(row.get("nama"))].append(row)

    data = load_profiles()
    kode_votes: dict[str, Counter] = defaultdict(Counter)
    for row in rows[1:]:
        if not row or not text(cell(row, "Trx ID", 0)):
            continue
        lokasi = text(cell(row, "lokasi", 11))
        remark = text(cell(row, "Remark", 5))
        prefix = remark.rsplit(" ", 1)[0] if " " in remark else remark
        if lokasi and prefix:
            kode_votes[lokasi][prefix] += 1
        nama = text(cell(row, "Receiver Name", 3))
        if not nama:
            continue
        roster_row = by_loc_name.get((norm(lokasi), norm(nama)))
        if roster_row is None and len(by_name.get(norm(nama), [])) == 1:
            roster_row = by_name[norm(nama)][0]
        username = (roster_row or {}).get("username") or ""
        home = (roster_row or {}).get("lokasi") or lokasi
        key = profile_key(home, username, (roster_row or {}).get("nama") or nama)
        current = data["people"].get(key, {})
        if current.get("manual") and current.get("rekening"):
            continue
        hris_nama = (roster_row or {}).get("nama") or nama
        data["people"][key] = {
            "lokasi": home,
            "nama": hris_nama,
            "username": username,
            "nama_rekening": nama,
            "bank": text(cell(row, "Transfer Type", 1)) or "BCA",
            "rekening": re.sub(r"\s+", "", text(cell(row, "Credited Account", 2))),
            "catatan": "" if text(cell(row, "update", 9)) in {"", "-"} else text(cell(row, "update", 9)),
            "manual": False,
            "sumber": path.name,
        }

    for lokasi, votes in kode_votes.items():
        current = data["locations"].get(lokasi, {})
        if current.get("manual") and current.get("kode"):
            continue
        data["locations"][lokasi] = {"kode": votes.most_common(1)[0][0], "manual": False}
    save_profiles(data)
    return data


def remark_period(period: str, month: str) -> str:
    end = ""
    if "s/d" in (period or ""):
        end = period.split("s/d")[-1].strip()[:10]
    elif month and len(month) >= 7:
        end = month[:7] + "-01"
    try:
        parsed = datetime.strptime(end, "%Y-%m-%d")
    except ValueError:
        parsed = datetime.now()
    return f"{MONTHS[parsed.month - 1]}{parsed.strftime('%y')}"


def _slip_fields(officer) -> dict:
    return {
        "basic": officer.basic,
        "jabatan": officer.amounts["jabatan"],
        "bpjs_kes": officer.amounts["bpjs_kes"],
        "bpjs_tk": officer.amounts["bpjs_tk"],
        "prorate": officer.amounts["prorate"],
        "lain_tunjangan": officer.amounts["lain_tunjangan"],
        "jam_lembur": round(officer.jam_lembur, 2),
        "lembur_hris": officer.lembur_hris,
        "lembur_hitung": officer.lembur_hitung,
        "alpha": officer.alpha,
        "denda_hris": officer.denda_alpa_hris,
        "denda_hitung": officer.denda_alpa_hitung,
        "potongan": officer.potongan_tetap,
        "paid": officer.paid,
        "hitung": officer.gaji_hitung,
        "selisih": officer.selisih_gaji,
    }


def transfer_candidates(payroll: Payroll, roster: dict | None = None) -> list[dict]:
    """Satu baris per petugas di slip, lengkap dengan calon baris Accounting."""
    from presence import payroll_blocked_usernames

    data = load_profiles()
    stored = roster if roster is not None else load_roster()
    locations = stored.get("locations") or {}
    people = stored.get("people") or {}
    blocked = payroll_blocked_usernames()
    label = remark_period(payroll.period, payroll.month)
    rows = []
    for officer in payroll.officers:
        if officer.username.strip().lower() in blocked:
            continue
        profile = find_profile(officer.lokasi, officer.username, officer.nama, data) or {}
        kode = location_kode(officer.lokasi, data)
        rekening = text(profile.get("rekening"))
        receiver = text(profile.get("nama_rekening")) or officer.nama
        place = locations.get(officer.lokasi)
        reasons = []
        person = people.get(person_key(officer.lokasi, officer.username, officer.nama)) or {}
        if officer.paid <= 0:
            reasons.append("Gaji nol")
        if place is not None and not place.get("active", True):
            reasons.append("Lokasi tidak aktif")
        if person.get("testing"):
            reasons.append("Akun testing")
        if person.get("status") == "resign":
            reasons.append("Sudah resign")
        rows.append(
            {
                "key": person_key(officer.lokasi, officer.username, officer.nama),
                "username": officer.username,
                "nama": officer.nama,
                "lokasi": officer.lokasi,
                "bank": text(profile.get("bank")) or ("BCA" if rekening else ""),
                "rekening": rekening,
                "penerima": receiver,
                "amount": int(round(officer.paid)),
                "remark": f"{kode} {label}".strip(),
                "kode": kode,
                "label": label,
                "include": not reasons,
                "reason": " · ".join(reasons),
                **_slip_fields(officer),
            }
        )
    return rows


def _system_notes(row: dict) -> list[str]:
    notes = []
    if not row["rekening"]:
        notes.append("rekening kosong")
    elif norm(row["penerima"]) != norm(row["nama"]):
        notes.append("beda nama")
    if not row["kode"]:
        notes.append("kode lokasi kosong")
    return notes


def present_transfer(payroll: Payroll, roster: dict | None = None, edits: dict | None = None) -> list[dict]:
    """Gabungkan slip, accounting, dan suntingan. Trx ID sama dengan file transfer."""
    edits = edits or {}
    rows = []
    for source in transfer_candidates(payroll, roster):
        edit = edits.get(source["key"]) or {}
        row = dict(source)
        row["deleted"] = bool(edit.get("deleted"))
        row["catatan"] = text(edit.get("catatan"))
        overrides = []
        for field in ("bank", "rekening", "penerima", "remark"):
            if field in edit and edit[field] is not None:
                row[field] = text(edit[field])
                overrides.append(field)
        if edit.get("amount") is not None and edit.get("amount") != "":
            row["amount"] = int(edit["amount"])
            overrides.append("amount")
        if row["catatan"]:
            overrides.append("catatan")
        notes = _system_notes(row)
        row["system"] = "-" if not notes else " / ".join(notes)
        bits = list(notes)
        if row["catatan"]:
            bits.append(row["catatan"])
        row["update"] = " / ".join(bits) if bits else "-"
        row["check"] = not notes and row["include"] and not row["deleted"]
        row["siap"] = "ok" if row["rekening"] and row["kode"] and row["include"] and not row["deleted"] else ""
        row["overrides"] = overrides
        if row["deleted"]:
            row["state"] = "keluar"
            row["status"] = "Dikeluarkan"
        elif row["include"]:
            row["state"] = "ikut"
            row["status"] = "Ikut transfer"
        else:
            row["state"] = "luar"
            row["status"] = row["reason"] or "Tidak masuk transfer"
        rows.append(row)

    included = [row for row in rows if row["include"] and not row["deleted"]]
    included.sort(key=lambda row: (row["remark"].lower(), row["penerima"].lower(), row["lokasi"].lower()))
    numbered = {}
    for number, row in enumerate(included, start=1):
        numbered[row["key"]] = f"SITE-{row['label']}-{number:03d}"
    for row in rows:
        row["trx"] = numbered.get(row["key"], "")
    return rows


def transfer_rows(payroll: Payroll, roster: dict | None = None, edits: dict | None = None) -> list[dict]:
    rows = [row for row in present_transfer(payroll, roster, edits) if row["include"] and not row["deleted"]]
    rows.sort(key=lambda row: (row["remark"].lower(), row["penerima"].lower(), row["lokasi"].lower()))
    return rows


def write_transfer(payroll: Payroll, output: Path, roster: dict | None = None, edits: dict | None = None) -> list[dict]:
    rows = transfer_rows(payroll, roster, edits)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Accounting"
    headers = [
        "Trx ID",
        "Transfer Type",
        "Credited Account",
        "Receiver Name",
        "Amount",
        "Remark",
        None,
        "rek",
        "check 2",
        "update",
        "check 3",
        "lokasi",
        "check 4",
    ]
    sheet.append(headers)
    for row in rows:
        sheet.append(
            [
                row["trx"],
                row["bank"],
                row["rekening"],
                row["penerima"],
                row["amount"],
                row["remark"],
                None,
                row["rekening"],
                row["check"],
                row["update"],
                None,
                row["lokasi"],
                row["siap"],
            ]
        )
    summary = workbook.create_sheet("Rekap")
    summary.append(["Row Labels", "Sum of Amount"])
    totals: dict[str, int] = defaultdict(int)
    for row in rows:
        totals[row["remark"]] += row["amount"]
    for remark, amount in sorted(totals.items(), key=lambda item: item[0].lower()):
        summary.append([remark, amount])
    summary.append(["Grand Total", sum(totals.values())])
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return rows
