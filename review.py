"""Ringkasan yang bisa dibaca dari hasil hitung payroll."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from calculate_payroll import (
    Officer,
    Payroll,
    load_officers,
    money,
    rupiah,
    sheet_rows,
    text,
)
from roster import roster_for_view

PAID_HEADERS = (
    "paid salary",
    "gaji hris",
    "gaji dibayar",
    "gaji bersih",
    "gaji",
)
NAME_HEADERS = ("nama", "petugas", "name")
LOCATION_HEADERS = ("lokasi", "location")
USER_HEADERS = ("username", "user")


def is_testing(officer: Officer) -> bool:
    blob = f"{officer.lokasi} {officer.nama} {officer.username}".lower()
    return "testing" in blob


def person_rows(officers: list[Officer]) -> list[dict]:
    rows = []
    for officer in officers:
        notes = []
        lembur_gap = round(officer.lembur_hitung - officer.lembur_hris, 2)
        denda_gap = round(officer.denda_alpa_hitung - officer.denda_alpa_hris, 2)
        if abs(lembur_gap) >= 1:
            notes.append(
                f"Lembur slip {rupiah(officer.lembur_hris)}, hasil durasi {rupiah(officer.lembur_hitung)}."
            )
        if abs(denda_gap) >= 1:
            if officer.alpha == 0:
                notes.append(f"Slip memotong alpa {rupiah(officer.denda_alpa_hris)} tanpa status A di lokasi ini.")
            elif officer.denda_alpa_hris == 0:
                notes.append(
                    f"{officer.alpha} status A, slip belum memotong. Hitung {rupiah(officer.denda_alpa_hitung)}."
                )
            else:
                notes.append(
                    f"Denda alpa slip {rupiah(officer.denda_alpa_hris)}, hitung {rupiah(officer.denda_alpa_hitung)} dari {officer.alpha} status A."
                )
        if is_testing(officer) and (officer.paid or officer.basic):
            notes.append("Akun atau lokasi testing tetap punya gaji.")
        if abs(officer.selisih_gaji) < 1 and not notes:
            continue
        rows.append(
            {
                "lokasi": officer.lokasi,
                "nama": officer.nama,
                "username": officer.username,
                "hris": officer.paid,
                "hitung": officer.gaji_hitung,
                "selisih": officer.selisih_gaji,
                "note": " ".join(notes),
            }
        )
    rows.sort(key=lambda row: (-abs(row["selisih"]), row["lokasi"].lower(), row["nama"].lower()))
    return rows


def duplicate_users(officers: list[Officer]) -> list[dict]:
    grouped: dict[str, list[Officer]] = defaultdict(list)
    for officer in officers:
        if officer.username and (officer.paid or officer.basic):
            grouped[officer.username.lower()].append(officer)
    rows = []
    for people in grouped.values():
        if len(people) < 2:
            continue
        paid = sum(officer.paid for officer in people)
        detail = "; ".join(f"{officer.lokasi} {rupiah(officer.paid)}" for officer in people)
        rows.append(
            {
                "lokasi": people[0].lokasi,
                "nama": people[0].nama,
                "username": people[0].username,
                "amount": paid,
                "note": detail,
            }
        )
    rows.sort(key=lambda row: -row["amount"])
    return rows


def unmatched_people(payroll: Payroll) -> list[dict]:
    grouped: dict[tuple[str, str], dict] = {}
    for record in payroll.unmatched_overtime:
        key = (record.lokasi.lower(), (record.username or record.nama).lower())
        bucket = grouped.setdefault(
            key,
            {
                "lokasi": record.lokasi,
                "nama": record.nama,
                "username": record.username,
                "amount": 0.0,
                "jam": 0.0,
                "n": 0,
            },
        )
        bucket["amount"] += record.fee_hitung
        bucket["jam"] += record.jam
        bucket["n"] += 1
    rows = []
    for bucket in grouped.values():
        rows.append(
            {
                "lokasi": bucket["lokasi"],
                "nama": bucket["nama"],
                "username": bucket["username"],
                "amount": bucket["amount"],
                "note": f"{bucket['n']} record, {bucket['jam']:.1f} jam, tidak ada slip di lokasi yang sama.",
            }
        )
    rows.sort(key=lambda row: -row["amount"])
    return rows


def lead_paragraphs(payroll: Payroll) -> list[str]:
    officers = payroll.officers
    hris = sum(officer.paid for officer in officers)
    hitung = sum(officer.gaji_hitung for officer in officers)
    selisih = hitung - hris
    lembur = sum(officer.lembur_hitung - officer.lembur_hris for officer in officers)
    denda = sum(officer.denda_alpa_hitung - officer.denda_alpa_hris for officer in officers)
    paid_count = sum(1 for officer in officers if officer.paid or officer.basic)
    zero_count = len(officers) - paid_count
    period = payroll.period or "periode di file"
    paragraphs = [
        f"Periode {period}. {paid_count} petugas masuk slip, {zero_count} baris gaji nol dipisahkan dari angka transfer. "
        f"Gaji yang tertulis di HRIS {rupiah(hris)}. Hasil hitung {rupiah(hitung)}."
    ]
    if abs(selisih) < 1:
        paragraphs.append("Gaji HRIS dan hasil hitung sama. Komponen master disalin dari slip, lembur mengikuti durasi, denda alpa mengikuti gaji pokok dibagi 25.")
    else:
        direction = "di atas" if selisih > 0 else "di bawah"
        bits = []
        if abs(lembur) >= 1:
            if lembur > 0:
                bits.append(f"fee lembur dari durasi lebih besar {rupiah(lembur)} daripada tunjangan di slip")
            else:
                bits.append(f"tunjangan lembur di slip lebih besar {rupiah(abs(lembur))} daripada fee dari durasi")
        if abs(denda) >= 1:
            if denda > 0:
                bits.append(f"denda alpa hasil hitung lebih besar {rupiah(denda)} daripada potongan di slip")
            else:
                bits.append(f"potongan alpa di slip lebih besar {rupiah(abs(denda))} daripada hitung pokok/25")
        joined = " dan ".join(bits) if bits else "ada selisih komponen"
        paragraphs.append(
            f"Hasil hitung {rupiah(abs(selisih))} {direction} gaji HRIS karena {joined}. "
            "Jabatan, BPJS, dan prorate tidak diubah."
        )
    unmatched = sum(record.fee_hitung for record in payroll.unmatched_overtime)
    if unmatched:
        paragraphs.append(
            f"Ada fee lembur {rupiah(unmatched)} yang record-nya tidak menempel ke slip di lokasi yang sama. "
            "Angka itu tidak masuk gaji dihitung."
        )
    return paragraphs


def build_decisions(payroll: Payroll) -> list[dict]:
    decisions = []
    dupes = duplicate_users(payroll.officers)
    if dupes:
        decisions.append(
            {
                "level": "tinggi",
                "title": "Username yang sama digaji di lebih dari satu lokasi",
                "text": (
                    f"{len(dupes)} username muncul dua kali. Kalau keduanya ditransfer, gajinya ikut dua kali. "
                    "Tandai mutasi kantor, lalu resign di kantor yang lama."
                ),
                "rows": dupes,
            }
        )
    loose = unmatched_people(payroll)
    if loose:
        decisions.append(
            {
                "level": "tinggi",
                "title": "Lembur tidak menempel ke slip",
                "text": "Record ini punya fee, tetapi tidak ada baris payroll di lokasi yang sama.",
                "rows": loose,
            }
        )
    from roster import load_roster, person_key

    flagged = load_roster().get("people") or {}
    testing = []
    for officer in payroll.officers:
        if not (is_testing(officer) and (officer.paid or officer.basic)):
            continue
        marked = bool((flagged.get(person_key(officer.lokasi, officer.username, officer.nama)) or {}).get("testing"))
        testing.append(
            {
                "lokasi": officer.lokasi,
                "nama": officer.nama,
                "username": officer.username,
                "amount": officer.paid,
                "note": "Tidak masuk transfer." if marked else "Belum ditandai. Buka profil petugas untuk mengabaikan.",
            }
        )
    if testing:
        decisions.append(
            {
                "level": "tinggi",
                "title": "Akun testing ikut digaji",
                "text": "Baris ini sebaiknya tidak masuk transfer sebelum dicek.",
                "rows": testing,
            }
        )
    late_people = [officer for officer in payroll.officers if officer.late and officer.denda_telat == 0 and (officer.paid or officer.basic)]
    late_status = sum(officer.late for officer in late_people)
    if late_people:
        decisions.append(
            {
                "level": "info",
                "title": "Keterlambatan tidak dipotong di slip",
                "text": (
                    f"{late_status} status terlambat pada {len(late_people)} petugas. "
                    "Kolom denda keterlambatan di file ini kosong, jadi tidak menambah atau mengurangi gaji dihitung."
                ),
                "rows": [],
            }
        )
    return decisions


def location_rows(payroll: Payroll) -> list[dict]:
    grouped: dict[str, list[Officer]] = defaultdict(list)
    for officer in payroll.officers:
        grouped[officer.lokasi].append(officer)
    rows = []
    for lokasi, people in grouped.items():
        hris = sum(officer.paid for officer in people)
        hitung = sum(officer.gaji_hitung for officer in people)
        rows.append(
            {
                "lokasi": lokasi,
                "n": len(people),
                "hris": hris,
                "hitung": hitung,
                "selisih": round(hitung - hris, 2),
            }
        )
    rows.sort(key=lambda row: (-abs(row["selisih"]), row["lokasi"].lower()))
    return rows


def _header_map(headers: list) -> dict[str, int]:
    mapping = {}
    for pos, header in enumerate(headers):
        name = " ".join(text(header).lower().split())
        if name and name not in mapping:
            mapping[name] = pos
    return mapping


def _pick(mapping: dict[str, int], options: tuple[str, ...]) -> int | None:
    for option in options:
        if option in mapping:
            return mapping[option]
    return None


def load_prior_rows(path: Path) -> list[dict]:
    """Baca gaji final bulan lalu dari tarikan HRIS atau sheet yang punya nama dan gaji."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Payroll" in wb.sheetnames:
            rows = sheet_rows(wb, "Payroll")
            headers = _header_map(rows[0] if rows else [])
            if _pick(headers, ("paid salary",)) is not None and _pick(headers, NAME_HEADERS) is not None:
                officers = load_officers(rows)
                return [
                    {
                        "lokasi": officer.lokasi,
                        "nama": officer.nama,
                        "username": officer.username,
                        "paid": officer.paid,
                    }
                    for officer in officers
                    if officer.paid or officer.basic
                ]
        found = []
        for sheet in wb.sheetnames:
            rows = list(wb[sheet].iter_rows(values_only=True))
            header_at = None
            mapping = {}
            for index, row in enumerate(rows[:25]):
                mapping = _header_map(row)
                if _pick(mapping, NAME_HEADERS) is not None and _pick(mapping, PAID_HEADERS) is not None:
                    header_at = index
                    break
            if header_at is None:
                continue
            name_at = _pick(mapping, NAME_HEADERS)
            paid_at = _pick(mapping, PAID_HEADERS)
            location_at = _pick(mapping, LOCATION_HEADERS)
            user_at = _pick(mapping, USER_HEADERS)
            for row in rows[header_at + 1 :]:
                if not row or name_at >= len(row) or not text(row[name_at]):
                    continue
                paid = money(row[paid_at]) if paid_at < len(row) else 0
                if paid == 0:
                    continue
                found.append(
                    {
                        "lokasi": text(row[location_at]) if location_at is not None and location_at < len(row) else "",
                        "nama": text(row[name_at]),
                        "username": text(row[user_at]) if user_at is not None and user_at < len(row) else "",
                        "paid": paid,
                    }
                )
            if found:
                return found
    finally:
        wb.close()
    raise ValueError("File bulan lalu tidak punya kolom nama dan gaji yang dikenali.")


def _match_key(lokasi: str, username: str, nama: str, use_location: bool) -> tuple[str, str]:
    identity = text(username).lower() or text(nama).lower()
    place = text(lokasi).lower() if use_location else ""
    return (place, identity)


def compare_prior(payroll: Payroll, prior_rows: list[dict]) -> dict:
    use_location = any(row.get("lokasi") for row in prior_rows)
    prior_map: dict[tuple[str, str], dict] = {}
    for row in prior_rows:
        prior_map[_match_key(row.get("lokasi", ""), row.get("username", ""), row.get("nama", ""), use_location)] = row
    current_map = {}
    for officer in payroll.officers:
        if not (officer.paid or officer.basic):
            continue
        current_map[_match_key(officer.lokasi, officer.username, officer.nama, use_location)] = officer

    changes = []
    same = 0
    for key, officer in current_map.items():
        previous = prior_map.get(key)
        if previous is None:
            changes.append(
                {
                    "status": "Baru",
                    "lokasi": officer.lokasi,
                    "nama": officer.nama,
                    "username": officer.username,
                    "lalu": 0,
                    "ini": officer.paid,
                    "selisih": officer.paid,
                }
            )
            continue
        gap = round(officer.paid - previous["paid"], 2)
        if abs(gap) < 1:
            same += 1
            continue
        changes.append(
            {
                "status": "Naik" if gap > 0 else "Turun",
                "lokasi": officer.lokasi,
                "nama": officer.nama,
                "username": officer.username,
                "lalu": previous["paid"],
                "ini": officer.paid,
                "selisih": gap,
            }
        )
    left = []
    for key, previous in prior_map.items():
        if key in current_map:
            continue
        left.append(
            {
                "status": "Tidak ada bulan ini",
                "lokasi": previous.get("lokasi") or "",
                "nama": previous.get("nama") or "",
                "username": previous.get("username") or "",
                "lalu": previous["paid"],
                "ini": 0,
                "selisih": -previous["paid"],
            }
        )
    rows = changes + left
    rows.sort(key=lambda row: -abs(row["selisih"]))
    prior_total = sum(row["paid"] for row in prior_rows)
    current_total = sum(officer.paid for officer in payroll.officers if officer.paid or officer.basic)
    return {
        "prior_total": prior_total,
        "current_total": current_total,
        "selisih": round(current_total - prior_total, 2),
        "same": same,
        "naik": sum(1 for row in rows if row["status"] == "Naik"),
        "turun": sum(1 for row in rows if row["status"] == "Turun"),
        "baru": sum(1 for row in rows if row["status"] == "Baru"),
        "lepas": sum(1 for row in rows if row["status"] == "Tidak ada bulan ini"),
        "rows": rows,
    }


def review(payroll: Payroll, prior_rows: list[dict] | None = None, prior_name: str = "", roster: dict | None = None) -> dict:
    officers = payroll.officers
    paid_count = sum(1 for officer in officers if officer.paid or officer.basic)
    return {
        "source_name": payroll.source_name,
        "prior_name": prior_name,
        "month": payroll.month,
        "period": payroll.period,
        "headcount": len(officers),
        "paid_count": paid_count,
        "zero_count": len(officers) - paid_count,
        "hris": sum(officer.paid for officer in officers),
        "hitung": sum(officer.gaji_hitung for officer in officers),
        "selisih": sum(officer.gaji_hitung - officer.paid for officer in officers),
        "lembur_hris": sum(officer.lembur_hris for officer in officers),
        "lembur_hitung": sum(officer.lembur_hitung for officer in officers),
        "denda_hris": sum(officer.denda_alpa_hris for officer in officers),
        "denda_hitung": sum(officer.denda_alpa_hitung for officer in officers),
        "paragraphs": lead_paragraphs(payroll),
        "decisions": build_decisions(payroll),
        "locations": location_rows(payroll),
        "people": person_rows(officers),
        "prior": compare_prior(payroll, prior_rows) if prior_rows else None,
        "kelola": roster_for_view(payroll, roster) if roster is not None else None,
        "rules": [
            "Gaji pokok, tunjangan jabatan, BPJS, dan prorate disalin dari slip HRIS.",
            "Lembur dihitung dari durasi: sampai 4 jam Rp 50.000, sampai 8 jam Rp 75.000, di atas 8 jam Rp 100.000. Di bawah 1 jam tetap masuk blok pertama.",
            "Denda alpa = floor(gaji pokok / 25) × jumlah status A. Sel 'A | A' dihitung dua.",
            "Denda keterlambatan mengikuti slip. Di file contoh angkanya nol, jadi tidak menambah potongan baru.",
        ],
    }
