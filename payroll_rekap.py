#!/usr/bin/env python3
"""Rekap tutup payroll dari file Excel yang sudah ditarik HRIS.

Tool ini tidak mengubah HRIS dan tidak menarik data baru. Ia membaca
workbook hasil tarikan (tab Payroll, Attendance, Overtime, Anomaly) lalu
menulis satu file yang dipakai saat tutup gaji:

  Ringkasan, Daftar Transfer, Rekap Lokasi, Rekap Komponen,
  Rekonsiliasi Lembur, Rekap Absensi, Cek Denda, Perlu Dicek, Tidak Digaji.

Contoh:
  python3 payroll_rekap.py "../payroll_report/payroll_2026-09 oktober.xlsx"
  python3 payroll_rekap.py tarikan.xlsx -o rekap.xlsx
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

IDENTITY = {
    "Lokasi",
    "Nama",
    "Spot ID",
    "Officer ID",
    "Numeric ID",
    "Username",
    "Basic Salary",
    "Paid Salary",
}
DEDUCTION_HEADERS = {
    "Denda Keterlambatan",
    "Denda Ketidakhadiran",
    "Potongan Asuransi Kesehatan BPJS",
    "Potongan Asuransi Karyawan BPJS",
    "Fraud",
    "Deposit",
}

# Pembanding denda. Ini aturan default tool tarikan lama, bukan bukti
# bahwa spot memakai angka yang sama. Ditampilkan supaya selisihnya kelihatan.
DEFAULT_LATE_PENALTY = 50_000

GROUPS = [
    ("jabatan", "Tunjangan Jabatan", "penambah"),
    ("lembur", "Tunjangan Lembur", "penambah"),
    ("bpjs_kes", "BPJS Kesehatan (tunjangan)", "penambah"),
    ("bpjs_tk", "BPJS Ketenagakerjaan (tunjangan)", "penambah"),
    ("prorate", "Prorate", "penambah"),
    ("lain_tunjangan", "Tunjangan Lain", "penambah"),
    ("denda_telat", "Denda Keterlambatan", "potongan"),
    ("denda_alpa", "Denda Ketidakhadiran", "potongan"),
    ("pot_bpjs_kes", "Potongan BPJS Kesehatan", "potongan"),
    ("pot_bpjs_tk", "Potongan BPJS Ketenagakerjaan", "potongan"),
    ("lain_potongan", "Potongan Lain", "potongan"),
]
GROUP_LABEL = {key: label for key, label, _side in GROUPS}
PENAMBAH = [key for key, _label, side in GROUPS if side == "penambah"]
POTONGAN = [key for key, _label, side in GROUPS if side == "potongan"]

NAVY = "1F4E79"
TEAL = "0F6E6E"
AMBER = "8A5A00"
RED = "9C2F2F"
GREEN = "1E6B3A"
PALE = "F4F7FB"
PALE_TEAL = "E6F3F3"
PALE_AMBER = "FFF6E5"
PALE_RED = "FDECEC"
PALE_GREEN = "E7F6EC"
WHITE = "FFFFFF"
LINE = "D9E2EF"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_pale = PatternFill("solid", fgColor=PALE)
fill_total = PatternFill("solid", fgColor="D6E3F0")
fill_section = PatternFill("solid", fgColor="E8EEF6")
font_header = Font(name="Calibri", bold=True, color=WHITE, size=11)
font_base = Font(name="Calibri", size=11, color="1A1A1A")
font_bold = Font(name="Calibri", size=11, color="1A1A1A", bold=True)
font_title = Font(name="Calibri", size=16, color=NAVY, bold=True)
font_section = Font(name="Calibri", size=12, color=NAVY, bold=True)
font_muted = Font(name="Calibri", size=11, color="4A5568")
font_red = Font(name="Calibri", size=11, color=RED, bold=True)
font_green = Font(name="Calibri", size=11, color=GREEN, bold=True)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
right_align = Alignment(horizontal="right", vertical="center")


def money(value) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def key_of(lokasi, username, nama="") -> tuple[str, str]:
    user = text(username).lower()
    if not user:
        user = text(nama).lower()
    return (text(lokasi).lower(), user)


def tokens(cell) -> list[str]:
    raw = text(cell)
    if not raw:
        return []
    return [part.strip() for part in raw.split("|") if part.strip()]


def is_testing(*parts) -> bool:
    blob = " ".join(text(part) for part in parts).lower()
    return "testing" in blob or "test officer" in blob


def group_for(header: str, side: str) -> str:
    name = text(header)
    if name.startswith("Tunjangan Lembur") or name == "Overtime":
        return "lembur"
    if name == "Tunjangan Jabatan":
        return "jabatan"
    if name == "Asuransi Kesehatan BPJS":
        return "bpjs_kes"
    if name == "Asuransi Karyawan BPJS":
        return "bpjs_tk"
    if name == "Prorate":
        return "prorate" if side == "allow" else "lain_potongan"
    if name == "Denda Keterlambatan":
        return "denda_telat"
    if name == "Denda Ketidakhadiran":
        return "denda_alpa"
    if name == "Potongan Asuransi Kesehatan BPJS":
        return "pot_bpjs_kes"
    if name == "Potongan Asuransi Karyawan BPJS":
        return "pot_bpjs_tk"
    return "lain_tunjangan" if side == "allow" else "lain_potongan"


@dataclass
class Officer:
    lokasi: str
    nama: str
    username: str
    numeric_id: str
    basic: float
    paid: float
    amounts: dict[str, float] = field(default_factory=lambda: defaultdict(float))

    @property
    def penambah(self) -> float:
        return self.basic + sum(self.amounts[key] for key in PENAMBAH)

    @property
    def potongan(self) -> float:
        return sum(abs(self.amounts[key]) for key in POTONGAN)

    @property
    def selisih(self) -> float:
        return round(self.penambah - self.potongan - self.paid, 2)

    @property
    def digaji(self) -> bool:
        return abs(self.paid) > 0.5 or abs(self.basic) > 0.5 or self.penambah > self.basic + 0.5 or self.potongan > 0.5


def sheet_rows(wb, name: str) -> list[tuple]:
    if name not in wb.sheetnames:
        return []
    return list(wb[name].iter_rows(values_only=True))


def read_meta(rows: list[tuple]) -> dict[str, str]:
    meta = {}
    for row in rows:
        if not row or row[0] is None:
            continue
        label = text(row[0])
        value = text(row[1]) if len(row) > 1 else ""
        if label and value:
            meta[label] = value
    return meta


def load_payroll(rows: list[tuple]) -> list[Officer]:
    if not rows:
        raise SystemExit("Tab Payroll kosong atau tidak ada.")
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    for required in ("Lokasi", "Nama", "Username", "Basic Salary", "Paid Salary"):
        if required not in index:
            raise SystemExit(f"Kolom '{required}' tidak ada di tab Payroll.")

    deduction_at = None
    for pos, name in enumerate(headers):
        if pos < 8:
            continue
        if name in DEDUCTION_HEADERS:
            deduction_at = pos
            break
    if deduction_at is None:
        deduction_at = len(headers)

    officers = []
    for row in rows[1:]:
        if not row or not text(row[index["Lokasi"]]):
            continue
        officer = Officer(
            lokasi=text(row[index["Lokasi"]]),
            nama=text(row[index["Nama"]]),
            username=text(row[index["Username"]]),
            numeric_id=text(row[index["Numeric ID"]]) if "Numeric ID" in index else "",
            basic=money(row[index["Basic Salary"]]),
            paid=money(row[index["Paid Salary"]]),
        )
        for pos, name in enumerate(headers):
            if pos < 8 or not name or name in IDENTITY:
                continue
            amount = money(row[pos]) if pos < len(row) else 0.0
            if amount == 0:
                continue
            side = "allow" if pos < deduction_at else "deduct"
            officer.amounts[group_for(name, side)] += amount
        officers.append(officer)
    return officers


def load_attendance(rows: list[tuple]) -> dict[tuple[str, str], dict]:
    stats: dict[tuple[str, str], dict] = {}
    if not rows:
        return stats
    headers = [text(cell) for cell in rows[0]]
    for row in rows[1:]:
        if not row or not text(row[0]):
            continue
        lokasi, nama, username = text(row[0]), text(row[1]), text(row[3]) if len(row) > 3 else ""
        bucket = stats.setdefault(
            key_of(lokasi, username, nama),
            {
                "lokasi": lokasi,
                "nama": nama,
                "username": username,
                "hari": 0,
                "hari_kode": defaultdict(int),
                "status": defaultdict(int),
                "ganda": 0,
            },
        )
        for pos, header in enumerate(headers):
            if pos < 4 or not header:
                continue
            parts = tokens(row[pos] if pos < len(row) else None)
            if not parts:
                continue
            bucket["hari"] += 1
            if len(parts) > 1:
                bucket["ganda"] += 1
            seen = set()
            for part in parts:
                bucket["status"][part] += 1
                seen.add(part)
            for part in seen:
                bucket["hari_kode"][part] += 1
    return stats


def load_overtime(rows: list[tuple]) -> dict[tuple[str, str], dict]:
    grouped: dict[tuple[str, str], dict] = {}
    if not rows:
        return grouped
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    for row in rows[1:]:
        if not row or not any(row):
            continue
        lokasi = text(row[index.get("Lokasi", 2)])
        nama = text(row[index.get("Petugas", 3)])
        username = text(row[index.get("Username", 4)])
        if not lokasi and not nama:
            continue
        bucket = grouped.setdefault(
            key_of(lokasi, username, nama),
            {
                "lokasi": lokasi,
                "nama": nama,
                "username": username,
                "fee": 0.0,
                "jam": 0.0,
                "record": 0,
            },
        )
        bucket["fee"] += money(row[index.get("Overtime Fee", 9)]) if index.get("Overtime Fee", 9) < len(row) else 0
        jam_col = index.get("Durasi Lembur (Jam)", 8)
        bucket["jam"] += money(row[jam_col]) if jam_col < len(row) else 0
        bucket["record"] += 1
        if nama:
            bucket["nama"] = nama
    return grouped


def load_resign(rows: list[tuple]) -> dict[tuple[str, str], str]:
    found = {}
    if not rows:
        return found
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    if "Tipe Anomali" not in index:
        return found
    for row in rows[1:]:
        if not row or text(row[index["Tipe Anomali"]]) != "resigned_in_period":
            continue
        lokasi = text(row[index.get("Lokasi", 1)])
        nama = text(row[index.get("Nama", 5)])
        username = text(row[index.get("Username", 4)])
        tanggal = text(row[index.get("Tanggal", 6)])
        found[key_of(lokasi, username, nama)] = tanggal
    return found


def load_anomaly_counts(rows: list[tuple]) -> dict[str, int]:
    counts: dict[str, int] = defaultdict(int)
    if not rows:
        return counts
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    if "Tipe Anomali" not in index:
        return counts
    for row in rows[1:]:
        if row and text(row[index["Tipe Anomali"]]):
            counts[text(row[index["Tipe Anomali"]])] += 1
    return counts


@dataclass
class Report:
    source_name: str
    month: str
    period: str
    officers: list[Officer]
    attendance: dict
    overtime: dict
    resign: dict
    anomaly_counts: dict


def build_report(path: Path) -> Report:
    wb = load_workbook(path, read_only=True, data_only=True)
    meta = read_meta(sheet_rows(wb, "Anomaly Summary"))
    report = Report(
        source_name=path.name,
        month=meta.get("Bulan payroll", ""),
        period=meta.get("Periode pemeriksaan", ""),
        officers=load_payroll(sheet_rows(wb, "Payroll")),
        attendance=load_attendance(sheet_rows(wb, "Attendance")),
        overtime=load_overtime(sheet_rows(wb, "Overtime")),
        resign=load_resign(sheet_rows(wb, "Anomaly")),
        anomaly_counts=load_anomaly_counts(sheet_rows(wb, "Anomaly")),
    )
    wb.close()
    return report


def officer_map(report: Report) -> dict[tuple[str, str], Officer]:
    return {key_of(o.lokasi, o.username, o.nama): o for o in report.officers}


def location_issues(report: Report) -> list[dict]:
    """Username yang gajinya, absensinya, atau lemburnya tidak di lokasi yang sama."""
    by_user: dict[str, list[Officer]] = defaultdict(list)
    for officer in report.officers:
        user = text(officer.username).lower()
        if user:
            by_user[user].append(officer)

    attendance_at: dict[str, set[str]] = defaultdict(set)
    overtime_at: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    names: dict[str, str] = {}
    for bucket in report.attendance.values():
        user = text(bucket["username"]).lower()
        if not user:
            continue
        attendance_at[user].add(bucket["lokasi"])
        names.setdefault(user, bucket["nama"])
    for item in report.overtime.values():
        user = text(item["username"]).lower()
        if not user or item["fee"] <= 0:
            continue
        overtime_at[user][item["lokasi"]] += item["fee"]
        names.setdefault(user, item["nama"])

    issues = []
    for user in sorted(set(by_user) | set(attendance_at) | set(overtime_at)):
        people = by_user.get(user, [])
        paid_people = [officer for officer in people if officer.digaji]
        pay_locs = {officer.lokasi for officer in paid_people}
        att_locs = attendance_at.get(user, set())
        ot_locs = overtime_at.get(user, {})
        att_outside = sorted(att_locs - pay_locs)
        ot_outside = {lokasi: fee for lokasi, fee in ot_locs.items() if lokasi not in pay_locs}
        missing_att = sorted(pay_locs - att_locs) if len(paid_people) > 1 else []
        if len(paid_people) < 2 and not ot_outside and not (paid_people and att_outside):
            continue
        notes = []
        if len(paid_people) > 1:
            detail = "; ".join(f"{officer.lokasi} {rupiah(officer.paid)}" for officer in paid_people)
            notes.append(f"Digaji di {len(paid_people)} lokasi ({detail}).")
        if missing_att:
            notes.append("Absensi lokasi penggajian tidak ada di: " + ", ".join(missing_att) + ".")
        if att_outside:
            notes.append("Absensi tercatat di " + ", ".join(att_outside) + ", tanpa baris payroll di lokasi itu.")
        if ot_outside:
            bits = ", ".join(f"{lokasi} {rupiah(fee)}" for lokasi, fee in sorted(ot_outside.items()))
            notes.append("Lembur tercatat di " + bits + ", tanpa baris payroll di lokasi itu.")
        if not people:
            notes.append("Username ini tidak ada di tab Payroll.")
        elif not paid_people:
            notes.append("Ada baris payroll, tetapi gaji pokok dan gaji dibayarnya 0.")
        nama = people[0].nama if people else names.get(user, "")
        issues.append(
            {
                "username": people[0].username if people else user,
                "nama": nama,
                "n_payroll": len(paid_people),
                "paid": sum(officer.paid for officer in paid_people),
                "pay_locs": sorted(pay_locs),
                "att_locs": sorted(att_locs),
                "ot_outside": sum(ot_outside.values()),
                "notes": " ".join(notes),
            }
        )
    return issues


def rupiah(value: float) -> str:
    rounded = int(round(value))
    sign = "-" if rounded < 0 else ""
    return f"{sign}Rp {abs(rounded):,}".replace(",", ".")


def findings(report: Report) -> list[tuple[str, str]]:
    officers = report.officers
    paid = [o for o in officers if o.digaji]
    zero = [o for o in officers if not o.digaji]
    lembur_payroll = sum(o.amounts["lembur"] for o in officers)
    fee = sum(item["fee"] for item in report.overtime.values())
    mapped = officer_map(report)
    by_user = defaultdict(list)
    for officer in officers:
        if text(officer.username):
            by_user[text(officer.username).lower()].append(officer)
    tanpa_baris = []
    beda_lokasi = []
    for key, item in report.overtime.items():
        if key in mapped or item["fee"] <= 0:
            continue
        if by_user.get(text(item["username"]).lower()):
            beda_lokasi.append(item)
        else:
            tanpa_baris.append(item)
    issues = location_issues(report)
    ganda = [item for item in issues if item["n_payroll"] > 1]
    late_people = 0
    late_status = 0
    alpha_people = 0
    alpha_status = 0
    alpha_without_penalty = 0
    late_without_penalty = 0
    for officer in paid:
        att = report.attendance.get(key_of(officer.lokasi, officer.username, officer.nama), {})
        status = att.get("status", {})
        late = status.get("L", 0)
        alpha = status.get("A", 0)
        if late:
            late_people += 1
            late_status += late
            if abs(officer.amounts["denda_telat"]) < 0.5:
                late_without_penalty += 1
        if alpha:
            alpha_people += 1
            alpha_status += alpha
            if abs(officer.amounts["denda_alpa"]) < 0.5:
                alpha_without_penalty += 1
    denda_telat = sum(abs(o.amounts["denda_telat"]) for o in officers)
    denda_alpa = sum(abs(o.amounts["denda_alpa"]) for o in officers)
    testing = [o for o in paid if is_testing(o.lokasi, o.nama, o.username)]
    selisih = round(sum(o.selisih for o in officers), 2)
    items = []
    if abs(selisih) < 1:
        items.append(
            (
                "Hitungan gaji di file sumber sudah cocok",
                f"Gaji dibayar {rupiah(sum(o.paid for o in officers))} = gaji pokok + tunjangan - potongan, selisih {rupiah(selisih)}.",
            )
        )
    else:
        items.append(
            (
                "Ada selisih hitung gaji",
                f"Gaji dibayar tidak sama dengan pokok + tunjangan - potongan. Selisih {rupiah(selisih)}.",
            )
        )
    if ganda:
        names = ", ".join(item["username"] for item in ganda)
        extra = sum(item["paid"] for item in ganda)
        items.append(
            (
                "Username yang sama digaji di lebih dari satu lokasi",
                f"{len(ganda)} username muncul di dua baris payroll: {names}. Jika keduanya ditransfer, totalnya {rupiah(extra)}. Rincian ada di sheet Beda Lokasi.",
            )
        )
    items.append(
        (
            "Tunjangan lembur tidak sama dengan fee lembur",
            f"Payroll mencatat tunjangan lembur {rupiah(lembur_payroll)}. Tab Overtime mencatat fee {rupiah(fee)}. Selisih fee dikurangi tunjangan {rupiah(fee - lembur_payroll)}. {len(beda_lokasi)} username lemburnya tercatat di lokasi yang berbeda dari baris payroll. {len(tanpa_baris)} username lembur tidak ada di tab Payroll.",
        )
    )
    items.append(
        (
            "Denda keterlambatan tidak terisi",
            f"Ada {late_status} status terlambat pada {late_people} petugas yang digaji. Kolom Denda Keterlambatan berjumlah {rupiah(denda_telat)}. {late_without_penalty} petugas terlambat tanpa potongan.",
        )
    )
    items.append(
        (
            "Denda alpa hanya menutup sebagian",
            f"Ada {alpha_status} status alpa pada {alpha_people} petugas yang digaji. Potongan tercatat {rupiah(denda_alpa)}. {alpha_without_penalty} petugas alpa tanpa potongan. Pembanding default (gaji pokok / 30 per status A) ada di sheet Cek Denda dan belum tentu aturan spot tersebut.",
        )
    )
    if zero:
        items.append(
            (
                "Baris gaji nol masih ikut tarikan",
                f"{len(zero)} petugas gaji pokok dan gaji dibayarnya 0. Mereka dipisah ke sheet Tidak Digaji supaya tidak masuk daftar transfer.",
            )
        )
    if report.resign:
        items.append(
            (
                "Ada resign di tengah periode",
                f"{len(report.resign)} petugas resign di dalam cut-off. Cek prorate dan absensi sesudah tanggal resign sebelum ditransfer.",
            )
        )
    if testing:
        names = ", ".join(f"{o.nama} ({o.lokasi})" for o in testing)
        items.append(
            (
                "Akun testing ikut digaji",
                f"{len(testing)} baris testing punya gaji: {names}.",
            )
        )
    alpha_shift = report.anomaly_counts.get("shift_but_alpha", 0)
    if alpha_shift:
        items.append(
            (
                "Daftar anomali lama terlalu ramai untuk tutup gaji",
                f"Tarikan menandai {alpha_shift} 'shift aktif tetapi alpa' sebagai HIGH. Itu sudah kelihatan sebagai status A di absensi, bukan selisih rumus gaji. Sheet Perlu Dicek hanya memuat hal yang mengubah uang atau orang yang dibayar.",
            )
        )
    return items


def style_header(ws: Worksheet, ncol: int, fill=fill_navy) -> None:
    for col in range(1, ncol + 1):
        cell = ws.cell(1, col)
        cell.fill = fill
        cell.font = font_header
        cell.alignment = center
        cell.border = thin
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncol)}{max(ws.max_row, 1)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "Rekap Tutup Payroll"
    ws.oddFooter.right.text = "Halaman &P / &N"
    ws.print_title_rows = "1:1"
    ws.page_setup.horizontalDpi = 300
    ws.page_setup.verticalDpi = 300
    ws.sheet_view.showGridLines = False
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.5
    ws.sheet_view.zoomScale = 110


def paint_body(ws: Worksheet, money_cols: set[int], start_row: int = 2) -> None:
    red_font = Font(name="Calibri", size=11, color=RED)
    for row in range(start_row, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            cell.font = font_base
            cell.border = thin
            cell.alignment = right_align if col in money_cols else left_align
            if col in money_cols and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
    if money_cols and ws.max_row >= start_row:
        for col in money_cols:
            letter = get_column_letter(col)
            ws.conditional_formatting.add(
                f"{letter}{start_row}:{letter}{ws.max_row}",
                CellIsRule(operator="lessThan", formula=["0"], font=red_font),
            )


def set_widths(ws: Worksheet, widths: list[float]) -> None:
    for pos, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(pos)].width = width


def write_row(ws: Worksheet, row_no: int, values: list) -> None:
    for col, value in enumerate(values, start=1):
        ws.cell(row_no, col, value)


def write_ringkasan(wb: Workbook, report: Report) -> None:
    ws = wb.active
    ws.title = "Ringkasan"
    officers = report.officers
    paid = [o for o in officers if o.digaji]
    zero = [o for o in officers if not o.digaji]
    basic = sum(o.basic for o in officers)
    tunjangan = sum(sum(o.amounts[key] for key in PENAMBAH) for o in officers)
    potongan = sum(o.potongan for o in officers)
    dibayar = sum(o.paid for o in officers)

    ws.merge_cells("A1:C1")
    ws["A1"] = "Rekap Tutup Payroll"
    ws["A1"].font = font_title
    ws.merge_cells("A2:C2")
    period = report.period or "periode tidak tertulis di file sumber"
    month = report.month or "-"
    ws["A2"] = f"Bulan payroll {month}  ·  {period}  ·  sumber {report.source_name}"
    ws["A2"].font = font_muted
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 20

    ws["A4"] = "Posisi gaji"
    ws["A4"].font = font_section
    ws.merge_cells("A4:C4")
    headers = ("Komponen", "Jumlah", "Keterangan")
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(5, col, header)
        cell.fill = fill_navy
        cell.font = font_header
        cell.alignment = center
        cell.border = thin

    rows = [
        ("Petugas di tarikan", len(officers), "Semua baris di tab Payroll"),
        ("Petugas yang digaji", len(paid), "Masuk daftar transfer"),
        ("Petugas gaji nol", len(zero), "Dipisah, tidak ditransfer"),
        ("Gaji pokok", basic, "Basic salary"),
        ("Total tunjangan", tunjangan, "Jabatan, lembur, BPJS, prorate, lain-lain"),
        ("Total potongan", potongan, "Ditampilkan sebagai angka positif"),
        ("Gaji dibayar", dibayar, "Paid salary, angka yang ditransfer"),
        ("Selisih hitung", round(basic + tunjangan - potongan - dibayar, 2), "Harus 0. Pokok + tunjangan - potongan - dibayar"),
    ]
    for offset, (label, value, note) in enumerate(rows):
        row_no = 6 + offset
        write_row(ws, row_no, [label, value, note])
        for col in range(1, 4):
            cell = ws.cell(row_no, col)
            cell.font = font_bold if label in ("Gaji dibayar", "Selisih hitung") else font_base
            cell.border = thin
            cell.fill = fill_total if label in ("Gaji dibayar", "Selisih hitung") else fill_pale
            cell.alignment = left_align if col != 2 else right_align
        if isinstance(value, float):
            ws.cell(row_no, 2).number_format = "#,##0"
        if label == "Selisih hitung":
            ws.cell(row_no, 2).font = font_green if abs(value) < 1 else font_red

    start = 16
    ws.cell(start, 1, "Yang perlu dilihat sebelum transfer").font = font_section
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=3)
    for col, header in enumerate(("Temuan", "Artinya untuk payroll", ""), start=1):
        cell = ws.cell(start + 1, col, header if col < 3 else "")
        cell.fill = fill_teal
        cell.font = font_header
        cell.alignment = center
        cell.border = thin
    ws.merge_cells(start_row=start + 1, start_column=2, end_row=start + 1, end_column=3)
    for offset, (title, detail) in enumerate(findings(report)):
        row_no = start + 2 + offset
        ws.merge_cells(start_row=row_no, start_column=2, end_row=row_no, end_column=3)
        ws.cell(row_no, 1, title)
        ws.cell(row_no, 2, detail)
        for col in range(1, 4):
            cell = ws.cell(row_no, col)
            cell.font = font_bold if col == 1 else font_base
            cell.alignment = left_align
            cell.border = thin
            cell.fill = PatternFill("solid", fgColor=WHITE)
        ws.row_dimensions[row_no].height = 58

    note_row = start + 2 + len(findings(report)) + 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=3)
    ws.cell(
        note_row,
        1,
        "Angka di file ini dihitung ulang dari tarikan HRIS yang sama. Tidak ada komponen yang diubah, ditambah, atau dihapus dari sistem.",
    ).font = font_muted
    ws.cell(note_row, 1).alignment = left_align

    set_widths(ws, [42, 28, 88])
    ws.row_dimensions[5].height = 22
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.5
    ws.sheet_view.showGridLines = False
    ws.oddHeader.left.text = "Rekap Tutup Payroll"
    ws.oddFooter.right.text = "Halaman &P / &N"
    ws.sheet_view.zoomScale = 120
    ws.auto_filter.ref = None


def transfer_headers() -> list[str]:
    return [
        "Lokasi",
        "Nama",
        "Numeric ID",
        "Username",
        "Gaji Pokok",
        "Tunjangan Jabatan",
        "Tunjangan Lembur",
        "BPJS Kesehatan",
        "BPJS Ketenagakerjaan",
        "Prorate",
        "Tunjangan Lain",
        "Denda Keterlambatan",
        "Denda Ketidakhadiran",
        "Potongan BPJS Kesehatan",
        "Potongan BPJS Ketenagakerjaan",
        "Potongan Lain",
        "Gaji Dibayar",
        "Selisih Hitung",
    ]


def transfer_values(officer: Officer) -> list:
    return [
        officer.lokasi,
        officer.nama,
        officer.numeric_id,
        officer.username,
        officer.basic,
        officer.amounts["jabatan"],
        officer.amounts["lembur"],
        officer.amounts["bpjs_kes"],
        officer.amounts["bpjs_tk"],
        officer.amounts["prorate"],
        officer.amounts["lain_tunjangan"],
        abs(officer.amounts["denda_telat"]),
        abs(officer.amounts["denda_alpa"]),
        abs(officer.amounts["pot_bpjs_kes"]),
        abs(officer.amounts["pot_bpjs_tk"]),
        abs(officer.amounts["lain_potongan"]),
        officer.paid,
        officer.selisih,
    ]


def write_transfer(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Daftar Transfer")
    headers = transfer_headers()
    write_row(ws, 1, headers)
    paid = sorted((o for o in report.officers if o.digaji), key=lambda o: (o.lokasi.lower(), o.nama.lower()))
    for row_no, officer in enumerate(paid, start=2):
        write_row(ws, row_no, transfer_values(officer))
    total_row = len(paid) + 2
    totals = ["", "TOTAL", "", "", *[0] * (len(headers) - 4)]
    for officer in paid:
        values = transfer_values(officer)
        for pos in range(4, len(headers)):
            totals[pos] += values[pos]
    write_row(ws, total_row, totals)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(total_row, col)
        cell.font = Font(name="Calibri", bold=True, color=NAVY, size=11)
        cell.fill = fill_total
        cell.border = thin
        if col >= 5:
            cell.number_format = "#,##0"
    money_cols = set(range(5, len(headers) + 1))
    style_header(ws, len(headers))
    paint_body(ws, money_cols)
    for col in range(1, len(headers) + 1):
        ws.cell(total_row, col).fill = fill_total
        ws.cell(total_row, col).font = Font(name="Calibri", bold=True, color=NAVY, size=11)
    set_widths(ws, [32, 32, 14, 22] + [18] * (len(headers) - 4))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{total_row - 1}"
    ws.auto_filter.add_sort_condition(f"A2:A{total_row - 1}")


def write_lokasi(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Rekap Lokasi")
    headers = ["Lokasi", "Digaji", "Gaji Nol"] + transfer_headers()[4:]
    write_row(ws, 1, headers)
    grouped: dict[str, list[Officer]] = defaultdict(list)
    for officer in report.officers:
        grouped[officer.lokasi].append(officer)
    row_no = 2
    for lokasi in sorted(grouped, key=str.lower):
        people = grouped[lokasi]
        paid = [o for o in people if o.digaji]
        zero = len(people) - len(paid)
        sums = [0.0] * (len(transfer_headers()) - 4)
        for officer in paid:
            values = transfer_values(officer)
            for pos, value in enumerate(values[4:]):
                sums[pos] += value
        write_row(ws, row_no, [lokasi, len(paid), zero, *sums])
        row_no += 1
    total = [0.0] * (len(headers) - 1)
    for lokasi_row in range(2, row_no):
        for col in range(2, len(headers) + 1):
            total[col - 2] += money(ws.cell(lokasi_row, col).value)
    write_row(ws, row_no, ["TOTAL", *total])
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row_no, col)
        cell.fill = fill_total
        cell.font = Font(name="Calibri", bold=True, color=NAVY, size=11)
        cell.border = thin
        if col >= 2:
            cell.number_format = "#,##0"
    style_header(ws, len(headers))
    paint_body(ws, set(range(4, len(headers) + 1)))
    for col in range(1, len(headers) + 1):
        ws.cell(row_no, col).fill = fill_total
        ws.cell(row_no, col).font = Font(name="Calibri", bold=True, color=NAVY, size=11)
    set_widths(ws, [36, 12, 12] + [18] * (len(headers) - 3))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{row_no - 1}"


def write_komponen(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Rekap Komponen")
    write_row(ws, 1, ["Komponen", "Arah", "Jumlah", "Petugas terisi"])
    rows = [("Gaji Pokok", "penambah", sum(o.basic for o in report.officers), sum(1 for o in report.officers if abs(o.basic) > 0.5))]
    for key, label, side in GROUPS:
        amount = sum(o.amounts[key] for o in report.officers)
        shown = abs(amount) if side == "potongan" else amount
        people = sum(1 for o in report.officers if abs(o.amounts[key]) > 0.5)
        if shown == 0 and people == 0:
            continue
        rows.append((label, side, shown, people))
    penambah = sum(value for _label, side, value, _people in rows if side == "penambah")
    potongan = sum(value for _label, side, value, _people in rows if side == "potongan")
    dibayar = sum(o.paid for o in report.officers)
    for row_no, values in enumerate(rows, start=2):
        write_row(ws, row_no, list(values))
    foot = len(rows) + 2
    footer = [
        ("Total penambah", penambah),
        ("Total potongan", potongan),
        ("Gaji dibayar", dibayar),
        ("Selisih (penambah - potongan - dibayar)", round(penambah - potongan - dibayar, 2)),
    ]
    for offset, (label, value) in enumerate(footer):
        write_row(ws, foot + offset, [label, "", value, ""])
        for col in range(1, 5):
            cell = ws.cell(foot + offset, col)
            cell.fill = fill_total
            cell.font = font_bold
            cell.border = thin
        ws.cell(foot + offset, 3).number_format = "#,##0"
        ws.cell(foot + offset, 3).font = font_green if "Selisih" in label and abs(value) < 1 else font_bold
    style_header(ws, 4)
    paint_body(ws, {3})
    for offset in range(len(footer)):
        for col in range(1, 5):
            ws.cell(foot + offset, col).fill = fill_total
    set_widths(ws, [52, 16, 22, 20])
    ws.auto_filter.ref = f"A1:D{len(rows) + 1}"


LEMBUR_FILL = {
    "Cocok": PALE_GREEN,
    "Tidak ada di payroll": PALE_RED,
    "Tunjangan tanpa record lembur": PALE_AMBER,
    "Record lembur belum masuk tunjangan": PALE_RED,
    "Selisih": PALE_AMBER,
    "Lembur di lokasi lain": PALE_RED,
}


def write_lembur(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Rekonsiliasi Lembur")
    headers = [
        "Lokasi",
        "Nama",
        "Username",
        "Record Lembur",
        "Jam Lembur",
        "Fee di Tab Overtime",
        "Tunjangan di Payroll",
        "Selisih (Fee - Tunjangan)",
        "Status",
        "Keterangan",
    ]
    write_row(ws, 1, headers)
    mapped = officer_map(report)
    rows = []
    seen = set()
    for officer in report.officers:
        key = key_of(officer.lokasi, officer.username, officer.nama)
        ot = report.overtime.get(key, {})
        fee = ot.get("fee", 0.0)
        allow = officer.amounts["lembur"]
        if abs(fee) < 0.5 and abs(allow) < 0.5:
            continue
        gap = round(fee - allow, 2)
        if abs(gap) < 1:
            status = "Cocok"
        elif abs(fee) < 0.5:
            status = "Tunjangan tanpa record lembur"
        elif abs(allow) < 0.5:
            status = "Record lembur belum masuk tunjangan"
        else:
            status = "Selisih"
        rows.append(
            [
                officer.lokasi,
                officer.nama,
                officer.username,
                ot.get("record", 0),
                round(ot.get("jam", 0.0), 2),
                fee,
                allow,
                gap,
                status,
                "",
            ]
        )
        seen.add(key)
    payroll_users: dict[str, list[Officer]] = defaultdict(list)
    for officer in report.officers:
        if text(officer.username):
            payroll_users[text(officer.username).lower()].append(officer)
    for key, item in report.overtime.items():
        if key in seen or item["fee"] == 0:
            continue
        matches = payroll_users.get(text(item["username"]).lower(), [])
        if matches:
            lokasi_gaji = ", ".join(sorted({officer.lokasi for officer in matches}))
            status = "Lembur di lokasi lain"
            note = f"Username ini digaji di {lokasi_gaji}."
        else:
            status = "Tidak ada di payroll"
            note = "Username tidak ada di tab Payroll."
        rows.append(
            [
                item["lokasi"],
                item["nama"],
                item["username"],
                item["record"],
                round(item["jam"], 2),
                item["fee"],
                0,
                item["fee"],
                status,
                note,
            ]
        )
    rows.sort(key=lambda item: (-abs(item[7]), item[0].lower(), item[1].lower()))
    for row_no, values in enumerate(rows, start=2):
        write_row(ws, row_no, values)
        status = values[-2]
        for col in range(1, len(headers) + 1):
            ws.cell(row_no, col).fill = PatternFill("solid", fgColor=LEMBUR_FILL[status])
    style_header(ws, len(headers), fill_teal)
    paint_body(ws, {6, 7, 8})
    for row_no in range(2, len(rows) + 2):
        status = ws.cell(row_no, 9).value
        for col in range(1, len(headers) + 1):
            ws.cell(row_no, col).fill = PatternFill("solid", fgColor=LEMBUR_FILL[status])
    set_widths(ws, [32, 32, 22, 16, 14, 22, 24, 26, 32, 55])
    ws.auto_filter.ref = f"A1:J{max(len(rows) + 1, 1)}"


def write_absensi(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Rekap Absensi")
    headers = [
        "Lokasi",
        "Nama",
        "Username",
        "Hari tercatat",
        "Hari hadir",
        "Hari terlambat",
        "Hari alpa",
        "Hari izin",
        "Hari sakit",
        "Hari cuti",
        "Hari off",
        "Status hadir",
        "Status terlambat",
        "Status alpa",
        "Hari dengan lebih dari 1 status",
    ]
    write_row(ws, 1, headers)
    rows = []
    for officer in report.officers:
        att = report.attendance.get(key_of(officer.lokasi, officer.username, officer.nama))
        if not att:
            continue
        hari = att["hari_kode"]
        status = att["status"]
        rows.append(
            [
                officer.lokasi,
                officer.nama,
                officer.username,
                att["hari"],
                hari.get("OT", 0),
                hari.get("L", 0),
                hari.get("A", 0),
                hari.get("P", 0),
                hari.get("S", 0),
                hari.get("LEAVE", 0),
                hari.get("OFF", 0),
                status.get("OT", 0),
                status.get("L", 0),
                status.get("A", 0),
                att["ganda"],
            ]
        )
    rows.sort(key=lambda item: (item[0].lower(), item[1].lower()))
    for row_no, values in enumerate(rows, start=2):
        write_row(ws, row_no, values)
    style_header(ws, len(headers))
    paint_body(ws, set())
    for row_no in range(2, len(rows) + 2):
        for col in range(4, len(headers) + 1):
            ws.cell(row_no, col).alignment = center
            ws.cell(row_no, col).number_format = "#,##0"
    set_widths(ws, [32, 32, 22] + [14] * (len(headers) - 3))


def denda_status(late: int, alpha: int, late_fee: float, alpha_fee: float, late_ref: float, alpha_ref: float) -> str:
    notes = []
    if late and late_fee < 0.5:
        notes.append("Belum ada denda keterlambatan")
    elif late and abs(late_fee - late_ref) > 1:
        notes.append("Denda telat beda dari asumsi 50.000")
    if alpha and alpha_fee < 0.5:
        notes.append("Belum ada denda alpa")
    elif alpha and abs(alpha_fee - alpha_ref) > 1:
        notes.append("Denda alpa beda dari asumsi pokok/30")
    if not notes:
        return "Sesuai asumsi default"
    return "; ".join(notes)


def write_denda(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Cek Denda")
    headers = [
        "Lokasi",
        "Nama",
        "Username",
        "Gaji Pokok",
        "Status terlambat",
        "Hari terlambat",
        "Denda keterlambatan",
        "Pembanding 50.000 x status",
        "Status alpa",
        "Hari alpa",
        "Denda ketidakhadiran",
        "Pembanding pokok/30 x status",
        "Status",
    ]
    write_row(ws, 1, headers)
    rows = []
    for officer in report.officers:
        if not officer.digaji:
            continue
        att = report.attendance.get(key_of(officer.lokasi, officer.username, officer.nama), {})
        status = att.get("status", {})
        hari = att.get("hari_kode", {})
        late = status.get("L", 0)
        alpha = status.get("A", 0)
        late_fee = abs(officer.amounts["denda_telat"])
        alpha_fee = abs(officer.amounts["denda_alpa"])
        if late == 0 and alpha == 0 and late_fee < 0.5 and alpha_fee < 0.5:
            continue
        late_ref = late * DEFAULT_LATE_PENALTY
        alpha_ref = alpha * (officer.basic / 30.0) if officer.basic else 0
        rows.append(
            [
                officer.lokasi,
                officer.nama,
                officer.username,
                officer.basic,
                late,
                hari.get("L", 0),
                late_fee,
                late_ref,
                alpha,
                hari.get("A", 0),
                alpha_fee,
                round(alpha_ref, 2),
                denda_status(late, alpha, late_fee, alpha_fee, late_ref, alpha_ref),
            ]
        )
    rows.sort(key=lambda item: (item[12] == "Sesuai asumsi default", item[0].lower(), item[1].lower()))
    for row_no, values in enumerate(rows, start=2):
        write_row(ws, row_no, values)
    style_header(ws, len(headers), PatternFill("solid", fgColor=AMBER))
    paint_body(ws, {4, 7, 8, 11, 12})
    set_widths(ws, [32, 32, 22, 16, 16, 16, 22, 28, 14, 12, 22, 30, 42])
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(rows) + 1, 1)}"


def write_perlu_dicek(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Perlu Dicek")
    headers = ["Prioritas", "Lokasi", "Nama", "Username", "Jenis", "Nominal", "Keterangan"]
    write_row(ws, 1, headers)
    mapped = officer_map(report)
    rows = []
    for officer in report.officers:
        if not officer.digaji:
            rows.append(
                [
                    "Info",
                    officer.lokasi,
                    officer.nama,
                    officer.username,
                    "Gaji nol",
                    0,
                    "Tetap ada di tarikan payroll, tidak masuk daftar transfer.",
                ]
            )
        elif is_testing(officer.lokasi, officer.nama, officer.username):
            rows.append(
                [
                    "Tinggi",
                    officer.lokasi,
                    officer.nama,
                    officer.username,
                    "Akun testing digaji",
                    officer.paid,
                    "Nama atau lokasi mengandung testing, tetapi gaji dibayar terisi.",
                ]
            )
        key = key_of(officer.lokasi, officer.username, officer.nama)
        if key in report.resign:
            rows.append(
                [
                    "Sedang",
                    officer.lokasi,
                    officer.nama,
                    officer.username,
                    "Resign di tengah periode",
                    officer.paid,
                    f"Tanggal resign {report.resign[key]}. Cek prorate dan hari sesudah resign.",
                ]
            )
        ot = report.overtime.get(key, {})
        gap = round(ot.get("fee", 0) - officer.amounts["lembur"], 2)
        if abs(officer.amounts["lembur"]) > 0.5 or ot.get("fee", 0) > 0.5:
            if abs(gap) >= 1:
                rows.append(
                    [
                        "Tinggi",
                        officer.lokasi,
                        officer.nama,
                        officer.username,
                        "Selisih lembur",
                        gap,
                        f"Fee lembur {rupiah(ot.get('fee', 0))} dibanding tunjangan payroll {rupiah(officer.amounts['lembur'])}.",
                    ]
                )
        att = report.attendance.get(key, {})
        status = att.get("status", {})
        if status.get("L", 0) and abs(officer.amounts["denda_telat"]) < 0.5:
            rows.append(
                [
                    "Sedang",
                    officer.lokasi,
                    officer.nama,
                    officer.username,
                    "Telat tanpa denda",
                    status.get("L", 0) * DEFAULT_LATE_PENALTY,
                    f"{status.get('L', 0)} status terlambat, denda keterlambatan 0. Nominal adalah pembanding 50.000 per status, bukan angka yang sudah dipotong.",
                ]
            )
        if status.get("A", 0) and abs(officer.amounts["denda_alpa"]) < 0.5:
            rows.append(
                [
                    "Sedang",
                    officer.lokasi,
                    officer.nama,
                    officer.username,
                    "Alpa tanpa denda",
                    round(status.get("A", 0) * officer.basic / 30.0, 2) if officer.basic else 0,
                    f"{status.get('A', 0)} status alpa, denda ketidakhadiran 0. Nominal adalah pembanding gaji pokok/30 per status.",
                ]
            )
    for issue in location_issues(report):
        if issue["n_payroll"] > 1:
            rows.append(
                [
                    "Tinggi",
                    ", ".join(issue["pay_locs"]),
                    issue["nama"],
                    issue["username"],
                    "Digaji di lebih dari satu lokasi",
                    issue["paid"],
                    issue["notes"],
                ]
            )
    payroll_users = {text(officer.username).lower() for officer in report.officers if text(officer.username)}
    for key, item in report.overtime.items():
        if key in mapped or item["fee"] <= 0:
            continue
        if text(item["username"]).lower() in payroll_users:
            jenis = "Lembur di lokasi lain"
            note = "Fee ini tidak menempel ke baris payroll di lokasi yang sama. Lihat sheet Beda Lokasi."
        else:
            jenis = "Lembur tanpa payroll"
            note = f"{item['record']} record lembur, {item['jam']:.1f} jam. Username tidak ada di tab Payroll."
        rows.append(
            [
                "Tinggi",
                item["lokasi"],
                item["nama"],
                item["username"],
                jenis,
                item["fee"],
                note,
            ]
        )
    rank = {"Tinggi": 0, "Sedang": 1, "Info": 2}
    rows.sort(key=lambda item: (rank.get(item[0], 9), item[1].lower(), item[2].lower()))
    for row_no, values in enumerate(rows, start=2):
        write_row(ws, row_no, values)
        color = {"Tinggi": PALE_RED, "Sedang": PALE_AMBER, "Info": PALE}[values[0]]
        for col in range(1, len(headers) + 1):
            ws.cell(row_no, col).fill = PatternFill("solid", fgColor=color)
    style_header(ws, len(headers), PatternFill("solid", fgColor=RED))
    paint_body(ws, {6})
    for row_no in range(2, len(rows) + 2):
        color = {"Tinggi": PALE_RED, "Sedang": PALE_AMBER, "Info": PALE}[ws.cell(row_no, 1).value]
        for col in range(1, len(headers) + 1):
            ws.cell(row_no, col).fill = PatternFill("solid", fgColor=color)
        ws.cell(row_no, 7).alignment = left_align
    set_widths(ws, [14, 32, 32, 22, 28, 18, 78])
    ws.auto_filter.ref = f"A1:G{max(len(rows) + 1, 1)}"


def write_beda_lokasi(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Beda Lokasi")
    headers = [
        "Username",
        "Nama",
        "Jumlah lokasi digaji",
        "Total gaji dibayar",
        "Lokasi payroll",
        "Lokasi absensi",
        "Fee lembur di luar lokasi payroll",
        "Keterangan",
    ]
    write_row(ws, 1, headers)
    issues = location_issues(report)
    issues.sort(key=lambda item: (-item["n_payroll"], -item["paid"], item["username"].lower()))
    for row_no, issue in enumerate(issues, start=2):
        write_row(
            ws,
            row_no,
            [
                issue["username"],
                issue["nama"],
                issue["n_payroll"],
                issue["paid"],
                ", ".join(issue["pay_locs"]) or "-",
                ", ".join(issue["att_locs"]) or "-",
                issue["ot_outside"],
                issue["notes"],
            ],
        )
    style_header(ws, len(headers), PatternFill("solid", fgColor=RED))
    paint_body(ws, {4, 7})
    for row_no in range(2, len(issues) + 2):
        ws.cell(row_no, 8).alignment = left_align
        ws.row_dimensions[row_no].height = 48
    set_widths(ws, [22, 32, 18, 20, 42, 42, 28, 88])
    ws.auto_filter.ref = f"A1:H{max(len(issues) + 1, 1)}"


def write_tidak_digaji(wb: Workbook, report: Report) -> None:
    ws = wb.create_sheet("Tidak Digaji")
    headers = ["Lokasi", "Nama", "Numeric ID", "Username", "Keterangan"]
    write_row(ws, 1, headers)
    rows = [
        [o.lokasi, o.nama, o.numeric_id, o.username, "Gaji pokok dan gaji dibayar 0 di tarikan ini."]
        for o in report.officers
        if not o.digaji
    ]
    rows.sort(key=lambda item: (item[0].lower(), item[1].lower()))
    for row_no, values in enumerate(rows, start=2):
        write_row(ws, row_no, values)
    style_header(ws, len(headers))
    paint_body(ws, set())
    set_widths(ws, [36, 36, 14, 24, 62])


def write_workbook(report: Report, output: Path) -> None:
    wb = Workbook()
    write_ringkasan(wb, report)
    write_transfer(wb, report)
    write_lokasi(wb, report)
    write_komponen(wb, report)
    write_lembur(wb, report)
    write_beda_lokasi(wb, report)
    write_absensi(wb, report)
    write_denda(wb, report)
    write_perlu_dicek(wb, report)
    write_tidak_digaji(wb, report)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def print_summary(report: Report, output: Path) -> None:
    paid = sum(o.paid for o in report.officers if o.digaji)
    print(f"Sumber     : {report.source_name}")
    print(f"Bulan      : {report.month or '-'}")
    print(f"Periode    : {report.period or '-'}")
    print(f"Digaji     : {sum(1 for o in report.officers if o.digaji)} petugas")
    print(f"Gaji nol   : {sum(1 for o in report.officers if not o.digaji)} petugas")
    print(f"Dibayar    : {rupiah(paid)}")
    print(f"File       : {output}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Buat rekap tutup payroll dari file tarikan HRIS.")
    parser.add_argument("source", type=Path, help="File .xlsx hasil tarikan payroll HRIS")
    parser.add_argument("-o", "--output", type=Path, help="File rekap yang ditulis")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source = args.source.expanduser().resolve()
    if not source.exists():
        raise SystemExit(f"File tidak ditemukan: {source}")
    output = args.output or source.with_name(f"rekap_{source.stem}.xlsx")
    report = build_report(source)
    write_workbook(report, output)
    print_summary(report, output)


if __name__ == "__main__":
    main()
