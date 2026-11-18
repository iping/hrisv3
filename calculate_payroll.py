#!/usr/bin/env python3
"""Hitung payroll dari file tarikan HRIS, lalu bandingkan dengan gaji yang sudah tercatat.

Alur:
  1. Baca tab Payroll, Attendance, dan Overtime.
  2. Hitung fee lembur dari durasi (blok jam yang cocok dengan file ini).
  3. Hitung denda alpa = floor(gaji pokok / 25) x jumlah status A.
  4. Susun slip: komponen master tetap dari HRIS, lembur dan denda alpa hasil hitung.
  5. Tulis trial balance: gaji HRIS vs gaji dihitung.

Tidak mengubah HRIS dan tidak menarik data baru.

  python3 calculate_payroll.py "../payroll_report/payroll_2026-09-trialbalance 1 hour.xlsx"
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

DEDUCTION_HEADERS = {
    "Denda Keterlambatan",
    "Denda Ketidakhadiran",
    "Potongan Asuransi Kesehatan BPJS",
    "Potongan Asuransi Karyawan BPJS",
    "Fraud",
    "Deposit",
}

NAVY = "1F4E79"
TEAL = "0F6E6E"
RED = "9C2F2F"
LINE = "D9E2EF"
PALE = "F4F7FB"
PALE_RED = "FDECEC"
PALE_AMBER = "FFF6E5"
PALE_GREEN = "E7F6EC"
TOTAL = "D6E3F0"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_pale = PatternFill("solid", fgColor=PALE)
fill_total = PatternFill("solid", fgColor=TOTAL)
font_header = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
font_base = Font(name="Calibri", size=11, color="1A1A1A")
font_bold = Font(name="Calibri", size=11, color="1A1A1A", bold=True)
font_title = Font(name="Calibri", size=16, color=NAVY, bold=True)
font_section = Font(name="Calibri", size=12, color=NAVY, bold=True)
font_muted = Font(name="Calibri", size=11, color="4A5568")
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
right_align = Alignment(horizontal="right", vertical="center")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)


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


def person_key(lokasi, username, nama="") -> tuple[str, str]:
    user = text(username).lower() or text(nama).lower()
    return (text(lokasi).lower(), user)


def tokens(cell) -> list[str]:
    raw = text(cell)
    if not raw:
        return []
    return [part.strip() for part in raw.split("|") if part.strip()]


def rupiah(value: float) -> str:
    rounded = int(round(value))
    sign = "-" if rounded < 0 else ""
    return f"{sign}Rp {abs(rounded):,}".replace(",", ".")


def overtime_fee(hours: float) -> int:
    """Blok fee yang cocok dengan seluruh record Overtime di file sumber.

    Lebih dari 0 jam sampai 4 jam = 50.000.
    Lebih dari 4 jam sampai 8 jam = 75.000.
    Lebih dari 8 jam = 100.000.
    Durasi di bawah 1 jam tetap masuk blok pertama.
    """
    if hours <= 0:
        return 0
    if hours <= 4:
        return 50_000
    if hours <= 8:
        return 75_000
    return 100_000


ATTEND_CODES = (
    ("OT", "Hadir"),
    ("L", "Terlambat"),
    ("A", "Alpa"),
    ("P", "Izin"),
    ("S", "Sakit"),
    ("LEAVE", "Cuti"),
    ("OFF", "Off"),
    ("O", "Tanpa jadwal"),
)
ATTEND_KEYS = {code for code, _label in ATTEND_CODES}


def day_iso(value) -> str:
    raw = text(value)
    if len(raw) >= 10 and raw[4] == "-" and raw[7] == "-":
        return raw[:10]
    return ""


def day_label(iso: str) -> str:
    if len(iso) == 10:
        return f"{iso[8:10]}/{iso[5:7]}"
    return iso


def overtime_block(hours: float) -> str:
    if hours <= 0:
        return "0 jam"
    if hours <= 4:
        return "0–4 jam"
    if hours <= 8:
        return "4–8 jam"
    return ">8 jam"


def absence_penalty(basic: float, alpha_status: int) -> int:
    """Denda satu status alpa = gaji pokok dibagi 25, dibulatkan ke bawah.

    Pembagi 25 cocok dengan mayoritas baris di file ini. Pembagi 30
    (asumsi tool tarikan lama) tidak cocok.
    """
    if basic <= 0 or alpha_status <= 0:
        return 0
    return int(basic // 25) * alpha_status


def alpha_token_count(raw: str) -> int:
    """How many Alpa marks are in one attendance cell (A | A counts as 2)."""
    return sum(1 for part in tokens(raw) if part == "A")


def alpha_day_labels(line: "AttendanceLine", date_labels: list[str]) -> str:
    """Human list of days that still carry Alpa after approved banding."""
    bits = []
    for label, raw in zip(date_labels, line.days):
        count = alpha_token_count(raw)
        if count <= 0:
            continue
        bits.append(f"{label} (×{count})" if count > 1 else label)
    return ", ".join(bits)


def banding_alpa_refund(before: str, after: str, basic: float) -> int:
    """Amount returned to payroll when approved banding removes Alpa marks."""
    removed = max(0, alpha_token_count(before) - alpha_token_count(after))
    if removed <= 0 or basic <= 0:
        return 0
    return int(basic // 25) * removed


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
    if name == "Deposit":
        return "deposit"
    if name == "Potongan Asuransi Kesehatan BPJS":
        return "pot_bpjs_kes"
    if name == "Potongan Asuransi Karyawan BPJS":
        return "pot_bpjs_tk"
    return "lain_tunjangan" if side == "allow" else "lain_potongan"


@dataclass
class OvertimeRecord:
    lokasi: str
    nama: str
    username: str
    mulai: str
    akhir: str
    menit: float
    jam: float
    fee_hris: float
    fee_hitung: int
    ada_slip: bool = False

    @property
    def blok(self) -> str:
        return overtime_block(self.jam)


WORK_DAY_DIVISOR = 25


def roster_day(raw: str) -> bool:
    """Day counts as on-roster once the officer has started (not blank / O = no schedule)."""
    return bool((raw or "").strip()) and (raw or "").strip() != "O"


def first_roster_iso(days: list[str], dates: list[str]) -> str:
    for iso, raw in zip(dates, days):
        if roster_day(raw):
            return iso
    return ""


def count_roster_days(days: list[str], dates: list[str], start_iso: str = "") -> int:
    total = 0
    for iso, raw in zip(dates, days):
        if start_iso and iso < start_iso:
            continue
        if roster_day(raw):
            total += 1
    return total


def new_joiner_prorate(basic: float, worked: int) -> float:
    """Express pay for worked roster days as a Prorate delta on full basic (divisor 25, same as denda alpa)."""
    if basic <= 0:
        return 0.0
    days = max(0, min(int(worked or 0), WORK_DAY_DIVISOR))
    earned = round(basic * days / WORK_DAY_DIVISOR)
    return round(earned - basic, 2)


@dataclass
class Officer:
    lokasi: str
    nama: str
    username: str
    numeric_id: str
    basic: float
    paid: float
    amounts: dict[str, float] = field(default_factory=lambda: defaultdict(float))
    alpha: int = 0
    late: int = 0
    hadir: int = 0
    ot_records: list[OvertimeRecord] = field(default_factory=list)
    new_joiner: bool = False
    first_attendance: str = ""
    worked_days: int = 0
    prorate_hris: float = 0.0
    period_start: str = ""
    period_end: str = ""

    @property
    def lembur_hris(self) -> float:
        return self.amounts["lembur"]

    @property
    def lembur_hitung(self) -> int:
        return sum(record.fee_hitung for record in self.ot_records)

    @property
    def jam_lembur(self) -> float:
        return sum(record.jam for record in self.ot_records)

    @property
    def denda_alpa_hris(self) -> float:
        return abs(self.amounts["denda_alpa"])

    @property
    def denda_alpa_hitung(self) -> int:
        return absence_penalty(self.basic, self.alpha)

    @property
    def denda_telat(self) -> float:
        return abs(self.amounts["denda_telat"])

    @property
    def master_penambah(self) -> float:
        return self.basic + sum(
            self.amounts[key]
            for key in ("jabatan", "bpjs_kes", "bpjs_tk", "prorate", "lain_tunjangan", "manual_tambahan")
        )

    @property
    def potongan_tetap(self) -> float:
        return sum(
            abs(self.amounts[key])
            for key in ("denda_telat", "pot_bpjs_kes", "pot_bpjs_tk", "deposit", "lain_potongan", "manual_potongan")
        )

    @property
    def gaji_hris_susun(self) -> float:
        return self.master_penambah + self.lembur_hris - self.denda_alpa_hris - self.potongan_tetap

    @property
    def gaji_hitung(self) -> float:
        return self.master_penambah + self.lembur_hitung - self.denda_alpa_hitung - self.potongan_tetap

    @property
    def selisih_gaji(self) -> float:
        return round(self.gaji_hitung - self.paid, 2)

    def _money_label(self, amount: float, signed: bool = True) -> str:
        value = round(float(amount or 0), 2)
        prefix = ""
        if signed:
            if value > 0.009:
                prefix = "+ "
            elif value < -0.009:
                prefix = "− "
                value = abs(value)
            else:
                return "Rp 0"
        else:
            value = abs(value)
        return prefix + "Rp " + f"{int(round(value)):,}".replace(",", ".")

    def new_joiner_prorate_detail(self) -> dict | None:
        """Full New Joiner prorate explanation for the payroll drawer."""
        if not self.new_joiner:
            return None
        money = self._money_label
        days_raw = int(self.worked_days or 0)
        days_used = max(0, min(days_raw, WORK_DAY_DIVISOR))
        basic = float(self.basic or 0)
        earned = round(basic * days_used / WORK_DAY_DIVISOR) if basic else 0
        prorate = round(float(self.amounts["prorate"] or 0), 2)
        prorate_hris = round(float(self.prorate_hris or 0), 2)
        gap = round(prorate - prorate_hris, 2)
        period_start = self.period_start or ""
        period_end = self.period_end or ""
        start = self.first_attendance or ""
        steps = [
            {
                "label": "Periode payroll",
                "value": (
                    f"{day_label(period_start)} s/d {day_label(period_end)}"
                    if period_start and period_end
                    else "—"
                ),
                "note": "Periode absensi file (biasanya 26 bulan lalu s/d 25 bulan ini)",
            },
            {
                "label": "Mulai absensi",
                "value": day_label(start) if start else "—",
                "note": "Hari pertama kode absensi bukan kosong/O setelah awal periode → New Joiner",
            },
            {
                "label": "Hari kerja dihitung",
                "value": f"{days_raw} hari" + (f" (dipakai {days_used})" if days_raw != days_used else ""),
                "note": (
                    f"Hari dengan jadwal/kode absensi sejak mulai sampai akhir periode. "
                    f"Maksimal {WORK_DAY_DIVISOR} hari = gaji penuh (sama pembagi denda alpa)"
                ),
            },
            {
                "label": "Gaji pokok penuh",
                "value": money(basic, False),
                "note": f"Asumsi kerja penuh {WORK_DAY_DIVISOR} hari",
            },
            {
                "label": "Gaji pokok prorate",
                "value": money(earned, False),
                "note": (
                    f"{money(basic, False)} × {days_used} ÷ {WORK_DAY_DIVISOR}"
                    if basic
                    else "Gaji pokok 0, tidak ada prorate dari pokok"
                ),
            },
            {
                "label": "Prorate dihitung",
                "value": money(prorate),
                "note": "Gaji pokok prorate − gaji pokok penuh (masuk kolom Prorate slip)",
            },
            {
                "label": "Prorate HRIS",
                "value": money(prorate_hris),
                "note": "Nilai Prorate dari file sumber sebelum diganti untuk New Joiner",
            },
            {
                "label": "Dampak ke selisih",
                "value": money(gap),
                "note": "Prorate dihitung − Prorate HRIS (bagian selisih gaji dari New Joiner)",
            },
        ]
        return {
            "active": True,
            "period_label": (
                f"{day_label(period_start)} s/d {day_label(period_end)}"
                if period_start and period_end
                else "—"
            ),
            "start_label": day_label(start) if start else "—",
            "worked_days": days_raw,
            "days_used": days_used,
            "divisor": WORK_DAY_DIVISOR,
            "basic": basic,
            "basic_label": money(basic, False),
            "earned": earned,
            "earned_label": money(earned, False),
            "prorate": prorate,
            "prorate_label": money(prorate),
            "prorate_hris": prorate_hris,
            "prorate_hris_label": money(prorate_hris),
            "gap": gap,
            "gap_label": money(gap),
            "formula": (
                f"Prorate = ({money(basic, False)} × {days_used} ÷ {WORK_DAY_DIVISOR}) − {money(basic, False)} "
                f"= {money(prorate)}"
            ),
            "steps": steps,
            "cls": "is-add" if gap > 1 else ("is-cut" if gap < -1 else "is-empty"),
        }

    def selisih_breakdown(self, banding: list | None = None) -> dict:
        """Explain Selisih = Gaji Dihitung − Gaji HRIS (Paid) for the drawer."""
        money = self._money_label

        lembur_gap = round(self.lembur_hitung - self.lembur_hris, 2)
        # Higher denda hitung lowers gaji dihitung; express as effect on selisih.
        denda_gap = round(self.denda_alpa_hris - self.denda_alpa_hitung, 2)
        prorate_gap = round(float(self.amounts["prorate"] or 0) - float(self.prorate_hris or 0), 2)
        manual_add = round(float(self.amounts["manual_tambahan"] or 0), 2)
        manual_cut = round(-abs(float(self.amounts["manual_potongan"] or 0)), 2)
        deposit = round(abs(float(self.amounts["deposit"] or 0)), 2)
        explained = lembur_gap + denda_gap + prorate_gap + manual_add + manual_cut
        residual = round(self.selisih_gaji - explained, 2)
        lines = []
        if abs(lembur_gap) >= 1:
            lines.append(
                {
                    "kind_label": "Lembur",
                    "title": "Selisih lembur",
                    "amount": lembur_gap,
                    "amount_label": money(lembur_gap),
                    "source": (
                        f"Lembur HRIS {money(self.lembur_hris, False)} → "
                        f"Lembur dihitung {money(self.lembur_hitung, False)}"
                    ),
                    "cls": "is-add" if lembur_gap > 0 else "is-cut",
                }
            )
        if abs(denda_gap) >= 1:
            lines.append(
                {
                    "kind_label": "Denda alpa",
                    "title": "Selisih denda alpa",
                    "amount": denda_gap,
                    "amount_label": money(denda_gap),
                    "source": (
                        f"Denda HRIS {money(self.denda_alpa_hris, False)} → "
                        f"Denda dihitung {money(self.denda_alpa_hitung, False)} "
                        f"({self.alpha} status A × tarif {money(int(self.basic // 25) if self.basic else 0, False)})"
                    ),
                    "cls": "is-add" if denda_gap > 0 else "is-cut",
                }
            )
        if self.new_joiner and abs(prorate_gap) >= 1:
            start_label = day_label(self.first_attendance) if self.first_attendance else "—"
            lines.append(
                {
                    "kind_label": "New Joiner",
                    "title": "Prorate hari kerja",
                    "amount": prorate_gap,
                    "amount_label": money(prorate_gap),
                    "source": (
                        f"Mulai absensi {start_label} · {self.worked_days} hari kerja "
                        f"(gaji pokok × hari ÷ {WORK_DAY_DIVISOR} − gaji pokok)"
                    ),
                    "cls": "is-add" if prorate_gap > 0 else "is-cut",
                }
            )
        if abs(manual_add) >= 1:
            lines.append(
                {
                    "kind_label": "Penambahan",
                    "title": "Tambahan manual",
                    "amount": manual_add,
                    "amount_label": money(manual_add),
                    "source": "Dari tombol Penambahan di payroll (Banding / Refund Deposit / Lainnya)",
                    "cls": "is-add",
                }
            )
        if abs(manual_cut) >= 1:
            lines.append(
                {
                    "kind_label": "Pengurangan",
                    "title": "Potongan manual",
                    "amount": manual_cut,
                    "amount_label": money(manual_cut),
                    "source": "Dari tombol Pengurangan di payroll (Deposit override / Fraud / Lainnya)",
                    "cls": "is-cut",
                }
            )
        if deposit >= 1:
            lines.append(
                {
                    "kind_label": "Deposit",
                    "title": "Deposit dipotong",
                    "amount": -deposit,
                    "amount_label": money(-deposit),
                    "source": "Deposit mengurangi gaji dihitung; biasanya sudah tercermin juga di Gaji HRIS",
                    "cls": "is-cut",
                }
            )
        if abs(residual) >= 1:
            lines.append(
                {
                    "kind_label": "Lainnya",
                    "title": "Sisa selisih",
                    "amount": residual,
                    "amount_label": money(residual),
                    "source": "Sisa setelah lembur, denda alpa, dan penyesuaian manual (beda susunan komponen HRIS vs Paid)",
                    "cls": "is-add" if residual > 0 else "is-cut",
                }
            )
        key = (self.username or "").strip().lower()
        banding_rows = []
        for row in banding or []:
            if (row.username or "").strip().lower() != key:
                continue
            if "APPROVED" not in (row.status or "").upper():
                continue
            if not (row.refund or alpha_token_count(row.before)):
                continue
            banding_rows.append(
                {
                    "day": row.day,
                    "before": row.before or "—",
                    "after": row.after or "—",
                    "effect": row.effect or "",
                    "refund": int(row.refund or 0),
                    "refund_label": money(row.refund, False) if row.refund else "",
                }
            )
        return {
            "paid": self.paid,
            "hitung": self.gaji_hitung,
            "selisih": self.selisih_gaji,
            "paid_label": money(self.paid, False),
            "hitung_label": money(self.gaji_hitung, False),
            "selisih_label": money(self.selisih_gaji),
            "alpha": self.alpha,
            "lines": lines,
            "banding": banding_rows,
            "new_joiner": self.new_joiner_prorate_detail(),
        }


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


def load_officers(rows: list[tuple]) -> list[Officer]:
    if not rows:
        raise SystemExit("Tab Payroll kosong atau tidak ada.")
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    for required in ("Lokasi", "Nama", "Username", "Basic Salary", "Paid Salary"):
        if required not in index:
            raise SystemExit(f"Kolom '{required}' tidak ada di tab Payroll.")
    deduction_at = len(headers)
    for pos, name in enumerate(headers):
        if pos >= 8 and name in DEDUCTION_HEADERS:
            deduction_at = pos
            break
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
            if pos < 8 or not name:
                continue
            amount = money(row[pos]) if pos < len(row) else 0.0
            if amount == 0:
                continue
            side = "allow" if pos < deduction_at else "deduct"
            officer.amounts[group_for(name, side)] += amount
        officers.append(officer)
    return officers


def _cells_by_day(row: tuple, header: tuple) -> dict[str, str]:
    found = {}
    if not row or not header:
        return found
    for pos, cell in enumerate(header):
        if pos < 4:
            continue
        iso = day_iso(cell)
        if not iso:
            continue
        found[iso] = text(row[pos]) if pos < len(row) else ""
    return found


def _recount_days(days: list[str]) -> dict[str, int]:
    counts = {code: 0 for code, _label in ATTEND_CODES}
    for raw in days:
        parts = tokens(raw)
        if parts:
            for token in parts:
                if token in ATTEND_KEYS:
                    counts[token] += 1
        elif raw == "O":
            counts["O"] += 1
    return counts


def load_attendance_grid(attendance_rows: list[tuple], shift_rows: list[tuple]) -> tuple[list[str], list[AttendanceLine], dict]:
    """Daily codes from the upload. An empty day with no shift is O, tidak ada jadwal."""
    header = attendance_rows[0] if attendance_rows else ()
    dates = [day_iso(cell) for cell in header[4:]]
    dates = [iso for iso in dates if iso]
    shifts: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
    if shift_rows:
        shift_header = shift_rows[0]
        for row in shift_rows[1:]:
            if not row or not text(row[0]):
                continue
            shifts[person_key(row[0], row[3] if len(row) > 3 else "", row[1])].append(_cells_by_day(row, shift_header))
    lines = []
    summed: dict[tuple[str, str], dict[str, int]] = defaultdict(lambda: defaultdict(int))
    seen: dict[tuple[str, str], int] = defaultdict(int)
    for row in attendance_rows[1:]:
        if not row or not text(row[0]):
            continue
        key = person_key(row[0], row[3] if len(row) > 3 else "", row[1])
        bucket = shifts.get(key, [])
        slot = seen[key]
        seen[key] += 1
        shift_days = bucket[slot] if slot < len(bucket) else {}
        own = _cells_by_day(row, header)
        shown = []
        for iso in dates:
            raw = own.get(iso, "")
            parts = tokens(raw)
            if parts:
                shown.append(raw)
            elif not shift_days.get(iso, ""):
                shown.append("O")
            else:
                shown.append("")
        counts = _recount_days(shown)
        for code, amount in counts.items():
            if amount:
                summed[key][code] += amount
        lines.append(
            AttendanceLine(
                lokasi=text(row[0]),
                nama=text(row[1]) if len(row) > 1 else "",
                username=text(row[3]) if len(row) > 3 else "",
                days=shown,
                counts=counts,
            )
        )
    lines.sort(key=lambda line: (line.lokasi.lower(), line.nama.lower(), line.username.lower()))
    return dates, lines, summed


def load_banding(rows: list[tuple]) -> list["BandingRecord"]:
    if not rows:
        return []
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    records = []
    for row in rows[1:]:
        if not row or not text(row[0] if len(row) > 0 else ""):
            continue
        status = text(row[index["Status"]]) if "Status" in index and index["Status"] < len(row) else ""
        day_raw = row[index["Tanggal Attendance"]] if "Tanggal Attendance" in index and index["Tanggal Attendance"] < len(row) else ""
        if hasattr(day_raw, "strftime"):
            day = day_raw.strftime("%Y-%m-%d")
        else:
            day = day_iso(day_raw) or text(day_raw)[:10]
        records.append(
            BandingRecord(
                lokasi=text(row[index["Lokasi"]]) if "Lokasi" in index else "",
                nama=text(row[index["Nama"]]) if "Nama" in index else "",
                username=text(row[index["Username"]]) if "Username" in index else "",
                submitted=text(row[index["Waktu Pengajuan"]]) if "Waktu Pengajuan" in index and index["Waktu Pengajuan"] < len(row) else "",
                day=day,
                label=text(row[index["Label"]]) if "Label" in index and index["Label"] < len(row) else "",
                reason=text(row[index["Reason"]]) if "Reason" in index and index["Reason"] < len(row) else "",
                status=status,
                before=text(row[index["Attendance Sebelum"]]) if "Attendance Sebelum" in index and index["Attendance Sebelum"] < len(row) else "",
                after=text(row[index["Attendance Sesudah"]]) if "Attendance Sesudah" in index and index["Attendance Sesudah"] < len(row) else "",
            )
        )
    return records


def apply_approved_banding(dates: list[str], lines: list[AttendanceLine], records: list["BandingRecord"]) -> dict:
    """Use approved banding as the attendance code for that day, then rebuild monthly counts."""
    date_at = {iso: index for index, iso in enumerate(dates)}
    by_key: dict[tuple[str, str], list[AttendanceLine]] = defaultdict(list)
    for line in lines:
        by_key[person_key(line.lokasi, line.username, line.nama)].append(line)
    approved = [row for row in records if "APPROVED" in row.status.upper()]
    approved.sort(key=lambda row: (row.submitted, row.day, row.username))
    for row in approved:
        if row.day not in date_at or not row.after:
            row.effect = "Di luar periode absensi"
            continue
        bucket = by_key.get(person_key(row.lokasi, row.username, row.nama))
        if not bucket:
            row.effect = "Petugas tidak ada di absensi"
            continue
        line = bucket[0]
        at = date_at[row.day]
        current = line.days[at]
        parts = tokens(current)
        if row.after in parts or current == row.after:
            row.effect = "Sudah di absensi"
            continue
        before = current or row.before or "—"
        line.days[at] = row.after
        row.applied = True
        row.effect = f"{before} → {row.after}"
    summed: dict[tuple[str, str], dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for line in lines:
        line.counts = _recount_days(line.days)
        key = person_key(line.lokasi, line.username, line.nama)
        for code, amount in line.counts.items():
            if amount:
                summed[key][code] += amount
    return summed


def load_overtime(rows: list[tuple]) -> list[OvertimeRecord]:
    if not rows:
        return []
    headers = [text(cell) for cell in rows[0]]
    index = {name: pos for pos, name in enumerate(headers)}
    records = []
    for row in rows[1:]:
        if not row or not any(row):
            continue
        menit = money(row[index["Durasi Lembur (Menit)"]]) if "Durasi Lembur (Menit)" in index else 0
        jam = money(row[index["Durasi Lembur (Jam)"]]) if "Durasi Lembur (Jam)" in index else 0
        if jam == 0 and menit:
            jam = menit / 60.0
        records.append(
            OvertimeRecord(
                lokasi=text(row[index.get("Lokasi", 2)]),
                nama=text(row[index.get("Petugas", 3)]),
                username=text(row[index.get("Username", 4)]),
                mulai=text(row[index.get("Mulai Lembur", 5)]),
                akhir=text(row[index.get("Akhir Lembur", 6)]),
                menit=menit,
                jam=jam,
                fee_hris=money(row[index.get("Overtime Fee", 9)]) if index.get("Overtime Fee", 9) < len(row) else 0,
                fee_hitung=overtime_fee(jam),
            )
        )
    return records


@dataclass
class AttendanceLine:
    lokasi: str
    nama: str
    username: str
    days: list[str]
    counts: dict[str, int]


@dataclass
class BandingRecord:
    lokasi: str
    nama: str
    username: str
    submitted: str
    day: str
    label: str
    reason: str
    status: str
    before: str
    after: str
    applied: bool = False
    effect: str = ""
    refund: int = 0


@dataclass
class Payroll:
    source_name: str
    month: str
    period: str
    officers: list[Officer]
    overtime: list[OvertimeRecord]
    unmatched_overtime: list[OvertimeRecord]
    attendance: list[AttendanceLine] = field(default_factory=list)
    attendance_dates: list[str] = field(default_factory=list)
    banding: list[BandingRecord] = field(default_factory=list)


def apply_new_joiners(dates: list[str], lines: list[AttendanceLine], officers: list[Officer]) -> None:
    """Flag officers whose first roster day is after period start, and set Prorate from hari kerja ÷ 25."""
    if not dates:
        return
    period_start = dates[0]
    by_key: dict[tuple[str, str], list[AttendanceLine]] = defaultdict(list)
    for line in lines:
        by_key[person_key(line.lokasi, line.username, line.nama)].append(line)
    period_end = dates[-1]
    for officer in officers:
        officer.period_start = period_start
        officer.period_end = period_end
        officer.prorate_hris = float(officer.amounts["prorate"] or 0)
        bucket = by_key.get(person_key(officer.lokasi, officer.username, officer.nama)) or []
        if not bucket:
            continue
        line = bucket[0]
        first = first_roster_iso(line.days, dates)
        officer.worked_days = count_roster_days(line.days, dates, first) if first else 0
        if not first or first <= period_start:
            continue
        officer.new_joiner = True
        officer.first_attendance = first
        officer.amounts["prorate"] = new_joiner_prorate(officer.basic, officer.worked_days)


def build(path: Path) -> Payroll:
    wb = load_workbook(path, read_only=True, data_only=True)
    meta = read_meta(sheet_rows(wb, "Anomaly Summary"))
    officers = load_officers(sheet_rows(wb, "Payroll"))
    attendance_rows = sheet_rows(wb, "Attendance")
    dates, attendance_lines, _raw_attendance = load_attendance_grid(attendance_rows, sheet_rows(wb, "Shift"))
    banding = load_banding(sheet_rows(wb, "Pengajuan Banding"))
    attendance = apply_approved_banding(dates, attendance_lines, banding)
    overtime = load_overtime(sheet_rows(wb, "Overtime"))
    wb.close()
    period_start, period_end = attendance_bounds(attendance_rows)

    by_key = {person_key(officer.lokasi, officer.username, officer.nama): officer for officer in officers}
    for officer in officers:
        counts = attendance.get(person_key(officer.lokasi, officer.username, officer.nama), {})
        officer.alpha = counts.get("A", 0)
        officer.late = counts.get("L", 0)
        officer.hadir = counts.get("OT", 0)
    apply_new_joiners(dates, attendance_lines, officers)
    unmatched = []
    for record in overtime:
        officer = by_key.get(person_key(record.lokasi, record.username, record.nama))
        if officer is None:
            unmatched.append(record)
            continue
        record.ada_slip = True
        officer.ot_records.append(record)
    # Approved banding that removes Alpa adds the avoided penalty back into the banding detail.
    for row in banding:
        if "APPROVED" not in row.status.upper():
            continue
        officer = by_key.get(person_key(row.lokasi, row.username, row.nama))
        basic = officer.basic if officer else 0
        row.refund = banding_alpa_refund(row.before, row.after, basic)
        if row.applied and row.refund > 0 and not row.effect:
            row.effect = f"Alpa dikembalikan Rp {row.refund:,}".replace(",", ".")
        elif row.applied and row.refund > 0:
            row.effect = f"{row.effect} · dikembalikan Rp {row.refund:,}".replace(",", ".")
    month = meta.get("Bulan payroll", "") or (period_end[:7] if period_end else "")
    period = meta.get("Periode pemeriksaan", "")
    if not period and period_start and period_end:
        period = f"{period_start} s/d {period_end}"
    return Payroll(
        source_name=path.name,
        month=month,
        period=period,
        officers=officers,
        overtime=overtime,
        unmatched_overtime=unmatched,
        attendance=attendance_lines,
        attendance_dates=[day_label(iso) for iso in dates],
        banding=banding,
    )


def attendance_bounds(rows: list[tuple]) -> tuple[str, str]:
    """Tanggal pertama dan terakhir di header tab Attendance."""
    if not rows:
        return "", ""
    dates = []
    for cell in rows[0][4:]:
        raw = text(cell)[:10]
        if len(raw) == 10 and raw[4] == "-" and raw[7] == "-":
            dates.append(raw)
    if not dates:
        return "", ""
    return dates[0], dates[-1]


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
    ws.print_title_rows = "1:1"
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.5
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.oddHeader.left.text = "Hitung Payroll"
    ws.oddFooter.right.text = "Halaman &P / &N"


def paint(ws: Worksheet, money_cols: set[int]) -> None:
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            cell.font = font_base
            cell.border = thin
            cell.alignment = right_align if col in money_cols else left_align
            if col in money_cols and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"


def mark_reducing_bpjs(ws: Worksheet, cols: set[int], only_bpjs_rows: bool = False, match: str = "BPJS") -> None:
    """A negative amount reduces gaji pokok: parentheses and red, including on the total row."""
    plain = Font(name="Calibri", size=11, color=RED)
    bold = Font(name="Calibri", size=11, color=RED, bold=True)
    needle = match.upper()
    for row in range(2, ws.max_row + 1):
        if only_bpjs_rows and needle not in text(ws.cell(row, 1).value).upper():
            continue
        for col in cols:
            cell = ws.cell(row, col)
            if isinstance(cell.value, (int, float)) and cell.value < 0:
                cell.font = bold if cell.font and cell.font.bold else plain
                cell.number_format = '#,##0;(#,##0);"-"'


def widths(ws: Worksheet, sizes: list[float]) -> None:
    for pos, size in enumerate(sizes, start=1):
        ws.column_dimensions[get_column_letter(pos)].width = size


def write_alur(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.active
    ws.title = "Alur"
    officers = payroll.officers
    paid = sum(officer.paid for officer in officers)
    calculated = sum(officer.gaji_hitung for officer in officers)
    fee_match = sum(1 for record in payroll.overtime if abs(record.fee_hitung - record.fee_hris) < 1)
    alpha_rows = [officer for officer in officers if officer.alpha or officer.denda_alpa_hris]
    alpha_match = sum(
        1 for officer in alpha_rows if abs(officer.denda_alpa_hitung - officer.denda_alpa_hris) < 1
    )
    rebuild_gap = max((abs(officer.gaji_hris_susun - officer.paid) for officer in officers), default=0)

    ws.merge_cells("A1:B1")
    ws["A1"] = "Alur hitung payroll"
    ws["A1"].font = font_title
    ws.merge_cells("A2:B2")
    ws["A2"] = (
        f"Bulan {payroll.month or '-'}  ·  {payroll.period or 'periode tidak tertulis'}  ·  {payroll.source_name}"
    )
    ws["A2"].font = font_muted
    ws.row_dimensions[1].height = 26

    banding_applied = sum(1 for row in payroll.banding if row.applied)
    banding_approved = sum(1 for row in payroll.banding if "APPROVED" in row.status.upper())
    deposit_total = sum(abs(officer.amounts["deposit"]) for officer in officers)
    steps = [
        (
            "1. Baca tarikan",
            "Payroll, Attendance, Overtime, Deposit, dan Pengajuan Banding dari file HRIS. "
            "Komponen master selain lembur, denda alpa, dan deposit dipakai apa adanya.",
        ),
        (
            "2. Banding disetujui",
            "Kode absensi harian mengikuti Attendance Sesudah pada banding yang disetujui. "
            f"{banding_applied} dari {banding_approved} banding disetujui mengubah hitungan di periode ini.",
        ),
        (
            "3. Hitung lembur",
            "Fee per record dari durasi. 0–4 jam = 50.000, lebih dari 4 sampai 8 jam = 75.000, lebih dari 8 jam = 100.000. "
            f"Aturan ini cocok dengan {fee_match} dari {len(payroll.overtime)} record, termasuk durasi di bawah 1 jam.",
        ),
        (
            "4. Hitung denda alpa",
            "floor(gaji pokok / 25) dikali jumlah status A setelah banding disetujui. Satu sel 'A | A' dihitung dua. "
            f"Cocok dengan angka HRIS pada {alpha_match} dari {len(alpha_rows)} petugas yang punya alpa atau denda.",
        ),
        (
            "5. Susun slip",
            "Gaji dihitung = gaji pokok + tunjangan jabatan + BPJS + prorate + tunjangan lain + lembur dihitung "
            f"− denda alpa dihitung − deposit ({rupiah(deposit_total)}) − denda telat HRIS − potongan BPJS − potongan lain.",
        ),
        (
            "6. Trial balance",
            f"Gaji HRIS {rupiah(paid)}. Gaji dihitung {rupiah(calculated)}. Selisih {rupiah(calculated - paid)}. "
            "Selisih datang dari lembur dan denda alpa setelah banding. Deposit mengikuti kolom Deposit di file sumber.",
        ),
    ]
    ws["A4"] = "Langkah"
    ws["B4"] = "Yang dilakukan"
    for col in (1, 2):
        ws.cell(4, col).fill = fill_navy
        ws.cell(4, col).font = font_header
        ws.cell(4, col).alignment = center
        ws.cell(4, col).border = thin
    for offset, (title, detail) in enumerate(steps):
        row = 5 + offset
        ws.cell(row, 1, title).font = font_bold
        ws.cell(row, 2, detail).font = font_base
        for col in (1, 2):
            ws.cell(row, col).alignment = left_align
            ws.cell(row, col).border = thin
            ws.cell(row, col).fill = fill_pale
        ws.row_dimensions[row].height = 48

    notes = [
        ("Petugas di slip", len(officers)),
        ("Gaji HRIS", paid),
        ("Gaji dihitung", calculated),
        ("Selisih (dihitung − HRIS)", round(calculated - paid, 2)),
        ("Fee lembur tanpa slip di lokasi yang sama", sum(record.fee_hitung for record in payroll.unmatched_overtime)),
        ("Selisih susun ulang komponen HRIS", round(rebuild_gap, 2)),
    ]
    start = 12
    ws.cell(start, 1, "Hasil").font = font_section
    for offset, (label, value) in enumerate(notes):
        row = start + 1 + offset
        ws.cell(row, 1, label).font = font_bold
        ws.cell(row, 2, value).font = font_base
        ws.cell(row, 2).alignment = right_align
        ws.cell(row, 2).number_format = "#,##0" if not isinstance(value, int) or abs(value) > 1000 else "0"
        if isinstance(value, float):
            ws.cell(row, 2).number_format = "#,##0"
        for col in (1, 2):
            ws.cell(row, col).border = thin
            ws.cell(row, col).fill = fill_total if "Selisih (dihitung" in label else PatternFill("solid", fgColor="FFFFFF")
    new_joiners = sum(1 for officer in officers if officer.new_joiner)
    ws.merge_cells("A20:B20")
    ws["A20"] = (
        "Denda keterlambatan di file ini kosong, jadi tidak dimasukkan ke gaji dihitung sebagai angka baru. "
        "Jabatan dan BPJS mengikuti HRIS. Prorate mengikuti HRIS kecuali New Joiner "
        f"(absensi mulai setelah hari pertama periode): {new_joiners} petugas, "
        f"prorate = pokok × hari kerja ÷ {WORK_DAY_DIVISOR} − pokok."
    )
    ws["A20"].font = font_muted
    ws["A20"].alignment = left_align
    ws.row_dimensions[20].height = 36
    widths(ws, [42, 120])
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.sheet_view.zoomScale = 120


SLIP_HEADERS = [
    "Lokasi",
    "Nama",
    "Username",
    "Gaji Pokok",
    "Tunjangan Jabatan",
    "BPJS Kesehatan",
    "BPJS Ketenagakerjaan",
    "Prorate",
    "Tunjangan Lain",
    "Tambahan Manual",
    "Jam Lembur",
    "Lembur HRIS",
    "Lembur Dihitung",
    "Status Alpa",
    "Denda Alpa HRIS",
    "Denda Alpa Dihitung",
    "Deposit",
    "Potongan Manual",
    "Potongan Lainnya",
    "Gaji HRIS",
    "Gaji Dihitung",
    "Selisih",
    "New Joiner",
    "Mulai Absensi",
    "Hari Kerja",
]

SLIP_SKIP_TOTAL = {"Lokasi", "Nama", "Username", "New Joiner", "Mulai Absensi"}


def slip_bpjs(officer: Officer, allow: str, deduct: str) -> float:
    """Allowance adds to gaji pokok. A potongan is returned negative so the slip can show it in parentheses."""
    return officer.amounts[allow] - abs(officer.amounts[deduct])


def slip_deposit(officer: Officer) -> float:
    """Fixed deposit from the source file. Returned negative so the slip shows a deduction."""
    return -abs(officer.amounts["deposit"])


def slip_potongan(officer: Officer) -> float:
    """Potongan besides BPJS, Deposit, and manual payroll cuts."""
    return abs(officer.amounts["denda_telat"]) + abs(officer.amounts["lain_potongan"])


def slip_manual_potongan(officer: Officer) -> float:
    """On-the-spot payroll cut (Fraud / Lainnya). Deposit reason stays in Deposit."""
    return -abs(officer.amounts["manual_potongan"])


def slip_manual_tambahan(officer: Officer) -> float:
    """On-the-spot payroll addition (Banding / Refund Deposit / Lainnya)."""
    return abs(officer.amounts["manual_tambahan"])


def slip_row(officer: Officer) -> list:
    return [
        officer.lokasi,
        officer.nama,
        officer.username,
        officer.basic,
        officer.amounts["jabatan"],
        slip_bpjs(officer, "bpjs_kes", "pot_bpjs_kes"),
        slip_bpjs(officer, "bpjs_tk", "pot_bpjs_tk"),
        officer.amounts["prorate"],
        officer.amounts["lain_tunjangan"],
        slip_manual_tambahan(officer),
        round(officer.jam_lembur, 2),
        officer.lembur_hris,
        officer.lembur_hitung,
        officer.alpha,
        officer.denda_alpa_hris,
        officer.denda_alpa_hitung,
        slip_deposit(officer),
        slip_manual_potongan(officer),
        slip_potongan(officer),
        officer.paid,
        officer.gaji_hitung,
        officer.selisih_gaji,
        "Ya" if officer.new_joiner else "",
        day_label(officer.first_attendance) if officer.new_joiner and officer.first_attendance else "",
        officer.worked_days if officer.new_joiner else "",
    ]


def write_slip(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.create_sheet("Slip Gaji")
    for col, header in enumerate(SLIP_HEADERS, start=1):
        ws.cell(1, col, header)
    officers = sorted(payroll.officers, key=lambda officer: (officer.lokasi.lower(), officer.nama.lower()))
    for row_no, officer in enumerate(officers, start=2):
        for col, value in enumerate(slip_row(officer), start=1):
            ws.cell(row_no, col, value)
    total_row = len(officers) + 2
    totals = [""] * len(SLIP_HEADERS)
    totals[1] = "TOTAL"
    for pos, header in enumerate(SLIP_HEADERS):
        if header in SLIP_SKIP_TOTAL:
            continue
        totals[pos] = 0
    for officer in officers:
        values = slip_row(officer)
        for pos, header in enumerate(SLIP_HEADERS):
            if header in SLIP_SKIP_TOTAL:
                continue
            value = values[pos]
            if value == "" or value is None:
                continue
            totals[pos] += value
    for col, value in enumerate(totals, start=1):
        cell = ws.cell(total_row, col, value)
        cell.font = Font(name="Calibri", bold=True, color=NAVY, size=11)
        cell.fill = fill_total
        cell.border = thin
        header = SLIP_HEADERS[col - 1]
        if header not in SLIP_SKIP_TOTAL and header != "Hari Kerja":
            cell.number_format = "#,##0"
    money_cols = {
        col
        for col, header in enumerate(SLIP_HEADERS, start=1)
        if header not in SLIP_SKIP_TOTAL and header not in ("Jam Lembur", "Status Alpa", "Hari Kerja")
    }
    style_header(ws, len(SLIP_HEADERS))
    paint(ws, money_cols)
    for col in range(1, len(SLIP_HEADERS) + 1):
        ws.cell(total_row, col).fill = fill_total
        ws.cell(total_row, col).font = Font(name="Calibri", bold=True, color=NAVY, size=11)
    reducing = {
        col
        for col, header in enumerate(SLIP_HEADERS, start=1)
        if "BPJS" in header or header in ("Deposit", "Potongan Manual")
    }
    mark_reducing_bpjs(ws, reducing)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(SLIP_HEADERS))}{total_row - 1}"
    widths(
        ws,
        [32, 32, 22, 16, 18, 18, 22, 14, 16, 16, 14, 16, 18, 14, 18, 20, 16, 16, 18, 16, 16, 14, 12, 14, 12],
    )


def write_trial_balance(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.create_sheet("Trial Balance")
    headers = ["Komponen", "Cara hitung", "HRIS", "Dihitung", "Selisih"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col, header)
    officers = payroll.officers

    def total(reader) -> float:
        return sum(reader(officer) for officer in officers)

    rows = [
        ("Gaji pokok", "Disalin dari HRIS", total(lambda o: o.basic), total(lambda o: o.basic)),
        ("Tunjangan jabatan", "Disalin dari HRIS", total(lambda o: o.amounts["jabatan"]), total(lambda o: o.amounts["jabatan"])),
        ("BPJS Kesehatan", "Menambah dari gaji pokok", total(lambda o: o.amounts["bpjs_kes"]), total(lambda o: o.amounts["bpjs_kes"])),
        ("BPJS Ketenagakerjaan", "Menambah dari gaji pokok", total(lambda o: o.amounts["bpjs_tk"]), total(lambda o: o.amounts["bpjs_tk"])),
        ("BPJS Kesehatan", "Mengurangi dari gaji pokok", total(lambda o: -abs(o.amounts["pot_bpjs_kes"])), total(lambda o: -abs(o.amounts["pot_bpjs_kes"]))),
        ("BPJS Ketenagakerjaan", "Mengurangi dari gaji pokok", total(lambda o: -abs(o.amounts["pot_bpjs_tk"])), total(lambda o: -abs(o.amounts["pot_bpjs_tk"]))),
        (
            "Prorate",
            "HRIS disalin; New Joiner diganti ke (pokok × hari kerja ÷ 25 − pokok)",
            total(lambda o: o.prorate_hris),
            total(lambda o: o.amounts["prorate"]),
        ),
        ("Tunjangan lain", "Disalin dari HRIS", total(lambda o: o.amounts["lain_tunjangan"]), total(lambda o: o.amounts["lain_tunjangan"])),
        ("Tambahan manual", "Banding / refund deposit / lainnya", total(lambda o: o.amounts["manual_tambahan"]), total(lambda o: o.amounts["manual_tambahan"])),
        ("Tunjangan lembur", "Jumlah fee blok jam", total(lambda o: o.lembur_hris), total(lambda o: o.lembur_hitung)),
        ("Denda ketidakhadiran", "floor(pokok / 25) x status A", total(lambda o: o.denda_alpa_hris), total(lambda o: o.denda_alpa_hitung)),
        ("Denda keterlambatan", "Mengikuti HRIS (di file ini 0)", total(lambda o: o.denda_telat), total(lambda o: o.denda_telat)),
        ("Deposit", "Mengurangi dari gaji pokok, tarif tetap", total(lambda o: -abs(o.amounts["deposit"])), total(lambda o: -abs(o.amounts["deposit"]))),
        ("Potongan manual", "Fraud / lainnya dari syarat payroll", total(lambda o: -abs(o.amounts["manual_potongan"])), total(lambda o: -abs(o.amounts["manual_potongan"]))),
        ("Potongan lain", "Disalin dari HRIS", total(lambda o: abs(o.amounts["lain_potongan"])), total(lambda o: abs(o.amounts["lain_potongan"]))),
        ("Gaji dibayar", "Penambah − potongan", total(lambda o: o.paid), total(lambda o: o.gaji_hitung)),
    ]
    for row_no, (name, how, hris, calc) in enumerate(rows, start=2):
        ws.cell(row_no, 1, name)
        ws.cell(row_no, 2, how)
        ws.cell(row_no, 3, hris)
        ws.cell(row_no, 4, calc)
        ws.cell(row_no, 5, round(calc - hris, 2))
    style_header(ws, 5, fill_teal)
    paint(ws, {3, 4, 5})
    last = 1 + len(rows)
    for col in range(1, 6):
        ws.cell(last, col).fill = fill_total
        ws.cell(last, col).font = font_bold
    mark_reducing_bpjs(ws, {3, 4, 5}, only_bpjs_rows=True)
    mark_reducing_bpjs(ws, {3, 4, 5}, only_bpjs_rows=True, match="Deposit")
    mark_reducing_bpjs(ws, {3, 4, 5}, only_bpjs_rows=True, match="Potongan manual")
    widths(ws, [36, 42, 20, 20, 18])
    ws.auto_filter.ref = f"A1:E{last - 1}"


def write_lokasi(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.create_sheet("Rekap Lokasi")
    headers = ["Lokasi", "Petugas", "Gaji HRIS", "Gaji Dihitung", "Selisih", "Selisih Lembur", "Selisih Denda Alpa"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col, header)
    grouped: dict[str, list[Officer]] = defaultdict(list)
    for officer in payroll.officers:
        grouped[officer.lokasi].append(officer)
    row_no = 2
    for lokasi in sorted(grouped):
        people = grouped[lokasi]
        hris = sum(officer.paid for officer in people)
        calc = sum(officer.gaji_hitung for officer in people)
        lembur = sum(officer.lembur_hitung - officer.lembur_hris for officer in people)
        denda = sum(officer.denda_alpa_hitung - officer.denda_alpa_hris for officer in people)
        values = [lokasi, len(people), hris, calc, round(calc - hris, 2), round(lembur, 2), round(denda, 2)]
        for col, value in enumerate(values, start=1):
            ws.cell(row_no, col, value)
        row_no += 1
    totals = ["TOTAL", 0, 0, 0, 0, 0, 0]
    for source in range(2, row_no):
        for col in range(2, 8):
            totals[col - 1] += money(ws.cell(source, col).value)
    for col, value in enumerate(totals, start=1):
        cell = ws.cell(row_no, col, value)
        cell.fill = fill_total
        cell.font = font_bold
        cell.border = thin
        if col >= 2:
            cell.number_format = "#,##0"
    style_header(ws, len(headers))
    paint(ws, {3, 4, 5, 6, 7})
    for col in range(1, 8):
        ws.cell(row_no, col).fill = fill_total
        ws.cell(row_no, col).font = font_bold
    ws.auto_filter.ref = f"A1:G{row_no - 1}"
    widths(ws, [36, 12, 18, 18, 16, 18, 20])


def write_overtime(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.create_sheet("Rincian Lembur")
    headers = [
        "Lokasi",
        "Nama",
        "Username",
        "Mulai",
        "Akhir",
        "Menit",
        "Jam",
        "Blok",
        "Fee HRIS",
        "Fee Dihitung",
        "Cocok",
        "Masuk slip",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col, header)
    records = sorted(payroll.overtime, key=lambda record: (record.lokasi.lower(), record.nama.lower(), record.mulai))
    for row_no, record in enumerate(records, start=2):
        match = abs(record.fee_hitung - record.fee_hris) < 1
        values = [
            record.lokasi,
            record.nama,
            record.username,
            record.mulai,
            record.akhir,
            record.menit,
            round(record.jam, 2),
            record.blok,
            record.fee_hris,
            record.fee_hitung,
            "Cocok" if match else "Beda",
            "Ya" if record.ada_slip else "Tidak",
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row_no, col, value)
        color = PALE_GREEN if match and record.ada_slip else PALE_AMBER if match else PALE_RED
        for col in range(1, len(headers) + 1):
            ws.cell(row_no, col).fill = PatternFill("solid", fgColor=color)
    style_header(ws, len(headers), fill_teal)
    paint(ws, {6, 7, 9, 10})
    for row_no, record in enumerate(records, start=2):
        match = abs(record.fee_hitung - record.fee_hris) < 1
        color = PALE_GREEN if match and record.ada_slip else PALE_AMBER if match else PALE_RED
        for col in range(1, len(headers) + 1):
            ws.cell(row_no, col).fill = PatternFill("solid", fgColor=color)
    widths(ws, [32, 32, 22, 22, 22, 12, 10, 12, 14, 16, 12, 14])


def write_denda(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.create_sheet("Denda Alpa")
    headers = [
        "Lokasi",
        "Nama",
        "Username",
        "Gaji Pokok",
        "Status Alpa",
        "Tarif per Alpa",
        "Denda HRIS",
        "Denda Dihitung",
        "Selisih",
        "Tanggal Alpa dipotong",
        "Status",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col, header)
    attendance_by_user: dict[str, list[AttendanceLine]] = defaultdict(list)
    for line in payroll.attendance:
        key = (line.username or "").strip().lower()
        if key:
            attendance_by_user[key].append(line)
    rows = []
    for officer in payroll.officers:
        if officer.alpha == 0 and officer.denda_alpa_hris == 0:
            continue
        gap = round(officer.denda_alpa_hitung - officer.denda_alpa_hris, 2)
        if abs(gap) < 1:
            status = "Cocok"
        elif officer.denda_alpa_hris == 0:
            status = "HRIS belum memotong"
        elif officer.alpha == 0:
            status = "HRIS memotong tanpa status A"
        else:
            status = "Jumlah hari yang dipotong berbeda"
        day_bits = []
        for line in attendance_by_user.get((officer.username or "").strip().lower(), []):
            label = alpha_day_labels(line, payroll.attendance_dates)
            if label:
                day_bits.append(label)
        rows.append(
            [
                officer.lokasi,
                officer.nama,
                officer.username,
                officer.basic,
                officer.alpha,
                int(officer.basic // 25) if officer.basic else 0,
                officer.denda_alpa_hris,
                officer.denda_alpa_hitung,
                gap,
                " · ".join(day_bits) if day_bits else ("—" if officer.alpha == 0 else "Lihat Detail Absensi"),
                status,
            ]
        )
    rows.sort(key=lambda item: (item[10] == "Cocok", item[0].lower(), item[1].lower()))
    for row_no, values in enumerate(rows, start=2):
        for col, value in enumerate(values, start=1):
            ws.cell(row_no, col, value)
    style_header(ws, len(headers))
    paint(ws, {4, 6, 7, 8, 9})
    widths(ws, [28, 28, 20, 14, 12, 14, 14, 16, 12, 42, 34])


def write_banding(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.create_sheet("Banding")
    headers = [
        "Lokasi",
        "Nama",
        "Username",
        "Tanggal",
        "Sebelum",
        "Sesudah",
        "Status",
        "Dipakai hitung",
        "Nilai dikembalikan",
        "Efek",
        "Alasan",
        "Label",
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col, header)
    rows = sorted(
        payroll.banding,
        key=lambda row: (
            0 if "APPROVED" in row.status.upper() else 1 if "PENDING" in row.status.upper() else 2,
            row.lokasi.lower(),
            row.nama.lower(),
            row.day,
        ),
    )
    for row_no, row in enumerate(rows, start=2):
        status = "Disetujui" if "APPROVED" in row.status.upper() else "Menunggu" if "PENDING" in row.status.upper() else "Ditolak" if "REJECTED" in row.status.upper() else row.status
        values = [
            row.lokasi,
            row.nama,
            row.username,
            row.day,
            row.before,
            row.after,
            status,
            "Ya" if row.applied else "Tidak",
            row.refund if row.refund else ("" if "APPROVED" not in row.status.upper() else 0),
            row.effect,
            row.reason,
            row.label,
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row_no, col, value)
    style_header(ws, len(headers))
    paint(ws, {9})
    widths(ws, [26, 26, 18, 12, 12, 12, 12, 12, 18, 36, 18, 26])
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"


def write_attendance(wb: Workbook, payroll: Payroll) -> None:
    ws = wb.create_sheet("Detail Absensi")
    headers = ["Lokasi", "Nama", "Username", *payroll.attendance_dates]
    headers.extend(label for _code, label in ATTEND_CODES)
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col, header)
    count_at = 4 + len(payroll.attendance_dates)
    for row_no, line in enumerate(payroll.attendance, start=2):
        values = [line.lokasi, line.nama, line.username, *line.days]
        values.extend(line.counts[code] for code, _label in ATTEND_CODES)
        for col, value in enumerate(values, start=1):
            ws.cell(row_no, col, value)
    total_row = len(payroll.attendance) + 2
    totals = ["", "TOTAL", ""] + [""] * len(payroll.attendance_dates)
    for code, _label in ATTEND_CODES:
        totals.append(sum(line.counts[code] for line in payroll.attendance))
    for col, value in enumerate(totals, start=1):
        cell = ws.cell(total_row, col, value)
        cell.font = Font(name="Calibri", bold=True, color=NAVY, size=11)
        cell.fill = fill_total
        cell.border = thin
        if col >= count_at and isinstance(value, int):
            cell.number_format = "#,##0"
    style_header(ws, len(headers))
    paint(ws, set(range(count_at, len(headers) + 1)))
    for col in range(1, len(headers) + 1):
        ws.cell(total_row, col).fill = fill_total
        ws.cell(total_row, col).font = Font(name="Calibri", bold=True, color=NAVY, size=11)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(total_row - 1, 1)}"
    ws.freeze_panes = "D2"
    widths(ws, [32, 32, 22] + [8] * len(payroll.attendance_dates) + [14] * len(ATTEND_CODES))


def write_excluded(wb: Workbook, rows: list[dict]) -> None:
    """Replace the Tidak Dihitung sheet with the current exclusion list."""
    title = "Tidak Dihitung"
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    headers = ["Nama", "Username", "Lokasi", "Alasan", "Status"]
    for col, header in enumerate(headers, start=1):
        ws.cell(1, col, header)
    ordered = sorted(rows, key=lambda item: ((item.get("nama") or "").casefold(), (item.get("username") or "").casefold()))
    for row_no, item in enumerate(ordered, start=2):
        values = [item.get("nama") or "", item.get("username") or "", item.get("lokasi") or "", item.get("alasan") or "", item.get("status_label") or ""]
        for col, value in enumerate(values, start=1):
            ws.cell(row_no, col, value)
    style_header(ws, len(headers))
    paint(ws, set())
    widths(ws, [32, 22, 36, 28, 18])
    ws.auto_filter.ref = f"A1:E{max(len(ordered) + 1, 1)}"
    note = ws.cell(len(ordered) + 3, 1, "Petugas berstatus Blacklist atau Resign, dan akun testing, tidak masuk prosedur payroll.")
    note.font = font_muted


def filter_payroll(payroll: Payroll, blocked: set[str]) -> Payroll:
    """Keep only people who are eligible for the payroll run."""
    blocked = {name.strip().lower() for name in blocked if name and name.strip()}
    if not blocked:
        return payroll

    def allowed(username: str) -> bool:
        return username.strip().lower() not in blocked

    payroll.officers = [officer for officer in payroll.officers if allowed(officer.username)]
    payroll.overtime = [record for record in payroll.overtime if allowed(record.username)]
    payroll.unmatched_overtime = [record for record in payroll.unmatched_overtime if allowed(record.username)]
    payroll.attendance = [line for line in payroll.attendance if allowed(line.username)]
    payroll.banding = [row for row in payroll.banding if allowed(row.username)]
    return payroll


def write_workbook(payroll: Payroll, output: Path) -> None:
    wb = Workbook()
    write_alur(wb, payroll)
    write_slip(wb, payroll)
    write_trial_balance(wb, payroll)
    write_lokasi(wb, payroll)
    write_overtime(wb, payroll)
    write_denda(wb, payroll)
    write_banding(wb, payroll)
    write_attendance(wb, payroll)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def print_summary(payroll: Payroll, output: Path) -> None:
    paid = sum(officer.paid for officer in payroll.officers)
    calculated = sum(officer.gaji_hitung for officer in payroll.officers)
    fee_match = sum(1 for record in payroll.overtime if abs(record.fee_hitung - record.fee_hris) < 1)
    print(f"Sumber      : {payroll.source_name}")
    print(f"Periode     : {payroll.period or '-'}")
    print(f"Petugas     : {len(payroll.officers)}")
    print(f"Gaji HRIS   : {rupiah(paid)}")
    print(f"Gaji hitung : {rupiah(calculated)}")
    print(f"Selisih     : {rupiah(calculated - paid)}")
    print(f"Fee lembur  : {fee_match}/{len(payroll.overtime)} record cocok dengan blok jam")
    print(f"File        : {output}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Hitung payroll dari file tarikan HRIS.")
    parser.add_argument("source", type=Path, help="File .xlsx hasil tarikan payroll")
    parser.add_argument("-o", "--output", type=Path, help="File hasil hitung")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source = args.source.expanduser().resolve()
    if not source.exists():
        raise SystemExit(f"File tidak ditemukan: {source}")
    output = args.output or Path("output") / f"hitung_{source.stem}.xlsx"
    payroll = build(source)
    write_workbook(payroll, output)
    print_summary(payroll, output)


if __name__ == "__main__":
    main()
