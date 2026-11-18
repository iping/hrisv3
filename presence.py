"""Absensi yang dihitung di aplikasi ini.

Lokasi, petugas, dan penempatan adalah master. Jadwal harian ditambah jam
masuk/keluar menghasilkan kode. Banding yang disetujui mengganti kode itu.
Payroll nanti membaca kode yang tersisa, bukan pengajuannya.
"""

from __future__ import annotations

import calendar
import re
from datetime import date, datetime, time, timedelta

from auth import actor_snapshot, format_actor_label
from db import connect, stamp
from roster import load_roster

LABELS = {
    "OT": "Hadir",
    "L": "Terlambat",
    "A": "Alpa",
    "P": "Izin",
    "S": "Sakit",
    "LEAVE": "Cuti",
    "OFF": "Off",
    "O": "Tanpa jadwal",
}
GPS_RADII = (20, 50, 100, 300, 500)
PLANNED = ("OFF", "LEAVE", "P", "S")
REQUESTS = ("OT", "L", "A", "P", "S", "LEAVE", "OFF")
GENDERS = (("laki", "Laki-laki"), ("perempuan", "Perempuan"))
MARITAL = (("married", "Married"), ("divorce", "Divorce"), ("single", "Not Married"))
# code, label, planned attendance code, genders or None, marital statuses or None
OFF_REASONS = (
    ("cuti_tahunan", "Cuti tahunan", "LEAVE", None, None),
    ("sakit", "Sakit", "S", None, None),
    ("izin", "Izin", "P", None, None),
    ("kedukaan", "Kedukaan", "LEAVE", None, None),
    ("haid", "Cuti haid", "LEAVE", ("perempuan",), None),
    ("melahirkan", "Cuti melahirkan", "LEAVE", ("perempuan",), None),
    ("istri_melahirkan", "Cuti istri melahirkan", "LEAVE", ("laki",), ("married",)),
    ("menikah", "Cuti menikah", "LEAVE", None, ("single", "divorce")),
)
ATTEND_CODES = ("OT", "L")
OFF_CODES = ("OFF", "LEAVE", "P", "S")
MONTH_NAMES = (
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
)
EMPLOYMENT = (
    ("onboard", "Onboard"),
    ("approved", "Approved"),
    ("aktif", "Aktif"),
    ("resign", "Resign"),
    ("ojt", "OJT"),
    ("lulus", "Pass OJT"),
    ("blacklist", "Blacklist"),
    ("baru", "New Join"),
    ("mutasi", "Mutasi"),
    ("tunjangan", "Tunjangan"),
)
EMPLOYMENT_CODES = {code for code, _name in EMPLOYMENT}
ACCOUNT_CHOICES = (
    ("approved", "Approved"),
    ("aktif", "Aktif"),
    ("ojt", "OJT"),
    ("lulus", "Pass OJT"),
    ("resign", "Resign"),
    ("blacklist", "Blacklist"),
)
STATUS_DATE_LABELS = {
    "approved": "Tanggal disetujui",
    "aktif": "Tanggal mulai aktif",
    "ojt": "Tanggal mulai OJT",
    "lulus": "Tanggal lulus OJT",
    "resign": "Tanggal resign",
    "blacklist": "Tanggal blacklist",
}
STATUS_HELP = (
    ("onboard", "Onboard", "Register awal. Belum disetujui, dan belum masuk absensi."),
    ("approved", "Approved", "Register sudah disetujui. Belum bisa dipakai untuk absensi."),
    ("aktif", "Aktif", "Sudah bisa dipakai untuk absensi."),
    ("resign", "Resign", "Sudah resign. Penempatan ditutup dan tidak masuk absensi."),
    ("new", "New Joiner", "Baru aktif, 25 hari atau kurang. Ini menempel pada Aktif, OJT, atau Pass OJT."),
    ("ojt", "OJT", "Memakai profil gaji Selama OJT, dan bisa masuk absensi."),
    ("lulus", "Pass OJT", "Memakai profil gaji Setelah OJT, dan bisa masuk absensi."),
    ("blacklist", "Blacklist", "Tidak dihitung payroll dan tidak masuk absensi."),
)
ATTENDANCE_STATUSES = ("aktif", "ojt", "lulus", "mutasi", "tunjangan", "baru")
NEW_JOINER_DAYS = 25
OJT_MONTHS = (3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
PAYROLL_POTONGAN_REASONS = (
    ("deposit", "Deposit"),
    ("fraud", "Fraud"),
    ("lainnya", "Lainnya"),
)
PAYROLL_TAMBAHAN_REASONS = (
    ("banding", "Banding disetujui"),
    ("refund_deposit", "Refund Deposit"),
    ("lainnya", "Lainnya"),
)
PAYROLL_POTONGAN_CODES = {code for code, _name in PAYROLL_POTONGAN_REASONS}
PAYROLL_TAMBAHAN_CODES = {code for code, _name in PAYROLL_TAMBAHAN_REASONS}


def derive_code(planned: str | None, start: time | None, grace: int, clock_in: time | None, ends_next_day: bool = False) -> str:
    """Satu slot, satu kode. Cuti yang sudah direncanakan menang atas jam absen."""
    if planned in PLANNED:
        return planned
    if start is None and clock_in is None:
        return "O"
    if start is not None and clock_in is None:
        return "A"
    if start is None:
        return "OT"
    limit = _minutes(start) + int(grace or 0)
    clock = _minutes(clock_in)
    if ends_next_day and clock < _minutes(start):
        clock += 24 * 60
    if clock <= limit:
        return "OT"
    return "L"


def worked_minutes(clock_in: time | None, clock_out: time | None) -> int | None:
    if clock_in is None or clock_out is None:
        return None
    start = _minutes(clock_in)
    end = _minutes(clock_out)
    if end <= start:
        end += 24 * 60
    return end - start


def hours_text(minutes) -> str:
    if minutes is None:
        return ""
    hours = int(minutes) / 60
    if hours == int(hours):
        return str(int(hours))
    return f"{hours:.1f}"


def parse_days(value: str | None) -> tuple[int | None, str]:
    raw = (value or "").strip()
    if raw == "":
        return None, ""
    if not raw.isdigit():
        return None, "Jumlah hari harus angka."
    number = int(raw)
    if number > 31:
        return None, "Jumlah hari dalam sebulan paling banyak 31."
    return number, ""


def reasons_for(gender: str, marital: str) -> list[dict]:
    found = []
    for code, label, planned, genders, statuses in OFF_REASONS:
        if genders and gender not in genders:
            continue
        if statuses and marital not in statuses:
            continue
        found.append({"code": code, "label": label, "planned": planned})
    return found


def reason_record(code: str) -> dict | None:
    for item_code, label, planned, _genders, _statuses in OFF_REASONS:
        if item_code == code:
            return {"code": item_code, "label": label, "planned": planned}
    return None


def plan_value(planned: str | None, off_reason: str | None) -> str:
    if off_reason:
        return f"reason:{off_reason}"
    if planned == "OFF":
        return "OFF"
    if planned == "S":
        return "reason:sakit"
    if planned == "P":
        return "reason:izin"
    if planned == "LEAVE":
        return "LEAVE"
    return ""


def _resolve_plan(raw: str, gender: str, marital: str, current_reason: str = "") -> tuple[str | None, str, str]:
    raw = raw or ""
    if raw in ("", "kerja"):
        return None, "", ""
    if raw == "OFF":
        return "OFF", "", ""
    if raw in PLANNED:
        return raw, "", ""
    if not raw.startswith("reason:"):
        return None, "", ""
    code = raw.split(":", 1)[1]
    record = reason_record(code)
    if not record:
        return None, "", "Pilih alasan off yang tersedia."
    allowed = {item["code"] for item in reasons_for(gender, marital)}
    if code not in allowed and code != (current_reason or ""):
        return None, "", "Alasan off tidak sesuai jenis kelamin atau status nikah."
    return record["planned"], code, ""


def parse_hours(value: str | None) -> tuple[int | None, str]:
    raw = (value or "").strip().replace(",", ".")
    if raw == "":
        return None, ""
    try:
        hours = float(raw)
    except ValueError:
        return None, "Jam kehadiran harus angka."
    if hours < 0:
        return None, "Jam kehadiran tidak boleh minus."
    return int(round(hours * 60)), ""


def _below_minimum(row, location_minutes=None) -> bool:
    minimum = row.get("min_attend_minutes")
    if minimum is None:
        minimum = location_minutes
    if not row.get("slot_id") or minimum is None:
        return False
    worked = worked_minutes(row.get("clock_in"), row.get("clock_out"))
    if worked is None:
        return False
    return worked < int(minimum)


def overtime_from_max(clock_in, clock_out, max_minutes, shift_overtime: float) -> float:
    """A personal maximum replaces the shift end. Time past that maximum is overtime."""
    if max_minutes is None:
        return shift_overtime
    worked = worked_minutes(clock_in, clock_out)
    if worked is None:
        return 0.0
    extra = worked - int(max_minutes)
    if extra <= 0:
        return 0.0
    return round(extra / 60, 2)


def overtime_hours(end: time | None, clock_out: time | None, ends_next_day: bool = False, start: time | None = None) -> float:
    if end is None or clock_out is None:
        return 0.0
    end_at = _minutes(end)
    out_at = _minutes(clock_out)
    if ends_next_day:
        end_at += 24 * 60
        if start is not None and out_at < _minutes(start):
            out_at += 24 * 60
    extra = out_at - end_at
    if extra <= 0:
        return 0.0
    return round(extra / 60, 2)


def _minutes(value: time) -> int:
    return value.hour * 60 + value.minute


def timebox_text(start, end, ends_next_day: bool) -> str:
    start_text = clock_text(start)
    end_text = clock_text(end)
    if ends_next_day:
        return f"{start_text} hari ini – {end_text} besok"
    return f"{start_text}–{end_text} hari ini"


def _shift_values(name: str, start: str, end: str, grace: str, hari: str):
    clean = " ".join((name or "").split())
    start_time = parse_time(start)
    end_time = parse_time(end)
    if not clean or start_time is None or end_time is None:
        return None, "Isi nama shift, jam mulai, dan jam selesai."
    raw_grace = (grace or "").strip()
    if raw_grace == "":
        grace_minutes = None
    else:
        try:
            grace_minutes = max(0, int(raw_grace))
        except ValueError:
            return None, "Tenggang terlambat harus angka, atau dikosongkan supaya ikut lokasi."
    ends_next_day = hari == "besok"
    if not ends_next_day and _minutes(end_time) <= _minutes(start_time):
        return None, "Jam selesai lebih awal dari jam mulai. Pilih Besok kalau shift lanjut ke hari berikutnya."
    return (clean, start_time, end_time, grace_minutes, ends_next_day), ""


def parse_month(value: str | None) -> date:
    raw = (value or "")[:7]
    try:
        return datetime.strptime(raw, "%Y-%m").date().replace(day=1)
    except ValueError:
        today = date.today()
        return today.replace(day=1)


def add_months(day: date, months: int) -> date:
    index = day.month - 1 + months
    year = day.year + index // 12
    month = index % 12 + 1
    last = month_last(date(year, month, 1)).day
    return date(year, month, min(day.day, last))


def month_last(first: date) -> date:
    next_month = (first.replace(day=28) + timedelta(days=4)).replace(day=1)
    return next_month - timedelta(days=1)


def parse_day(value: str | None) -> date:
    raw = (value or "")[:10]
    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError:
        return date.today()


def parse_time(value: str | None) -> time | None:
    raw = (value or "").strip()
    if not raw:
        return None
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(raw[:8], fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            continue
    return None


def clock_text(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, timedelta):
        total = int(value.total_seconds())
        return f"{total // 3600:02d}:{(total % 3600) // 60:02d}"
    return str(value)[:5]


def label(code: str) -> str:
    if not code:
        return ""
    return LABELS.get(code, code)


def ensure_seeded(db) -> None:
    existing = db.execute("SELECT COUNT(*) AS n FROM attendance_locations").fetchone()
    if existing and existing["n"]:
        return
    roster = load_roster()
    locations = roster.get("locations") or {}
    if not locations:
        return
    loc_ids: dict[str, int] = {}
    for name, row in locations.items():
        active = bool(row.get("active", True))
        created = db.execute(
            """
            INSERT INTO attendance_locations (name, active, payable)
            VALUES (?, ?, ?)
            ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
            RETURNING id
            """,
            (name, active, active),
        ).fetchone()
        loc_ids[name] = created["id"]
        db.execute(
            """
            INSERT INTO attendance_shifts (location_id, name, start_time, end_time)
            VALUES (?, 'Shift 1', '08:00', '17:00')
            """,
            (created["id"],),
        )
    officers: dict[str, int] = {}
    for row in (roster.get("people") or {}).values():
        username = (row.get("username") or "").strip()
        nama = (row.get("nama") or "").strip() or username
        lokasi = row.get("lokasi") or ""
        if not username or lokasi not in loc_ids:
            continue
        key = username.lower()
        officer_id = officers.get(key)
        if officer_id is None:
            created = db.execute(
                """
                INSERT INTO attendance_officers (username, nama, testing)
                VALUES (?, ?, ?)
                ON CONFLICT (username) DO UPDATE SET nama = attendance_officers.nama
                RETURNING id
                """,
                (username, nama, bool(row.get("testing"))),
            ).fetchone()
            officer_id = created["id"]
            officers[key] = officer_id
        status = row.get("status") or "aktif"
        tanggal = (row.get("tanggal") or "")[:10]
        start = tanggal if status == "baru" and len(tanggal) == 10 else "2020-01-01"
        end = tanggal if status == "resign" and len(tanggal) == 10 else None
        db.execute(
            """
            INSERT INTO attendance_placements (officer_id, location_id, start_date, end_date)
            VALUES (?, ?, ?, ?)
            ON CONFLICT (officer_id, location_id, start_date) DO NOTHING
            """,
            (officer_id, loc_ids[lokasi], start, end),
        )


def locations() -> list[dict]:
    with connect() as db:
        ensure_seeded(db)
        return db.execute(
            """
            SELECT id, name, active, payable, grace_minutes
            FROM attendance_locations
            ORDER BY active DESC, name
            """
        ).fetchall()


def _default_shift(db, location_id: int) -> int | None:
    row = db.execute(
        """
        SELECT id FROM attendance_shifts
        WHERE location_id = ?
        ORDER BY id
        LIMIT 1
        """,
        (location_id,),
    ).fetchone()
    if row:
        return row["id"]
    created = db.execute(
        """
        INSERT INTO attendance_shifts (location_id, name, start_time, end_time)
        VALUES (?, 'Shift 1', '08:00', '17:00')
        RETURNING id
        """,
        (location_id,),
    ).fetchone()
    return created["id"]


def shifts_for(location_id: int) -> list[dict]:
    with connect() as db:
        ensure_seeded(db)
        rows = db.execute(
            """
            SELECT s.id, s.name, s.start_time, s.end_time, s.grace_minutes, s.ends_next_day,
                   l.grace_minutes AS location_grace
            FROM attendance_shifts s
            JOIN attendance_locations l ON l.id = s.location_id
            WHERE s.location_id = ?
            ORDER BY s.start_time, s.name
            """,
            (location_id,),
        ).fetchall()
    for row in rows:
        row["start_text"] = clock_text(row["start_time"])
        row["end_text"] = clock_text(row["end_time"])
        row["span_text"] = timebox_text(row["start_time"], row["end_time"], bool(row["ends_next_day"]))
        row["follows_location"] = row["grace_minutes"] is None
        row["grace_used"] = row["location_grace"] if row["follows_location"] else row["grace_minutes"]
    return rows


def add_location(name: str, grace: str = "15", minimum: str = "", work_days: str = "", off_days: str = "") -> str:
    clean = " ".join((name or "").split())
    if not clean:
        return "Nama lokasi kosong."
    minutes, error = _location_grace(grace)
    if error:
        return error
    minimum_minutes, min_error = parse_hours(minimum)
    if min_error:
        return min_error
    work_day_count, work_error = parse_days(work_days)
    if work_error:
        return work_error
    off_day_count, off_error = parse_days(off_days)
    if off_error:
        return off_error
    with connect() as db:
        ensure_seeded(db)
        found = db.execute(
            "SELECT id FROM attendance_locations WHERE lower(name) = lower(?)",
            (clean,),
        ).fetchone()
        if found:
            return "Lokasi itu sudah ada."
        created = db.execute(
            """
            INSERT INTO attendance_locations
                (name, grace_minutes, min_attend_minutes, min_work_days, min_off_days)
            VALUES (?, ?, ?, ?, ?) RETURNING id
            """,
            (clean, minutes, minimum_minutes, work_day_count, off_day_count),
        ).fetchone()
        _default_shift(db, created["id"])
        facts = [(label, value) for label, value in _location_facts(db, created["id"]) if value != "Belum diatur"]
        _log_location_fields(db, created["id"], "profil", [(label, "") for label, _value in facts], facts)
        shift = db.execute(
            """
            SELECT name, start_time, end_time, grace_minutes, ends_next_day
            FROM attendance_shifts WHERE location_id = ? ORDER BY id LIMIT 1
            """,
            (created["id"],),
        ).fetchone()
        if shift:
            _add_location_log(
                db,
                created["id"],
                "shift",
                "Shift",
                "",
                _shift_story(shift["name"], shift["start_time"], shift["end_time"], shift["grace_minutes"], shift["ends_next_day"]),
            )
    return ""


def add_shift(location_id: int, name: str, start: str, end: str, grace: str, hari: str = "ini") -> str:
    values, error = _shift_values(name, start, end, grace, hari)
    if error:
        return error
    clean, start_time, end_time, grace_minutes, ends_next_day = values
    with connect() as db:
        db.execute(
            """
            INSERT INTO attendance_shifts (location_id, name, start_time, end_time, grace_minutes, ends_next_day)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (location_id, clean, start_time, end_time, grace_minutes, ends_next_day),
        )
        _add_location_log(
            db,
            location_id,
            "shift",
            "Shift",
            "",
            _shift_story(clean, start_time, end_time, grace_minutes, ends_next_day),
        )
    return ""


def add_officer(nama: str, username: str, location_id: int, start: str, testing: bool) -> str:
    clean_name = " ".join((nama or "").split())
    clean_user = (username or "").strip()
    if not clean_name or not clean_user or not location_id:
        return "Isi nama, username, dan lokasi."
    start_day = parse_day(start).isoformat()
    with connect() as db:
        ensure_seeded(db)
        found = db.execute(
            "SELECT id FROM attendance_officers WHERE lower(username) = lower(?)",
            (clean_user,),
        ).fetchone()
        if found:
            officer_id = found["id"]
            if testing:
                db.execute("UPDATE attendance_officers SET testing = TRUE WHERE id = ?", (officer_id,))
        else:
            created = db.execute(
                """
                INSERT INTO attendance_officers (username, nama, joined_date, testing)
                VALUES (?, ?, ?, ?)
                RETURNING id
                """,
                (clean_user, clean_name, start_day, testing),
            ).fetchone()
            officer_id = created["id"]
        open_row = db.execute(
            """
            SELECT id FROM attendance_placements
            WHERE officer_id = ? AND location_id = ? AND (end_date IS NULL OR end_date >= ?)
            """,
            (officer_id, location_id, start_day),
        ).fetchone()
        if open_row:
            return "Petugas ini sudah ditempatkan di lokasi itu."
        db.execute(
            """
            INSERT INTO attendance_placements (officer_id, location_id, start_date)
            VALUES (?, ?, ?)
            """,
            (officer_id, location_id, start_day),
        )
    return ""


def _recalc(db, slot_id: int) -> None:
    row = db.execute(
        """
        SELECT s.planned_code, s.clock_in, sh.start_time, sh.ends_next_day,
               COALESCE(sh.grace_minutes, l.grace_minutes, 0) AS grace_minutes
        FROM attendance_slots s
        LEFT JOIN attendance_shifts sh ON sh.id = s.shift_id
        LEFT JOIN attendance_locations l ON l.id = s.location_id
        WHERE s.id = ?
        """,
        (slot_id,),
    ).fetchone()
    if not row:
        return
    calculated = derive_code(
        row["planned_code"],
        row["start_time"],
        row["grace_minutes"] or 0,
        row["clock_in"],
        bool(row["ends_next_day"]),
    )
    approved = db.execute(
        """
        SELECT after_code FROM attendance_banding
        WHERE slot_id = ? AND status = 'approved'
        ORDER BY id DESC
        LIMIT 1
        """,
        (slot_id,),
    ).fetchone()
    if approved and approved["after_code"]:
        code = approved["after_code"]
        source = "banding"
    else:
        code = calculated
        source = "hitung"
    db.execute(
        """
        UPDATE attendance_slots
        SET code = ?, calculated_code = ?, code_source = ?
        WHERE id = ?
        """,
        (code, calculated, source, slot_id),
    )


def add_second_slot(location_id: int, officer_id: int, work_date: date, shift_id: int | None = None) -> str:
    with connect() as db:
        existing = db.execute(
            """
            SELECT COALESCE(MAX(slot_no), 0) AS n
            FROM attendance_slots
            WHERE officer_id = ? AND location_id = ? AND work_date = ?
            """,
            (officer_id, location_id, work_date),
        ).fetchone()
        if not existing or existing["n"] < 1:
            return "Pilih shift hari ini dulu, lalu simpan."
        if existing["n"] >= 2:
            return "Satu hari paling banyak dua slot."
        if not shift_id:
            shift_id = _default_shift(db, location_id)
        created = db.execute(
            """
            INSERT INTO attendance_slots (officer_id, location_id, work_date, shift_id, slot_no)
            VALUES (?, ?, ?, ?, ?)
            RETURNING id
            """,
            (officer_id, location_id, work_date, shift_id, existing["n"] + 1),
        ).fetchone()
        _recalc(db, created["id"])
    return ""


def _officer_identity(db, officer_id: int) -> tuple[str, str]:
    row = db.execute(
        "SELECT gender, marital_status FROM attendance_officers WHERE id = ?",
        (officer_id,),
    ).fetchone()
    if not row:
        return "", ""
    return row["gender"] or "", row["marital_status"] or ""


class _PlanError(Exception):
    def __init__(self, message: str):
        self.message = message


def save_slots(location_id: int, work_date: date, form) -> str:
    try:
        return _save_slots(location_id, work_date, form)
    except _PlanError as exc:
        return exc.message


def _save_slots(location_id: int, work_date: date, form) -> str:
    with connect() as db:
        for key in form._data:
            if not key.startswith("slot-"):
                continue
            slot_id = int(key.split("-", 1)[1])
            current = db.execute(
                """
                SELECT officer_id, off_reason
                FROM attendance_slots
                WHERE id = ? AND location_id = ? AND work_date = ?
                """,
                (slot_id, location_id, work_date),
            ).fetchone()
            if not current:
                continue
            gender, marital = _officer_identity(db, current["officer_id"])
            planned, reason, error = _resolve_plan(
                form.get(f"plan-{slot_id}") or "",
                gender,
                marital,
                current["off_reason"] or "",
            )
            if error:
                raise _PlanError(error)
            shift_raw = form.get(f"shift-{slot_id}") or ""
            shift_id = int(shift_raw) if shift_raw.isdigit() else None
            db.execute(
                """
                UPDATE attendance_slots
                SET planned_code = ?, off_reason = ?, shift_id = ?, clock_in = ?, clock_out = ?
                WHERE id = ? AND location_id = ? AND work_date = ?
                """,
                (
                    planned,
                    reason,
                    shift_id,
                    parse_time(form.get(f"in-{slot_id}")),
                    parse_time(form.get(f"out-{slot_id}")),
                    slot_id,
                    location_id,
                    work_date,
                ),
            )
            _recalc(db, slot_id)
        allowed = {
            row["id"]
            for row in db.execute(
                "SELECT id FROM attendance_shifts WHERE location_id = ?",
                (location_id,),
            ).fetchall()
        }
        for key in form._data:
            if not key.startswith("baru-"):
                continue
            officer_raw = key.split("-", 1)[1]
            shift_raw = form.get(key) or ""
            if not officer_raw.isdigit() or not shift_raw.isdigit() or int(shift_raw) not in allowed:
                continue
            officer_id = int(officer_raw)
            placed = db.execute(
                """
                SELECT 1 FROM attendance_placements
                WHERE officer_id = ? AND location_id = ?
                  AND start_date <= ? AND (end_date IS NULL OR end_date >= ?)
                """,
                (officer_id, location_id, work_date, work_date),
            ).fetchone()
            if not placed:
                continue
            exists = db.execute(
                """
                SELECT id FROM attendance_slots
                WHERE officer_id = ? AND location_id = ? AND work_date = ? AND slot_no = 1
                """,
                (officer_id, location_id, work_date),
            ).fetchone()
            if exists:
                continue
            gender, marital = _officer_identity(db, officer_id)
            planned, reason, error = _resolve_plan(
                form.get(f"planb-{officer_id}") or "",
                gender,
                marital,
            )
            if error:
                raise _PlanError(error)
            created = db.execute(
                """
                INSERT INTO attendance_slots
                    (officer_id, location_id, work_date, shift_id, slot_no,
                     planned_code, off_reason, clock_in, clock_out)
                VALUES (?, ?, ?, ?, 1, ?, ?, ?, ?)
                RETURNING id
                """,
                (
                    officer_id,
                    location_id,
                    work_date,
                    int(shift_raw),
                    planned,
                    reason,
                    parse_time(form.get(f"inb-{officer_id}")),
                    parse_time(form.get(f"outb-{officer_id}")),
                ),
            ).fetchone()
            _recalc(db, created["id"])
    return ""


def _assignment_state(db, location_id: int, work_date: date) -> dict[int, tuple[str, str]]:
    rows = db.execute(
        """
        SELECT o.id, o.nama, s.name AS shift_name, sl.planned_code
        FROM attendance_placements p
        JOIN attendance_officers o ON o.id = p.officer_id
        LEFT JOIN attendance_slots sl
          ON sl.officer_id = o.id AND sl.location_id = p.location_id
         AND sl.work_date = ? AND sl.slot_no = 1
        LEFT JOIN attendance_shifts s ON s.id = sl.shift_id
        WHERE p.location_id = ?
          AND p.start_date <= ? AND (p.end_date IS NULL OR p.end_date >= ?)
        ORDER BY o.nama
        """,
        (work_date, location_id, work_date, work_date),
    ).fetchall()
    state = {}
    for row in rows:
        if row["planned_code"] == "OFF":
            label = "Off"
        elif row["shift_name"]:
            label = row["shift_name"]
        else:
            label = "Kosong"
        state[row["id"]] = (row["nama"], label)
    return state


def _log_assignment(db, location_id: int, work_date: date, before: dict, after: dict) -> None:
    changes = []
    for officer_id, (nama, new) in after.items():
        old = before.get(officer_id, (nama, "Kosong"))[1]
        if old != new:
            changes.append(f"{nama}: {old} → {new}")
    if not changes:
        return
    _add_location_log(
        db,
        location_id,
        "penempatan",
        f"Penempatan {day_label(work_date)}",
        "",
        "; ".join(changes),
    )


def save_shift_assignment(location_id: int, work_date: date, form) -> str:
    """Set which shift each officer must attend on this date."""
    with connect() as db:
        before = _assignment_state(db, location_id, work_date)
        allowed = {
            row["id"]
            for row in db.execute(
                "SELECT id FROM attendance_shifts WHERE location_id = ?",
                (location_id,),
            ).fetchall()
        }
        for key in form._data:
            if not key.startswith("tugas-"):
                continue
            officer_raw = key.split("-", 1)[1]
            if not officer_raw.isdigit():
                continue
            officer_id = int(officer_raw)
            placed = db.execute(
                """
                SELECT 1 FROM attendance_placements
                WHERE officer_id = ? AND location_id = ?
                  AND start_date <= ? AND (end_date IS NULL OR end_date >= ?)
                """,
                (officer_id, location_id, work_date, work_date),
            ).fetchone()
            if not placed:
                continue
            raw = form.get(key) or ""
            existing = db.execute(
                """
                SELECT id, clock_in, clock_out FROM attendance_slots
                WHERE officer_id = ? AND location_id = ? AND work_date = ? AND slot_no = 1
                """,
                (officer_id, location_id, work_date),
            ).fetchone()
            if raw == "off":
                if existing and (existing["clock_in"] or existing["clock_out"]):
                    continue
                if existing and _slot_locked(db, existing["id"]):
                    continue
                if existing:
                    db.execute(
                        "UPDATE attendance_slots SET shift_id = NULL, planned_code = 'OFF', off_reason = '' WHERE id = ?",
                        (existing["id"],),
                    )
                    _recalc(db, existing["id"])
                else:
                    created = db.execute(
                        """
                        INSERT INTO attendance_slots
                            (officer_id, location_id, work_date, slot_no, planned_code)
                        VALUES (?, ?, ?, 1, 'OFF')
                        RETURNING id
                        """,
                        (officer_id, location_id, work_date),
                    ).fetchone()
                    _recalc(db, created["id"])
                continue
            if raw == "":
                if not existing or existing["clock_in"] or existing["clock_out"]:
                    continue
                pending = db.execute(
                    "SELECT id FROM attendance_banding WHERE slot_id = ? AND status = 'pending'",
                    (existing["id"],),
                ).fetchone()
                if pending:
                    continue
                db.execute("DELETE FROM attendance_slots WHERE id = ?", (existing["id"],))
                continue
            if not raw.isdigit() or int(raw) not in allowed:
                return "Pilih shift yang ada di lokasi ini."
            shift_id = int(raw)
            if existing:
                db.execute(
                    "UPDATE attendance_slots SET shift_id = ? WHERE id = ?",
                    (shift_id, existing["id"]),
                )
                _recalc(db, existing["id"])
            else:
                created = db.execute(
                    """
                    INSERT INTO attendance_slots (officer_id, location_id, work_date, shift_id, slot_no)
                    VALUES (?, ?, ?, ?, 1)
                    RETURNING id
                    """,
                    (officer_id, location_id, work_date, shift_id),
                ).fetchone()
                _recalc(db, created["id"])
        _log_assignment(db, location_id, work_date, before, _assignment_state(db, location_id, work_date))
    return ""


def _slot_locked(db, slot_id: int) -> bool:
    row = db.execute(
        "SELECT clock_in, clock_out FROM attendance_slots WHERE id = ?",
        (slot_id,),
    ).fetchone()
    if not row or row["clock_in"] or row["clock_out"]:
        return True
    banding = db.execute(
        "SELECT id FROM attendance_banding WHERE slot_id = ?",
        (slot_id,),
    ).fetchone()
    return bool(banding)


def copy_shift_pattern(location_id: int, source: date, target: date) -> tuple[str, str]:
    """Copy who attends which shift onto the same weekday of the target month.

    A day that already has a clock-in, clock-out, or banding stays as saved.
    The copied day can still be changed one person at a time afterwards.
    """
    source_first = source.replace(day=1)
    target_first = target.replace(day=1)
    if (source_first.year, source_first.month) == (target_first.year, target_first.month):
        return "Pilih bulan lain sebagai sumber. Bulan yang sedang dibuka tetap bisa diubah per tanggal.", ""
    source_last = month_last(source_first)
    target_last = month_last(target_first)
    copied = 0
    kept = 0
    with connect() as db:
        source_rows = db.execute(
            """
            SELECT officer_id, work_date, shift_id, planned_code, off_reason
            FROM attendance_slots
            WHERE location_id = ? AND slot_no = 1
              AND work_date >= ? AND work_date <= ?
            """,
            (location_id, source_first, source_last),
        ).fetchall()
        if not source_rows:
            return "Bulan sumber belum punya penempatan shift.", ""
        pattern = {(row["officer_id"], row["work_date"]): row for row in source_rows}
        day = target_first
        while day <= target_last:
            origin = _same_weekday(day, source_first, source_last)
            placed = db.execute(
                """
                SELECT p.officer_id FROM attendance_placements p
                JOIN attendance_officers o ON o.id = p.officer_id
                WHERE p.location_id = ?
                  AND p.start_date <= ? AND (p.end_date IS NULL OR p.end_date >= ?)
                  AND o.attendance_enabled
                  AND o.employment_status IN ('aktif', 'ojt', 'lulus', 'mutasi', 'tunjangan', 'baru')
                """,
                (location_id, day, day),
            ).fetchall()
            for person in placed:
                officer_id = person["officer_id"]
                existing = db.execute(
                    """
                    SELECT id FROM attendance_slots
                    WHERE officer_id = ? AND location_id = ? AND work_date = ? AND slot_no = 1
                    """,
                    (officer_id, location_id, day),
                ).fetchone()
                if existing and _slot_locked(db, existing["id"]):
                    kept += 1
                    continue
                sample = pattern.get((officer_id, origin)) if origin else None
                if sample is None:
                    if existing:
                        db.execute("DELETE FROM attendance_slots WHERE id = ?", (existing["id"],))
                        copied += 1
                    continue
                planned = sample["planned_code"] if sample["planned_code"] in PLANNED else None
                reason = (sample["off_reason"] or "") if planned in ("LEAVE", "P", "S") else ""
                shift_id = sample["shift_id"]
                if existing:
                    db.execute(
                        """
                        UPDATE attendance_slots
                        SET shift_id = ?, planned_code = ?, off_reason = ?, clock_in = NULL, clock_out = NULL
                        WHERE id = ?
                        """,
                        (shift_id, planned, reason, existing["id"]),
                    )
                    _recalc(db, existing["id"])
                else:
                    created = db.execute(
                        """
                        INSERT INTO attendance_slots
                            (officer_id, location_id, work_date, shift_id, slot_no, planned_code, off_reason)
                        VALUES (?, ?, ?, ?, 1, ?, ?)
                        RETURNING id
                        """,
                        (officer_id, location_id, day, shift_id, planned, reason),
                    ).fetchone()
                    _recalc(db, created["id"])
                copied += 1
            day += timedelta(days=1)
        if copied:
            _add_location_log(
                db,
                location_id,
                "penempatan",
                "Pola shift",
                "",
                f"{month_label(source_first)} → {month_label(target_first)} · {copied} penempatan",
            )
    if copied == 0:
        return "", f"Penempatan {month_label(target_first)} tetap. Hari yang sudah ada jam masuk tidak diubah."
    return "", (
        f"Pola {month_label(source_first)} disalin ke {month_label(target_first)}. "
        f"{copied} penempatan mengikuti hari yang sama. "
        "Hari yang sudah ada jam masuk tetap, dan tiap tanggal masih bisa diubah manual."
    )


def file_banding(slot_id: int, requested: str, reason: str, detail: str) -> str:
    if requested not in REQUESTS:
        return "Pilih kode yang diminta."
    if not (reason or "").strip():
        return "Isi alasan banding."
    with connect() as db:
        slot = db.execute("SELECT id, code FROM attendance_slots WHERE id = ?", (slot_id,)).fetchone()
        if not slot:
            return "Slot absensi tidak ditemukan."
        pending = db.execute(
            "SELECT id FROM attendance_banding WHERE slot_id = ? AND status = 'pending'",
            (slot_id,),
        ).fetchone()
        if pending:
            return "Masih ada banding yang menunggu untuk slot ini."
        previous = db.execute(
            """
            SELECT id, sequence FROM attendance_banding
            WHERE slot_id = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (slot_id,),
        ).fetchone()
        db.execute(
            """
            INSERT INTO attendance_banding
                (slot_id, requested_code, reason, detail, before_code, parent_id, sequence)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                slot_id,
                requested,
                reason.strip(),
                (detail or "").strip(),
                slot["code"],
                previous["id"] if previous else None,
                (previous["sequence"] + 1) if previous else 1,
            ),
        )
    return ""


def decide_banding(banding_id: int, approved: bool, note: str) -> str:
    with connect() as db:
        row = db.execute(
            "SELECT id, slot_id, requested_code, status FROM attendance_banding WHERE id = ?",
            (banding_id,),
        ).fetchone()
        if not row or row["status"] != "pending":
            return "Banding ini sudah diputuskan."
        if approved:
            db.execute(
                """
                UPDATE attendance_banding
                SET status = 'approved', after_code = requested_code, decided_note = ?, decided_at = now()
                WHERE id = ?
                """,
                ((note or "").strip(), banding_id),
            )
            _recalc(db, row["slot_id"])
        else:
            db.execute(
                """
                UPDATE attendance_banding
                SET status = 'rejected', decided_note = ?, decided_at = now()
                WHERE id = ?
                """,
                ((note or "").strip(), banding_id),
            )
    return ""


def day_board(location_id: int, work_date: date) -> dict:
    with connect() as db:
        ensure_seeded(db)
        location = db.execute(
            "SELECT id, name, active, min_attend_minutes FROM attendance_locations WHERE id = ?",
            (location_id,),
        ).fetchone()
        shift_rows = db.execute(
            """
            SELECT id, name, start_time, end_time, grace_minutes, ends_next_day
            FROM attendance_shifts WHERE location_id = ? ORDER BY start_time, id
            """,
            (location_id,),
        ).fetchall()
        rows = db.execute(
            """
            SELECT o.id AS officer_id, o.nama, o.username, o.testing,
                   o.gender, o.marital_status,
                   o.min_attend_minutes, o.max_attend_minutes,
                   s.id AS slot_id, s.slot_no, s.planned_code, s.off_reason, s.clock_in, s.clock_out,
                   s.code, s.calculated_code, s.code_source, s.shift_id,
                   sh.name AS shift_name, sh.start_time, sh.end_time, sh.grace_minutes, sh.ends_next_day,
                   b.id AS pending_id
            FROM attendance_placements p
            JOIN attendance_officers o ON o.id = p.officer_id
            LEFT JOIN attendance_slots s
              ON s.officer_id = o.id AND s.location_id = p.location_id AND s.work_date = ?
            LEFT JOIN attendance_shifts sh ON sh.id = s.shift_id
            LEFT JOIN attendance_banding b ON b.slot_id = s.id AND b.status = 'pending'
            WHERE p.location_id = ?
              AND p.start_date <= ?
              AND (p.end_date IS NULL OR p.end_date >= ?)
              AND o.attendance_enabled
              AND o.employment_status IN ('aktif', 'ojt', 'lulus', 'mutasi', 'tunjangan', 'baru')
            ORDER BY o.nama, s.slot_no NULLS FIRST
            """,
            (work_date, location_id, work_date, work_date),
        ).fetchall()
    people = []
    counts = {code: 0 for code in LABELS}
    scheduled = 0
    location_minimum = location["min_attend_minutes"] if location else None
    for row in rows:
        if row["slot_id"]:
            scheduled += 1
            code = row["code"] or ""
            counts[code] = counts.get(code, 0) + 1
        else:
            code = ""
        reasons = reasons_for(row["gender"] or "", row["marital_status"] or "")
        stored_reason = row["off_reason"] or ""
        if stored_reason and stored_reason not in {item["code"] for item in reasons}:
            extra = reason_record(stored_reason)
            if extra:
                reasons = [extra, *reasons]
        if not stored_reason and row["planned_code"] == "LEAVE":
            reasons = [{"code": "", "label": "Cuti", "planned": "LEAVE", "legacy": True}, *reasons]
        people.append(
            {
                "officer_id": row["officer_id"],
                "nama": row["nama"],
                "username": row["username"],
                "testing": row["testing"],
                "slot_id": row["slot_id"],
                "slot_no": row["slot_no"] or 0,
                "planned": row["planned_code"] or "",
                "plan_value": plan_value(row["planned_code"], stored_reason),
                "reasons": reasons,
                "clock_in": clock_text(row["clock_in"]),
                "clock_out": clock_text(row["clock_out"]),
                "code": code,
                "code_label": label(code),
                "calculated": row["calculated_code"] or "",
                "calculated_label": label(row["calculated_code"] or ""),
                "source": row["code_source"] or "",
                "differs": bool(row["slot_id"] and row["code_source"] == "banding" and row["code"] != row["calculated_code"]),
                "shift_id": row["shift_id"],
                "span_text": timebox_text(row["start_time"], row["end_time"], bool(row["ends_next_day"])) if row["slot_id"] else "",
                "min_jam": hours_text(row["min_attend_minutes"] if row["min_attend_minutes"] is not None else location_minimum),
                "max_jam": hours_text(row["max_attend_minutes"]),
                "below_min": _below_minimum(row, location_minimum),
                "overtime": overtime_from_max(
                    row["clock_in"],
                    row["clock_out"],
                    row["max_attend_minutes"],
                    overtime_hours(row["end_time"], row["clock_out"], bool(row["ends_next_day"]), row["start_time"]) if row["slot_id"] else 0,
                ) if row["slot_id"] else 0,
                "pending": bool(row["pending_id"]),
            }
        )
    for row in shift_rows:
        row["start_text"] = clock_text(row["start_time"])
        row["end_text"] = clock_text(row["end_time"])
        row["span_text"] = timebox_text(row["start_time"], row["end_time"], bool(row["ends_next_day"]))
    return {
        "location": location,
        "date": work_date.isoformat(),
        "shifts": shift_rows,
        "people": people,
        "scheduled": scheduled,
        "unscheduled": sum(1 for person in people if not person["slot_id"]),
        "hadir": counts["OT"],
        "telat": counts["L"],
        "alpa": counts["A"],
        "off": counts["OFF"] + counts["LEAVE"] + counts["P"] + counts["S"],
    }


def banding_queue(location_id: int | None) -> list[dict]:
    with connect() as db:
        ensure_seeded(db)
        sql = """
            SELECT b.id, b.requested_code, b.reason, b.detail, b.before_code, b.sequence,
                   b.created_at, s.work_date, s.slot_no, o.nama, o.username, l.name AS lokasi, l.id AS location_id
            FROM attendance_banding b
            JOIN attendance_slots s ON s.id = b.slot_id
            JOIN attendance_officers o ON o.id = s.officer_id
            JOIN attendance_locations l ON l.id = s.location_id
            WHERE b.status = 'pending'
        """
        params: tuple = ()
        if location_id:
            sql += " AND s.location_id = ?"
            params = (location_id,)
        sql += " ORDER BY b.created_at"
        rows = db.execute(sql, params).fetchall()
    for row in rows:
        row["before_label"] = label(row["before_code"])
        row["requested_label"] = label(row["requested_code"])
        row["when"] = row["work_date"].isoformat() if hasattr(row["work_date"], "isoformat") else str(row["work_date"])
    return rows


def slot_detail(slot_id: int) -> dict | None:
    with connect() as db:
        row = db.execute(
            """
            SELECT s.id, s.code, s.work_date, s.slot_no, o.nama, o.username, l.name AS lokasi, l.id AS location_id
            FROM attendance_slots s
            JOIN attendance_officers o ON o.id = s.officer_id
            JOIN attendance_locations l ON l.id = s.location_id
            WHERE s.id = ?
            """,
            (slot_id,),
        ).fetchone()
    if not row:
        return None
    row["code_label"] = label(row["code"])
    row["when"] = row["work_date"].isoformat() if hasattr(row["work_date"], "isoformat") else str(row["work_date"])
    return row


def employment_label(code: str) -> str:
    return dict(EMPLOYMENT).get(code or "", code or "")


def _as_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def account_view(row) -> dict:
    """One account status, plus New Joiner when attendance started within 25 days."""
    status = row.get("employment_status") or ""
    approved = (row.get("onboard_state") or "") == "disetujui"
    enabled = bool(row.get("attendance_enabled"))
    if status == "blacklist":
        code = "blacklist"
    elif status == "resign":
        code = "resign"
    elif not approved:
        code = "onboard"
    elif not enabled or status in ("approved", "onboard"):
        code = "approved"
    elif status == "ojt":
        code = "ojt"
    elif status == "lulus":
        code = "lulus"
    else:
        code = "aktif"
    activated = _as_date(row.get("activated_on"))
    age = (date.today() - activated).days if activated else None
    new_joiner = code in ("aktif", "ojt", "lulus") and age is not None and 0 <= age <= NEW_JOINER_DAYS
    help_text = dict((item[0], item[2]) for item in STATUS_HELP).get(code, "")
    choice = code if code in dict(ACCOUNT_CHOICES) else "aktif"
    return {
        "code": code,
        "label": dict((item[0], item[1]) for item in STATUS_HELP).get(code, code),
        "help": help_text,
        "new_joiner": new_joiner,
        "choice": choice,
    }


def _text(value) -> str:
    if value is None:
        return ""
    return str(value)


def _money_text(value) -> str:
    number = int(round(float(value or 0)))
    return "Rp " + f"{number:,}".replace(",", ".")


def _add_log(db, officer_id: int, bagian: str, label: str, before, after) -> None:
    old = _text(before).strip()
    new = _text(after).strip()
    if old == new:
        return
    actor_id, actor_username, actor_nama = actor_snapshot()
    db.execute(
        """
        INSERT INTO attendance_change_log (
            officer_id, bagian, field_label, before_text, after_text,
            actor_user_id, actor_username, actor_nama
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (officer_id, bagian, label, old, new, actor_id, actor_username, actor_nama),
    )


def _write_changes(db, officer_id: int, bagian: str, before: list[tuple[str, str]], after: list[tuple[str, str]]) -> None:
    current = dict(after)
    for label, old in before:
        _add_log(db, officer_id, bagian, label, old, current.get(label, ""))


def _work_facts(db, officer_id: int) -> dict:
    row = db.execute(
        """
        SELECT o.nama, o.joined_date, o.posisi, o.gender, o.marital_status, o.testing,
               o.employment_status, o.onboard_state, o.attendance_enabled, o.activated_on,
               o.status_on,
               o.min_attend_minutes, o.max_attend_minutes, o.min_work_days, o.min_off_days,
               o.ojt_months, o.ojt_start, o.ojt_decision, o.resigned_date,
               l.name AS lokasi, s.name AS lokasi_kedua, p.start_date
        FROM attendance_officers o
        LEFT JOIN LATERAL (
            SELECT location_id, start_date
            FROM attendance_placements
            WHERE officer_id = o.id
            ORDER BY (end_date IS NULL) DESC, start_date DESC
            LIMIT 1
        ) p ON TRUE
        LEFT JOIN attendance_locations l ON l.id = COALESCE(p.location_id, o.pending_location_id)
        LEFT JOIN attendance_locations s ON s.id = o.secondary_location_id
        WHERE o.id = ?
        """,
        (officer_id,),
    ).fetchone()
    if not row:
        return {"status": "", "fields": []}
    account = account_view(row)
    status = account["label"]
    if account["new_joiner"]:
        status = f"{status} · New Joiner"
    def show_day(value):
        found = _as_date(value)
        return day_label(found) if found else ""
    ojt = ""
    if row["ojt_months"] or row["ojt_start"]:
        ojt = f"{row['ojt_months'] or ''} bulan dari {show_day(row['ojt_start'])}".strip()
        if row["ojt_decision"] == "gagal":
            ojt += " · tidak lulus"
        elif row["ojt_decision"] == "lulus":
            ojt += " · lulus"
    def jam(minutes):
        if minutes is None:
            return "Ikut lokasi"
        return f"{hours_text(minutes)} jam"
    def hari(value):
        return str(value) if value is not None else "Ikut lokasi"
    return {
        "status": status,
        "fields": [
            ("Nama", row["nama"] or ""),
            ("Tanggal masuk", show_day(row["joined_date"])),
            ("Posisi", position_label(row["posisi"] or "")),
            ("Jenis kelamin", dict(GENDERS).get(row["gender"] or "", "")),
            ("Status nikah", dict(MARITAL).get(row["marital_status"] or "", "")),
            ("Lokasi", row["lokasi"] or ""),
            ("Lokasi kedua", row["lokasi_kedua"] or ""),
            ("Mulai di lokasi", show_day(row["start_date"])),
            ("Tanggal resign", show_day(row["resigned_date"])),
            ("Tanggal status", show_day(row.get("status_on"))),
            ("Akun testing", "Ya" if row["testing"] else "Tidak"),
            ("Minimum kehadiran", jam(row["min_attend_minutes"])),
            ("Maksimum kehadiran", "" if row["max_attend_minutes"] is None else f"{hours_text(row['max_attend_minutes'])} jam"),
            ("Minimum hari hadir", hari(row["min_work_days"])),
            ("Minimum hari off", hari(row["min_off_days"])),
            ("Masa OJT", ojt),
        ],
    }


def _record_work(db, officer_id: int, before: dict) -> None:
    after = _work_facts(db, officer_id)
    _write_changes(db, officer_id, "kerja", before["fields"], after["fields"])
    _add_log(db, officer_id, "status", "Status", before["status"], after["status"])


def _benefit_amount(db, officer_id: int, profile_id, jenis: str) -> int:
    if profile_id:
        row = db.execute(
            """
            SELECT nilai FROM attendance_benefits
            WHERE pay_profile_id = ? AND struktur = 'ojt' AND jenis = ?
            ORDER BY id LIMIT 1
            """,
            (profile_id, jenis),
        ).fetchone()
    else:
        row = db.execute(
            """
            SELECT nilai FROM attendance_benefits
            WHERE officer_id = ? AND pay_profile_id IS NULL AND struktur = 'ojt' AND jenis = ?
            ORDER BY id LIMIT 1
            """,
            (officer_id, jenis),
        ).fetchone()
    return int(round(float(row["nilai"] or 0))) if row else 0


def _sync_user_profile(db, officer_id: int) -> None:
    """Keep the new-system profile row aligned with this officer."""
    officer = db.execute(
        """
        SELECT username, secondary_location_id, joined_date, resigned_date,
               pending_location_id, gaji_ojt
        FROM attendance_officers WHERE id = ?
        """,
        (officer_id,),
    ).fetchone()
    if not officer:
        return
    place = db.execute(
        """
        SELECT location_id FROM attendance_placements
        WHERE officer_id = ?
        ORDER BY (end_date IS NULL) DESC, start_date DESC
        LIMIT 1
        """,
        (officer_id,),
    ).fetchone()
    primary = place["location_id"] if place else officer["pending_location_id"]
    pay = db.execute(
        """
        SELECT id, gaji_ojt FROM attendance_pay_profiles
        WHERE officer_id = ? AND active
        ORDER BY id DESC LIMIT 1
        """,
        (officer_id,),
    ).fetchone()
    fee = int(pay["gaji_ojt"] or 0) if pay else int(officer["gaji_ojt"] or 0)
    secondary = officer["secondary_location_id"]
    if secondary == primary:
        secondary = None
    db.execute(
        """
        INSERT INTO attendance_user_profiles (
            officer_id, username, primary_location_id, secondary_location_id,
            basic_salary, bpjs_kesehatan, bpjs_ketenagakerjaan,
            joined_date, resigned_date, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, now())
        ON CONFLICT (officer_id) DO UPDATE SET
            username = EXCLUDED.username,
            primary_location_id = EXCLUDED.primary_location_id,
            secondary_location_id = EXCLUDED.secondary_location_id,
            basic_salary = EXCLUDED.basic_salary,
            bpjs_kesehatan = EXCLUDED.bpjs_kesehatan,
            bpjs_ketenagakerjaan = EXCLUDED.bpjs_ketenagakerjaan,
            joined_date = EXCLUDED.joined_date,
            resigned_date = EXCLUDED.resigned_date,
            updated_at = now()
        """,
        (
            officer_id,
            officer["username"],
            primary,
            secondary,
            fee,
            _benefit_amount(db, officer_id, pay["id"] if pay else None, "bpjs_kes"),
            _benefit_amount(db, officer_id, pay["id"] if pay else None, "bpjs_tk"),
            officer["joined_date"],
            officer["resigned_date"],
        ),
    )


def _pay_facts(db, officer_id: int, struktur: str) -> list[tuple[str, str]]:
    title = "Selama OJT" if struktur == "ojt" else "Setelah OJT"
    profile = db.execute(
        """
        SELECT p.id, p.gaji_ojt, p.jabatan_ojt, p.gaji_lanjut, p.jabatan_lanjut, l.name AS lokasi
        FROM attendance_pay_profiles p
        JOIN attendance_locations l ON l.id = p.location_id
        WHERE p.officer_id = ? AND p.active
        ORDER BY p.id DESC LIMIT 1
        """,
        (officer_id,),
    ).fetchone()
    if profile:
        fee = profile["gaji_ojt"] if struktur == "ojt" else profile["gaji_lanjut"]
        jabatan = profile["jabatan_ojt"] if struktur == "ojt" else profile["jabatan_lanjut"]
        location = profile["lokasi"]
        benefits = db.execute(
            """
            SELECT jenis, nilai, penanggung FROM attendance_benefits
            WHERE pay_profile_id = ? AND struktur = ?
            """,
            (profile["id"], struktur),
        ).fetchall()
    else:
        officer = db.execute(
            "SELECT gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut FROM attendance_officers WHERE id = ?",
            (officer_id,),
        ).fetchone()
        fee = (officer["gaji_ojt"] if struktur == "ojt" else officer["gaji_lanjut"]) if officer else 0
        jabatan = (officer["jabatan_ojt"] if struktur == "ojt" else officer["jabatan_lanjut"]) if officer else 0
        location = ""
        benefits = db.execute(
            """
            SELECT jenis, nilai, penanggung FROM attendance_benefits
            WHERE officer_id = ? AND pay_profile_id IS NULL AND struktur = ?
            """,
            (officer_id, struktur),
        ).fetchall()
    by_kind = {item["jenis"]: item for item in benefits}
    def component(code):
        item = by_kind.get(code)
        if not item:
            if code in ("bpjs_tk", "bpjs_kes"):
                return "Rp 0 · None"
            return "Rp 0"
        amount = _money_text(item["nilai"])
        if code in ("bpjs_tk", "bpjs_kes"):
            return f"{amount} · {dict(BPJS_MODES).get(bpjs_mode(item), 'None')}"
        return amount
    return [
        (f"Lokasi profil {title}", location),
        (f"Fee {title}", _money_text(fee)),
        (f"Tunjangan Jabatan {title}", _money_text(jabatan)),
        (f"Insentif {title}", component("insentif")),
        (f"BPJS Ketenagakerjaan {title}", component("bpjs_tk")),
        (f"BPJS Kesehatan {title}", component("bpjs_kes")),
        (f"THR {title}", component("thr")),
    ]


def _account_facts(db, officer_id: int) -> list[tuple[str, str]]:
    row = db.execute(
        """
        SELECT bank, rekening, pemilik_sendiri, pemilik_rekening
        FROM attendance_officers WHERE id = ?
        """,
        (officer_id,),
    ).fetchone()
    if not row:
        return []
    owner = "Sama dengan nama karyawan" if row["pemilik_sendiri"] else (row["pemilik_rekening"] or "")
    return [
        ("Bank", row["bank"] or ""),
        ("Nomor rekening", row["rekening"] or ""),
        ("Pemilik rekening", owner),
    ]


def profile_log(officer_id: int) -> list[dict]:
    labels = {
        "kerja": "Data kerja",
        "gaji": "Gaji",
        "rekening": "Rekening",
        "status": "Status",
        "payroll": "Payroll",
    }
    with connect() as db:
        rows = db.execute(
            """
            SELECT bagian, field_label, before_text, after_text, created_at,
                   actor_username, actor_nama
            FROM attendance_change_log
            WHERE officer_id = ?
            ORDER BY created_at DESC, id DESC
            """,
            (officer_id,),
        ).fetchall()
    for row in rows:
        row["bagian_label"] = labels.get(row["bagian"], row["bagian"])
        row["when"] = stamp(row["created_at"])
        row["actor_label"] = format_actor_label(row.get("actor_nama"), row.get("actor_username"))
        if row["before_text"]:
            row["change"] = f"{row['before_text']} → {row['after_text']}"
        else:
            row["change"] = row["after_text"]
    return rows


def _ojt_end(start, months) -> date | None:
    if not start or not months:
        return None
    if not isinstance(start, date):
        start = datetime.strptime(str(start)[:10], "%Y-%m-%d").date()
    return add_months(start, int(months))


def _ojt_waiting(status: str, start, months, decision: str) -> bool:
    end = _ojt_end(start, months)
    return status == "ojt" and not (decision or "").strip() and end is not None and date.today() >= end


def active_pay_structure(status: str) -> str:
    return "ojt" if status == "ojt" else "lanjut"


def _sync_active_pay(db, officer_id: int) -> None:
    row = db.execute(
        """
        SELECT employment_status, gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut
        FROM attendance_officers WHERE id = ?
        """,
        (officer_id,),
    ).fetchone()
    if not row:
        return
    if active_pay_structure(row["employment_status"]) == "ojt":
        pokok, jabatan = row["gaji_ojt"] or 0, row["jabatan_ojt"] or 0
    else:
        pokok, jabatan = row["gaji_lanjut"] or 0, row["jabatan_lanjut"] or 0
    db.execute(
        "UPDATE attendance_officers SET gaji_pokok = ?, tunjangan_jabatan = ? WHERE id = ?",
        (pokok, jabatan, officer_id),
    )


POSITIONS = (
    ("spv", "SPV"),
    ("leader", "Leader"),
    ("admin", "Admin"),
    ("gaskir", "Gaskir"),
)
PAY_COMPONENTS = (
    ("insentif", "Insentif", "fee"),
    ("bpjs_tk", "BPJS Ketenagakerjaan", "bpjs"),
    ("bpjs_kes", "BPJS Kesehatan", "bpjs"),
    ("thr", "THR / Tunjangan Hari Raya", "thr"),
)
BPJS_MODES = (
    ("deduction", "Deduction"),
    ("include", "Include"),
    ("none", "None"),
)
BENEFIT_KINDS = tuple((code, label) for code, label, _group in PAY_COMPONENTS)
BENEFIT_PAYERS = (
    ("perusahaan", "Ditanggung perusahaan"),
    ("petugas", "Ditanggung petugas"),
)


def benefit_amount(pokok: int, mode: str, nilai) -> int:
    """Nominal stays as rupiah. Persen is taken from gaji pokok only."""
    number = float(nilai or 0)
    if mode == "persen":
        return int(round((pokok or 0) * number / 100))
    return int(round(number))


def position_label(code: str) -> str:
    return dict(POSITIONS).get(code or "", "")


def benefit_title(item: dict) -> str:
    jenis = item.get("jenis") or ""
    return dict(BENEFIT_KINDS).get(jenis, item.get("nama") or "Benefit")


def benefit_effect(jenis: str, penanggung: str) -> str:
    """Company BPJS is a company contribution. Company allowances are paid to the officer. Officer-paid rows are deductions."""
    if penanggung == "petugas":
        return "potongan"
    if jenis in ("kesehatan", "ketenagakerjaan"):
        return "iuran"
    return "tambah"


def rate_input(mode: str, nilai) -> str:
    number = float(nilai or 0)
    if mode == "persen":
        return f"{number:.2f}".rstrip("0").rstrip(".")
    return str(int(round(number)))


def parse_rate(value: str | None, mode: str) -> tuple[float, str]:
    raw = (value or "").strip().replace(" ", "").replace("%", "")
    if mode == "persen":
        raw = raw.replace(",", ".")
        if raw == "":
            return 0.0, ""
        try:
            number = float(raw)
        except ValueError:
            return 0.0, "Persen harus angka."
        if number < 0 or number > 100:
            return 0.0, "Persen diisi dari 0 sampai 100."
        return round(number, 2), ""
    return float(parse_money(raw)), ""


def bpjs_mode(item) -> str:
    raw = (item or {}).get("penanggung") or "none"
    if raw == "overhead":
        return "deduction"
    if raw in dict(BPJS_MODES):
        return raw
    return "none"


def _component_rows(benefits: list[dict], jenis: str) -> list[dict]:
    names = ("insentif", "insentif_operasional", "insentif_operational") if jenis == "insentif" else (jenis,)
    return [item for item in benefits if item.get("jenis") in names and item.get("ikut", True)]


def _component_nilai(benefits: list[dict], jenis: str, pokok: int) -> int:
    return sum(benefit_amount(pokok, item.get("mode") or "nominal", item.get("nilai") or 0) for item in _component_rows(benefits, jenis))


def _bpjs_line(benefits: list[dict], jenis: str, nama: str, pokok: int) -> tuple[int, str]:
    rows = _component_rows(benefits, jenis)
    mode = bpjs_mode(rows[0] if rows else None)
    if mode == "none":
        return 0, "None"
    nilai = sum(benefit_amount(pokok, item.get("mode") or "nominal", item.get("nilai") or 0) for item in rows)
    return nilai, dict(BPJS_MODES)[mode]


def compose_pay(pokok: int, benefits: list[dict], jabatan: int = 0) -> dict:
    """Fee + Tunjangan Jabatan + Insentif. Include adds a BPJS amount. Deduction subtracts it. None leaves it out."""
    fee = int(pokok or 0)
    jabatan = int(jabatan or 0)
    insentif = _component_nilai(benefits, "insentif", fee)
    thr = _component_nilai(benefits, "thr", fee)
    bpjs_tk, tk_label = _bpjs_line(benefits, "bpjs_tk", "BPJS Ketenagakerjaan", fee)
    bpjs_kes, kes_label = _bpjs_line(benefits, "bpjs_kes", "BPJS Kesehatan", fee)
    take_home = fee + jabatan + insentif
    deducted = 0
    for nilai, label in ((bpjs_tk, tk_label), (bpjs_kes, kes_label)):
        if label == "Include":
            take_home += nilai
        elif label == "Deduction":
            take_home -= nilai
            deducted += nilai
    lines = [
        {"nama": "Fee", "hasil": fee, "kelompok": "Komponen Fee & Insentif", "catatan": ""},
        {"nama": "Tunjangan Jabatan", "hasil": jabatan, "kelompok": "Komponen Fee & Insentif", "catatan": "Masuk gaji dihitung (payroll)"},
        {"nama": "Insentif", "hasil": insentif, "kelompok": "Komponen Fee & Insentif", "catatan": ""},
        {"nama": "BPJS Ketenagakerjaan", "hasil": bpjs_tk, "kelompok": "BPJS & THR", "catatan": tk_label},
        {"nama": "BPJS Kesehatan", "hasil": bpjs_kes, "kelompok": "BPJS & THR", "catatan": kes_label},
        {"nama": "THR / Tunjangan Hari Raya", "hasil": thr, "kelompok": "BPJS & THR", "catatan": "THR (Prorata) 1x Fee"},
    ]
    return {
        "lines": lines,
        "take_home": take_home,
        "overhead": deducted,
        "gross": take_home + deducted,
        "total": take_home,
        "diterima": take_home,
        "iuran": deducted,
    }


def parse_money(value: str | None) -> int:
    digits = "".join(ch for ch in (value or "") if ch.isdigit())
    return int(digits) if digits else 0


def _day_text(value) -> str:
    if value is None or value == "":
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)[:10]


def location_record(location_id: int) -> dict | None:
    with connect() as db:
        ensure_seeded(db)
        row = db.execute(
            """
            SELECT l.id, l.name, l.active, l.grace_minutes, l.min_attend_minutes,
                   l.min_work_days, l.min_off_days, l.latitude, l.longitude, l.gps_radius_meters,
                (SELECT COUNT(*) FROM attendance_placements p
                 WHERE p.location_id = l.id AND (p.end_date IS NULL OR p.end_date >= CURRENT_DATE)) AS people,
                (SELECT COUNT(*) FROM attendance_shifts s WHERE s.location_id = l.id) AS shifts
            FROM attendance_locations l
            WHERE l.id = ?
            """,
            (location_id,),
        ).fetchone()
    if row:
        row["min_jam"] = hours_text(row["min_attend_minutes"])
        row["latitude_text"] = _coord_text(row["latitude"])
        row["longitude_text"] = _coord_text(row["longitude"])
    return row


def location_team(location_id: int) -> list[dict]:
    today = date.today()
    first = today.replace(day=1)
    last = month_last(first)
    with connect() as db:
        place = db.execute(
            "SELECT min_work_days, min_off_days FROM attendance_locations WHERE id = ?",
            (location_id,),
        ).fetchone()
        rows = db.execute(
            """
            SELECT o.id, o.nama, o.username, o.employment_status, o.testing,
                   o.onboard_state, o.attendance_enabled, o.activated_on,
                   o.min_work_days, o.min_off_days,
                   p.start_date, p.end_date
            FROM attendance_placements p
            JOIN attendance_officers o ON o.id = p.officer_id
            WHERE p.location_id = ?
              AND (p.end_date IS NULL OR p.end_date >= CURRENT_DATE)
            ORDER BY o.nama
            """,
            (location_id,),
        ).fetchall()
        counted = db.execute(
            """
            SELECT officer_id, work_date,
                   BOOL_OR(code IN ('OT', 'L')) AS hadir,
                   BOOL_OR(code IN ('OFF', 'LEAVE', 'P', 'S')) AS off
            FROM attendance_slots
            WHERE location_id = ? AND work_date >= ? AND work_date <= ?
            GROUP BY officer_id, work_date
            """,
            (location_id, first, last),
        ).fetchall()
    tally: dict[int, dict[str, int]] = {}
    for item in counted:
        bucket = tally.setdefault(item["officer_id"], {"hadir": 0, "off": 0})
        if item["hadir"]:
            bucket["hadir"] += 1
        elif item["off"]:
            bucket["off"] += 1
    location_work = place["min_work_days"] if place else None
    location_off = place["min_off_days"] if place else None
    for row in rows:
        row["mulai"] = _day_text(row["start_date"])
        row["sampai"] = _day_text(row["end_date"])
        account = account_view(row)
        row["status_label"] = account["label"]
        row["new_joiner"] = account["new_joiner"]
        stats = tally.get(row["id"], {"hadir": 0, "off": 0})
        row["hadir_bulan"] = stats["hadir"]
        row["off_bulan"] = stats["off"]
        row["bulan"] = month_label(first)
        own_work = row["min_work_days"]
        own_off = row["min_off_days"]
        row["work_target"] = own_work if own_work is not None else location_work
        row["off_target"] = own_off if own_off is not None else location_off
        row["short_work"] = row["work_target"] is not None and stats["hadir"] < row["work_target"]
        row["short_off"] = row["off_target"] is not None and stats["off"] < row["off_target"]
    return rows


def day_label(value: date) -> str:
    return f"{value.day} {MONTH_NAMES[value.month - 1]} {value.year}"


def month_label(value: date) -> str:
    return f"{MONTH_NAMES[value.month - 1]} {value.year}"


def _same_weekday(target: date, source_first: date, source_last: date) -> date | None:
    """Match Senin to Senin, using the same week of the month when that week exists."""
    delta = (target.weekday() - source_first.weekday()) % 7
    first = source_first + timedelta(days=delta)
    week = (target.day - 1) // 7
    candidate = first + timedelta(days=7 * week)
    if candidate > source_last:
        candidate = candidate - timedelta(days=7)
    if source_first <= candidate <= source_last and candidate.weekday() == target.weekday():
        return candidate
    return None


def _month_shift(anchor: date, delta: int) -> date:
    month_index = anchor.month - 1 + delta
    year = anchor.year + month_index // 12
    month = month_index % 12 + 1
    last = month_last(date(year, month, 1))
    return date(year, month, min(anchor.day, last.day))


def assignment_month(location_id: int, anchor: date) -> dict:
    """Month grid of who is already placed on a shift, counted per shift."""
    first = anchor.replace(day=1)
    weeks_raw = calendar.Calendar(firstweekday=0).monthdatescalendar(first.year, first.month)
    start, end = weeks_raw[0][0], weeks_raw[-1][-1]
    with connect() as db:
        shifts = db.execute(
            """
            SELECT id, name FROM attendance_shifts
            WHERE location_id = ? ORDER BY start_time, id
            """,
            (location_id,),
        ).fetchall()
        counts = db.execute(
            """
            SELECT work_date, shift_id, COUNT(*) AS n
            FROM attendance_slots
            WHERE location_id = ? AND work_date >= ? AND work_date <= ?
              AND slot_no = 1 AND shift_id IS NOT NULL
            GROUP BY work_date, shift_id
            """,
            (location_id, start, end),
        ).fetchall()
    tones = {row["id"]: index % 6 for index, row in enumerate(shifts)}
    names = {row["id"]: row["name"] for row in shifts}
    by_day: dict[date, list[dict]] = {}
    for row in counts:
        day = row["work_date"]
        if not isinstance(day, date):
            day = datetime.strptime(str(day)[:10], "%Y-%m-%d").date()
        by_day.setdefault(day, []).append(
            {
                "id": row["shift_id"],
                "name": names.get(row["shift_id"], "Shift"),
                "n": row["n"],
                "tone": tones.get(row["shift_id"], 0),
            }
        )
    today = date.today()
    weeks = []
    for week in weeks_raw:
        cells = []
        for day in week:
            chips = sorted(by_day.get(day, []), key=lambda item: item["tone"])
            cells.append(
                {
                    "iso": day.isoformat(),
                    "day": day.day,
                    "in_month": day.month == first.month,
                    "today": day == today,
                    "selected": day == anchor,
                    "people": sum(item["n"] for item in chips),
                    "chips": chips[:3],
                    "more": max(0, len(chips) - 3),
                }
            )
        weeks.append(cells)
    return {
        "label": f"{MONTH_NAMES[first.month - 1]} {first.year}",
        "prev": _month_shift(anchor, -1).isoformat(),
        "next": _month_shift(anchor, 1).isoformat(),
        "weeks": weeks,
        "weekdays": ("Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"),
        "legend": [{"id": row["id"], "name": row["name"], "tone": tones[row["id"]]} for row in shifts],
    }


def shift_groups(board: dict) -> tuple[list[dict], dict]:
    """People on the selected day, grouped by the shift they must attend."""
    groups = []
    by_id = {}
    for index, shift in enumerate(board["shifts"]):
        group = {
            "id": shift["id"],
            "name": shift["name"],
            "span_text": shift["span_text"],
            "tone": index % 6,
            "people": [],
        }
        groups.append(group)
        by_id[shift["id"]] = group
    loose = {
        "id": None,
        "name": "Belum ditugaskan",
        "span_text": "Belum masuk shift hari ini",
        "tone": None,
        "people": [],
    }
    off = {
        "id": None,
        "name": "Off",
        "span_text": "Tidak dijadwalkan masuk",
        "tone": None,
        "people": [],
    }
    for person in board["people"]:
        if person["slot_no"] > 1:
            continue
        if person["planned"] == "OFF" and not person["shift_id"]:
            off["people"].append(person)
            continue
        target = by_id.get(person["shift_id"])
        (target if target is not None else loose)["people"].append(person)
    shown = [group for group in groups if group["people"]]
    if off["people"]:
        shown.append(off)
    if loose["people"]:
        shown.append(loose)
    summary = {
        "people": sum(len(group["people"]) for group in (*groups, loose, off)),
        "open": len(loose["people"]),
        "off": len(off["people"]),
        "shifts_used": sum(1 for group in groups if group["people"]),
        "lines": [
            {"name": group["name"], "n": len(group["people"]), "tone": group["tone"]}
            for group in groups
            if group["people"]
        ],
    }
    return shown, summary


def location_filters(source) -> dict:
    def bound(key: str) -> str:
        raw = (source.get(key) or "").strip()
        return raw if raw.isdigit() else ""

    status = (source.get("status") or "").strip()
    if status not in ("aktif", "nonaktif"):
        status = ""
    minimum = (source.get("minimum") or "").strip()
    if minimum not in ("ada", "belum"):
        minimum = ""
    return {
        "q": " ".join((source.get("q") or "").split()),
        "status": status,
        "petugas_dari": bound("petugas_dari"),
        "petugas_sampai": bound("petugas_sampai"),
        "shift_dari": bound("shift_dari"),
        "shift_sampai": bound("shift_sampai"),
        "tenggang_dari": bound("tenggang_dari"),
        "tenggang_sampai": bound("tenggang_sampai"),
        "minimum": minimum,
    }


def _within(number, low: str, high: str) -> bool:
    if low and number < int(low):
        return False
    if high and number > int(high):
        return False
    return True


def location_directory(filters: dict | None = None) -> list[dict]:
    with connect() as db:
        ensure_seeded(db)
        rows = db.execute(
            """
            SELECT l.id, l.name, l.active, l.grace_minutes, l.min_attend_minutes,
                   l.min_work_days, l.min_off_days,
                (SELECT COUNT(*) FROM attendance_placements p
                 WHERE p.location_id = l.id AND (p.end_date IS NULL OR p.end_date >= CURRENT_DATE)) AS people,
                (SELECT COUNT(*) FROM attendance_shifts s WHERE s.location_id = l.id) AS shifts
            FROM attendance_locations l
            ORDER BY l.name
            """
        ).fetchall()
    for row in rows:
        row["min_jam"] = hours_text(row["min_attend_minutes"])
    if not filters:
        return rows
    query = filters["q"].lower()
    kept = []
    for row in rows:
        if query and query not in (row["name"] or "").lower():
            continue
        if filters["status"] == "aktif" and not row["active"]:
            continue
        if filters["status"] == "nonaktif" and row["active"]:
            continue
        if not _within(int(row["people"] or 0), filters["petugas_dari"], filters["petugas_sampai"]):
            continue
        if not _within(int(row["shifts"] or 0), filters["shift_dari"], filters["shift_sampai"]):
            continue
        if not _within(int(row["grace_minutes"] or 0), filters["tenggang_dari"], filters["tenggang_sampai"]):
            continue
        if filters["minimum"] == "ada" and row["min_attend_minutes"] is None:
            continue
        if filters["minimum"] == "belum" and row["min_attend_minutes"] is not None:
            continue
        kept.append(row)
    return kept


def _coord_text(value) -> str:
    if value is None or value == "":
        return ""
    text = f"{float(value):.7f}".rstrip("0").rstrip(".")
    return text


def _parse_coordinate(value: str, low: float, high: float, label: str) -> tuple[float | None, str]:
    raw = (value or "").strip().replace(",", ".")
    if raw == "":
        return None, ""
    try:
        number = float(raw)
    except ValueError:
        return None, f"{label} harus angka, misalnya -6.200000."
    if number < low or number > high:
        return None, f"{label} di luar jangkauan."
    return round(number, 7), ""


def _parse_point(latitude: str, longitude: str) -> tuple[float | None, float | None, str]:
    lat_raw = (latitude or "").strip()
    lng_raw = (longitude or "").strip()
    if lng_raw == "":
        found = re.findall(r"[+-]?\d+(?:[.,]\d+)?", lat_raw)
        if len(found) == 2:
            lat_raw, lng_raw = found
    lat, lat_error = _parse_coordinate(lat_raw, -90, 90, "Lintang")
    if lat_error:
        return None, None, lat_error
    lng, lng_error = _parse_coordinate(lng_raw, -180, 180, "Bujur")
    if lng_error:
        return None, None, lng_error
    if (lat is None) != (lng is None):
        return None, None, "Isi lintang dan bujur bersamaan. Boleh menempel koordinat dari peta ke kolom lintang."
    return lat, lng, ""


def _parse_radius(value: str, has_point: bool) -> tuple[int | None, str]:
    raw = (value or "").strip()
    if raw == "":
        if has_point:
            return None, "Pilih jarak absensi: 20, 50, 100, 300, atau 500 meter."
        return None, ""
    if not has_point:
        return None, "Isi koordinat dulu, lalu pilih jarak absensi."
    if not raw.isdigit() or int(raw) not in GPS_RADII:
        return None, "Jarak absensi hanya 20, 50, 100, 300, atau 500 meter."
    return int(raw), ""


def _location_grace(value: str) -> tuple[int | None, str]:
    raw = (value or "").strip()
    if raw == "":
        return None, "Isi tenggang terlambat untuk lokasi ini."
    try:
        return max(0, int(raw)), ""
    except ValueError:
        return None, "Tenggang terlambat lokasi harus angka."


def _add_location_log(db, location_id: int, bagian: str, label: str, before, after) -> None:
    old = _text(before).strip()
    new = _text(after).strip()
    if old == new:
        return
    actor_id, actor_username, actor_nama = actor_snapshot()
    db.execute(
        """
        INSERT INTO attendance_location_log (
            location_id, bagian, field_label, before_text, after_text,
            actor_user_id, actor_username, actor_nama
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (location_id, bagian, label, old, new, actor_id, actor_username, actor_nama),
    )


def _log_location_fields(db, location_id: int, bagian: str, before: list[tuple[str, str]], after: list[tuple[str, str]]) -> None:
    current = dict(after)
    for label, old in before:
        _add_location_log(db, location_id, bagian, label, old, current.get(label, ""))


def _location_facts(db, location_id: int) -> list[tuple[str, str]]:
    row = db.execute(
        """
        SELECT name, active, grace_minutes, min_attend_minutes, min_work_days, min_off_days,
               latitude, longitude, gps_radius_meters
        FROM attendance_locations WHERE id = ?
        """,
        (location_id,),
    ).fetchone()
    if not row:
        return []
    return [
        ("Nama", row["name"] or ""),
        ("Status", "Aktif" if row["active"] else "Nonaktif"),
        ("Tenggang terlambat", f"{row['grace_minutes']} menit" if row["grace_minutes"] is not None else "Belum diatur"),
        ("Minimum kehadiran", f"{hours_text(row['min_attend_minutes'])} jam" if row["min_attend_minutes"] is not None else "Belum diatur"),
        ("Minimum hari hadir", f"{row['min_work_days']} hari" if row["min_work_days"] is not None else "Belum diatur"),
        ("Minimum hari off", f"{row['min_off_days']} hari" if row["min_off_days"] is not None else "Belum diatur"),
        ("Lintang", _coord_text(row["latitude"]) or "Belum diatur"),
        ("Bujur", _coord_text(row["longitude"]) or "Belum diatur"),
        ("Jarak GPS", f"{row['gps_radius_meters']} m" if row["gps_radius_meters"] else "Belum diatur"),
    ]


def _shift_story(name, start, end, grace, ends_next_day) -> str:
    span = timebox_text(start, end, bool(ends_next_day))
    extra = "tenggang ikut lokasi" if grace is None else f"tenggang {grace} menit"
    return f"{name} · {span} · {extra}"


def location_log(location_id: int) -> list[dict]:
    labels = {"profil": "Profil", "shift": "Shift", "penempatan": "Penempatan"}
    with connect() as db:
        rows = db.execute(
            """
            SELECT bagian, field_label, before_text, after_text, created_at,
                   actor_username, actor_nama
            FROM attendance_location_log
            WHERE location_id = ?
            ORDER BY created_at DESC, id DESC
            """,
            (location_id,),
        ).fetchall()
    for row in rows:
        row["bagian_label"] = labels.get(row["bagian"], row["bagian"])
        row["when"] = stamp(row["created_at"])
        row["actor_label"] = format_actor_label(row.get("actor_nama"), row.get("actor_username"))
        if row["before_text"]:
            row["change"] = f"{row['before_text']} → {row['after_text']}"
        else:
            row["change"] = row["after_text"]
    return rows


def save_location(
    location_id: int,
    name: str,
    active: bool,
    grace: str = "15",
    minimum: str = "",
    work_days: str = "",
    off_days: str = "",
    latitude: str | None = None,
    longitude: str | None = None,
    radius: str | None = None,
) -> str:
    clean = " ".join((name or "").split())
    if not clean:
        return "Nama lokasi kosong."
    minutes, error = _location_grace(grace)
    if error:
        return error
    minimum_minutes, min_error = parse_hours(minimum)
    if min_error:
        return min_error
    work_day_count, work_error = parse_days(work_days)
    if work_error:
        return work_error
    off_day_count, off_error = parse_days(off_days)
    if off_error:
        return off_error
    point = None
    if latitude is not None or longitude is not None or radius is not None:
        lat, lng, point_error = _parse_point(latitude or "", longitude or "")
        if point_error:
            return point_error
        meters, radius_error = _parse_radius(radius or "", lat is not None)
        if radius_error:
            return radius_error
        point = (lat, lng, meters)
    with connect() as db:
        found = db.execute(
            "SELECT id FROM attendance_locations WHERE lower(name) = lower(?) AND id <> ?",
            (clean, location_id),
        ).fetchone()
        if found:
            return "Nama lokasi itu sudah dipakai."
        before = _location_facts(db, location_id)
        if point is None:
            db.execute(
                """
                UPDATE attendance_locations
                SET name = ?, active = ?, payable = ?, grace_minutes = ?, min_attend_minutes = ?,
                    min_work_days = ?, min_off_days = ?
                WHERE id = ?
                """,
                (clean, active, active, minutes, minimum_minutes, work_day_count, off_day_count, location_id),
            )
        else:
            db.execute(
                """
                UPDATE attendance_locations
                SET name = ?, active = ?, payable = ?, grace_minutes = ?, min_attend_minutes = ?,
                    min_work_days = ?, min_off_days = ?,
                    latitude = ?, longitude = ?, gps_radius_meters = ?
                WHERE id = ?
                """,
                (
                    clean, active, active, minutes, minimum_minutes, work_day_count, off_day_count,
                    point[0], point[1], point[2], location_id,
                ),
            )
        slots = db.execute(
            """
            SELECT s.id
            FROM attendance_slots s
            JOIN attendance_shifts sh ON sh.id = s.shift_id
            WHERE s.location_id = ? AND sh.grace_minutes IS NULL
            """,
            (location_id,),
        ).fetchall()
        for slot in slots:
            _recalc(db, slot["id"])
        _log_location_fields(db, location_id, "profil", before, _location_facts(db, location_id))
    return ""


def delete_location(location_id: int) -> str:
    with connect() as db:
        used = db.execute(
            """
            SELECT
                (SELECT COUNT(*) FROM attendance_placements WHERE location_id = ?) AS people,
                (SELECT COUNT(*) FROM attendance_slots WHERE location_id = ?) AS slots
            """,
            (location_id, location_id),
        ).fetchone()
        if used and (used["people"] or used["slots"]):
            return "Lokasi masih punya petugas atau absensi. Nonaktifkan dari ubah data."
        db.execute("DELETE FROM attendance_locations WHERE id = ?", (location_id,))
    return ""


def save_shift(shift_id: int, name: str, start: str, end: str, grace: str, hari: str = "ini") -> str:
    values, error = _shift_values(name, start, end, grace, hari)
    if error:
        return error
    clean, start_time, end_time, grace_minutes, ends_next_day = values
    with connect() as db:
        current = db.execute(
            """
            SELECT location_id, name, start_time, end_time, grace_minutes, ends_next_day
            FROM attendance_shifts WHERE id = ?
            """,
            (shift_id,),
        ).fetchone()
        if not current:
            return "Shift tidak ditemukan."
        db.execute(
            """
            UPDATE attendance_shifts
            SET name = ?, start_time = ?, end_time = ?, grace_minutes = ?, ends_next_day = ?
            WHERE id = ?
            """,
            (clean, start_time, end_time, grace_minutes, ends_next_day, shift_id),
        )
        _add_location_log(
            db,
            current["location_id"],
            "shift",
            "Shift",
            _shift_story(current["name"], current["start_time"], current["end_time"], current["grace_minutes"], current["ends_next_day"]),
            _shift_story(clean, start_time, end_time, grace_minutes, ends_next_day),
        )
        slots = db.execute(
            "SELECT id FROM attendance_slots WHERE shift_id = ?",
            (shift_id,),
        ).fetchall()
        for slot in slots:
            _recalc(db, slot["id"])
    return ""


def delete_shift(shift_id: int) -> str:
    with connect() as db:
        used = db.execute(
            "SELECT COUNT(*) AS n FROM attendance_slots WHERE shift_id = ?",
            (shift_id,),
        ).fetchone()
        planned = db.execute(
            "SELECT COUNT(*) AS n FROM attendance_shift_months WHERE shift_id = ?",
            (shift_id,),
        ).fetchone()
        if (used and used["n"]) or (planned and planned["n"]):
            return "Shift ini sudah dipakai petugas. Kosongkan penempatan bulannya dulu, atau ubah jamnya."
        current = db.execute(
            """
            SELECT location_id, name, start_time, end_time, grace_minutes, ends_next_day
            FROM attendance_shifts WHERE id = ?
            """,
            (shift_id,),
        ).fetchone()
        db.execute("DELETE FROM attendance_shifts WHERE id = ?", (shift_id,))
        if current:
            _add_location_log(
                db,
                current["location_id"],
                "shift",
                "Shift",
                _shift_story(current["name"], current["start_time"], current["end_time"], current["grace_minutes"], current["ends_next_day"]),
                "Dihapus",
            )
    return ""


def _employee_query(db, where: str, params: tuple) -> list[dict]:
    rows = db.execute(
        f"""
        SELECT o.id, o.nama, o.username, o.joined_date, o.resigned_date, o.gaji_pokok,
               o.employment_status, o.onboard_state, o.attendance_enabled, o.activated_on, o.testing,
               o.ojt_months, o.ojt_start, o.ojt_decision, o.posisi,
               l.name AS lokasi, l.id AS location_id,
               s.name AS lokasi_kedua,
               COALESCE(u.basic_salary, 0) AS basic_salary,
               COALESCE(u.bpjs_kesehatan, 0) AS bpjs_kesehatan,
               COALESCE(u.bpjs_ketenagakerjaan, 0) AS bpjs_ketenagakerjaan
        FROM attendance_officers o
        LEFT JOIN LATERAL (
            SELECT location_id FROM attendance_placements
            WHERE officer_id = o.id
            ORDER BY (end_date IS NULL) DESC, start_date DESC
            LIMIT 1
        ) p ON TRUE
        LEFT JOIN attendance_locations l ON l.id = COALESCE(p.location_id, o.pending_location_id)
        LEFT JOIN attendance_user_profiles u ON u.officer_id = o.id
        LEFT JOIN attendance_locations s ON s.id = COALESCE(u.secondary_location_id, o.secondary_location_id)
        {where}
        ORDER BY o.nama
        """,
        params,
    ).fetchall()
    for row in rows:
        joined = _as_date(row["joined_date"])
        resigned = _as_date(row["resigned_date"])
        row["joined"] = day_label(joined) if joined else ""
        row["resigned"] = day_label(resigned) if resigned else ""
        account = account_view(row)
        row["status_label"] = account["label"]
        row["new_joiner"] = account["new_joiner"]
        row["account_code"] = account["code"]
        row["posisi_label"] = position_label(row.get("posisi") or "")
        row["ojt_menunggu"] = _ojt_waiting(row["employment_status"], row["ojt_start"], row["ojt_months"], row["ojt_decision"])
        row["masa_hari"] = None
        row["masa_hari_label"] = ""
        row["masa_tahun_label"] = ""
        if account["code"] in WORKING_CODES and joined:
            end = resigned if account["code"] == "resign" and resigned else date.today()
            days, years = _tenure(joined, end)
            row["masa_hari"] = days
            row["masa_hari_label"], row["masa_tahun_label"] = _tenure_labels(days, years)
    return rows


OFFICER_SORTS = ("nama", "lokasi", "kedua", "posisi", "status", "masuk", "resign", "masa", "gaji", "bpjs_kes", "bpjs_tk")
WORKING_CODES = ("aktif", "ojt", "lulus", "resign")


def _tenure(start: date, end: date) -> tuple[int, int]:
    """Whole days from start to end, and how many full years fit in that span."""
    if end < start:
        return 0, 0
    days = (end - start).days
    years = 0
    cursor = start
    while True:
        try:
            nxt = date(cursor.year + 1, cursor.month, cursor.day)
        except ValueError:
            nxt = date(cursor.year + 1, cursor.month, 28)
        if nxt > end:
            break
        years += 1
        cursor = nxt
    return days, years


def _tenure_labels(days: int, years: int) -> tuple[str, str]:
    day_text = f"{days:,}".replace(",", ".") + " hari"
    if years <= 0:
        year_text = "kurang dari 1 tahun"
    elif years == 1:
        year_text = "1 tahun"
    else:
        year_text = f"{years} tahun"
    return day_text, year_text


def _filter_day(value: str) -> str:
    raw = (value or "").strip()[:10]
    if not raw:
        return ""
    try:
        return date.fromisoformat(raw).isoformat()
    except ValueError:
        return ""


def _filter_money(value: str) -> str:
    raw = (value or "").strip()
    if not any(ch.isdigit() for ch in raw):
        return ""
    return str(parse_money(raw))


def officer_filters(source) -> dict:
    """Keep only the filter and sort values the petugas grid understands."""
    posisi = (source.get("posisi") or "").strip()
    if posisi not in {code for code, _name in POSITIONS} | {"kosong"}:
        posisi = ""
    lokasi = (source.get("lokasi") or "").strip()
    if lokasi and not lokasi.isdigit():
        lokasi = ""
    kedua = (source.get("lokasi_kedua") or "").strip()
    if kedua not in ("", "kosong") and not kedua.isdigit():
        kedua = ""
    resign = (source.get("resign") or "").strip()
    if resign not in ("ada", "belum"):
        resign = ""
    sort = (source.get("sort") or "nama").strip()
    if sort not in OFFICER_SORTS:
        sort = "nama"
    arah = (source.get("arah") or "naik").strip()
    if arah not in ("naik", "turun"):
        arah = "naik"
    return {
        "lokasi": lokasi,
        "lokasi_kedua": kedua,
        "posisi": posisi,
        "masuk_dari": _filter_day(source.get("masuk_dari") or ""),
        "masuk_sampai": _filter_day(source.get("masuk_sampai") or ""),
        "resign": resign,
        "resign_dari": _filter_day(source.get("resign_dari") or ""),
        "resign_sampai": _filter_day(source.get("resign_sampai") or ""),
        "gaji_dari": _filter_money(source.get("gaji_dari") or ""),
        "gaji_sampai": _filter_money(source.get("gaji_sampai") or ""),
        "bpjs_kes_dari": _filter_money(source.get("bpjs_kes_dari") or ""),
        "bpjs_kes_sampai": _filter_money(source.get("bpjs_kes_sampai") or ""),
        "bpjs_tk_dari": _filter_money(source.get("bpjs_tk_dari") or ""),
        "bpjs_tk_sampai": _filter_money(source.get("bpjs_tk_sampai") or ""),
        "sort": sort,
        "arah": arah,
    }


def _sort_directory(rows: list[dict], sort: str, arah: str) -> list[dict]:
    desc = arah == "turun"

    def parts(row: dict):
        if sort == "lokasi":
            value = (row.get("lokasi") or "").casefold()
            return (not value, value)
        if sort == "kedua":
            value = (row.get("lokasi_kedua") or "").casefold()
            return (not value, value)
        if sort == "posisi":
            value = (row.get("posisi_label") or "").casefold()
            return (not value, value)
        if sort == "status":
            return (False, (row.get("status_label") or "").casefold())
        if sort == "masuk":
            day = row.get("joined_date")
            return (day is None, day or date.min)
        if sort == "resign":
            day = row.get("resigned_date")
            return (day is None, day or date.min)
        if sort == "masa":
            hari = row.get("masa_hari")
            return (hari is None, hari if hari is not None else 0)
        if sort == "gaji":
            return (False, int(round(float(row.get("basic_salary") or 0))))
        if sort == "bpjs_kes":
            return (False, int(round(float(row.get("bpjs_kesehatan") or 0))))
        if sort == "bpjs_tk":
            return (False, int(round(float(row.get("bpjs_ketenagakerjaan") or 0))))
        value = (row.get("nama") or "").casefold()
        return (not value, value)

    blanks_last = sort in ("lokasi", "kedua", "posisi", "masuk", "resign", "masa")
    if blanks_last:
        filled = [row for row in rows if not parts(row)[0]]
        blank = [row for row in rows if parts(row)[0]]
        filled.sort(key=lambda row: parts(row)[1], reverse=desc)
        return filled + blank
    ordered = list(rows)
    ordered.sort(key=lambda row: parts(row)[1], reverse=desc)
    return ordered


def employee_directory(query: str = "", status: str = "", tahap: str = "", filters: dict | None = None, sort: str = "", arah: str = "") -> list[dict]:
    clauses = []
    params: list = []
    chosen = filters or {}
    if query.strip():
        like = f"%{query.strip()}%"
        if query.strip().isdigit():
            clauses.append("(o.nama ILIKE ? OR o.username ILIKE ? OR o.id = ?)")
            params.extend([like, like, int(query.strip())])
        else:
            clauses.append("(o.nama ILIKE ? OR o.username ILIKE ?)")
            params.extend([like, like])
    if status == "onboard":
        clauses.append("o.onboard_state <> 'disetujui' AND o.employment_status NOT IN ('blacklist', 'resign')")
    elif status == "approved":
        clauses.append(
            "o.onboard_state = 'disetujui' AND NOT o.attendance_enabled AND o.employment_status NOT IN ('resign', 'blacklist')"
        )
    elif status == "new":
        clauses.append(
            """
            o.attendance_enabled
            AND o.activated_on IS NOT NULL
            AND o.activated_on >= CURRENT_DATE - INTERVAL '25 days'
            AND o.employment_status IN ('aktif', 'ojt', 'lulus', 'mutasi', 'tunjangan', 'baru')
            """
        )
    elif status == "aktif":
        clauses.append("o.attendance_enabled AND o.employment_status IN ('aktif', 'mutasi', 'tunjangan', 'baru')")
    elif status in ("ojt", "lulus", "resign", "blacklist"):
        clauses.append("o.employment_status = ?")
        params.append(status)
    if tahap in ("pengguna", "menunggu", "disetujui", "ditolak"):
        clauses.append("o.onboard_state = ?")
        params.append(tahap)
    if chosen.get("lokasi"):
        clauses.append("l.id = ?")
        params.append(int(chosen["lokasi"]))
    if chosen.get("lokasi_kedua") == "kosong":
        clauses.append("COALESCE(u.secondary_location_id, o.secondary_location_id) IS NULL")
    elif chosen.get("lokasi_kedua"):
        clauses.append("COALESCE(u.secondary_location_id, o.secondary_location_id) = ?")
        params.append(int(chosen["lokasi_kedua"]))
    if chosen.get("posisi") == "kosong":
        clauses.append("(o.posisi IS NULL OR btrim(o.posisi) = '')")
    elif chosen.get("posisi"):
        clauses.append("o.posisi = ?")
        params.append(chosen["posisi"])
    if chosen.get("masuk_dari"):
        clauses.append("o.joined_date >= ?")
        params.append(chosen["masuk_dari"])
    if chosen.get("masuk_sampai"):
        clauses.append("o.joined_date <= ?")
        params.append(chosen["masuk_sampai"])
    if chosen.get("resign") == "ada":
        clauses.append("o.resigned_date IS NOT NULL")
    elif chosen.get("resign") == "belum":
        clauses.append("o.resigned_date IS NULL")
    if chosen.get("resign_dari"):
        clauses.append("o.resigned_date >= ?")
        params.append(chosen["resign_dari"])
    if chosen.get("resign_sampai"):
        clauses.append("o.resigned_date <= ?")
        params.append(chosen["resign_sampai"])
    for column, start_key, end_key in (
        ("COALESCE(u.basic_salary, 0)", "gaji_dari", "gaji_sampai"),
        ("COALESCE(u.bpjs_kesehatan, 0)", "bpjs_kes_dari", "bpjs_kes_sampai"),
        ("COALESCE(u.bpjs_ketenagakerjaan, 0)", "bpjs_tk_dari", "bpjs_tk_sampai"),
    ):
        if chosen.get(start_key):
            clauses.append(f"{column} >= ?")
            params.append(int(chosen[start_key]))
        if chosen.get(end_key):
            clauses.append(f"{column} <= ?")
            params.append(int(chosen[end_key]))
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with connect() as db:
        ensure_seeded(db)
        rows = _employee_query(db, where, tuple(params))
    if sort:
        return _sort_directory(rows, sort if sort in OFFICER_SORTS else "nama", arah if arah in ("naik", "turun") else "naik")
    return rows


def officer_status_index() -> dict[str, dict]:
    index = {}
    for row in employee_directory():
        key = (row.get("username") or "").strip().lower()
        if not key:
            continue
        index[key] = {
            "id": row["id"],
            "label": row["status_label"],
            "code": row["account_code"],
            "testing": bool(row.get("testing")),
        }
    return index


def payroll_blocked(row) -> list[str]:
    """Reasons this officer stays out of the payroll run."""
    reasons = []
    code = row.get("account_code") or ""
    if code == "blacklist":
        reasons.append("Blacklist")
    elif code == "resign":
        reasons.append("Resign")
    if row.get("testing"):
        reasons.append("Akun testing")
    return reasons


def _payroll_money(raw: str, label: str):
    text = (raw or "").strip()
    if text == "":
        return None, ""
    try:
        value = int(float(text.replace(".", "").replace(",", ".")))
    except ValueError:
        return None, f"{label} harus angka."
    if value < 0:
        return None, f"{label} tidak boleh negatif."
    if value == 0:
        return None, ""
    return value, ""


def _payroll_reason_label(kind: str, code: str) -> str:
    choices = PAYROLL_POTONGAN_REASONS if kind == "potongan" else PAYROLL_TAMBAHAN_REASONS
    return dict(choices).get(code or "", code or "")


def _empty_payroll_condition() -> dict:
    return {
        "deposit": None,
        "potongan": None,
        "potongan_reason": "",
        "tambahan": None,
        "tambahan_reason": "",
        "excluded": False,
        "note": "",
    }


def payroll_conditions_index() -> dict[str, dict]:
    """Saved payroll conditions by username. Used again on the next payroll run."""
    with connect() as db:
        rows = db.execute(
            """
            SELECT username, officer_id, deposit, potongan, potongan_reason,
                   tambahan, tambahan_reason, excluded, note
            FROM attendance_payroll_conditions
            """
        ).fetchall()
    index = {}
    for row in rows:
        key = (row.get("username") or "").strip().lower()
        if not key:
            continue
        potongan = row.get("potongan")
        potongan_reason = (row.get("potongan_reason") or "").strip()
        # Legacy rows that only stored deposit become a Deposit potongan.
        if potongan is None and row.get("deposit") is not None:
            potongan = row.get("deposit")
            potongan_reason = potongan_reason or "deposit"
        index[key] = {
            "username": key,
            "officer_id": row.get("officer_id"),
            "deposit": row.get("deposit"),
            "potongan": potongan,
            "potongan_reason": potongan_reason,
            "tambahan": row.get("tambahan"),
            "tambahan_reason": (row.get("tambahan_reason") or "").strip(),
            "excluded": bool(row.get("excluded")),
            "note": row.get("note") or "",
        }
    return index


def payroll_condition_for(username: str) -> dict:
    key = (username or "").strip().lower()
    if not key:
        return _empty_payroll_condition()
    return payroll_conditions_index().get(key) or _empty_payroll_condition()


def officer_payroll_info(officer_id: int, source_deposit: int | float | None = None) -> dict:
    """Deposit already cut from payroll plus saved pengurangan/penambahan for profile/detail drawers."""
    person = employee_profile(officer_id)
    empty = {
        "source_deposit": 0,
        "effective_deposit": 0,
        "has_deposit": False,
        "deposit_overridden": False,
        "has_adjustments": False,
        "condition": _empty_payroll_condition(),
        "potongan_label": "",
        "tambahan_label": "",
        "note": "",
        "lines": [],
        "log": [],
    }
    if person is None:
        return empty
    condition = payroll_condition_for(person.get("username") or "")
    file_deposit = int(round(float(source_deposit or 0)))
    if file_deposit < 0:
        file_deposit = abs(file_deposit)
    effective = file_deposit
    deposit_overridden = False
    if condition.get("potongan") is not None and (condition.get("potongan_reason") or "") == "deposit":
        effective = int(condition["potongan"])
        deposit_overridden = True
    potongan_label = ""
    if condition.get("potongan") is not None:
        potongan_label = (
            f"{_payroll_reason_label('potongan', condition.get('potongan_reason') or '')} · "
            f"Rp {int(condition['potongan']):,}".replace(",", ".")
        )
    tambahan_label = ""
    if condition.get("tambahan") is not None:
        tambahan_label = (
            f"{_payroll_reason_label('tambahan', condition.get('tambahan_reason') or '')} · "
            f"Rp {int(condition['tambahan']):,}".replace(",", ".")
        )

    def money(amount: int) -> str:
        return "Rp " + f"{abs(int(amount)):,}".replace(",", ".")

    lines = []
    if effective > 0:
        lines.append(
            {
                "kind": "potongan",
                "kind_label": "Pengurangan",
                "title": "Deposit",
                "amount": effective,
                "amount_label": "− " + money(effective),
                "source": (
                    "Override syarat payroll (mengganti deposit file)"
                    if deposit_overridden
                    else "Kolom Deposit di file payroll — sudah dipotong dari gaji dihitung"
                ),
                "cls": "is-cut",
            }
        )
    elif file_deposit > 0 and deposit_overridden and effective == 0:
        lines.append(
            {
                "kind": "potongan",
                "kind_label": "Pengurangan",
                "title": "Deposit",
                "amount": 0,
                "amount_label": "− " + money(0),
                "source": f"Override syarat menghapus/mengganti deposit file (semula {money(file_deposit)})",
                "cls": "is-empty",
            }
        )
    if condition.get("potongan") is not None and (condition.get("potongan_reason") or "") != "deposit":
        reason = _payroll_reason_label("potongan", condition.get("potongan_reason") or "")
        lines.append(
            {
                "kind": "potongan",
                "kind_label": "Pengurangan",
                "title": reason or "Pengurangan manual",
                "amount": int(condition["potongan"]),
                "amount_label": "− " + money(int(condition["potongan"])),
                "source": "Disimpan dari tombol Pengurangan di daftar payroll (kolom Potongan Manual)",
                "cls": "is-cut",
            }
        )
    if condition.get("tambahan") is not None:
        reason = _payroll_reason_label("tambahan", condition.get("tambahan_reason") or "")
        lines.append(
            {
                "kind": "tambahan",
                "kind_label": "Penambahan",
                "title": reason or "Penambahan manual",
                "amount": int(condition["tambahan"]),
                "amount_label": "+ " + money(int(condition["tambahan"])),
                "source": "Disimpan dari tombol Penambahan di daftar payroll (kolom Tambahan Manual)",
                "cls": "is-add",
            }
        )
    return {
        "source_deposit": file_deposit,
        "effective_deposit": effective,
        "has_deposit": effective > 0,
        "deposit_overridden": deposit_overridden,
        "has_adjustments": bool(lines) or bool(condition.get("note")),
        "condition": condition,
        "potongan_label": potongan_label,
        "tambahan_label": tambahan_label,
        "note": condition.get("note") or "",
        "lines": lines,
        "log": payroll_adjustment_log(officer_id),
    }


def payroll_adjustment_log(officer_id: int) -> list[dict]:
    """Trackable history of potongan/tambahan saved for this officer."""
    with connect() as db:
        rows = db.execute(
            """
            SELECT kind, reason, amount, note, created_at, actor_username, actor_nama
            FROM attendance_payroll_adjustment_log
            WHERE officer_id = ?
            ORDER BY created_at DESC, id DESC
            LIMIT 40
            """,
            (officer_id,),
        ).fetchall()
    found = []
    for row in rows:
        kind = row.get("kind") or ""
        reason = row.get("reason") or ""
        amount = int(row.get("amount") or 0)
        sign = "-" if kind == "potongan" else "+"
        found.append(
            {
                "kind": kind,
                "kind_label": "Potongan" if kind == "potongan" else "Tambahan",
                "reason": reason,
                "reason_label": _payroll_reason_label(kind, reason),
                "amount": amount,
                "amount_label": f"{sign} Rp " + f"{amount:,}".replace(",", "."),
                "note": row.get("note") or "",
                "when": stamp(row.get("created_at")),
                "actor_label": format_actor_label(row.get("actor_nama"), row.get("actor_username")),
            }
        )
    return found


def _append_payroll_adjustment_log(db, officer_id: int, username: str, kind: str, reason: str, amount: int, note: str) -> None:
    actor_id, actor_username, actor_nama = actor_snapshot()
    db.execute(
        """
        INSERT INTO attendance_payroll_adjustment_log
            (officer_id, username, kind, reason, amount, note, created_at,
             actor_user_id, actor_username, actor_nama)
        VALUES (?, ?, ?, ?, ?, ?, now(), ?, ?, ?)
        """,
        (
            officer_id,
            username,
            kind,
            reason or "",
            int(amount or 0),
            note or "",
            actor_id,
            actor_username,
            actor_nama,
        ),
    )


def save_payroll_condition(
    officer_id: int,
    note: str | None = None,
    excluded: bool | None = None,
    potongan_raw: str | None = None,
    potongan_reason: str | None = None,
    tambahan_raw: str | None = None,
    tambahan_reason: str | None = None,
    deposit_raw: str | None = None,
    section: str = "all",
) -> str:
    """Store potongan/tambahan/catatan. section: potongan|tambahan|catatan|all."""
    person = employee_profile(officer_id)
    if person is None:
        return "Petugas tidak ditemukan."
    username = (person.get("username") or "").strip()
    if not username:
        return "Username kosong."
    section = (section or "all").strip().lower()
    if section == "syarat":
        section = "catatan"
    current = payroll_conditions_index().get(username.lower()) or _empty_payroll_condition()
    return _save_payroll_condition_impl(
        officer_id=officer_id,
        person=person,
        username=username,
        current=current,
        section=section,
        note=note,
        excluded=excluded,
        potongan_raw=potongan_raw,
        potongan_reason=potongan_reason,
        tambahan_raw=tambahan_raw,
        tambahan_reason=tambahan_reason,
        deposit_raw=deposit_raw,
    )


def _save_payroll_condition_impl(
    officer_id: int,
    person: dict,
    username: str,
    current: dict,
    section: str,
    note: str | None,
    excluded: bool | None,
    potongan_raw: str | None,
    potongan_reason: str | None,
    tambahan_raw: str | None,
    tambahan_reason: str | None,
    deposit_raw: str | None,
) -> str:
    touch_potongan = section in ("all", "potongan")
    touch_tambahan = section in ("all", "tambahan")
    touch_note = section in ("all", "catatan", "potongan", "tambahan")

    potongan = current.get("potongan")
    potongan_code = current.get("potongan_reason") or ""
    tambahan = current.get("tambahan")
    tambahan_code = current.get("tambahan_reason") or ""
    clean_note = current.get("note") or ""
    deposit = current.get("deposit")

    if touch_potongan and potongan_raw is not None:
        potongan, err = _payroll_money(potongan_raw, "Pengurangan")
        if err:
            return err
        potongan_code = (potongan_reason or "").strip()
        if potongan is not None and potongan_code not in PAYROLL_POTONGAN_CODES:
            return "Pilih alasan pengurangan."
        if potongan is None:
            potongan_code = ""

    if touch_tambahan and tambahan_raw is not None:
        tambahan, err = _payroll_money(tambahan_raw, "Penambahan")
        if err:
            return err
        tambahan_code = (tambahan_reason or "").strip()
        if tambahan is not None and tambahan_code not in PAYROLL_TAMBAHAN_CODES:
            return "Pilih alasan penambahan."
        if tambahan is None:
            tambahan_code = ""

    if touch_note and note is not None:
        clean_note = " ".join((note or "").split())

    # Legacy deposit-only callers.
    if deposit_raw is not None and touch_potongan and potongan_raw is None:
        deposit, err = _payroll_money(deposit_raw, "Deposit")
        if err:
            return err
        if deposit is not None:
            potongan = deposit
            potongan_code = "deposit"
    if potongan is not None and potongan_code == "deposit":
        deposit = potongan
    elif touch_potongan:
        deposit = potongan if potongan_code == "deposit" else None

    keep_excluded = bool(current.get("excluded")) if excluded is None else bool(excluded)
    if (person.get("account_code") or "") in ("resign", "blacklist") or person.get("testing"):
        keep_excluded = False

    before_potongan = current.get("potongan")
    before_potongan_reason = current.get("potongan_reason") or ""
    before_tambahan = current.get("tambahan")
    before_tambahan_reason = current.get("tambahan_reason") or ""
    before_note = current.get("note") or ""

    with connect() as db:
        db.execute(
            """
            INSERT INTO attendance_payroll_conditions (
                username, officer_id, deposit, potongan, potongan_reason,
                tambahan, tambahan_reason, excluded, note, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, now())
            ON CONFLICT (username) DO UPDATE SET
                officer_id = EXCLUDED.officer_id,
                deposit = EXCLUDED.deposit,
                potongan = EXCLUDED.potongan,
                potongan_reason = EXCLUDED.potongan_reason,
                tambahan = EXCLUDED.tambahan,
                tambahan_reason = EXCLUDED.tambahan_reason,
                excluded = EXCLUDED.excluded,
                note = EXCLUDED.note,
                updated_at = now()
            """,
            (
                username,
                officer_id,
                deposit,
                potongan,
                potongan_code,
                tambahan,
                tambahan_code,
                keep_excluded,
                clean_note,
            ),
        )
        if (
            deposit is None
            and potongan is None
            and tambahan is None
            and not keep_excluded
            and not clean_note
        ):
            db.execute("DELETE FROM attendance_payroll_conditions WHERE lower(username) = lower(?)", (username,))

        potongan_changed = (before_potongan != potongan) or (before_potongan_reason != potongan_code)
        tambahan_changed = (before_tambahan != tambahan) or (before_tambahan_reason != tambahan_code)
        log_note = clean_note
        if potongan_changed and potongan is not None:
            _append_payroll_adjustment_log(db, officer_id, username, "potongan", potongan_code, potongan, log_note)
        if tambahan_changed and tambahan is not None:
            _append_payroll_adjustment_log(db, officer_id, username, "tambahan", tambahan_code, tambahan, log_note)
        if potongan_changed and potongan is None and before_potongan is not None:
            _append_payroll_adjustment_log(
                db, officer_id, username, "potongan", before_potongan_reason or "lainnya", 0, "Pengurangan dihapus"
            )
        if tambahan_changed and tambahan is None and before_tambahan is not None:
            _append_payroll_adjustment_log(
                db, officer_id, username, "tambahan", before_tambahan_reason or "lainnya", 0, "Penambahan dihapus"
            )

        def money_text(value, reason_code, kind):
            if value is None:
                return "Tidak ada"
            label = _payroll_reason_label(kind, reason_code)
            return f"{label} · Rp " + f"{int(value):,}".replace(",", ".")

        if touch_potongan:
            potongan_label = "Deposit" if (potongan_code == "deposit" or before_potongan_reason == "deposit") else "Pengurangan payroll"
            _add_log(
                db,
                officer_id,
                "payroll",
                potongan_label,
                money_text(before_potongan, before_potongan_reason, "potongan"),
                money_text(potongan, potongan_code, "potongan"),
            )
        if touch_tambahan:
            _add_log(
                db,
                officer_id,
                "payroll",
                "Penambahan payroll",
                money_text(before_tambahan, before_tambahan_reason, "tambahan"),
                money_text(tambahan, tambahan_code, "tambahan"),
            )
        if touch_note:
            _add_log(db, officer_id, "payroll", "Catatan payroll", before_note, clean_note)
    return ""


def save_payroll_syarat(
    officer_id: int,
    note: str = "",
    potongan_raw: str | None = None,
    potongan_reason: str | None = None,
    tambahan_raw: str | None = None,
    tambahan_reason: str | None = None,
    section: str = "all",
) -> str:
    """Save one payroll section without touching unrelated fields."""
    return save_payroll_condition(
        officer_id,
        note=note,
        excluded=False,
        potongan_raw=potongan_raw,
        potongan_reason=potongan_reason,
        tambahan_raw=tambahan_raw,
        tambahan_reason=tambahan_reason,
        section=section,
    )


def payroll_blocked_usernames() -> set[str]:
    blocked = set()
    for row in employee_directory():
        if not payroll_blocked(row):
            continue
        key = (row.get("username") or "").strip().lower()
        if key:
            blocked.add(key)
    for key, row in payroll_conditions_index().items():
        if row.get("excluded"):
            blocked.add(key)
    return blocked


def payroll_exclusions() -> list[dict]:
    """Officers who stay out of the payroll run: blacklist, testing, resign, or saved exclude."""
    found = []
    conditions = payroll_conditions_index()
    seen = set()
    for row in employee_directory():
        reasons = payroll_blocked(row)
        key = (row.get("username") or "").strip().lower()
        cond = conditions.get(key) or {}
        if cond.get("excluded") and "Tidak dihitung (disimpan)" not in reasons:
            reasons.append("Tidak dihitung (disimpan)")
        if not reasons:
            continue
        seen.add(key)
        found.append(
            {
                "id": row["id"],
                "nama": row.get("nama") or "",
                "username": row.get("username") or "",
                "lokasi": row.get("lokasi") or "",
                "alasan": " · ".join(reasons),
                "status_label": row.get("status_label") or "",
            }
        )
    for key, cond in conditions.items():
        if not cond.get("excluded") or key in seen:
            continue
        found.append(
            {
                "id": cond.get("officer_id") or 0,
                "nama": key,
                "username": key,
                "lokasi": "",
                "alasan": "Tidak dihitung (disimpan)",
                "status_label": "",
            }
        )
    found.sort(key=lambda item: ((item["nama"] or "").casefold(), (item["username"] or "").casefold()))
    return found


def apply_payroll_conditions(payroll) -> None:
    """Overlay saved potongan/tambahan onto the current payroll object."""
    conditions = payroll_conditions_index()
    if not conditions:
        return
    for officer in payroll.officers:
        key = (officer.username or "").strip().lower()
        cond = conditions.get(key)
        if not cond:
            continue
        potongan = cond.get("potongan")
        reason = cond.get("potongan_reason") or ""
        if potongan is not None:
            if reason == "deposit":
                officer.amounts["deposit"] = float(potongan)
            else:
                officer.amounts["manual_potongan"] = float(potongan)
        elif cond.get("deposit") is not None:
            officer.amounts["deposit"] = float(cond["deposit"])
        tambahan = cond.get("tambahan")
        if tambahan is not None:
            officer.amounts["manual_tambahan"] = float(tambahan)


def employee_row(officer_id: int) -> dict | None:
    with connect() as db:
        rows = _employee_query(db, "WHERE o.id = ?", (officer_id,))
    return rows[0] if rows else None


def create_user(nama: str, username: str) -> tuple[str, int]:
    clean_name = " ".join((nama or "").split())
    clean_user = (username or "").strip()
    if not clean_name or not clean_user:
        return "Isi nama dan username.", 0
    with connect() as db:
        ensure_seeded(db)
        found = db.execute(
            "SELECT id FROM attendance_officers WHERE lower(username) = lower(?)",
            (clean_user,),
        ).fetchone()
        if found:
            return "Username itu sudah ada.", 0
        created = db.execute(
            """
            INSERT INTO attendance_officers (username, nama, employment_status, onboard_state, attendance_enabled)
            VALUES (?, ?, 'onboard', 'pengguna', FALSE)
            RETURNING id
            """,
            (clean_user, clean_name),
        ).fetchone()
        _add_log(db, created["id"], "status", "Status", "", "Onboard")
        _sync_user_profile(db, created["id"])
    return "", created["id"]


def register_employee(officer_id: int, location_id: int, joined: str, gaji: str, status: str) -> str:
    if not location_id:
        return "Pilih lokasi."
    if not joined:
        return "Isi tanggal masuk."
    start = parse_day(joined).isoformat()
    with connect() as db:
        officer = db.execute(
            "SELECT id, onboard_state FROM attendance_officers WHERE id = ?",
            (officer_id,),
        ).fetchone()
        if not officer:
            return "Karyawan tidak ditemukan."
        if officer["onboard_state"] not in ("pengguna", "ditolak"):
            return "Karyawan ini sudah didaftarkan."
        db.execute(
            """
            UPDATE attendance_officers
            SET joined_date = ?, gaji_pokok = ?, employment_status = 'approved',
                onboard_state = 'menunggu', attendance_enabled = FALSE, pending_location_id = ?
            WHERE id = ?
            """,
            (start, parse_money(gaji), location_id, officer_id),
        )
        place = _location_name(db, location_id)
        when = _as_date(start)
        when_text = day_label(when) if when else start
        _add_log(db, officer_id, "kerja", "Pendaftaran", "", f"Dikirim untuk persetujuan · {place} · {when_text}")
        _sync_user_profile(db, officer_id)
    return ""


def decide_onboarding(officer_id: int, approved: bool) -> str:
    with connect() as db:
        officer = db.execute(
            """
            SELECT id, joined_date, employment_status, onboard_state, pending_location_id
            FROM attendance_officers WHERE id = ?
            """,
            (officer_id,),
        ).fetchone()
        if not officer or officer["onboard_state"] != "menunggu":
            return "Pengajuan ini sudah diputuskan."
        if not approved:
            db.execute(
                "UPDATE attendance_officers SET onboard_state = 'ditolak' WHERE id = ?",
                (officer_id,),
            )
            _add_log(db, officer_id, "status", "Status", "Onboard", "Onboard · pendaftaran ditolak")
            _sync_user_profile(db, officer_id)
            return ""
        if not officer["pending_location_id"] or not officer["joined_date"]:
            return "Lokasi atau tanggal masuk belum lengkap."
        end = officer["joined_date"] if officer["employment_status"] == "resign" else None
        db.execute(
            """
            INSERT INTO attendance_placements (officer_id, location_id, start_date, end_date)
            VALUES (?, ?, ?, ?)
            ON CONFLICT (officer_id, location_id, start_date) DO UPDATE SET end_date = EXCLUDED.end_date
            """,
            (officer_id, officer["pending_location_id"], officer["joined_date"], end),
        )
        db.execute(
            """
            UPDATE attendance_officers
            SET onboard_state = 'disetujui', employment_status = 'approved',
                attendance_enabled = FALSE, pending_location_id = NULL,
                resigned_date = CASE WHEN employment_status = 'resign' THEN joined_date ELSE resigned_date END
            WHERE id = ?
            """,
            (officer_id,),
        )
        if officer["employment_status"] != "resign":
            _ensure_pay_profile(db, officer_id, officer["pending_location_id"])
        _add_log(db, officer_id, "status", "Status", "Onboard", "Approved")
        _sync_user_profile(db, officer_id)
    return ""


def employee_profile(officer_id: int) -> dict | None:
    with connect() as db:
        ensure_seeded(db)
        row = db.execute(
            """
            SELECT o.id, o.nama, o.username, o.joined_date, o.resigned_date, o.gaji_pokok,
                   o.tunjangan_jabatan, o.bank, o.rekening, o.pemilik_rekening, o.pemilik_sendiri,
                   o.employment_status, o.onboard_state, o.attendance_enabled, o.activated_on, o.status_on,
                   o.testing, o.pending_location_id, o.secondary_location_id,
                   o.min_attend_minutes, o.max_attend_minutes,
                   o.min_work_days, o.min_off_days, o.gender, o.marital_status,
                   o.ojt_months, o.ojt_start, o.ojt_decision, o.posisi,
                   o.gaji_ojt, o.jabatan_ojt, o.gaji_lanjut, o.jabatan_lanjut,
                   p.id AS placement_id, p.location_id, p.start_date, p.end_date, l.name AS lokasi,
                   s.name AS lokasi_kedua,
                   l.min_attend_minutes AS location_min,
                   l.min_work_days AS location_work_days,
                   l.min_off_days AS location_off_days,
                   u.basic_salary, u.bpjs_kesehatan, u.bpjs_ketenagakerjaan
            FROM attendance_officers o
            LEFT JOIN LATERAL (
                SELECT id, location_id, start_date, end_date
                FROM attendance_placements
                WHERE officer_id = o.id
                ORDER BY (end_date IS NULL) DESC, start_date DESC
                LIMIT 1
            ) p ON TRUE
            LEFT JOIN attendance_locations l ON l.id = COALESCE(p.location_id, o.pending_location_id)
            LEFT JOIN attendance_locations s ON s.id = o.secondary_location_id
            LEFT JOIN attendance_user_profiles u ON u.officer_id = o.id
            WHERE o.id = ?
            """,
            (officer_id,),
        ).fetchone()
    if not row:
        return None
    row["min_jam"] = hours_text(row["min_attend_minutes"])
    row["max_jam"] = hours_text(row["max_attend_minutes"])
    row["location_min_jam"] = hours_text(row["location_min"])
    row["min_follows"] = row["min_attend_minutes"] is None
    row["work_follows"] = row["min_work_days"] is None
    row["off_follows"] = row["min_off_days"] is None
    row["genders"] = GENDERS
    row["maritals"] = MARITAL
    row["joined"] = _day_text(row["joined_date"])
    row["resigned"] = _day_text(row["resigned_date"])
    row["placement_start"] = _day_text(row["start_date"])
    row["placement_end"] = _day_text(row["end_date"])
    account = account_view(row)
    row["status_label"] = account["label"]
    row["account_code"] = account["code"]
    row["account_help"] = account["help"]
    row["new_joiner"] = account["new_joiner"]
    row["status_choice"] = account["choice"]
    row["account_choices"] = ACCOUNT_CHOICES
    row["posisi_label"] = position_label(row.get("posisi") or "")
    row["positions"] = POSITIONS
    row["location_id"] = row["location_id"] or row["pending_location_id"]
    profiles = _pay_profile_rows(officer_id, row["employment_status"])
    active_profile = next((item for item in profiles if item["active"]), None)
    if active_profile:
        row["gaji_ojt"] = active_profile["gaji_ojt"]
        row["jabatan_ojt"] = active_profile["jabatan_ojt"]
        row["gaji_lanjut"] = active_profile["gaji_lanjut"]
        row["jabatan_lanjut"] = active_profile["jabatan_lanjut"]
    row["pay_profiles"] = profiles
    row["pay_location"] = active_profile["lokasi"] if active_profile else (row["lokasi"] or "")
    benefits = _benefits(officer_id, active_profile["id"] if active_profile else None)
    row["benefits"] = benefits
    row["struktur_aktif"] = active_pay_structure(row["employment_status"])
    row["ojt_mulai"] = _day_text(row["ojt_start"])
    end = _ojt_end(row["ojt_start"], row["ojt_months"])
    row["ojt_selesai"] = end.isoformat() if end else ""
    row["ojt_selesai_label"] = day_label(end) if end else ""
    row["ojt_menunggu"] = _ojt_waiting(row["employment_status"], row["ojt_start"], row["ojt_months"], row["ojt_decision"])
    row["benefit_kinds"] = BENEFIT_KINDS
    row["benefit_payers"] = BENEFIT_PAYERS
    row["structures"] = [
        _structure_view("ojt", "Selama OJT", row["gaji_ojt"] or 0, row["jabatan_ojt"] or 0, benefits, row["struktur_aktif"] == "ojt"),
        _structure_view("lanjut", "Setelah OJT", row["gaji_lanjut"] or 0, row["jabatan_lanjut"] or 0, benefits, row["struktur_aktif"] == "lanjut"),
    ]
    active = next(item for item in row["structures"] if item["active"])
    row["pay"] = active["pay"]
    row["punya_rekening"] = bool((row.get("rekening") or "").strip())
    if row["pemilik_sendiri"] or not (row.get("pemilik_rekening") or "").strip():
        row["pemilik_tampil"] = row["nama"]
    else:
        row["pemilik_tampil"] = row["pemilik_rekening"]
    row["beda_nama"] = row["punya_rekening"] and row["pemilik_tampil"].strip().lower() != (row["nama"] or "").strip().lower()
    return row


def _structure_view(code: str, title: str, pokok: int, jabatan: int, benefits: list[dict], active: bool) -> dict:
    own = [item for item in benefits if (item.get("struktur") or "lanjut") == code]
    return {
        "code": code,
        "title": title,
        "active": active,
        "gaji": pokok,
        "jabatan": jabatan,
        "benefits": own,
        "inputs": {
            code: rate_input("nominal", _component_nilai(own, code, pokok))
            for code, _label, _group in PAY_COMPONENTS
        },
        "modes": {
            "bpjs_tk": bpjs_mode(next((item for item in own if item.get("jenis") == "bpjs_tk"), None)),
            "bpjs_kes": bpjs_mode(next((item for item in own if item.get("jenis") == "bpjs_kes"), None)),
        },
        "pay": compose_pay(pokok, own, jabatan),
    }


def _pay_profile_rows(officer_id: int, status: str) -> list[dict]:
    with connect() as db:
        rows = db.execute(
            """
            SELECT p.id, p.location_id, p.active, p.gaji_ojt, p.jabatan_ojt, p.gaji_lanjut, p.jabatan_lanjut,
                   l.name AS lokasi
            FROM attendance_pay_profiles p
            JOIN attendance_locations l ON l.id = p.location_id
            WHERE p.officer_id = ?
            ORDER BY p.active DESC, l.name
            """,
            (officer_id,),
        ).fetchall()
    structure = active_pay_structure(status)
    for item in rows:
        item["fee"] = item["gaji_ojt"] if structure == "ojt" else item["gaji_lanjut"]
    return rows


def _benefits(officer_id: int, profile_id: int | None = None) -> list[dict]:
    with connect() as db:
        if profile_id:
            rows = db.execute(
                """
                SELECT id, nama, ikut, mode, nilai, struktur, jenis, penanggung
                FROM attendance_benefits
                WHERE pay_profile_id = ?
                ORDER BY id
                """,
                (profile_id,),
            ).fetchall()
        else:
            rows = db.execute(
                """
                SELECT id, nama, ikut, mode, nilai, struktur, jenis, penanggung
                FROM attendance_benefits
                WHERE officer_id = ? AND pay_profile_id IS NULL
                ORDER BY id
                """,
                (officer_id,),
            ).fetchall()
    for item in rows:
        item["jenis"] = item.get("jenis") or "lain"
        item["penanggung"] = item.get("penanggung") or "perusahaan"
        item["nilai_tampil"] = rate_input(item.get("mode") or "nominal", item.get("nilai") or 0)
    return rows


def _location_name(db, location_id: int) -> str:
    row = db.execute("SELECT name FROM attendance_locations WHERE id = ?", (location_id,)).fetchone()
    return row["name"] if row else ""


def _ensure_pay_profile(db, officer_id: int, location_id: int) -> dict:
    found = db.execute(
        """
        SELECT id, active FROM attendance_pay_profiles
        WHERE officer_id = ? AND location_id = ?
        """,
        (officer_id, location_id),
    ).fetchone()
    if found:
        return found
    officer = db.execute(
        """
        SELECT gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut
        FROM attendance_officers WHERE id = ?
        """,
        (officer_id,),
    ).fetchone()
    others = db.execute(
        "SELECT id FROM attendance_pay_profiles WHERE officer_id = ? AND active LIMIT 1",
        (officer_id,),
    ).fetchone()
    created = db.execute(
        """
        INSERT INTO attendance_pay_profiles
            (officer_id, location_id, active, gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        RETURNING id, active
        """,
        (
            officer_id,
            location_id,
            others is None,
            officer["gaji_ojt"] or 0,
            officer["jabatan_ojt"] or 0,
            officer["gaji_lanjut"] or 0,
            officer["jabatan_lanjut"] or 0,
        ),
    ).fetchone()
    db.execute(
        """
        UPDATE attendance_benefits
        SET pay_profile_id = ?
        WHERE officer_id = ? AND pay_profile_id IS NULL
        """,
        (created["id"], officer_id),
    )
    return created


def _apply_pay_profile(db, officer_id: int, profile_id: int) -> None:
    db.execute("UPDATE attendance_pay_profiles SET active = FALSE WHERE officer_id = ?", (officer_id,))
    db.execute("UPDATE attendance_pay_profiles SET active = TRUE WHERE id = ?", (profile_id,))
    row = db.execute(
        """
        SELECT gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut
        FROM attendance_pay_profiles WHERE id = ?
        """,
        (profile_id,),
    ).fetchone()
    db.execute(
        """
        UPDATE attendance_officers
        SET gaji_ojt = ?, jabatan_ojt = ?, gaji_lanjut = ?, jabatan_lanjut = ?
        WHERE id = ?
        """,
        (row["gaji_ojt"], row["jabatan_ojt"], row["gaji_lanjut"], row["jabatan_lanjut"], officer_id),
    )
    _sync_active_pay(db, officer_id)


def _replicate_pay_profile(db, officer_id: int, source_location: int, target_location: int) -> dict:
    source = _ensure_pay_profile(db, officer_id, source_location)
    target = db.execute(
        """
        SELECT id FROM attendance_pay_profiles
        WHERE officer_id = ? AND location_id = ?
        """,
        (officer_id, target_location),
    ).fetchone()
    copied = target is None
    if copied:
        origin = db.execute(
            """
            SELECT gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut
            FROM attendance_pay_profiles WHERE id = ?
            """,
            (source["id"],),
        ).fetchone()
        target = db.execute(
            """
            INSERT INTO attendance_pay_profiles
                (officer_id, location_id, active, gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut)
            VALUES (?, ?, FALSE, ?, ?, ?, ?)
            RETURNING id
            """,
            (
                officer_id,
                target_location,
                origin["gaji_ojt"],
                origin["jabatan_ojt"],
                origin["gaji_lanjut"],
                origin["jabatan_lanjut"],
            ),
        ).fetchone()
        db.execute(
            """
            INSERT INTO attendance_benefits
                (officer_id, nama, ikut, mode, nilai, struktur, jenis, penanggung, pay_profile_id)
            SELECT officer_id, nama, ikut, mode, nilai, struktur, jenis, penanggung, ?
            FROM attendance_benefits
            WHERE pay_profile_id = ?
            """,
            (target["id"], source["id"]),
        )
    _apply_pay_profile(db, officer_id, target["id"])
    return {
        "dari": _location_name(db, source_location),
        "ke": _location_name(db, target_location),
        "salin": "1" if copied else "0",
    }


def _current_pay_location(db, officer_id: int) -> int | None:
    row = db.execute(
        """
        SELECT location_id FROM attendance_placements
        WHERE officer_id = ? AND (end_date IS NULL OR end_date >= CURRENT_DATE)
        ORDER BY start_date DESC LIMIT 1
        """,
        (officer_id,),
    ).fetchone()
    return row["location_id"] if row else None


def status_date_fields(person) -> list[dict]:
    """The date each account status needs: when that status happened."""
    helps = {code: text for code, _name, text in STATUS_HELP}
    current = person.get("employment_status") or ""
    rows = []
    for code, _name in ACCOUNT_CHOICES:
        if code == "aktif":
            value = _day_text(person.get("activated_on"))
        elif code == "ojt":
            value = _day_text(person.get("ojt_start"))
        elif code == "resign":
            value = _day_text(person.get("resigned_date"))
        elif current == code:
            value = _day_text(person.get("status_on"))
        else:
            value = ""
        help_text = helps.get(code, "")
        if person.get("new_joiner") and code == person.get("status_choice"):
            help_text = (help_text + " New Joiner: baru aktif, 25 hari atau kurang.").strip()
        rows.append({
            "code": code,
            "label": STATUS_DATE_LABELS[code],
            "value": value,
            "help": help_text,
        })
    return rows


def save_account_status(officer_id: int, status: str, happened: str) -> str:
    """Change only the account status, on the date that status happened."""
    if status not in dict(ACCOUNT_CHOICES):
        return "Pilih status akun."
    when = _as_date(happened)
    if when is None:
        return f"Isi {STATUS_DATE_LABELS.get(status, 'tanggal status')}."
    with connect() as db:
        current = db.execute(
            """
            SELECT onboard_state, joined_date, activated_on, status_on, ojt_months, ojt_start,
                   ojt_decision, resigned_date, employment_status, attendance_enabled
            FROM attendance_officers WHERE id = ?
            """,
            (officer_id,),
        ).fetchone()
        if not current:
            return "Karyawan tidak ditemukan."
        if current["onboard_state"] != "disetujui" and status not in ("resign", "blacklist"):
            return "Petugas ini masih Onboard. Setujui dulu sebelum mengganti status ini."
        resigned = when if status == "resign" else None
        activated = when if status == "aktif" else current["activated_on"]
        months = current["ojt_months"]
        start = current["ojt_start"]
        decision = current["ojt_decision"] or ""
        if status == "ojt":
            months = months if months in OJT_MONTHS else 3
            start = when
            decision = ""
        elif status == "lulus" and not decision:
            decision = "lulus"
        ready = current["onboard_state"] == "disetujui" and status in ("aktif", "ojt", "lulus")
        same = (
            status == (current["employment_status"] or "")
            and _as_date(current["status_on"]) == when
            and _as_date(current["resigned_date"]) == _as_date(resigned)
            and _as_date(current["activated_on"]) == _as_date(activated)
            and _as_date(current["ojt_start"]) == _as_date(start)
            and (current["ojt_months"] or None) == (months or None)
            and (current["ojt_decision"] or "") == decision
            and bool(current["attendance_enabled"]) == bool(ready)
        )
        if same:
            return ""
        before = _work_facts(db, officer_id)
        db.execute(
            """
            UPDATE attendance_officers
            SET employment_status = ?, attendance_enabled = ?, activated_on = ?, status_on = ?,
                resigned_date = ?, ojt_months = ?, ojt_start = ?, ojt_decision = ?
            WHERE id = ?
            """,
            (status, ready, activated, when, resigned, months or None, start or None, decision, officer_id),
        )
        _sync_active_pay(db, officer_id)
        if status in ("resign", "blacklist"):
            db.execute(
                """
                UPDATE attendance_placements
                SET end_date = ?
                WHERE officer_id = ? AND end_date IS NULL
                """,
                (when, officer_id),
            )
        _sync_user_profile(db, officer_id)
        _record_work(db, officer_id, before)
    return ""


def save_employee_profile(officer_id: int, form):
    nama = " ".join((form.get("nama") or "").split())
    status = form.get("status") or ""
    joined = (form.get("masuk") or "")[:10]
    location_raw = form.get("lokasi") or ""
    posisi = form.get("posisi") or ""
    gender = form.get("gender") or ""
    marital = form.get("marital") or ""
    if posisi not in dict(POSITIONS):
        return "Pilih posisi: SPV, Leader, Admin, atau Gaskir."
    if gender not in dict(GENDERS):
        return "Pilih jenis kelamin."
    if marital not in dict(MARITAL):
        return "Pilih status nikah."
    if not nama or status not in EMPLOYMENT_CODES or not joined or not location_raw.isdigit():
        return "Isi nama, tanggal masuk, lokasi, dan status."
    location_id = int(location_raw)
    second_raw = (form.get("lokasi_kedua") or "").strip()
    secondary_id = int(second_raw) if second_raw.isdigit() else None
    if secondary_id == location_id:
        return "Lokasi kedua harus berbeda dari lokasi utama."
    gaji = parse_money(form.get("gaji")) if "gaji" in form else None
    minimum, min_error = parse_hours(form.get("min_jam"))
    if min_error:
        return min_error
    maximum, max_error = parse_hours(form.get("max_jam"))
    if max_error:
        return max_error
    work_days, work_error = parse_days(form.get("min_hadir"))
    if work_error:
        return work_error
    off_days, off_error = parse_days(form.get("min_off"))
    if off_error:
        return off_error
    if minimum is not None and maximum is not None and maximum < minimum:
        return "Maksimum kehadiran tidak boleh lebih kecil dari minimum."
    if minimum is None and maximum is not None:
        with connect() as db:
            place = db.execute(
                "SELECT min_attend_minutes FROM attendance_locations WHERE id = ?",
                (location_id,),
            ).fetchone()
        if place and place["min_attend_minutes"] is not None and maximum < place["min_attend_minutes"]:
            return "Maksimum kehadiran tidak boleh lebih kecil dari minimum lokasi."
    testing = form.get("testing") == "1"
    move = (form.get("pindah") or joined)[:10]
    with connect() as db:
        current = db.execute(
            """
            SELECT joined_date, onboard_state, gaji_pokok, ojt_months, ojt_start, ojt_decision,
                   activated_on, status_on, attendance_enabled
            FROM attendance_officers WHERE id = ?
            """,
            (officer_id,),
        ).fetchone()
        if not current:
            return "Karyawan tidak ditemukan."
        if secondary_id is not None:
            second = db.execute(
                "SELECT id FROM attendance_locations WHERE id = ?",
                (secondary_id,),
            ).fetchone()
            if not second:
                return "Lokasi kedua tidak ditemukan."
        before = _work_facts(db, officer_id)
        if current["onboard_state"] != "disetujui" and status not in ("onboard", "resign", "blacklist"):
            status = "onboard"
        when = None
        if status in STATUS_DATE_LABELS:
            when = _as_date(form.get("tanggal"))
            if when is None:
                return f"Isi {STATUS_DATE_LABELS[status]}."
        ready = current["onboard_state"] == "disetujui" and status in ("aktif", "ojt", "lulus")
        activated = when if status == "aktif" and when else current["activated_on"]
        resigned = when if status == "resign" else None
        if gaji is None:
            gaji = current["gaji_pokok"] or 0
        months_raw = (form.get("ojt_bulan") or "").strip()
        months = int(months_raw) if months_raw.isdigit() and int(months_raw) in OJT_MONTHS else current["ojt_months"]
        start = (form.get("ojt_mulai") or "")[:10] or _day_text(current["ojt_start"])
        decision = current["ojt_decision"] or ""
        if status == "ojt":
            months = months if months in OJT_MONTHS else 3
            start = when.isoformat()
            decision = ""
        elif status == "lulus" and not decision:
            decision = "lulus"
        status_on = when if when else current["status_on"]
        db.execute(
            """
            UPDATE attendance_officers
            SET nama = ?, joined_date = ?, gaji_pokok = ?, employment_status = ?,
                testing = ?, resigned_date = ?, posisi = ?,
                gender = ?, marital_status = ?,
                attendance_enabled = ?, activated_on = ?, status_on = ?,
                min_attend_minutes = ?, max_attend_minutes = ?,
                min_work_days = ?, min_off_days = ?,
                ojt_months = ?, ojt_start = ?, ojt_decision = ?,
                secondary_location_id = ?
            WHERE id = ?
            """,
            (
                nama, joined, gaji, status, testing, resigned, posisi,
                gender, marital,
                ready, activated, status_on,
                minimum, maximum, work_days, off_days,
                months or None, start or None, decision, secondary_id, officer_id,
            ),
        )
        _sync_active_pay(db, officer_id)
        db.execute(
            """
            UPDATE attendance_officers
            SET pemilik_rekening = ?
            WHERE id = ? AND pemilik_sendiri
            """,
            (nama, officer_id),
        )
        if current["onboard_state"] != "disetujui":
            db.execute(
                "UPDATE attendance_officers SET pending_location_id = ? WHERE id = ?",
                (location_id, officer_id),
            )
            _sync_user_profile(db, officer_id)
            _record_work(db, officer_id, before)
            return "", {}
        open_row = db.execute(
            """
            SELECT id, location_id, start_date FROM attendance_placements
            WHERE officer_id = ? AND end_date IS NULL
            ORDER BY start_date DESC LIMIT 1
            """,
            (officer_id,),
        ).fetchone()
        if open_row and open_row["location_id"] == location_id:
            if status in ("resign", "blacklist"):
                db.execute(
                    "UPDATE attendance_placements SET end_date = ? WHERE id = ?",
                    (when, open_row["id"]),
                )
            else:
                _ensure_pay_profile(db, officer_id, location_id)
            _sync_user_profile(db, officer_id)
            _record_work(db, officer_id, before)
            return "", {}
        moved_from = open_row["location_id"] if open_row and open_row["location_id"] != location_id else None
        if open_row and open_row["location_id"] != location_id:
            previous = datetime.strptime(move, "%Y-%m-%d").date() - timedelta(days=1)
            started = open_row["start_date"]
            if hasattr(started, "isoformat"):
                started = started
            else:
                started = datetime.strptime(str(started)[:10], "%Y-%m-%d").date()
            closed = previous if previous >= started else started
            db.execute(
                "UPDATE attendance_placements SET end_date = ? WHERE id = ? AND end_date IS NULL",
                (closed, open_row["id"]),
            )
        if status not in ("resign", "blacklist"):
            db.execute(
                """
                INSERT INTO attendance_placements (officer_id, location_id, start_date)
                VALUES (?, ?, ?)
                ON CONFLICT (officer_id, location_id, start_date) DO UPDATE SET end_date = NULL
                """,
                (officer_id, location_id, move),
            )
        elif not open_row and status == "resign":
            db.execute(
                """
                INSERT INTO attendance_placements (officer_id, location_id, start_date, end_date)
                VALUES (?, ?, ?, ?)
                ON CONFLICT (officer_id, location_id, start_date) DO UPDATE SET end_date = EXCLUDED.end_date
                """,
                (officer_id, location_id, joined, when),
            )
        notice = {}
        if moved_from and status not in ("resign", "blacklist"):
            notice = _replicate_pay_profile(db, officer_id, moved_from, location_id)
        elif status not in ("resign", "blacklist"):
            _ensure_pay_profile(db, officer_id, location_id)
        _sync_user_profile(db, officer_id)
        _record_work(db, officer_id, before)
        if notice.get("dari") and notice.get("ke"):
            copied = "disalin dari lokasi sebelumnya" if notice.get("salin") == "1" else "profil yang sudah ada dipakai lagi"
            _add_log(db, officer_id, "gaji", "Profil gaji aktif", notice["dari"], f"{notice['ke']} · {copied}")
    return "", notice


def _benefit_rows(form) -> tuple[list[dict], str]:
    ids = form.getlist("benefit_id")
    kinds = form.getlist("benefit_jenis")
    names = form.getlist("benefit_nama")
    modes = form.getlist("benefit_mode")
    amounts = form.getlist("benefit_nilai")
    payers = form.getlist("benefit_penanggung")
    drop = form.get("hapus") or ""
    labels = dict(BENEFIT_KINDS)
    count = max(len(ids), len(kinds), len(names), len(modes), len(amounts), len(payers), 0)
    rows = []
    for index in range(count):
        raw_id = ids[index] if index < len(ids) else ""
        if drop.isdigit() and raw_id == drop:
            continue
        jenis = (kinds[index] if index < len(kinds) else "").strip()
        nama = " ".join(((names[index] if index < len(names) else "") or "").split())
        mode = modes[index] if index < len(modes) else "nominal"
        nilai_raw = amounts[index] if index < len(amounts) else ""
        penanggung = payers[index] if index < len(payers) else "perusahaan"
        if jenis not in labels:
            if raw_id.isdigit() or (nilai_raw or "").strip() or nama:
                return [], "Pilih jenis benefit."
            continue
        if mode not in ("nominal", "persen"):
            mode = "nominal"
        if penanggung not in dict(BENEFIT_PAYERS):
            penanggung = "perusahaan"
        nilai, error = parse_rate(nilai_raw, mode)
        if error:
            return [], error
        if jenis == "lain":
            if not nama:
                if not raw_id.isdigit() and nilai == 0:
                    continue
                return [], "Isi nama untuk benefit Lainnya."
        else:
            nama = labels[jenis]
        rows.append(
            {
                "id": int(raw_id) if raw_id.isdigit() else None,
                "jenis": jenis,
                "nama": nama,
                "mode": mode,
                "nilai": nilai,
                "penanggung": penanggung,
            }
        )
    return rows, ""


def _jabatan_cash(db, officer_id: int, struktur: str, pokok: int) -> int:
    found = db.execute(
        """
        SELECT mode, nilai, penanggung, ikut
        FROM attendance_benefits
        WHERE officer_id = ? AND struktur = ? AND jenis = 'jabatan'
        """,
        (officer_id, struktur),
    ).fetchall()
    total = 0
    for item in found:
        if item["penanggung"] == "perusahaan" and item["ikut"]:
            total += benefit_amount(pokok, item["mode"], item["nilai"])
    return total


def save_officer_deposit(officer_id: int, amount_raw: str | None) -> str:
    """Edit deposit from the officer profile; every change is logged on payroll + adjustment history."""
    return save_payroll_syarat(
        officer_id,
        note="Diubah dari profil petugas · Deposit",
        potongan_raw="" if amount_raw is None else amount_raw,
        potongan_reason="deposit",
        section="potongan",
    )


def save_officer_tunjangan(officer_id: int, form) -> str:
    """Edit Tunjangan Jabatan alone (active structure by default); every change is logged on Gaji."""
    person = employee_profile(officer_id)
    if person is None:
        return "Karyawan tidak ditemukan."
    struktur = (form.get("struktur") or person.get("struktur_aktif") or "lanjut").strip()
    if struktur not in ("ojt", "lanjut"):
        struktur = "lanjut"
    jabatan = parse_money(form.get("tunjangan_jabatan"))
    title = "Selama OJT" if struktur == "ojt" else "Setelah OJT"
    with connect() as db:
        before = _pay_facts(db, officer_id, struktur)
        before_jabatan = dict(before).get(f"Tunjangan Jabatan {title}", "Rp 0")
        location_id = _current_pay_location(db, officer_id)
        profile = db.execute(
            """
            SELECT id FROM attendance_pay_profiles
            WHERE officer_id = ? AND active
            ORDER BY id DESC LIMIT 1
            """,
            (officer_id,),
        ).fetchone()
        if not profile and location_id:
            profile = _ensure_pay_profile(db, officer_id, location_id)
        profile_id = profile["id"] if profile else None
        jabatan_col = "jabatan_ojt" if struktur == "ojt" else "jabatan_lanjut"
        if profile_id:
            db.execute(
                f"UPDATE attendance_pay_profiles SET {jabatan_col} = ? WHERE id = ?",
                (jabatan, profile_id),
            )
        else:
            db.execute(
                f"UPDATE attendance_officers SET {jabatan_col} = ? WHERE id = ?",
                (jabatan, officer_id),
            )
        if profile_id:
            _apply_pay_profile(db, officer_id, profile_id)
        else:
            _sync_active_pay(db, officer_id)
        after = _pay_facts(db, officer_id, struktur)
        after_jabatan = dict(after).get(f"Tunjangan Jabatan {title}", "Rp 0")
        _add_log(db, officer_id, "gaji", f"Tunjangan Jabatan {title}", before_jabatan, after_jabatan)
        _sync_user_profile(db, officer_id)
    return ""


def save_compensation(officer_id: int, form) -> str:
    struktur = form.get("struktur") or "lanjut"
    if struktur not in ("ojt", "lanjut"):
        struktur = "lanjut"
    pokok = parse_money(form.get("gaji"))
    jabatan = parse_money(form.get("tunjangan_jabatan"))
    amounts = {code: parse_money(form.get(code)) for code, _label, _group in PAY_COMPONENTS}
    labels = {code: label for code, label, _group in PAY_COMPONENTS}
    kept = tuple(labels)
    with connect() as db:
        officer = db.execute("SELECT id FROM attendance_officers WHERE id = ?", (officer_id,)).fetchone()
        if not officer:
            return "Karyawan tidak ditemukan."
        before = _pay_facts(db, officer_id, struktur)
        location_id = _current_pay_location(db, officer_id)
        profile = db.execute(
            """
            SELECT id FROM attendance_pay_profiles
            WHERE officer_id = ? AND active
            ORDER BY id DESC LIMIT 1
            """,
            (officer_id,),
        ).fetchone()
        if not profile and location_id:
            profile = _ensure_pay_profile(db, officer_id, location_id)
        profile_id = profile["id"] if profile else None
        if profile_id:
            fee_col = "gaji_ojt" if struktur == "ojt" else "gaji_lanjut"
            jabatan_col = "jabatan_ojt" if struktur == "ojt" else "jabatan_lanjut"
            db.execute(
                f"UPDATE attendance_pay_profiles SET {fee_col} = ?, {jabatan_col} = ? WHERE id = ?",
                (pokok, jabatan, profile_id),
            )
        elif struktur == "ojt":
            db.execute(
                "UPDATE attendance_officers SET gaji_ojt = ?, jabatan_ojt = ? WHERE id = ?",
                (pokok, jabatan, officer_id),
            )
        else:
            db.execute(
                "UPDATE attendance_officers SET gaji_lanjut = ?, jabatan_lanjut = ? WHERE id = ?",
                (pokok, jabatan, officer_id),
            )
        if profile_id:
            db.execute(
                f"DELETE FROM attendance_benefits WHERE pay_profile_id = ? AND struktur = ? AND jenis NOT IN ({', '.join('?' for _ in kept)})",
                (profile_id, struktur, *kept),
            )
        else:
            db.execute(
                f"DELETE FROM attendance_benefits WHERE officer_id = ? AND pay_profile_id IS NULL AND struktur = ? AND jenis NOT IN ({', '.join('?' for _ in kept)})",
                (officer_id, struktur, *kept),
            )
        for code, nilai in amounts.items():
            if code in ("bpjs_tk", "bpjs_kes"):
                penanggung = form.get(code + "_mode") or "none"
                if penanggung not in dict(BPJS_MODES):
                    penanggung = "none"
            else:
                penanggung = "perusahaan"
            if profile_id:
                existing = db.execute(
                    """
                    SELECT id FROM attendance_benefits
                    WHERE pay_profile_id = ? AND struktur = ? AND jenis = ?
                    ORDER BY id LIMIT 1
                    """,
                    (profile_id, struktur, code),
                ).fetchone()
            else:
                existing = db.execute(
                    """
                    SELECT id FROM attendance_benefits
                    WHERE officer_id = ? AND pay_profile_id IS NULL AND struktur = ? AND jenis = ?
                    ORDER BY id LIMIT 1
                    """,
                    (officer_id, struktur, code),
                ).fetchone()
            if existing:
                db.execute(
                    """
                    UPDATE attendance_benefits
                    SET nama = ?, ikut = TRUE, mode = 'nominal', nilai = ?, penanggung = ?
                    WHERE id = ?
                    """,
                    (labels[code], nilai, penanggung, existing["id"]),
                )
            elif nilai or code in ("bpjs_tk", "bpjs_kes"):
                db.execute(
                    """
                    INSERT INTO attendance_benefits
                        (officer_id, nama, ikut, mode, nilai, struktur, jenis, penanggung, pay_profile_id)
                    VALUES (?, ?, TRUE, 'nominal', ?, ?, ?, ?, ?)
                    """,
                    (officer_id, labels[code], nilai, struktur, code, penanggung, profile_id),
                )
        if profile_id:
            _apply_pay_profile(db, officer_id, profile_id)
        else:
            _sync_active_pay(db, officer_id)
        _write_changes(db, officer_id, "gaji", before, _pay_facts(db, officer_id, struktur))
        _sync_user_profile(db, officer_id)
    return ""


def decide_ojt(officer_id: int, passed: bool) -> str:
    with connect() as db:
        row = db.execute(
            """
            SELECT employment_status, ojt_months, ojt_start, ojt_decision
            FROM attendance_officers WHERE id = ?
            """,
            (officer_id,),
        ).fetchone()
        if not row:
            return "Karyawan tidak ditemukan."
        if row["employment_status"] != "ojt":
            return "Keputusan ini untuk petugas yang statusnya OJT."
        end = _ojt_end(row["ojt_start"], row["ojt_months"])
        if end is None:
            return "Isi lama OJT dan tanggal mulainya dulu."
        if date.today() < end:
            return f"Masa OJT masih berjalan sampai {day_label(end)}."
        if passed:
            db.execute(
                """
                UPDATE attendance_officers
                SET employment_status = 'lulus', ojt_decision = 'lulus', attendance_enabled = TRUE
                WHERE id = ?
                """,
                (officer_id,),
            )
            _add_log(db, officer_id, "status", "Status", "OJT", "Pass OJT")
        else:
            db.execute(
                "UPDATE attendance_officers SET ojt_decision = 'gagal' WHERE id = ?",
                (officer_id,),
            )
            _add_log(db, officer_id, "status", "Keputusan OJT", "", "Tidak lulus. Tetap OJT")
        _sync_active_pay(db, officer_id)
    return ""


def save_account(officer_id: int, form) -> str:
    rekening = "".join(ch for ch in (form.get("rekening") or "") if ch.isdigit())
    bank = (form.get("bank") or "").strip()
    sendiri = form.get("pemilik") != "orang"
    assigned = " ".join((form.get("pemilik_nama") or "").split())
    with connect() as db:
        officer = db.execute(
            "SELECT id, nama, bank, rekening, pemilik_sendiri, pemilik_rekening FROM attendance_officers WHERE id = ?",
            (officer_id,),
        ).fetchone()
        if not officer:
            return "Karyawan tidak ditemukan."
        if rekening and not sendiri and not assigned:
            return "Isi nama pemilik rekening. Boleh nama karyawan, atau nama orang yang ditunjuk."
        before = _account_facts(db, officer_id)
        owner = officer["nama"] if sendiri else assigned
        db.execute(
            """
            UPDATE attendance_officers
            SET bank = ?, rekening = ?, pemilik_sendiri = ?, pemilik_rekening = ?
            WHERE id = ?
            """,
            (bank, rekening, sendiri, owner, officer_id),
        )
        _write_changes(db, officer_id, "rekening", before, _account_facts(db, officer_id))
    return ""
