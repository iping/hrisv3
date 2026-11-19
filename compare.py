"""Bandingkan Officer Profile unggahan baru dengan data yang sudah ada."""

from __future__ import annotations

from pathlib import Path

from archive import file_status, load_profiles, person_key, split_period
from calculate_payroll import text
from db import connect, get_upload, list_uploads
from openpyxl import load_workbook

COMPARE_FIELDS = (
    ("nama", "Nama"),
    ("lokasi", "Lokasi"),
    ("numeric_id", "Numeric ID"),
    ("bank", "Bank"),
    ("rekening", "No. Rekening"),
    ("phone", "Phone"),
    ("email", "Email"),
    ("joined", "Tanggal masuk"),
    ("resigned", "Tanggal resign"),
    ("activated", "Activated"),
)

ROOT = Path(__file__).resolve().parent
SESSIONS = ROOT / "output" / "sessions"


def session_source(sid: str) -> Path | None:
    path = SESSIONS / sid / "source.xlsx"
    return path if path.is_file() else None


def period_from_source(path: Path) -> tuple[str, str, str]:
    """Return period_start, period_end, source_name hint from Payroll/Attendance."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        period = ""
        if "Payroll" in wb.sheetnames:
            rows = list(wb["Payroll"].iter_rows(max_row=3, values_only=True))
            for row in rows:
                for cell in row or ():
                    raw = text(cell)
                    if "s/d" in raw:
                        period = raw
                        break
                if period:
                    break
        start, end = split_period(period)
        if not start or not end:
            from calculate_payroll import attendance_bounds, sheet_rows

            start, end = attendance_bounds(sheet_rows(wb, "Attendance"))
        return start, end, period or (f"{start} s/d {end}" if start and end else "")
    finally:
        wb.close()


def profiles_from_upload(sid: str) -> tuple[dict[str, dict], tuple[str, str], Path | None]:
    path = session_source(sid)
    if path is None:
        return {}, ("", ""), None
    profiles = load_profiles(path)
    start, end, _ = period_from_source(path)
    upload = get_upload(sid) or {}
    if not start:
        start = text(upload.get("period_start") or "")
    if not end:
        end = text(upload.get("period_end") or "")
    return profiles, (start, end), path


def profiles_from_month(month_id: int) -> tuple[dict[str, dict], tuple[str, str], dict | None]:
    with connect() as db:
        month = db.execute("SELECT * FROM months WHERE id = ?", (month_id,)).fetchone()
        if month is None:
            return {}, ("", ""), None
        lines = db.execute(
            "SELECT * FROM payroll_lines WHERE month_id = ?",
            (month_id,),
        ).fetchall()
    profiles = {}
    for row in lines:
        key = row["person_key"] or person_key(row["lokasi"], row["username"], row["nama"])
        profiles[key] = {
            "nama": text(row["nama"]),
            "username": text(row["username"]),
            "lokasi": text(row["lokasi"]),
            "numeric_id": text(row["numeric_id"]),
            "bank": "",
            "rekening": "",
            "phone": "",
            "email": "",
            "joined": text(row["joined_date"]),
            "resigned": text(row["resigned_date"]),
            "activated": text(row["activated"]),
            "payroll_status": text(row["status"]) or "aktif",
            "source": "arsip",
        }
    return profiles, (text(month["period_start"]), text(month["period_end"])), dict(month)


def attendance_index() -> dict[str, dict]:
    with connect() as db:
        rows = db.execute(
            """
            SELECT id, username, nama, employment_status, joined_date, resigned_date,
                   attendance_enabled, gaji_pokok
            FROM attendance_officers
            """
        ).fetchall()
    found = {}
    for row in rows:
        user = text(row["username"]).lower()
        if user:
            found[user] = dict(row)
    return found


def by_username(profiles: dict[str, dict]) -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = {}
    for key, row in profiles.items():
        user = text(row.get("username")).lower()
        if not user:
            continue
        item = {**row, "person_key": key}
        grouped.setdefault(user, []).append(item)
    return grouped


def _value(row: dict, key: str) -> str:
    if key == "joined":
        return text(row.get("joined") or row.get("joined_date"))
    if key == "resigned":
        return text(row.get("resigned") or row.get("resigned_date"))
    return text(row.get(key))


def field_changes(old: dict, new: dict) -> list[dict]:
    changes = []
    for key, label in COMPARE_FIELDS:
        before = _value(old, key)
        after = _value(new, key)
        if before == after:
            continue
        changes.append({"field": key, "label": label, "before": before or "—", "after": after or "—"})
    return changes


def pick_baseline(upload: dict, uploads: list[dict], months: list[dict]) -> dict:
    """Default baseline: previous upload with source, else previous arsip month."""
    sid = upload.get("id") or ""
    for row in uploads:
        if row["id"] == sid:
            continue
        if session_source(row["id"]) is None:
            continue
        return {"kind": "upload", "id": row["id"], "label": row.get("label") or row.get("alias") or row["id"]}
    start = text(upload.get("period_start") or "")
    for month in months:
        if start and text(month.get("period_start")) == start:
            continue
        if start and text(month.get("period_end") or "") >= start:
            continue
        return {
            "kind": "month",
            "id": str(month["id"]),
            "label": f"{month['period_start']} s/d {month['period_end']} · {month.get('source_name') or 'arsip'}",
        }
    if months:
        month = months[0]
        if not start or text(month.get("period_start")) != start:
            return {
                "kind": "month",
                "id": str(month["id"]),
                "label": f"{month['period_start']} s/d {month['period_end']} · {month.get('source_name') or 'arsip'}",
            }
    return {"kind": "", "id": "", "label": "Tidak ada data lama"}


def resolve_baseline(kind: str, baseline_id: str, upload: dict, uploads: list[dict], months: list[dict]) -> tuple[dict[str, dict], tuple[str, str], dict]:
    if kind == "upload" and baseline_id:
        profiles, period, path = profiles_from_upload(baseline_id)
        row = get_upload(baseline_id) or next((u for u in uploads if u["id"] == baseline_id), {})
        meta = {
            "kind": "upload",
            "id": baseline_id,
            "label": (row or {}).get("label") or (row or {}).get("alias") or baseline_id,
            "period": f"{period[0]} s/d {period[1]}" if period[0] else "",
            "exists": path is not None,
        }
        return profiles, period, meta
    if kind == "month" and baseline_id:
        profiles, period, month = profiles_from_month(int(baseline_id))
        meta = {
            "kind": "month",
            "id": baseline_id,
            "label": (
                f"{month['period_start']} s/d {month['period_end']} · {month.get('source_name') or 'arsip'}"
                if month
                else baseline_id
            ),
            "period": f"{period[0]} s/d {period[1]}" if period[0] else "",
            "exists": month is not None,
        }
        return profiles, period, meta
    chosen = pick_baseline(upload, uploads, months)
    if not chosen.get("kind") or not chosen.get("id"):
        return {}, ("", ""), {
            "kind": "",
            "id": "",
            "label": chosen.get("label") or "Tidak ada data lama",
            "period": "",
            "exists": False,
        }
    return resolve_baseline(chosen["kind"], chosen["id"], upload, uploads, months)


def current_payroll_keys(period: tuple[str, str]) -> set[str]:
    start, end = period
    if not start or not end:
        return set()
    with connect() as db:
        month = db.execute(
            "SELECT id FROM months WHERE period_start = ? AND period_end = ?",
            (start, end),
        ).fetchone()
        if month is None:
            return set()
        rows = db.execute(
            "SELECT person_key, lokasi, username, nama FROM payroll_lines WHERE month_id = ?",
            (month["id"],),
        ).fetchall()
    keys = set()
    users = set()
    for row in rows:
        key = row["person_key"] or person_key(row["lokasi"], row["username"], row["nama"])
        keys.add(key)
        user = text(row["username"]).lower()
        if user:
            users.add(user)
    return keys | {f"user:{user}" for user in users}


def compare_profiles(
    current: dict[str, dict],
    baseline: dict[str, dict],
    period: tuple[str, str],
    baseline_kind: str = "",
) -> list[dict]:
    start, end = period
    attendance = attendance_index()
    baseline_users = by_username(baseline)
    payroll_marks = current_payroll_keys(period) if baseline_kind == "month" else set()
    rows = []
    matched_baseline: set[str] = set()

    for key, new in current.items():
        user = text(new.get("username")).lower()
        old = baseline.get(key)
        match_note = ""
        old_key = key if old is not None else ""
        if old is None and user and user in baseline_users:
            candidates = baseline_users[user]
            if len(candidates) == 1:
                old_key = candidates[0]["person_key"]
                old = {k: v for k, v in candidates[0].items() if k != "person_key"}
                match_note = "Cocok lewat username (lokasi berubah)"
            else:
                match_note = "Username muncul di beberapa lokasi lama"
        if old_key:
            matched_baseline.add(old_key)
        status, when, note = file_status(new.get("joined", ""), new.get("resigned", ""), start, end)
        changes = field_changes(old or {}, new) if old else []
        on_payroll = (not payroll_marks) or key in payroll_marks or (user and f"user:{user}" in payroll_marks)
        if old is None:
            if baseline_kind == "month" and not on_payroll and status not in {"baru", "resign"}:
                kind = "profil"
                match_note = match_note or "Hanya di Officer Profile, tidak ada di payroll periode ini / arsip lama"
            else:
                kind = "baru"
        elif changes or match_note.startswith("Cocok"):
            kind = "berubah"
        else:
            kind = "sama"
        if status == "baru" and kind == "sama":
            kind = "status"
        if status == "resign" and kind == "sama":
            kind = "status"
        account = attendance.get(user) if user else None
        rows.append(
            {
                "kind": kind,
                "person_key": key,
                "nama": new.get("nama") or "",
                "username": new.get("username") or "",
                "lokasi": new.get("lokasi") or "",
                "numeric_id": new.get("numeric_id") or "",
                "joined": new.get("joined") or "",
                "resigned": new.get("resigned") or "",
                "activated": new.get("activated") or "",
                "status": status,
                "status_note": note,
                "status_when": when,
                "changes": changes,
                "change_text": "; ".join(f"{c['label']}: {c['before']} → {c['after']}" for c in changes),
                "match_note": match_note,
                "old_lokasi": (old or {}).get("lokasi") or "",
                "old_joined": _value(old or {}, "joined"),
                "old_resigned": _value(old or {}, "resigned"),
                "old_activated": text((old or {}).get("activated")),
                "in_karyawan": bool(account),
                "karyawan_status": (account or {}).get("employment_status") or "",
                "karyawan_id": (account or {}).get("id"),
                "on_payroll": on_payroll if baseline_kind == "month" else True,
            }
        )

    for key, old in baseline.items():
        if key in matched_baseline or key in current:
            continue
        user = text(old.get("username")).lower()
        account = attendance.get(user) if user else None
        rows.append(
            {
                "kind": "hilang",
                "person_key": key,
                "nama": old.get("nama") or "",
                "username": old.get("username") or "",
                "lokasi": old.get("lokasi") or "",
                "numeric_id": old.get("numeric_id") or "",
                "joined": _value(old, "joined"),
                "resigned": _value(old, "resigned"),
                "activated": text(old.get("activated")),
                "status": old.get("payroll_status") or "aktif",
                "status_note": "Ada di data lama, tidak ada di unggahan ini",
                "status_when": "",
                "changes": [],
                "change_text": "Tidak muncul di Officer Profile unggahan",
                "match_note": "",
                "old_lokasi": old.get("lokasi") or "",
                "old_joined": _value(old, "joined"),
                "old_resigned": _value(old, "resigned"),
                "old_activated": text(old.get("activated")),
                "in_karyawan": bool(account),
                "karyawan_status": (account or {}).get("employment_status") or "",
                "karyawan_id": (account or {}).get("id"),
                "on_payroll": False,
            }
        )

    order = {"baru": 0, "status": 1, "berubah": 2, "hilang": 3, "profil": 4, "sama": 5}
    rows.sort(
        key=lambda row: (
            order.get(row["kind"], 9),
            0 if row["status"] == "baru" else 1 if row["status"] == "resign" else 2,
            row["lokasi"].lower(),
            row["nama"].lower(),
        )
    )
    return rows


def summarize(rows: list[dict]) -> dict:
    return {
        "total": len(rows),
        "baru": sum(1 for row in rows if row["kind"] == "baru"),
        "berubah": sum(1 for row in rows if row["kind"] == "berubah"),
        "status": sum(1 for row in rows if row["kind"] == "status"),
        "hilang": sum(1 for row in rows if row["kind"] == "hilang"),
        "profil": sum(1 for row in rows if row["kind"] == "profil"),
        "sama": sum(1 for row in rows if row["kind"] == "sama"),
        "new_joiner": sum(1 for row in rows if row["status"] == "baru"),
        "resign": sum(1 for row in rows if row["status"] == "resign"),
        "belum_karyawan": sum(
            1
            for row in rows
            if row["kind"] not in {"hilang", "profil"} and not row["in_karyawan"]
        ),
        "perhatian": sum(1 for row in rows if row["kind"] in {"baru", "berubah", "status", "hilang"}),
    }


def location_counts(profiles: dict[str, dict]) -> dict[str, dict]:
    found: dict[str, dict] = {}
    for row in profiles.values():
        name = text(row.get("lokasi"))
        if not name:
            continue
        key = name.lower()
        bucket = found.setdefault(
            key,
            {"name": name, "people": 0, "active_people": 0, "joined": 0, "resigned": 0},
        )
        bucket["people"] += 1
        if text(row.get("activated")).lower() in {"ya", "yes", "1", "true", "aktif"}:
            bucket["active_people"] += 1
        if row.get("joined"):
            bucket["joined"] += 1
        if row.get("resigned"):
            bucket["resigned"] += 1
    return found


def system_locations() -> dict[str, dict]:
    with connect() as db:
        rows = db.execute(
            """
            SELECT l.id, l.name, l.active, l.grace_minutes, l.min_attend_minutes,
                   l.min_work_days, l.min_off_days, l.latitude, l.longitude, l.gps_radius_meters,
                   (SELECT COUNT(*) FROM attendance_placements p
                    WHERE p.location_id = l.id AND (p.end_date IS NULL OR p.end_date >= CURRENT_DATE)) AS people,
                   (SELECT COUNT(*) FROM attendance_shifts s WHERE s.location_id = l.id) AS shifts
            FROM attendance_locations l
            ORDER BY l.name
            """
        ).fetchall()
    return {text(row["name"]).lower(): dict(row) for row in rows if text(row["name"])}


def summarize_locations(rows: list[dict]) -> dict:
    return {
        "total": len(rows),
        "baru": sum(1 for row in rows if row["kind"] == "baru"),
        "berubah": sum(1 for row in rows if row["kind"] == "berubah"),
        "hilang": sum(1 for row in rows if row["kind"] == "hilang"),
        "sama": sum(1 for row in rows if row["kind"] == "sama"),
        "belum_profil": sum(1 for row in rows if row["kind"] != "hilang" and not row["in_system"]),
        "perhatian": sum(1 for row in rows if row["kind"] in {"baru", "berubah", "hilang"}),
    }


def compare_locations(
    current_profiles: dict[str, dict],
    baseline_profiles: dict[str, dict],
) -> list[dict]:
    current = location_counts(current_profiles)
    baseline = location_counts(baseline_profiles)
    system = system_locations()
    rows = []

    for key, new in current.items():
        old = baseline.get(key)
        profile = system.get(key)
        changes = []
        if old and int(old.get("people") or 0) != int(new.get("people") or 0):
            changes.append(
                {
                    "label": "Jumlah petugas di file",
                    "before": str(old.get("people") or 0),
                    "after": str(new.get("people") or 0),
                }
            )
        if old is None:
            kind = "baru"
            note = "Lokasi muncul di unggahan, belum ada di data lama"
        elif changes:
            kind = "berubah"
            note = ""
        else:
            kind = "sama"
            note = ""
        if profile is None:
            if kind == "sama":
                kind = "baru"
            note = (note + " · " if note else "") + "Belum ada profil lokasi di sistem"
        elif not profile.get("active"):
            note = (note + " · " if note else "") + "Profil lokasi nonaktif"
        rows.append(
            {
                "kind": kind,
                "name": new["name"],
                "people": new["people"],
                "old_people": (old or {}).get("people"),
                "active_people": new.get("active_people") or 0,
                "in_system": profile is not None,
                "system_id": (profile or {}).get("id"),
                "system_active": bool((profile or {}).get("active")) if profile else False,
                "system_people": (profile or {}).get("people") or 0,
                "system_shifts": (profile or {}).get("shifts") or 0,
                "grace_minutes": (profile or {}).get("grace_minutes"),
                "has_gps": bool((profile or {}).get("latitude") is not None and (profile or {}).get("longitude") is not None),
                "changes": changes,
                "change_text": "; ".join(f"{c['label']}: {c['before']} → {c['after']}" for c in changes),
                "note": note,
            }
        )

    for key, old in baseline.items():
        if key in current:
            continue
        profile = system.get(key)
        rows.append(
            {
                "kind": "hilang",
                "name": old["name"],
                "people": 0,
                "old_people": old.get("people") or 0,
                "active_people": 0,
                "in_system": profile is not None,
                "system_id": (profile or {}).get("id"),
                "system_active": bool((profile or {}).get("active")) if profile else False,
                "system_people": (profile or {}).get("people") or 0,
                "system_shifts": (profile or {}).get("shifts") or 0,
                "grace_minutes": (profile or {}).get("grace_minutes"),
                "has_gps": bool((profile or {}).get("latitude") is not None and (profile or {}).get("longitude") is not None),
                "changes": [],
                "change_text": "Tidak muncul di unggahan ini",
                "note": "Ada di data lama, tidak ada di unggahan",
            }
        )

    # System profiles never seen in this upload or baseline — useful when reviewing setup gaps.
    seen = set(current) | set(baseline)
    for key, profile in system.items():
        if key in seen or not profile.get("active"):
            continue
        rows.append(
            {
                "kind": "hilang",
                "name": profile["name"],
                "people": 0,
                "old_people": None,
                "active_people": 0,
                "in_system": True,
                "system_id": profile["id"],
                "system_active": True,
                "system_people": profile.get("people") or 0,
                "system_shifts": profile.get("shifts") or 0,
                "grace_minutes": profile.get("grace_minutes"),
                "has_gps": profile.get("latitude") is not None and profile.get("longitude") is not None,
                "changes": [],
                "change_text": "Ada profil lokasi aktif, tidak ada di unggahan",
                "note": "Profil lokasi di sistem tanpa data di file ini",
            }
        )

    order = {"baru": 0, "berubah": 1, "hilang": 2, "sama": 3}
    rows.sort(key=lambda row: (order.get(row["kind"], 9), row["name"].lower()))
    return rows


LOW_ATTEND_LIMIT = 10  # show officers with fewer than this many hadir days (OT + L)


def attendance_hadir_map(path: Path) -> dict[str, dict]:
    """Hadir (OT+L) and alpa per username from the upload Attendance sheet."""
    from calculate_payroll import load_attendance_grid, sheet_rows
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        att_rows = sheet_rows(wb, "Attendance")
        shift_rows = sheet_rows(wb, "Shift")
    finally:
        wb.close()
    if not att_rows:
        return {}
    _dates, lines, _summed = load_attendance_grid(att_rows, shift_rows)
    by_user: dict[str, dict] = {}
    for line in lines:
        user = text(line.username).lower()
        if not user:
            continue
        hadir = int(line.counts.get("OT", 0) or 0) + int(line.counts.get("L", 0) or 0)
        alpa = int(line.counts.get("A", 0) or 0)
        bucket = by_user.get(user)
        if bucket is None:
            by_user[user] = {
                "username": text(line.username),
                "nama": text(line.nama),
                "lokasi": text(line.lokasi),
                "hadir": hadir,
                "alpa": alpa,
                "in_attendance": True,
            }
        else:
            bucket["hadir"] += hadir
            bucket["alpa"] += alpa
            if text(line.lokasi) and text(line.lokasi) not in bucket["lokasi"]:
                bucket["lokasi"] = f"{bucket['lokasi']}; {text(line.lokasi)}" if bucket["lokasi"] else text(line.lokasi)
            if text(line.nama):
                bucket["nama"] = text(line.nama)
    return by_user


def payroll_usernames(path: Path) -> set[str]:
    from calculate_payroll import sheet_rows
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = sheet_rows(wb, "Payroll")
    finally:
        wb.close()
    if not rows:
        return set()
    headers = [text(cell) for cell in rows[0]]
    try:
        at = headers.index("Username")
    except ValueError:
        return set()
    found = set()
    for row in rows[1:]:
        if not row or at >= len(row):
            continue
        user = text(row[at]).lower()
        if user:
            found.add(user)
    return found


def _is_deleted_username(username: str) -> bool:
    return "deleted at" in text(username).lower()


def low_attendance_rows(path: Path, profiles: dict[str, dict], limit: int = LOW_ATTEND_LIMIT) -> list[dict]:
    """Officers with account/profile whose OT+L hadir is below limit (including 0).

    Scope for Cek unggahan: must be tied to this file (payroll/attendance) or already
    have a Petugas account. Skip stale Officer Profile-only rows with no presence.
    """
    hadir_map = attendance_hadir_map(path)
    accounts = attendance_index()
    on_payroll = payroll_usernames(path)
    profiles_by_user = by_username(profiles)
    candidates: dict[str, dict] = {}

    def remember(user: str, *, profile: dict | None = None, account: dict | None = None) -> None:
        if not user or user in candidates or _is_deleted_username(user):
            return
        profile = profile or {}
        account = account or accounts.get(user)
        in_att = user in hadir_map
        paid = user in on_payroll
        has_account = bool(account)
        # Actionable: punya akun Petugas, atau ada di payroll periode ini
        if not has_account and not paid:
            return
        if text(profile.get("resigned")) and not has_account and not paid:
            return
        candidates[user] = {
            "nama": text(profile.get("nama") or (account or {}).get("nama")),
            "username": text(profile.get("username") or (account or {}).get("username") or user),
            "lokasi": text(profile.get("lokasi")),
            "activated": text(profile.get("activated")),
            "joined": text(profile.get("joined") or (account or {}).get("joined_date") or ""),
            "resigned": text(profile.get("resigned") or (account or {}).get("resigned_date") or ""),
            "has_profile": bool(profile),
            "on_payroll": paid,
            "in_karyawan": has_account,
            "karyawan_id": (account or {}).get("id"),
            "karyawan_status": (account or {}).get("employment_status") or "",
        }

    for user, group in profiles_by_user.items():
        profile = group[0]
        activated = text(profile.get("activated")).lower()
        active_profile = activated in {"ya", "yes", "1", "true", "aktif", ""}
        account = accounts.get(user)
        if not active_profile and not account and user not in on_payroll and user not in hadir_map:
            continue
        remember(user, profile=profile, account=account)

    for user, account in accounts.items():
        if user in candidates:
            continue
        # Account holders linked to this upload (profil / payroll / absensi)
        if user not in profiles_by_user and user not in hadir_map and user not in on_payroll:
            continue
        group = profiles_by_user.get(user) or []
        remember(user, profile=group[0] if group else None, account=account)

    # Anyone on payroll / in attendance with low hadir even if profile missing
    for user in set(hadir_map) | on_payroll:
        if user in candidates:
            continue
        group = profiles_by_user.get(user) or []
        remember(user, profile=group[0] if group else None)

    rows = []
    for user, base in candidates.items():
        att = hadir_map.get(user) or {}
        hadir = int(att.get("hadir") or 0)
        if hadir >= limit:
            continue
        lokasi = text(att.get("lokasi")) or base["lokasi"]
        nama = text(att.get("nama")) or base["nama"]
        if hadir == 0 and not att.get("in_attendance"):
            note = "Punya akun/payroll, tidak ada baris di sheet Attendance"
            level = "nol"
        elif hadir == 0:
            note = "Ada di Attendance, tapi hari hadir (OT/L) = 0"
            level = "nol"
        else:
            note = f"Hanya {hadir} hari hadir (OT+L), di bawah {limit}"
            level = "rendah"
        rows.append(
            {
                **base,
                "nama": nama,
                "lokasi": lokasi,
                "hadir": hadir,
                "alpa": int(att.get("alpa") or 0),
                "in_attendance": bool(att.get("in_attendance")),
                "level": level,
                "note": note,
                "kind": level,
            }
        )

    rows.sort(key=lambda row: (row["hadir"], row["lokasi"].lower(), row["nama"].lower()))
    return rows


def summarize_low_attendance(rows: list[dict]) -> dict:
    return {
        "total": len(rows),
        "nol": sum(1 for row in rows if row["hadir"] == 0),
        "rendah": sum(1 for row in rows if 0 < row["hadir"] < LOW_ATTEND_LIMIT),
        "punya_akun": sum(1 for row in rows if row["in_karyawan"]),
        "di_payroll": sum(1 for row in rows if row["on_payroll"]),
        "limit": LOW_ATTEND_LIMIT,
    }


def compare_page(
    sid: str = "",
    baseline_kind: str = "",
    baseline_id: str = "",
    focus: str = "perhatian",
    tab: str = "petugas",
) -> dict:
    from archive import list_months

    tab = tab if tab in {"petugas", "lokasi", "kehadiran"} else "petugas"
    uploads = [row for row in list_uploads() if session_source(row["id"])]
    months = list_months()
    if not sid and uploads:
        sid = uploads[0]["id"]
    upload = get_upload(sid) if sid else None
    if upload is None and uploads:
        upload = uploads[0]
        sid = upload["id"]
    empty = {
        "upload": upload,
        "uploads": uploads,
        "months": months,
        "baseline": {"kind": "", "id": "", "label": "", "period": "", "exists": False},
        "rows": [],
        "locations": [],
        "low_attendance": [],
        "summary": summarize([]),
        "location_summary": summarize_locations([]),
        "low_attendance_summary": summarize_low_attendance([]),
        "period": "",
        "focus": focus,
        "tab": tab,
        "error": "",
        "current_count": 0,
        "baseline_count": 0,
        "attend_limit": LOW_ATTEND_LIMIT,
    }
    if upload is None:
        empty["error"] = "Belum ada unggahan dengan file source.xlsx."
        return empty

    current, period, path = profiles_from_upload(sid)
    if path is None:
        empty["error"] = "File source.xlsx unggahan ini tidak ditemukan."
        return empty

    baseline_profiles, baseline_period, baseline_meta = resolve_baseline(
        baseline_kind, baseline_id, upload, uploads, months
    )
    rows = compare_profiles(current, baseline_profiles, period, baseline_kind=baseline_meta.get("kind") or "")
    locations = compare_locations(current, baseline_profiles)
    low_attendance = low_attendance_rows(path, current)
    period_label = f"{period[0]} s/d {period[1]}" if period[0] and period[1] else text(upload.get("period_start") or "")
    return {
        "upload": upload,
        "uploads": uploads,
        "months": months,
        "baseline": baseline_meta,
        "rows": rows,
        "locations": locations,
        "low_attendance": low_attendance,
        "summary": summarize(rows),
        "location_summary": summarize_locations(locations),
        "low_attendance_summary": summarize_low_attendance(low_attendance),
        "period": period_label,
        "focus": focus or "perhatian",
        "tab": tab,
        "error": "",
        "current_count": len(current),
        "baseline_count": len(baseline_profiles),
        "attend_limit": LOW_ATTEND_LIMIT,
    }
