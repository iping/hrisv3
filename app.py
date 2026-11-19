#!/usr/bin/env python3
"""Upload tarikan HRIS, lalu tampilkan analisis payroll.

  python3 app.py

Buka http://127.0.0.1:5055

Server lokal dari pustaka standar Python. Halaman tetap dirender Jinja2.
"""

from __future__ import annotations

import json
import re
import traceback
import uuid
from contextvars import ContextVar
from datetime import date, datetime
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlencode, urlparse

from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup, escape as html_escape
from openpyxl import load_workbook

from archive import legend_terms, list_months, month_page, person_history, save_month, split_period
from compare import compare_page
from db import (
    TERMS,
    add_activity as _add_activity,
    all_terms,
    connect,
    delete_activity,
    delete_upload,
    get_activity,
    get_upload,
    list_activities,
    list_uploads,
    remember_upload,
    rename_upload,
    seed_terms,
    term_map,
    update_activity as _update_activity,
)
from attendance import attach_counts, officer_attendance
from presence import (
    EMPLOYMENT,
    GPS_RADII,
    LABELS,
    STATUS_HELP,
    add_location,
    add_officer,
    add_second_slot,
    add_shift,
    banding_queue,
    copy_shift_pattern,
    create_user,
    day_board,
    day_label,
    decide_banding,
    decide_ojt,
    decide_onboarding,
    delete_location,
    delete_shift,
    employee_directory,
    employee_profile,
    employee_row,
    officer_filters,
    officer_status_index,
    min_work_days_index,
    apply_payroll_conditions,
    officer_payroll_info,
    payroll_adjustment_log,
    payroll_blocked_usernames,
    payroll_condition_for,
    payroll_conditions_index,
    payroll_exclusions,
    PAYROLL_POTONGAN_REASONS,
    PAYROLL_TAMBAHAN_REASONS,
    save_payroll_condition,
    save_payroll_syarat,
    apply_upload_profiles,
    file_banding,
    POSITIONS,
    profile_log,
    assignment_month,
    location_directory,
    location_filters,
    location_log,
    location_record,
    location_team,
    locations as attendance_locations,
    parse_day,
    register_employee,
    save_account,
    save_compensation,
    save_officer_deposit,
    save_officer_tunjangan,
    save_account_status,
    status_date_fields,
    save_employee_profile,
    save_location,
    save_shift,
    save_shift_assignment,
    shift_groups,
    save_slots,
    shifts_for,
    slot_detail,
)
from auth import (
    ROLES,
    SESSION_COOKIE,
    SESSION_DAYS,
    authenticate,
    bind_actor,
    create_app_user,
    create_session,
    current_user,
    delete_app_user,
    destroy_session,
    get_user,
    list_users,
    reset_actor,
    seed_auth,
    update_app_user,
)
from calculate_payroll import build, filter_payroll, rupiah, text, write_excluded, write_workbook
from closing import edits_from_form, load_closing, save_closing
from profiles import (
    attach_banks,
    coverage,
    find_profile,
    location_kode,
    present_transfer,
    remark_period,
    set_location_kode,
    set_profile,
    transfer_candidates,
    write_transfer,
)
from review import load_prior_rows, review
from roster import (
    apply_choices,
    load_roster,
    location_detail,
    merge_payroll,
    page_rows,
    people_page,
    set_employment,
    sync_flags,
)

ROOT = Path(__file__).resolve().parent
SESSIONS = ROOT / "output" / "sessions"
STATIC = ROOT / "static"
SESSION_ID = re.compile(r"^[0-9a-f]{32}$")
MAX_BODY = 30 * 1024 * 1024
HOST = "127.0.0.1"
PORT = 5055
XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
STATIC_TYPES = {
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
    ".webp": "image/webp",
}

PATHS = {
    "home": "/",
    "kegiatan": "/kegiatan",
    "istilah": "/istilah",
    "arsip": "/arsip",
    "arsip_bulan": "/arsip/<month_id>",
    "banding_unggah": "/banding-unggah",
    "analyze": "/analyze",
    "kelola": "/kelola",
    "save_kelola": "/kelola",
    "kelola_petugas": "/kelola/petugas",
    "save_petugas": "/kelola/petugas",
    "kelola_lokasi": "/kelola/lokasi/<path:lokasi>",
    "kelola_absensi": "/kelola/lokasi/<path:lokasi>/petugas/<path:username>",
    "kelola_profil": "/kelola/lokasi/<path:lokasi>/petugas/<path:username>/profil",
    "save_profil": "/kelola/lokasi/<path:lokasi>/petugas/<path:username>/profil",
    "report": "/report/<sid>",
    "view_hitung": "/report/<sid>/lihat/hitung",
    "view_transfer": "/report/<sid>/lihat/transfer",
    "view_final": "/report/<sid>/lihat/final",
    "save_final": "/report/<sid>/lihat/final",
    "download": "/report/<sid>/download",
    "download_transfer": "/report/<sid>/transfer",
    "save_kegiatan": "/kegiatan",
    "hapus_kegiatan": "/kegiatan/<activity_id>/hapus",
    "simpan_alias": "/unggahan/<sid>/nama",
    "hapus_unggahan": "/unggahan/<sid>/hapus",
    "data_lokasi": "/data/lokasi",
    "data_lokasi_profil": "/data/lokasi/<location_id>",
    "data_shift": "/data/shift",
    "penempatan": "/penempatan",
    "data_karyawan": "/data/karyawan",
    "data_profil": "/data/karyawan/<officer_id>",
    "data_pengguna": "/data/karyawan/pengguna",
    "data_daftar": "/data/karyawan/daftar",
    "data_setuju": "/data/karyawan/<officer_id>/setuju",
    "login": "/login",
    "logout": "/logout",
    "users": "/pengaturan/pengguna",
    "save_user": "/pengaturan/pengguna",
    "absensi": "/absensi",
    "absensi_shift": "/absensi/shift",
    "absensi_petugas": "/absensi/petugas",
    "absensi_banding": "/absensi/banding",
    "putus_banding": "/absensi/banding/<banding_id>",
}

BOOKS = {
    "hitung": {
        "title": "Hasil hitung",
        "default": "Slip Gaji",
        "blurb": "Cek gaji per petugas dulu. Kalau ada selisih, buka ringkasan atau rincian untuk melihat asal angkanya.",
        "sheets": ["Slip Gaji", "Trial Balance", "Rekap Lokasi", "Rincian Lembur", "Denda Alpa", "Banding", "Detail Absensi"],
        "groups": [
            {
                "id": "petugas",
                "label": "Petugas",
                "blurb": "Gaji per orang",
                "sheets": ["Slip Gaji", "Tidak Dihitung"],
            },
            {
                "id": "ringkasan",
                "label": "Ringkasan",
                "blurb": "Total & lokasi",
                "sheets": ["Trial Balance", "Rekap Lokasi"],
            },
            {
                "id": "rincian",
                "label": "Rincian",
                "blurb": "Lembur, alpa, banding",
                "sheets": ["Rincian Lembur", "Denda Alpa", "Banding", "Detail Absensi"],
            },
        ],
        "sheets_meta": {
            "Slip Gaji": {
                "title": "Slip gaji",
                "hint": "Gaji dihitung per petugas",
                "help": "Satu baris satu petugas. Hari Kerja = OT+L, Hari Off = kode OFF. Pakai kartu Kehadiran kurang untuk yang di bawah minimum hari hadir lokasi/petugas.",
                "terms": ["Selisih", "New Joiner", "Prorate", "Deposit", "Syarat", "Hari Kerja", "Hari Off", "Minimum hari hadir"],
            },
            "Tidak Dihitung": {
                "title": "Tidak dihitung",
                "hint": "Resign, blacklist, atau testing",
                "help": "Petugas yang sengaja tidak masuk prosedur payroll bulan ini: resign, blacklist, atau akun testing.",
                "terms": ["Resign", "Akun testing", "Tidak dihitung"],
            },
            "Trial Balance": {
                "title": "Trial balance",
                "hint": "Cocokkan total HRIS vs dihitung",
                "help": "Bandingkan total komponen dari file HRIS dengan hasil hitung. Angka biasa menambah; angka dalam kurung atau deposit mengurangi.",
                "terms": ["Trial balance", "Gaji HRIS", "Hasil hitung", "Deposit"],
            },
            "Rekap Lokasi": {
                "title": "Rekap lokasi",
                "hint": "Total gaji per lokasi",
                "help": "Jumlah gaji digroup per lokasi, untuk cek cepat sebaran pembayaran.",
                "terms": ["Hasil hitung"],
            },
            "Rincian Lembur": {
                "title": "Rincian lembur",
                "hint": "Detail fee lembur",
                "help": "Setiap baris satu sesi lembur. Cocok = Fee Dihitung sama dengan Fee HRIS. Masuk slip = Ya jika sesi ini sudah terikat ke petugas di Slip Gaji; Tidak berarti lembur tidak ketemu di payroll.",
                "terms": ["Lembur", "Cocok", "Masuk slip", "Fee HRIS", "Fee Dihitung"],
            },
            "Denda Alpa": {
                "title": "Denda alpa",
                "hint": "Potongan karena status A",
                "help": "Denda Alpa HRIS dari file. Denda Alpa Dihitung = pokok ÷ 25 × jumlah status A setelah banding; ini yang memotong gaji dihitung.",
                "terms": ["Denda alpa", "Alpa", "Banding"],
            },
            "Banding": {
                "title": "Banding",
                "hint": "Disetujui, menunggu, ditolak",
                "help": "Semua pengajuan banding dari file. Status Disetujui mengubah kode absensi sebelum denda dihitung; Menunggu/Ditolak tidak mengubah hitungan.",
                "terms": ["Banding", "Alpa"],
            },
            "Detail Absensi": {
                "title": "Detail absensi",
                "hint": "Kode harian setelah banding",
                "help": "Kode kehadiran harian setelah banding. Setiap status A mengurangi gaji lewat denda alpa dihitung.",
                "terms": ["Alpa", "Hadir", "Terlambat", "Izin", "Sakit", "Cuti"],
            },
        },
    },
    "transfer": {
        "title": "File transfer",
        "default": "Accounting",
        "blurb": "Siapkan file untuk bank. Accounting diunggah sebagai transfer; Rekap dipakai untuk cek total per remark.",
        "groups": [
            {
                "id": "transfer",
                "label": "Transfer",
                "blurb": "File untuk bank",
                "sheets": ["Accounting", "Rekap"],
            },
        ],
        "sheets_meta": {
            "Accounting": {
                "title": "Accounting",
                "hint": "Daftar transfer ke bank",
                "help": "Isi yang diunggah ke bank: rekening, nama penerima, nominal, dan remark.",
                "terms": ["Accounting", "Transfer", "Remark"],
            },
            "Rekap": {
                "title": "Rekap",
                "hint": "Total per remark",
                "help": "Jumlah transfer digroup per remark, untuk memastikan total sebelum diunggah.",
                "terms": ["Rekap", "Remark"],
            },
        },
    },
}


class Redirect(Exception):
    def __init__(self, location: str, headers: list[tuple[str, str]] | None = None):
        self.location = location
        self.headers = headers or []


class NotFound(Exception):
    pass


class Forbidden(Exception):
    pass


class Download:
    def __init__(self, path: Path, name: str):
        self.path = path
        self.name = name


class BytesResult:
    def __init__(self, data: bytes, content_type: str):
        self.data = data
        self.content_type = content_type


class Upload:
    def __init__(self, filename: str, data: bytes):
        self.filename = filename
        self.data = data

    def save(self, path: Path) -> None:
        Path(path).write_bytes(self.data)


class Form:
    def __init__(self, data: dict[str, list[str]]):
        self._data = data

    def get(self, key: str, default=None):
        values = self._data.get(key)
        if not values:
            return default
        return values[0]

    def getlist(self, key: str) -> list[str]:
        return list(self._data.get(key) or [])

    def __contains__(self, key: str) -> bool:
        return key in self._data


class Files:
    def __init__(self, data: dict[str, list[Upload]]):
        self._data = data

    def get(self, key: str) -> Upload | None:
        values = self._data.get(key) or []
        return values[0] if values else None


class Request:
    def __init__(self, args: Form, form: Form, files: Files, cookies: dict[str, str] | None = None):
        self.args = args
        self.form = form
        self.files = files
        self.cookies = cookies or {}
        self.user = None


_current: ContextVar[Request | None] = ContextVar("request", default=None)
_response_headers: ContextVar[list[tuple[str, str]] | None] = ContextVar("response_headers", default=None)


class _RequestProxy:
    def __getattr__(self, name: str):
        current = _current.get()
        if current is None:
            raise RuntimeError("request is only available while handling a page")
        return getattr(current, name)

    def __setattr__(self, name: str, value) -> None:
        current = _current.get()
        if current is None:
            raise RuntimeError("request is only available while handling a page")
        setattr(current, name, value)


request = _RequestProxy()


def redirect(location: str, headers: list[tuple[str, str]] | None = None):
    raise Redirect(location, headers)


def queue_header(name: str, value: str) -> None:
    items = _response_headers.get()
    if items is None:
        items = []
        _response_headers.set(items)
    items.append((name, value))


def set_session_cookie(token: str) -> None:
    queue_header(
        "Set-Cookie",
        f"{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age={SESSION_DAYS * 86400}",
    )


def clear_session_cookie() -> None:
    queue_header(
        "Set-Cookie",
        f"{SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0",
    )


def parse_cookies(header: str) -> dict[str, str]:
    jar = SimpleCookie()
    try:
        jar.load(header or "")
    except Exception:
        return {}
    return {key: morsel.value for key, morsel in jar.items()}


def abort(code: int):
    if code == 403:
        raise Forbidden()
    raise NotFound()


def url_for(endpoint: str, **values) -> str:
    if endpoint == "static":
        return "/static/" + quote(str(values["filename"]), safe="/")
    pattern = PATHS[endpoint]
    query = {}
    for key, value in values.items():
        token_path = f"<path:{key}>"
        token = f"<{key}>"
        if token_path in pattern:
            pattern = pattern.replace(token_path, quote(str(value), safe=""))
        elif token in pattern:
            pattern = pattern.replace(token, quote(str(value), safe=""))
        elif value is not None and value != "":
            query[key] = value
    if query:
        return pattern + "?" + urlencode(query)
    return pattern


env = Environment(
    loader=FileSystemLoader(ROOT / "templates"),
    autoescape=select_autoescape(["html"]),
    auto_reload=True,
)
env.filters["rp"] = lambda value: rupiah(float(value or 0))
env.filters["num"] = lambda value: plain_number(value)
env.globals["url_for"] = url_for
env.globals["tip"] = lambda label: tip(label)


def render_template(name: str, **context) -> str:
    if "current_user" not in context:
        context["current_user"] = getattr(request, "user", None)
    return env.get_template(name).render(**context)


_term_cache = None


def tip(label: str):
    global _term_cache
    if _term_cache is None:
        try:
            seed_terms(TERMS)
            _term_cache = term_map()
        except Exception:
            _term_cache = {}
    description = _term_cache.get(str(label).strip().lower())
    safe = html_escape(label)
    if not description:
        return Markup(str(safe))
    return Markup(
        f'<span class="term" tabindex="0">{safe}<span class="tip" role="tooltip">{html_escape(description)}</span></span>'
    )


def refresh_terms(path: Path | None = None) -> None:
    global _term_cache
    pairs = list(TERMS)
    if path is not None and path.exists():
        pairs.extend(legend_terms(path))
    seed_terms(pairs)
    _term_cache = term_map()


def sync_uploads() -> None:
    known = {row["id"] for row in list_uploads()}
    latest_source = None
    latest_mtime = 0
    if SESSIONS.exists():
        for meta_path in SESSIONS.glob("*/meta.json"):
            source = meta_path.parent / "source.xlsx"
            if source.exists() and source.stat().st_mtime >= latest_mtime:
                latest_mtime = source.stat().st_mtime
                latest_source = source
            sid = meta_path.parent.name
            if sid in known:
                continue
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
            alias = meta.get("source_name") or "payroll.xlsx"
            uploaded = datetime.fromtimestamp(meta_path.stat().st_mtime)
            remember_upload(sid, alias, "", "", alias, uploaded.isoformat())
    with connect() as db:
        month_count = db.execute("SELECT COUNT(*) AS n FROM months").fetchone()["n"]
    if month_count == 0 and latest_source is not None:
        meta = json.loads((latest_source.parent / "meta.json").read_text(encoding="utf-8"))
        payroll = build(latest_source)
        payroll.source_name = meta.get("source_name") or payroll.source_name
        save_month(latest_source, payroll)
        start, end = split_period(payroll.period)
        remember_upload(
            latest_source.parent.name,
            payroll.source_name,
            start,
            end,
            payroll.source_name,
            datetime.fromtimestamp((latest_source.parent / "meta.json").stat().st_mtime).isoformat(),
        )
    if latest_source is not None and len(all_terms()) <= len(TERMS):
        refresh_terms(latest_source)


def _disposition_param(header: str, name: str) -> str | None:
    star = re.search(rf"(?:^|;)\s*{name}\*=UTF-8''([^;]+)", header, re.I)
    if star:
        return unquote(star.group(1))
    quoted = re.search(rf'(?:^|;)\s*{name}="((?:\\.|[^"])*)"', header, re.I)
    if quoted:
        return quoted.group(1).replace('\\"', '"')
    plain = re.search(rf"(?:^|;)\s*{name}=([^;]+)", header, re.I)
    if plain:
        return plain.group(1).strip().strip('"')
    return None


def parse_multipart(content_type: str, body: bytes) -> tuple[Form, Files]:
    match = re.search(r'boundary="?([^";]+)"?', content_type, re.I)
    if not match:
        return Form({}), Files({})
    boundary = match.group(1).encode("utf-8")
    form: dict[str, list[str]] = {}
    files: dict[str, list[Upload]] = {}
    for chunk in body.split(b"--" + boundary)[1:]:
        if chunk.startswith(b"--"):
            break
        if chunk.startswith(b"\r\n"):
            chunk = chunk[2:]
        elif chunk.startswith(b"\n"):
            chunk = chunk[1:]
        if chunk.endswith(b"\r\n"):
            chunk = chunk[:-2]
        elif chunk.endswith(b"\n"):
            chunk = chunk[:-1]
        header_blob, sep, data = chunk.partition(b"\r\n\r\n")
        if not sep:
            header_blob, sep, data = chunk.partition(b"\n\n")
        if not sep:
            continue
        disposition = ""
        for line in header_blob.decode("utf-8", "replace").splitlines():
            if line.lower().startswith("content-disposition:"):
                disposition = line.split(":", 1)[1].strip()
                break
        if "form-data" not in disposition.lower():
            continue
        name = _disposition_param(disposition, "name")
        if not name:
            continue
        filename = _disposition_param(disposition, "filename")
        if filename:
            files.setdefault(name, []).append(Upload(Path(filename).name, data))
        elif filename is None:
            form.setdefault(name, []).append(data.decode("utf-8", "replace"))
    return Form(form), Files(files)


def query_form(raw: str) -> Form:
    parsed = parse_qs(raw, keep_blank_values=True)
    return Form({key: values for key, values in parsed.items()})


def parse_request(method: str, query: str, content_type: str, body: bytes, cookie_header: str = "") -> Request:
    args = query_form(query)
    cookies = parse_cookies(cookie_header)
    if method != "POST" or not body:
        return Request(args, Form({}), Files({}), cookies)
    kind = (content_type or "").split(";", 1)[0].strip().lower()
    if kind == "application/x-www-form-urlencoded":
        return Request(args, query_form(body.decode("utf-8", "replace")), Files({}), cookies)
    if kind == "multipart/form-data":
        form, files = parse_multipart(content_type, body)
        return Request(args, form, files, cookies)
    return Request(args, Form({}), Files({}), cookies)


PUBLIC_PATHS = {"/login", "/logout"}


def _actor() -> dict | None:
    try:
        user = request.user
    except RuntimeError:
        return None
    return user if user else None


def add_activity(kind: str, title: str, detail: str = "", upload_id: str | None = None) -> int:
    return _add_activity(kind, title, detail, upload_id, actor=_actor())


def update_activity(activity_id: int, title: str, detail: str) -> None:
    _update_activity(activity_id, title, detail, actor=_actor())


def login_page():
    if request.user:
        redirect(url_for("home"))
    return render_template(
        "login.html",
        salah=request.args.get("salah") or "",
        next_url=request.args.get("next") or "/",
    )


def save_login():
    if request.user:
        redirect(url_for("home"))
    username = request.form.get("username") or ""
    password = request.form.get("password") or ""
    next_url = request.form.get("next") or "/"
    if not next_url.startswith("/") or next_url.startswith("//"):
        next_url = "/"
    user = authenticate(username, password)
    if user is None:
        return render_template(
            "login.html",
            salah="Username atau password salah.",
            next_url=next_url,
            username=username,
        )
    token = create_session(user["id"])
    set_session_cookie(token)
    redirect(next_url)


def logout():
    destroy_session(request.cookies.get(SESSION_COOKIE))
    clear_session_cookie()
    redirect(url_for("login"))


def users_page():
    if not request.user or not request.user.get("is_admin"):
        abort(403)
    return render_template(
        "users.html",
        users=list_users(),
        roles=ROLES,
        editing=get_user(int(request.args.get("edit"))) if (request.args.get("edit") or "").isdigit() else None,
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
        nav="pengaturan",
    )


def save_users():
    if not request.user or not request.user.get("is_admin"):
        abort(403)
    action = request.form.get("aksi") or "buat"
    if action == "hapus":
        salah = delete_app_user(int(request.form.get("id") or 0), request.user["id"])
    elif action == "ubah":
        salah = update_app_user(int(request.form.get("id") or 0), request.form, request.user["id"])
    else:
        salah = create_app_user(
            request.form.get("nama") or "",
            request.form.get("username") or "",
            request.form.get("password") or "",
            request.form.get("role") or "staff",
        )
    redirect(url_for("users", salah=salah, simpan="" if salah else "1", edit=request.form.get("id") if salah and action == "ubah" else ""))


def content_disposition(name: str) -> str:
    cleaned = name.replace('"', "").replace("\r", "").replace("\n", "") or "download.xlsx"
    ascii_name = cleaned.encode("ascii", "ignore").decode() or "download.xlsx"
    return f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{quote(cleaned)}'


def session_dir(sid: str) -> Path:
    if not SESSION_ID.match(sid or ""):
        abort(404)
    folder = (SESSIONS / sid).resolve()
    if folder.parent != SESSIONS.resolve() or not folder.is_dir():
        abort(404)
    return folder


def open_session(folder: Path) -> dict:
    meta = json.loads((folder / "meta.json").read_text(encoding="utf-8"))
    payroll = build(folder / "source.xlsx")
    payroll.source_name = meta.get("source_name") or payroll.source_name
    prior_rows = None
    prior_path = folder / "prior.xlsx"
    prior_error = meta.get("prior_error") or ""
    if prior_path.exists() and not prior_error:
        try:
            prior_rows = load_prior_rows(prior_path)
        except ValueError as exc:
            prior_error = str(exc)
    merge_payroll(payroll, folder / "source.xlsx")
    view = review(payroll, prior_rows, meta.get("prior_name") or "", load_roster())
    view["prior_error"] = prior_error
    view["sid"] = folder.name
    view["archive"] = save_month(folder / "source.xlsx", payroll)
    view["upload"] = get_upload(folder.name)
    return view


def home():
    sync_uploads()
    uploads = list_uploads()
    if uploads and request.args.get("unggah") != "1":
        return redirect(url_for("view_hitung", sid=uploads[0]["id"]))
    months = list_months()
    return render_template(
        "upload.html",
        error=request.args.get("error", ""),
        uploads=uploads[:5],
        upload_count=len(uploads),
        latest=uploads[0] if uploads else None,
        month=months[0] if months else None,
    )


def kegiatan():
    upload_id = request.args.get("unggahan") or ""
    saved = request.args.get("saved") == "1"
    return render_template(
        "kegiatan.html",
        activities=list_activities(upload_id),
        uploads=list_uploads(),
        upload_id=upload_id,
        editing=get_activity(int(request.args.get("edit"))) if request.args.get("edit", "").isdigit() else None,
        saved=saved,
    )


def save_kegiatan():
    activity_id = (request.form.get("id") or "").strip()
    title = (request.form.get("title") or "").strip()
    detail = (request.form.get("detail") or "").strip()
    upload_id = (request.form.get("upload_id") or "").strip() or None
    if not title:
        return redirect(url_for("kegiatan", error="Isi judul kegiatan."))
    if activity_id.isdigit():
        update_activity(int(activity_id), title, detail)
    else:
        add_activity("catatan", title, detail, upload_id)
    return redirect(url_for("kegiatan", unggahan=upload_id or "", saved=1))


def hapus_kegiatan(activity_id: str):
    delete_activity(int(activity_id))
    return redirect(url_for("kegiatan", saved=1))


def simpan_alias(sid: str):
    alias = (request.form.get("alias") or "").strip()
    if alias:
        rename_upload(sid, alias)
        add_activity("alias", "Nama unggahan diubah", alias, sid)
    return redirect(url_for("kegiatan", saved=1))


def hapus_unggahan(sid: str):
    folder = SESSIONS / sid
    delete_upload(sid)
    if folder.is_dir() and folder.resolve().parent == SESSIONS.resolve():
        import shutil

        shutil.rmtree(folder)
    return redirect(url_for("kegiatan", saved=1))


def istilah():
    refresh_terms()
    return render_template("istilah.html", terms=all_terms())


def arsip():
    return render_template("arsip.html", months=list_months())


def arsip_bulan(month_id: str):
    page = month_page(int(month_id))
    if page is None:
        abort(404)
    return render_template("arsip_bulan.html", page=page)


def banding_unggah():
    sid = (request.args.get("unggah") or "").strip()
    dasar = (request.args.get("dasar") or "").strip()
    focus = (request.args.get("fokus") or "perhatian").strip() or "perhatian"
    tab = (request.args.get("tab") or "petugas").strip() or "petugas"
    baseline_kind = ""
    baseline_id = ""
    if ":" in dasar:
        baseline_kind, baseline_id = dasar.split(":", 1)
    page = compare_page(
        sid=sid,
        baseline_kind=baseline_kind,
        baseline_id=baseline_id,
        focus=focus,
        tab=tab,
    )
    sync = {}
    if page.get("upload"):
        meta_path = SESSIONS / page["upload"]["id"] / "meta.json"
        if meta_path.exists():
            try:
                sync = json.loads(meta_path.read_text(encoding="utf-8")).get("profile_sync") or {}
            except (json.JSONDecodeError, OSError):
                sync = {}
    page["profile_sync"] = sync
    return render_template("banding_unggah.html", page=page)


def analyze():
    upload = request.files.get("source")
    if upload is None or not upload.filename:
        return redirect(url_for("home", error="Pilih file tarikan HRIS terlebih dahulu."))
    if not upload.filename.lower().endswith(".xlsx"):
        return redirect(url_for("home", error="File tarikan harus .xlsx dari HRIS."))

    sid = uuid.uuid4().hex
    folder = SESSIONS / sid
    folder.mkdir(parents=True, exist_ok=True)
    source_path = folder / "source.xlsx"
    upload.save(source_path)

    meta = {"source_name": Path(upload.filename).name, "prior_name": "", "prior_error": ""}
    prior = request.files.get("prior")
    if prior and prior.filename:
        if not prior.filename.lower().endswith(".xlsx"):
            meta["prior_error"] = "File bulan lalu harus .xlsx."
        else:
            prior.save(folder / "prior.xlsx")
            meta["prior_name"] = Path(prior.filename).name
            try:
                load_prior_rows(folder / "prior.xlsx")
            except ValueError as exc:
                meta["prior_error"] = str(exc)

    try:
        payroll = build(source_path)
    except SystemExit as exc:
        return redirect(url_for("home", error=str(exc) or "File tidak bisa dibaca sebagai tarikan payroll."))

    payroll.source_name = meta["source_name"]
    merge_payroll(payroll, source_path)
    save_month(source_path, payroll)
    start, end = split_period(payroll.period)
    remember_upload(sid, meta["source_name"], start, end, meta["source_name"])
    sync = apply_upload_profiles(
        source_path,
        upload_id=sid,
        source_name=meta["source_name"],
        period=payroll.period or (f"{start} s/d {end}" if start and end else ""),
    )
    meta["profile_sync"] = {
        "updated": sync.get("updated") or 0,
        "fields": sync.get("fields") or 0,
        "error": sync.get("error") or "",
    }
    add_activity("unggah", f"Mengunggah {meta['source_name']}", payroll.period or "Periode mengikuti file", sid)
    if sync.get("updated"):
        add_activity(
            "profil",
            f"Memperbarui {sync['updated']} profil petugas dari unggahan",
            f"{sync.get('fields') or 0} field berubah · {payroll.period or meta['source_name']}",
            sid,
        )
    refresh_terms(source_path)
    write_eligible_book(folder, payroll)
    (folder / "meta.json").write_text(json.dumps(meta), encoding="utf-8")
    return redirect(url_for("banding_unggah", unggah=sid, fokus="perhatian"))


def latest_source() -> Path | None:
    found = sorted(SESSIONS.glob("*/source.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
    return found[0] if found else None


def file_deposit_for_username(username: str) -> int:
    """Deposit amount already deducted in the latest payroll source file."""
    key = (username or "").strip().lower()
    source = latest_source()
    if not key or source is None or not source.exists():
        return 0
    try:
        from openpyxl import load_workbook

        wb = load_workbook(source, read_only=True, data_only=True)
        if "Payroll" not in wb.sheetnames:
            wb.close()
            return 0
        ws = wb["Payroll"]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            wb.close()
            return 0
        names = [text(cell) for cell in header]
        if "Username" not in names or "Deposit" not in names:
            wb.close()
            return 0
        user_at = names.index("Username")
        deposit_at = names.index("Deposit")
        for row in rows:
            if not row or user_at >= len(row):
                continue
            if text(row[user_at]).lower() != key:
                continue
            raw = row[deposit_at] if deposit_at < len(row) else 0
            wb.close()
            try:
                return abs(int(round(float(raw or 0))))
            except (TypeError, ValueError):
                return 0
        wb.close()
    except Exception:
        return 0
    return 0


def ensure_roster() -> None:
    roster = load_roster()
    if roster["people"]:
        sync_flags(roster)
        return
    source = latest_source()
    if source is None:
        return
    try:
        payroll = build(source)
        meta_path = source.parent / "meta.json"
        if meta_path.exists():
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
            payroll.source_name = meta.get("source_name") or payroll.source_name
        merge_payroll(payroll, source)
    except SystemExit:
        pass


def posted_people() -> list[dict]:
    keys = request.form.getlist("pkey")
    statuses = request.form.getlist("pstatus")
    dates = request.form.getlist("pdate")
    notes = request.form.getlist("pnote")
    origins = request.form.getlist("pfrom")
    people = []
    for index, key in enumerate(keys):
        people.append(
            {
                "key": key,
                "status": statuses[index] if index < len(statuses) else "aktif",
                "tanggal": dates[index] if index < len(dates) else "",
                "catatan": notes[index] if index < len(notes) else "",
                "mutasi_dari": origins[index] if index < len(origins) else "",
            }
        )
    return people


def posted_flags():
    if request.form.get("flags") != "1":
        return None, None
    return set(request.form.getlist("ptesting")), set(request.form.getlist("pmutasi"))


def kelola():
    ensure_roster()
    saved = request.args.get("saved") == "1"
    page = page_rows()
    page["rekening"] = coverage()
    return render_template("kelola.html", page=page, saved=saved)


def kelola_petugas():
    ensure_roster()
    saved = request.args.get("saved") == "1"
    page = people_page()
    attach_banks(page["people"])
    page["rekening"] = coverage()
    return render_template(
        "kelola_petugas.html",
        page=page,
        saved=saved,
        status_filter=request.args.get("status") or "",
        lokasi_filter=request.args.get("lokasi") or "",
    )


def save_petugas():
    testing_keys, mutasi_keys = posted_flags()
    apply_choices([], set(), posted_people(), testing_keys, mutasi_keys)
    add_activity("petugas", "Status petugas disimpan", "Resign, petugas baru, atau tanggal diperbarui.")
    return redirect(url_for("kelola_petugas", saved=1))


def kelola_lokasi(lokasi: str):
    ensure_roster()
    detail = location_detail(lokasi)
    if detail is None:
        abort(404)
    attach_counts(detail["people"], latest_source())
    attach_banks(detail["people"])
    detail["kode"] = location_kode(detail["name"])
    saved = request.args.get("saved") == "1"
    return render_template("kelola_lokasi.html", page=detail, saved=saved)


def kelola_absensi(lokasi: str, username: str):
    source = latest_source()
    if source is None:
        abort(404)
    page = officer_attendance(source, lokasi, username)
    if page is None:
        abort(404)
    return render_template("kelola_absensi.html", page=page)


def roster_person(lokasi: str, username: str) -> dict | None:
    wanted_user = username.strip().lower()
    wanted_loc = lokasi.strip().lower()
    for row in load_roster().get("people", {}).values():
        if (row.get("username") or "").lower() == wanted_user and (row.get("lokasi") or "").lower() == wanted_loc:
            return row
    return None


def kelola_profil(lokasi: str, username: str):
    ensure_roster()
    person = roster_person(lokasi, username)
    if person is None:
        abort(404)
    profile = find_profile(person.get("lokasi") or lokasi, person.get("username") or username, person.get("nama") or "") or {}
    meta = load_roster().get("meta") or {}
    saved = request.args.get("saved") == "1"
    return render_template(
        "kelola_profil.html",
        person=person,
        profile=profile,
        kode=location_kode(person.get("lokasi") or lokasi),
        period=remark_period(meta.get("period") or "", meta.get("month") or ""),
        history=person_history(person.get("lokasi") or lokasi, person.get("username") or username),
        saved=saved,
    )


def save_profil(lokasi: str, username: str):
    person = roster_person(lokasi, username)
    if person is None:
        abort(404)
    if request.form.get("bagian") != "rekening":
        set_employment(
            person.get("lokasi") or lokasi,
            person.get("username") or username,
            {
                "status": request.form.get("status") or person.get("status") or "aktif",
                "tanggal": request.form.get("tanggal") or "",
                "testing": request.form.get("testing") == "1",
                "mutasi": request.form.get("mutasi") == "1",
                "mutasi_dari": request.form.get("mutasi_dari") or "",
            },
        )
    if request.form.get("bagian") == "status":
        add_activity(
            "profil",
            f"Status {person.get('nama') or username} disimpan",
            person.get("lokasi") or lokasi,
        )
        return redirect(url_for("kelola_profil", lokasi=person.get("lokasi") or lokasi, username=username, saved=1))
    set_profile(
        person.get("lokasi") or lokasi,
        person.get("username") or username,
        person.get("nama") or "",
        {
            "nama_rekening": request.form.get("nama_rekening"),
            "bank": request.form.get("bank"),
            "rekening": request.form.get("rekening"),
            "catatan": request.form.get("catatan"),
        },
    )
    add_activity(
        "profil",
        f"Rekening {person.get('nama') or username} disimpan",
        person.get("lokasi") or lokasi,
    )
    return redirect(url_for("kelola_profil", lokasi=person.get("lokasi") or lokasi, username=username, saved=1))


def save_kelola():
    name = (request.form.get("location") or "").strip()
    setting = request.form.get("set_active")
    if name and setting in {"0", "1"}:
        apply_choices([name], {name} if setting == "1" else set(), [])
        add_activity("lokasi", f"Lokasi {name} {'diaktifkan' if setting == '1' else 'dimatikan'}", name)
        if request.form.get("back") == "detail":
            return redirect(url_for("kelola_lokasi", lokasi=name, saved=1))
        return redirect(url_for("kelola", saved=1))

    lokasi = (request.form.get("lokasi") or "").strip()
    if lokasi and "kode" in request.form:
        set_location_kode(lokasi, request.form.get("kode") or "")
    testing_keys, mutasi_keys = posted_flags()
    apply_choices([], set(), posted_people(), testing_keys, mutasi_keys)
    add_activity("petugas", "Perubahan petugas disimpan", lokasi or "Semua lokasi")
    if lokasi:
        return redirect(url_for("kelola_lokasi", lokasi=lokasi, saved=1))
    return redirect(url_for("kelola", saved=1))


def report(sid: str):
    view = open_session(session_dir(sid))
    return render_template("report.html", view=view)


def sheet_slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "lembar"


def plain_number(value) -> str:
    number = float(value or 0)
    if abs(number - round(number)) > 0.001:
        whole, frac = f"{abs(number):.2f}".rstrip("0").rstrip(".").split(".")
        shown = f"{int(whole):,}".replace(",", ".") + "," + frac
    else:
        shown = f"{int(round(abs(number))):,}".replace(",", ".")
    if number < 0 and shown != "0":
        return "-" + shown
    return shown


def load_payroll(folder: Path):
    meta_path = folder / "meta.json"
    meta = json.loads(meta_path.read_text(encoding="utf-8")) if meta_path.exists() else {}
    payroll = build(folder / "source.xlsx")
    payroll.source_name = meta.get("source_name") or payroll.source_name
    return payroll, meta


def format_number(value: float, fmt: str) -> str:
    money = "#,##0" in (fmt or "")
    number = abs(float(value))
    if not money and abs(number - round(number)) > 0.001:
        whole, frac = f"{number:.2f}".rstrip("0").rstrip(".").split(".")
        text = f"{int(whole):,}".replace(",", ".") + "," + frac
    else:
        rounded = int(round(number))
        text = f"{rounded:,}".replace(",", ".") if money or rounded >= 1000 else str(rounded)
    if value < 0 and text != "0":
        return "-" + text
    return text


def present_cell(cell, selisih: bool, paren: bool = False) -> dict:
    value = cell.value
    classes: list[str] = []
    if value is None or value == "":
        return {"text": "", "cls": "", "digits": "", "kind": "text", "sort": ""}
    if isinstance(value, bool):
        return {"text": "Ya" if value else "Tidak", "cls": "tag-ok" if value else "tag-bad", "digits": "", "kind": "text", "sort": ""}
    if isinstance(value, datetime):
        text = value.strftime("%d/%m/%Y %H:%M")
        if text.endswith(" 00:00"):
            text = text[:-6]
        return {"text": text, "cls": "", "digits": "", "kind": "date", "sort": value.strftime("%Y%m%d%H%M")}
    if isinstance(value, date):
        return {"text": value.strftime("%d/%m/%Y"), "cls": "", "digits": "", "kind": "date", "sort": value.strftime("%Y%m%d")}
    if isinstance(value, (int, float)):
        if paren and value < 0:
            text = "(" + format_number(abs(float(value)), cell.number_format or "") + ")"
            classes.extend(["num", "minus"])
        else:
            text = format_number(float(value), cell.number_format or "")
            classes.append("num")
            if selisih and value > 1:
                classes.append("pos")
            elif selisih and value < -1:
                classes.append("neg")
            elif value < 0:
                classes.append("minus")
        digits = re.sub(r"\D", "", text)
        return {"text": text, "cls": " ".join(classes), "digits": digits, "kind": "num", "sort": f"{float(value):.6f}"}
    text = str(value).strip()
    low = text.lower()
    if low in {"cocok", "ya", "ok"}:
        classes.append("tag-ok")
    elif low in {"beda", "tidak"}:
        classes.append("tag-bad")
    elif any(word in low for word in ("kosong", "beda", "belum", "berbeda", "memotong")):
        classes.append("tag-warn")
    if len(text) > 48:
        classes.append("wrap")
    return {"text": text, "cls": " ".join(classes), "digits": "", "kind": "text", "sort": ""}


def headerish(cells) -> bool:
    texts = []
    for cell in cells:
        value = cell.value
        if isinstance(value, str) and value.strip():
            texts.append(value.strip())
        elif value not in (None, ""):
            return False
    return len(texts) >= 2 and all(len(text) <= 40 for text in texts)


def read_workbook(path: Path) -> list[dict]:
    workbook = load_workbook(path, data_only=True)
    sheets = []
    for ws in workbook.worksheets:
        matrix = [list(row) for row in ws.iter_rows(max_row=ws.max_row, max_col=ws.max_column)]
        if not matrix:
            continue
        keep = [
            col
            for col in range(ws.max_column)
            if any(matrix[row][col].value not in (None, "") for row in range(len(matrix)))
        ]
        if not keep:
            continue
        rows = [[matrix[row][col] for col in keep] for row in range(len(matrix))]
        header_at = 0
        for index, cells in enumerate(rows[:8]):
            if headerish(cells):
                header_at = index
                break
        intro = []
        for cells in rows[:header_at]:
            parts = [present_cell(cell, False)["text"] for cell in cells]
            line = " ".join(part for part in parts if part).strip()
            if line:
                intro.append(line)
        headers = [present_cell(cell, False)["text"] for cell in rows[header_at]]
        selisih = {index for index, header in enumerate(headers) if "selisih" in header.lower()}
        body = []
        for cells in rows[header_at + 1 :]:
            label = ""
            if cells and isinstance(cells[0].value, str):
                label = cells[0].value.lower()
            shown = []
            for index, cell in enumerate(cells):
                header = headers[index].lower() if index < len(headers) else ""
                shown.append(present_cell(cell, index in selisih, paren="bpjs" in header or "bpjs" in label or "deposit" in header or "deposit" in label))
            if not any(item["text"] for item in shown):
                continue
            bits = []
            for item in shown:
                if item["text"]:
                    bits.append(item["text"].lower())
                if item["digits"]:
                    bits.append(item["digits"])
            body.append(
                {
                    "cells": shown,
                    "search": " ".join(bits),
                    "total": any(item["text"].strip().lower() in {"total", "grand total"} for item in shown),
                }
            )
        numeric = []
        for index in range(len(headers)):
            sample = [
                row["cells"][index]["kind"]
                for row in body[:24]
                if index < len(row["cells"]) and not row["total"]
            ]
            numeric.append(bool(sample) and sample.count("num") + sample.count("date") >= max(1, len(sample) // 2))
        sheets.append(
            {
                "name": ws.title,
                "id": sheet_slug(ws.title),
                "intro": intro,
                "headers": headers,
                "rows": body,
                "numeric": numeric,
                "freeze": 2 if headers[:2] == ["Lokasi", "Nama"] else 1,
            }
        )
    return sheets


def _eligible_stamp(blocked: set[str]) -> str:
    conditions = payroll_conditions_index()
    bits = [f"x:{name}" for name in sorted(blocked)]
    for name in sorted(conditions):
        row = conditions[name]
        bits.append(
            "d:{name}:{deposit}:{potongan}:{preason}:{tambahan}:{treason}:{excluded}:{note}".format(
                name=name,
                deposit=row.get("deposit"),
                potongan=row.get("potongan"),
                preason=row.get("potongan_reason") or "",
                tambahan=row.get("tambahan"),
                treason=row.get("tambahan_reason") or "",
                excluded=int(bool(row.get("excluded"))),
                note=row.get("note") or "",
            )
        )
    return "|".join(bits)


def write_eligible_book(folder: Path, payroll=None) -> Path:
    """Write hitung.xlsx with only people allowed in this payroll run."""
    target = folder / "hitung.xlsx"
    blocked = payroll_blocked_usernames()
    source = payroll if payroll is not None else build(folder / "source.xlsx")
    apply_payroll_conditions(source)
    write_workbook(filter_payroll(source, blocked), target)
    (folder / "eligible.stamp").write_text(_eligible_stamp(blocked), encoding="utf-8")
    return target


def prepare_book(folder: Path, book: str) -> Path:
    if book == "hitung":
        target = folder / "hitung.xlsx"
        script = Path(__file__).with_name("calculate_payroll.py")
        stamp_path = folder / "eligible.stamp"
        blocked = payroll_blocked_usernames()
        stamp = _eligible_stamp(blocked)
        previous = stamp_path.read_text(encoding="utf-8") if stamp_path.exists() else ""
        stale = (
            not target.exists()
            or target.stat().st_mtime < script.stat().st_mtime
            or previous != stamp
        )
        if stale:
            write_eligible_book(folder)
        return target
    payroll, _meta = load_payroll(folder)
    target = folder / "transfer.xlsx"
    write_transfer(payroll, target, edits=load_closing(folder))
    return target


def _paint_total(cell: dict, amount: float, header: str, fractional: bool) -> None:
    paren = "bpjs" in header or "deposit" in header
    selisih = "selisih" in header
    text = format_number(amount, "" if fractional else "#,##0")
    classes = ["num"]
    if paren and amount < 0:
        text = "(" + format_number(abs(amount), "#,##0") + ")"
        classes.append("minus")
    elif selisih and amount > 1:
        classes.append("pos")
    elif selisih and amount < -1:
        classes.append("neg")
    elif amount < 0:
        classes.append("minus")
    cell["text"] = text
    cell["cls"] = " ".join(classes)
    cell["kind"] = "num"
    cell["sort"] = f"{float(amount):.6f}"
    cell["digits"] = re.sub(r"\D", "", text)


def _retotal(sheet: dict, rows: list[dict], total: dict) -> None:
    for index, cell in enumerate(total["cells"]):
        if cell.get("kind") != "num":
            continue
        samples = []
        amount = 0.0
        for row in rows:
            if index >= len(row["cells"]):
                continue
            item = row["cells"][index]
            if item.get("kind") != "num" or not item.get("sort"):
                continue
            samples.append(item["sort"])
            amount += float(item["sort"])
        header = sheet["headers"][index].lower() if index < len(sheet["headers"]) else ""
        fractional = any(abs(float(sample) - round(float(sample))) > 0.001 for sample in samples)
        _paint_total(cell, amount, header, fractional)
    bits = []
    for item in total["cells"]:
        if item["text"]:
            bits.append(item["text"].lower())
        if item.get("digits"):
            bits.append(item["digits"])
    total["search"] = " ".join(bits)


def _sheet_amount(row: dict, headers: list[str], name: str):
    if name not in headers:
        return None
    at = headers.index(name)
    if at >= len(row["cells"]):
        return None
    raw = row["cells"][at].get("sort") or ""
    if raw == "":
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _near_zero(value: float | None, limit: float = 1.0) -> bool:
    if value is None:
        return True
    try:
        return abs(float(value)) < limit
    except (TypeError, ValueError):
        return True


def _banding_usernames(sheets) -> set[str]:
    """Usernames that appear on the Banding sheet for this payroll book."""
    found: set[str] = set()
    for sheet in sheets:
        if sheet.get("name") != "Banding":
            continue
        headers = sheet.get("headers") or []
        if "Username" not in headers:
            continue
        at = headers.index("Username")
        for row in sheet.get("rows") or []:
            if row.get("total") or at >= len(row.get("cells") or []):
                continue
            username = (row["cells"][at].get("text") or "").strip().lower()
            if username:
                found.add(username)
    return found


def _mark_payroll_people(sheets) -> None:
    index = officer_status_index()
    min_days = min_work_days_index()
    blocked = payroll_blocked_usernames()
    conditions = payroll_conditions_index()
    banding_users = _banding_usernames(sheets)
    for sheet in sheets:
        headers = sheet["headers"]
        # Banding has Username but is not a payroll-person sheet
        sheet["person"] = "Username" in headers and sheet.get("name") != "Banding"
        sheet["excluded"] = False
        if not sheet["person"]:
            continue
        at = headers.index("Username")
        lokasi_at = headers.index("Lokasi") if "Lokasi" in headers else -1
        deposit_at = headers.index("Deposit") if "Deposit" in headers else -1
        hadir_at = headers.index("Hari Kerja") if "Hari Kerja" in headers else -1
        kept = []
        total = None
        for row in sheet["rows"]:
            row["officer_id"] = None
            row["status_label"] = ""
            row["account_code"] = ""
            row["file_deposit"] = ""
            row["has_condition"] = False
            row["has_potongan"] = False
            row["has_tambahan"] = False
            row["has_rincian"] = False
            row["has_selisih"] = False
            row["new_joiner"] = False
            row["normal_case"] = False
            row["filter_lokasi"] = ""
            row["filter_gaji_pokok"] = ""
            row["filter_gaji_hitung"] = ""
            row["filter_deposit"] = ""
            row["filter_selisih"] = ""
            row["filter_alpa"] = ""
            row["filter_hadir"] = ""
            row["filter_min_hadir"] = ""
            row["filter_hadir_ok"] = ""
            row["hadir_kurang"] = False
            row["filter_new_joiner"] = ""
            row["filter_normal"] = "0"
            if row["total"]:
                total = row
                continue
            if at >= len(row["cells"]):
                kept.append(row)
                continue
            username = row["cells"][at]["text"].strip().lower()
            if username in blocked:
                continue
            match = index.get(username)
            if match:
                row["officer_id"] = match["id"]
                row["status_label"] = match["label"]
                row["account_code"] = match.get("code") or ""
            if lokasi_at >= 0 and lokasi_at < len(row["cells"]):
                row["filter_lokasi"] = row["cells"][lokasi_at]["text"].strip()
            if deposit_at >= 0 and deposit_at < len(row["cells"]):
                raw = row["cells"][deposit_at].get("sort") or ""
                row["file_deposit"] = str(abs(int(round(float(raw))))) if raw else "0"
            gaji_pokok = _sheet_amount(row, headers, "Gaji Pokok")
            gaji_hitung = _sheet_amount(row, headers, "Gaji Dihitung")
            deposit = _sheet_amount(row, headers, "Deposit")
            selisih = _sheet_amount(row, headers, "Selisih")
            alpa = _sheet_amount(row, headers, "Status Alpa")
            hadir = _sheet_amount(row, headers, "Hari Kerja")
            if hadir is None and hadir_at >= 0 and hadir_at < len(row["cells"]):
                raw_hadir = row["cells"][hadir_at].get("sort") or row["cells"][hadir_at].get("text") or ""
                try:
                    hadir = float(raw_hadir) if str(raw_hadir).strip() != "" else None
                except ValueError:
                    hadir = None
            lembur_hris = _sheet_amount(row, headers, "Lembur HRIS")
            lembur_hitung = _sheet_amount(row, headers, "Lembur Dihitung")
            jam_lembur = _sheet_amount(row, headers, "Jam Lembur")
            denda_hitung = _sheet_amount(row, headers, "Denda Alpa Dihitung")
            denda_hris = _sheet_amount(row, headers, "Denda Alpa HRIS")
            potongan_manual = _sheet_amount(row, headers, "Potongan Manual")
            potongan_lain = _sheet_amount(row, headers, "Potongan Lainnya")
            tambahan_manual = _sheet_amount(row, headers, "Tambahan Manual")
            if gaji_pokok is not None:
                row["filter_gaji_pokok"] = f"{gaji_pokok:.6f}"
            if gaji_hitung is not None:
                row["filter_gaji_hitung"] = f"{gaji_hitung:.6f}"
            if deposit is not None:
                row["filter_deposit"] = f"{abs(deposit):.6f}"
            if selisih is not None:
                row["filter_selisih"] = f"{selisih:.6f}"
            if alpa is not None:
                row["filter_alpa"] = f"{alpa:.6f}"
            if hadir is not None:
                row["filter_hadir"] = f"{hadir:.6f}"
            min_info = min_days.get(username) or {}
            # Profil petugas/lokasi; kalau belum diatur, pakai 10 hari (sama batas Cek unggahan)
            min_hadir = min_info.get("min_work_days")
            if min_hadir is None:
                min_hadir = 10
                row["filter_min_hadir"] = "10"
                row["min_hadir_default"] = True
            else:
                row["filter_min_hadir"] = str(int(min_hadir))
                row["min_hadir_default"] = False
            days_hadir = int(hadir or 0)
            row["hadir_kurang"] = days_hadir < int(min_hadir)
            row["filter_hadir_ok"] = "0" if row["hadir_kurang"] else "1"
            cond = conditions.get(username) or {}
            row["has_potongan"] = cond.get("potongan") is not None or cond.get("deposit") is not None
            row["has_tambahan"] = cond.get("tambahan") is not None
            row["has_condition"] = bool(
                cond.get("excluded")
                or row["has_potongan"]
                or row["has_tambahan"]
                or cond.get("note")
            )
            row["has_rincian"] = bool(
                row["has_potongan"]
                or row["has_tambahan"]
                or (row.get("file_deposit") and row.get("file_deposit") != "0")
            )
            try:
                row["has_selisih"] = abs(float(row.get("filter_selisih") or 0)) > 1
            except ValueError:
                row["has_selisih"] = False
            joiner_at = headers.index("New Joiner") if "New Joiner" in headers else -1
            if joiner_at >= 0 and joiner_at < len(row["cells"]):
                flag = (row["cells"][joiner_at]["text"] or "").strip().lower()
                row["new_joiner"] = flag in ("ya", "yes", "1", "true")
                row["filter_new_joiner"] = "1" if row["new_joiner"] else "0"
            has_banding = username in banding_users
            has_overtime = not (
                _near_zero(lembur_hitung)
                and _near_zero(lembur_hris)
                and _near_zero(jam_lembur)
            )
            has_deduction = not (
                _near_zero(deposit)
                and (not row.get("file_deposit") or row.get("file_deposit") == "0")
                and _near_zero(alpa)
                and _near_zero(denda_hitung)
                and _near_zero(denda_hris)
                and _near_zero(potongan_manual)
                and _near_zero(potongan_lain)
                and not row["has_potongan"]
            )
            has_extra_pay = not _near_zero(tambahan_manual) or row["has_tambahan"]
            row["normal_case"] = bool(
                not row["has_selisih"]
                and not has_overtime
                and not has_banding
                and not has_deduction
                and not has_extra_pay
                and not row["new_joiner"]
                and not row["has_condition"]
                and not row["hadir_kurang"]
            )
            row["filter_normal"] = "1" if row["normal_case"] else "0"
            kept.append(row)
        if total:
            _retotal(sheet, kept, total)
            sheet["rows"] = kept + [total]
        else:
            sheet["rows"] = kept


def _payroll_filter_options(sheets) -> dict:
    locations = set()
    statuses = set()
    for sheet in sheets:
        if not sheet.get("person"):
            continue
        for row in sheet["rows"]:
            if row.get("total"):
                continue
            if row.get("filter_lokasi"):
                locations.add(row["filter_lokasi"])
            if row.get("status_label"):
                statuses.add(row["status_label"])
    return {
        "lokasi": sorted(locations, key=lambda item: item.casefold()),
        "status": sorted(statuses, key=lambda item: item.casefold()),
    }


def _plain_cell(value: str) -> dict:
    shown = (value or "").strip()
    return {"text": shown, "cls": "", "digits": "", "kind": "text", "sort": shown.casefold()}


def _excluded_sheet() -> dict:
    headers = ["Nama", "Username", "Lokasi", "Alasan"]
    rows = []
    for item in payroll_exclusions():
        cells = [
            _plain_cell(item["nama"]),
            _plain_cell(item["username"]),
            _plain_cell(item["lokasi"]),
            _plain_cell(item["alasan"]),
        ]
        bits = [cell["text"].lower() for cell in cells if cell["text"]]
        if item["status_label"]:
            bits.append(item["status_label"].lower())
        rows.append(
            {
                "cells": cells,
                "search": " ".join(bits),
                "total": False,
                "officer_id": item["id"],
                "status_label": item["status_label"],
                "account_code": "",
                "file_deposit": "",
                "has_condition": "disimpan" in (item.get("alasan") or "").lower(),
                "has_potongan": False,
                "has_tambahan": False,
                "has_rincian": False,
                "has_selisih": False,
                "new_joiner": False,
                "normal_case": False,
                "filter_lokasi": item.get("lokasi") or "",
                "filter_gaji_pokok": "",
                "filter_gaji_hitung": "",
                "filter_deposit": "",
                "filter_selisih": "",
                "filter_alpa": "",
                "filter_new_joiner": "0",
                "filter_normal": "0",
            }
        )
    return {
        "name": "Tidak Dihitung",
        "id": sheet_slug("Tidak Dihitung"),
        "intro": [
            "Petugas yang tidak berhak masuk payroll saat ini: Blacklist, Resign, atau akun testing. Perubahan status tetap bisa dilakukan dari kolom Status.",
        ],
        "headers": headers,
        "rows": rows,
        "numeric": [False, False, False, False],
        "freeze": 1,
        "person": True,
        "excluded": True,
    }


def _attach_sheet_meta(sheet: dict, meta: dict) -> None:
    info = meta.get(sheet["name"]) or {}
    sheet["title"] = info.get("title") or sheet["name"]
    sheet["hint"] = info.get("hint") or ""
    sheet["help"] = info.get("help") or ""
    sheet["terms"] = list(info.get("terms") or [])


def _sheet_nav(spec: dict, sheets: list[dict]) -> list[dict]:
    """Group sheets for payroll IA; unknown sheets fall into Lainnya."""
    by_name = {sheet["name"]: sheet for sheet in sheets}
    meta = spec.get("sheets_meta") or {}
    # Back-compat if older help map still present
    help_map = spec.get("help") or {}
    groups = []
    used = set()
    for group in spec.get("groups") or []:
        items = []
        for name in group.get("sheets") or []:
            sheet = by_name.get(name)
            if sheet is None:
                continue
            _attach_sheet_meta(sheet, meta)
            if not sheet["help"]:
                sheet["help"] = help_map.get(name) or ""
            sheet["group_id"] = group["id"]
            items.append(sheet)
            used.add(name)
        if items:
            groups.append(
                {
                    "id": group["id"],
                    "label": group["label"],
                    "blurb": group.get("blurb") or "",
                    "sheets": items,
                }
            )
    leftover = [sheet for sheet in sheets if sheet["name"] not in used]
    if leftover:
        for sheet in leftover:
            _attach_sheet_meta(sheet, meta)
            if not sheet["help"]:
                sheet["help"] = help_map.get(sheet["name"]) or ""
            sheet["group_id"] = "lainnya"
        groups.append({"id": "lainnya", "label": "Lainnya", "blurb": "Tabel lain", "sheets": leftover})
    ordered = [sheet for group in groups for sheet in group["sheets"]]
    return groups, ordered


def _normalize_banding_status(raw: str) -> str:
    text = (raw or "").strip().lower()
    if not text:
        return ""
    if "approve" in text or "disetujui" in text:
        return "disetujui"
    if "pending" in text or "menunggu" in text or "waiting" in text:
        return "menunggu"
    if "reject" in text or "tolak" in text or "ditolak" in text:
        return "ditolak"
    return text


def _payroll_banding_summary(sheets: list[dict]) -> dict:
    """Counts for Banding sheet statuses so payroll users see approve/reject/pending totals."""
    counts = {"total": 0, "disetujui": 0, "menunggu": 0, "ditolak": 0, "lain": 0}
    for sheet in sheets:
        if sheet.get("name") != "Banding":
            continue
        headers = sheet.get("headers") or []
        if "Status" not in headers:
            continue
        at = headers.index("Status")
        for row in sheet.get("rows") or []:
            if row.get("total"):
                continue
            cells = row.get("cells") or []
            raw = cells[at].get("text") if at < len(cells) else ""
            kind = _normalize_banding_status(raw)
            row["filter_banding"] = kind
            counts["total"] += 1
            if kind in {"disetujui", "menunggu", "ditolak"}:
                counts[kind] += 1
            else:
                counts["lain"] += 1
        sheet["banding"] = True
        break
    return counts


def _payroll_focus_summary(sheets: list[dict]) -> dict:
    """Quick-filter counts for Slip Gaji person rows (compare-cards)."""
    people = []
    for sheet in sheets:
        if sheet.get("name") != "Slip Gaji" or not sheet.get("person"):
            continue
        people = [row for row in sheet.get("rows") or [] if not row.get("total")]
        break
    if not people:
        for sheet in sheets:
            if not sheet.get("person") or sheet.get("excluded"):
                continue
            people = [row for row in sheet.get("rows") or [] if not row.get("total")]
            if people:
                break

    def _num(row, key):
        raw = row.get(key) or ""
        if raw == "":
            return None
        try:
            return float(raw)
        except ValueError:
            return None

    total = len(people)
    selisih = sum(1 for row in people if row.get("has_selisih"))
    normal = sum(1 for row in people if row.get("filter_normal") == "1")
    new_joiner = sum(1 for row in people if row.get("filter_new_joiner") == "1" or row.get("new_joiner"))
    deposit = sum(1 for row in people if (_num(row, "filter_deposit") or 0) > 0)
    hadir_kurang = sum(1 for row in people if row.get("hadir_kurang") or row.get("filter_hadir_ok") == "0")
    hadir_ok = sum(1 for row in people if row.get("filter_hadir_ok") == "1")
    perhatian = sum(
        1
        for row in people
        if row.get("has_selisih")
        or row.get("new_joiner")
        or row.get("has_condition")
        or row.get("hadir_kurang")
        or (_num(row, "filter_deposit") or 0) > 0
    )
    return {
        "total": total,
        "perhatian": perhatian,
        "selisih": selisih,
        "normal": normal,
        "new_joiner": new_joiner,
        "deposit": deposit,
        "hadir_kurang": hadir_kurang,
        "hadir_ok": hadir_ok,
    }


def view_book(sid: str, book: str):
    spec = BOOKS.get(book)
    if spec is None:
        abort(404)
    folder = session_dir(sid)
    meta_path = folder / "meta.json"
    meta = json.loads(meta_path.read_text(encoding="utf-8")) if meta_path.exists() else {}
    sheets = read_workbook(prepare_book(folder, book))
    wanted = spec.get("sheets") or []
    if wanted:
        order = {name: index for index, name in enumerate(wanted)}
        sheets = [sheet for sheet in sheets if sheet["name"] in order]
        sheets.sort(key=lambda sheet: order[sheet["name"]])
    if not sheets:
        abort(404)
    active = spec["default"]
    if not any(sheet["name"] == active for sheet in sheets):
        active = sheets[0]["name"]
    source_name = meta.get("source_name") or "payroll"
    heading = spec["title"]
    if book == "hitung" and "2026-10" in source_name:
        heading = "Payroll Oktober 2026"
    _mark_payroll_people(sheets)
    banding = _payroll_banding_summary(sheets) if book == "hitung" else {
        "total": 0, "disetujui": 0, "menunggu": 0, "ditolak": 0, "lain": 0
    }
    if book == "hitung":
        sheets.append(_excluded_sheet())
    filters = _payroll_filter_options(sheets)
    focus = _payroll_focus_summary(sheets)
    sheet_groups, sheets = _sheet_nav(spec, sheets)
    active_id = sheet_slug(active)
    guide_terms = []
    seen_terms = set()
    for sheet in sheets:
        for term in sheet.get("terms") or []:
            key = str(term).strip().lower()
            if not key or key in seen_terms:
                continue
            seen_terms.add(key)
            guide_terms.append(term)
    from compare import session_source

    uploads = [row for row in list_uploads() if session_source(row["id"])]
    upload = get_upload(folder.name)
    return render_template(
        "grid.html",
        book=book,
        spec=spec,
        sheets=sheets,
        sheet_groups=sheet_groups,
        active=active_id,
        active_group=next((sheet.get("group_id") for sheet in sheets if sheet["id"] == active_id), (sheet_groups[0]["id"] if sheet_groups else "")),
        guide_terms=guide_terms,
        sid=folder.name,
        source_name=source_name,
        heading=heading,
        filters=filters,
        focus=focus,
        banding=banding,
        uploads=uploads,
        upload=upload,
    )


def final_context(folder: Path, saved: bool = False) -> dict:
    payroll, meta = load_payroll(folder)
    edits = load_closing(folder)
    rows = present_transfer(payroll, edits=edits)
    rows.sort(key=lambda row: (row["lokasi"].lower(), row["nama"].lower(), row["username"].lower()))
    for row in rows:
        bits = [
            row["lokasi"],
            row["nama"],
            row["username"],
            row["trx"],
            row["bank"],
            row["rekening"],
            row["penerima"],
            row["remark"],
            row["catatan"],
            row["status"],
            row["system"],
            str(row["amount"]),
        ]
        row["search"] = " ".join(text(bit).lower() for bit in bits if text(bit))
    active = [row for row in rows if row["include"] and not row["deleted"]]
    return {
        "sid": folder.name,
        "source_name": meta.get("source_name") or payroll.source_name or "payroll",
        "rows": rows,
        "saved": saved,
        "stats": {
            "slip": len(rows),
            "transfer": len(active),
            "amount": sum(row["amount"] for row in active),
            "deleted": sum(1 for row in rows if row["deleted"]),
            "notes": sum(1 for row in rows if row["catatan"]),
            "hris": sum(row["paid"] for row in rows),
        },
    }


def view_final(sid: str):
    folder = session_dir(sid)
    return render_template("final.html", page=final_context(folder, request.args.get("saved") == "1"))


def save_final(sid: str):
    folder = session_dir(sid)
    payroll, _meta = load_payroll(folder)
    originals = {row["key"]: row for row in transfer_candidates(payroll)}
    save_closing(folder, edits_from_form(request.form, originals))
    add_activity("ringkasan", "Ringkasan akhir disimpan", "Catatan, nominal transfer, atau baris yang dikeluarkan.", folder.name)
    return redirect(url_for("view_final", sid=folder.name, saved=1))


def _attach_excluded_sheet(path: Path) -> None:
    workbook = load_workbook(path)
    write_excluded(workbook, payroll_exclusions())
    workbook.save(path)


def download(sid: str):
    folder = session_dir(sid)
    target = prepare_book(folder, "hitung")
    _attach_excluded_sheet(target)
    meta = json.loads((folder / "meta.json").read_text(encoding="utf-8"))
    stem = Path(meta.get("source_name") or "payroll").stem
    return Download(target, f"hitung_{stem}.xlsx")


def download_transfer(sid: str):
    folder = session_dir(sid)
    payroll, meta = load_payroll(folder)
    target = folder / "transfer.xlsx"
    write_transfer(payroll, target, edits=load_closing(folder))
    stem = Path(meta.get("source_name") or "payroll").stem
    return Download(target, f"transfer_{stem}.xlsx")


def _previous_month(day: date) -> str:
    year = day.year
    month = day.month - 1
    if month == 0:
        year -= 1
        month = 12
    return f"{year:04d}-{month:02d}"


def _absensi_choice():
    rows = attendance_locations()
    raw = request.args.get("lokasi") or request.form.get("lokasi") or ""
    if raw.isdigit() and any(row["id"] == int(raw) for row in rows):
        location_id = int(raw)
    elif rows:
        location_id = next((row["id"] for row in rows if row["active"]), rows[0]["id"])
    else:
        location_id = 0
    work_date = parse_day(request.args.get("tanggal") or request.form.get("tanggal"))
    return rows, location_id, work_date


def absensi():
    rows, location_id, work_date = _absensi_choice()
    board = day_board(location_id, work_date) if location_id else None
    ajukan_raw = request.args.get("ajukan") or ""
    ajukan = slot_detail(int(ajukan_raw)) if ajukan_raw.isdigit() else None
    if ajukan and ajukan["location_id"] != location_id:
        ajukan = None
    return render_template(
        "absensi.html",
        locations=rows,
        location_id=location_id,
        tanggal=work_date.isoformat(),
        board=board,
        ajukan=ajukan,
        codes=[{"code": code, "label": name} for code, name in LABELS.items()],
        saved=request.args.get("simpan") == "1",
        info=request.args.get("info") or "",
        jadwal=request.args.get("jadwal") or "",
        salah=request.args.get("salah") or "",
        pending=len(banding_queue(location_id)) if location_id else 0,
    )


def save_absensi():
    rows, location_id, work_date = _absensi_choice()
    if not location_id:
        redirect(url_for("absensi"))
    aksi = request.form.get("aksi") or "simpan"
    salah = ""
    if aksi == "slot2":
        officer = request.form.get("petugas") or ""
        shift_raw = request.form.get("shift") or ""
        if officer.isdigit():
            salah = add_second_slot(
                location_id,
                int(officer),
                work_date,
                int(shift_raw) if shift_raw.isdigit() else None,
            )
    elif aksi == "banding":
        slot = request.form.get("slot") or ""
        if slot.isdigit():
            salah = file_banding(
                int(slot),
                request.form.get("kode") or "",
                request.form.get("alasan") or "",
                request.form.get("catatan") or "",
            )
    else:
        salah = save_slots(location_id, work_date, request.form)
    redirect(
        url_for(
            "absensi",
            lokasi=location_id,
            tanggal=work_date.isoformat(),
            simpan="1" if aksi == "simpan" and not salah else "",
            info="Banding terkirim. Kode belum berubah sampai disetujui." if aksi == "banding" and not salah else ("Slot kedua ditambahkan." if aksi == "slot2" and not salah else ""),
            salah=salah,
        )
    )


def absensi_shift():
    rows, location_id, work_date = _absensi_choice()
    return render_template(
        "absensi_shift.html",
        locations=rows,
        location_id=location_id,
        tanggal=work_date.isoformat(),
        shifts=shifts_for(location_id) if location_id else [],
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
        pending=len(banding_queue(location_id)) if location_id else 0,
    )


def save_absensi_shift():
    _, location_id, work_date = _absensi_choice()
    if request.form.get("aksi") == "lokasi":
        salah = add_location(request.form.get("nama") or "")
        redirect(url_for("absensi_shift", tanggal=work_date.isoformat(), salah=salah, simpan="" if salah else "1"))
    salah = add_shift(
        location_id,
        request.form.get("nama") or "",
        request.form.get("mulai") or "",
        request.form.get("selesai") or "",
        request.form.get("tenggang") or "",
        request.form.get("hari") or "ini",
    )
    redirect(
        url_for(
            "absensi_shift",
            lokasi=location_id,
            tanggal=work_date.isoformat(),
            salah=salah,
            simpan="" if salah else "1",
        )
    )


def absensi_petugas():
    rows, location_id, work_date = _absensi_choice()
    return render_template(
        "absensi_petugas.html",
        locations=rows,
        location_id=location_id,
        tanggal=work_date.isoformat(),
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
        pending=len(banding_queue(location_id)) if location_id else 0,
    )


def save_absensi_petugas():
    _, location_id, work_date = _absensi_choice()
    target = request.form.get("lokasi") or ""
    salah = add_officer(
        request.form.get("nama") or "",
        request.form.get("username") or "",
        int(target) if target.isdigit() else location_id,
        request.form.get("mulai") or work_date.isoformat(),
        request.form.get("testing") == "1",
    )
    redirect(
        url_for(
            "absensi_petugas",
            lokasi=location_id,
            tanggal=work_date.isoformat(),
            salah=salah,
            simpan="" if salah else "1",
        )
    )


def absensi_banding():
    rows, location_id, work_date = _absensi_choice()
    return render_template(
        "absensi_banding.html",
        locations=rows,
        location_id=location_id,
        tanggal=work_date.isoformat(),
        queue=banding_queue(location_id) if location_id else [],
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
        pending=len(banding_queue(location_id)) if location_id else 0,
    )


def putus_banding(banding_id: str):
    _, location_id, work_date = _absensi_choice()
    salah = decide_banding(
        int(banding_id),
        request.form.get("keputusan") == "setuju",
        request.form.get("catatan") or "",
    )
    redirect(
        url_for(
            "absensi_banding",
            lokasi=location_id,
            tanggal=work_date.isoformat(),
            salah=salah,
            simpan="" if salah else "1",
        )
    )


def data_lokasi():
    raw = request.args.get("ubah") or ""
    if raw.isdigit():
        redirect(url_for("data_lokasi_profil", location_id=int(raw)))
    filters = location_filters(request.args)
    return render_template(
        "data_lokasi.html",
        rows=location_directory(filters),
        filters=filters,
        filtered=any(filters.values()),
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
    )


def save_data_lokasi():
    aksi = request.form.get("aksi") or "tambah"
    salah = ""
    if aksi == "ubah":
        raw = request.form.get("id") or ""
        if raw.isdigit():
            salah = save_location(
                int(raw),
                request.form.get("nama") or "",
                request.form.get("aktif") == "1",
                request.form.get("tenggang") or "",
                request.form.get("min_jam") or "",
                request.form.get("min_hadir") or "",
                request.form.get("min_off") or "",
            )
    elif aksi == "hapus":
        raw = request.form.get("id") or ""
        if raw.isdigit():
            salah = delete_location(int(raw))
    else:
        salah = add_location(
            request.form.get("nama") or "",
            request.form.get("tenggang") or "15",
            request.form.get("min_jam") or "",
            request.form.get("min_hadir") or "",
            request.form.get("min_off") or "",
        )
    raw = request.form.get("id") or ""
    if aksi == "ubah" and raw.isdigit():
        redirect(url_for("data_lokasi_profil", location_id=int(raw), salah=salah, simpan="" if salah else "1"))
    kept = {key: value for key, value in location_filters(request.args).items() if value}
    redirect(url_for("data_lokasi", salah=salah, simpan="" if salah else "1", **kept))


def _location_snapshot(record: dict) -> dict:
    minimum = f"{record['min_jam']} jam" if record.get("min_jam") else "—"
    active = bool(record.get("active"))
    return {
        "id": record["id"],
        "name": record["name"],
        "grace": f"{record['grace_minutes']} menit",
        "minimum": minimum,
        "active": active,
        "status": "Aktif" if active else "Nonaktif",
    }


def _render_location_panel(record: dict, saved: bool, salah: str):
    return render_template(
        "data_lokasi_panel.html",
        location=record,
        radii=GPS_RADII,
        saved=saved,
        salah=salah,
        snapshot=_location_snapshot(record),
    )


def data_lokasi_profil(location_id: str):
    record = location_record(int(location_id))
    if not record:
        abort(404)
    if request.args.get("panel") == "1":
        return _render_location_panel(record, request.args.get("simpan") == "1", request.args.get("salah") or "")
    tab = request.args.get("tab") or "profil"
    if tab not in ("profil", "karyawan", "penempatan", "riwayat"):
        tab = "profil"
    work_date = parse_day(request.args.get("tanggal"))
    board = groups = summary = month = None
    if tab == "penempatan":
        board = day_board(record["id"], work_date)
        groups, summary = shift_groups(board)
        month = assignment_month(record["id"], work_date)
    return render_template(
        "data_lokasi_profil.html",
        location=record,
        radii=GPS_RADII,
        team=location_team(record["id"]),
        tab=tab,
        tanggal=work_date.isoformat(),
        tanggal_label=day_label(work_date),
        board=board,
        groups=groups or [],
        summary=summary,
        month=month,
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
        info=request.args.get("info") or "",
        sumber=request.args.get("sumber") or _previous_month(work_date),
        log=location_log(record["id"]) if tab == "riwayat" else [],
    )


def save_data_lokasi_profil(location_id: str):
    record = location_record(int(location_id))
    if not record:
        abort(404)
    tab = request.form.get("tab") or "profil"
    if tab == "penempatan":
        work_date = parse_day(request.form.get("tanggal"))
        if request.form.get("aksi") == "salin":
            raw_source = (request.form.get("sumber") or "")[:7]
            try:
                source = datetime.strptime(raw_source, "%Y-%m").date()
            except ValueError:
                salah, info = "Pilih bulan sumber.", ""
            else:
                salah, info = copy_shift_pattern(record["id"], source, work_date)
            redirect(
                url_for(
                    "data_lokasi_profil",
                    location_id=record["id"],
                    tab="penempatan",
                    tanggal=work_date.isoformat(),
                    sumber=raw_source,
                    salah=salah,
                    info=info,
                    simpan="" if salah else "1",
                )
            )
        salah = save_shift_assignment(record["id"], work_date, request.form)
        redirect(
            url_for(
                "data_lokasi_profil",
                location_id=record["id"],
                tab="penempatan",
                tanggal=work_date.isoformat(),
                salah=salah,
                simpan="" if salah else "1",
            )
        )
    salah = save_location(
        record["id"],
        request.form.get("nama") or "",
        request.form.get("aktif") == "1",
        request.form.get("tenggang") or "",
        request.form.get("min_jam") or "",
        request.form.get("min_hadir") or "",
        request.form.get("min_off") or "",
        request.form.get("lintang") or "",
        request.form.get("bujur") or "",
        request.form.get("jarak") or "",
    )
    if request.form.get("panel") == "1":
        fresh = location_record(record["id"]) or record
        return _render_location_panel(fresh, not salah, salah or "")
    redirect(url_for("data_lokasi_profil", location_id=record["id"], salah=salah, simpan="" if salah else "1"))


def penempatan():
    _rows, location_id, work_date = _absensi_choice()
    if not location_id:
        redirect(url_for("data_lokasi"))
    redirect(
        url_for(
            "data_lokasi_profil",
            location_id=location_id,
            tab="penempatan",
            tanggal=work_date.isoformat(),
        )
    )


def save_penempatan():
    _rows, location_id, work_date = _absensi_choice()
    if not location_id:
        redirect(url_for("data_lokasi"))
    salah = save_shift_assignment(location_id, work_date, request.form)
    redirect(
        url_for(
            "data_lokasi_profil",
            location_id=location_id,
            tab="penempatan",
            tanggal=work_date.isoformat(),
            salah=salah,
            simpan="" if salah else "1",
        )
    )


def data_shift():
    rows, location_id, _work_date = _absensi_choice()
    shift_rows = shifts_for(location_id) if location_id else []
    raw = request.args.get("ubah") or ""
    editing = next((row for row in shift_rows if str(row["id"]) == raw), None)
    return render_template(
        "data_shift.html",
        locations=rows,
        location_id=location_id,
        shifts=shift_rows,
        editing=editing,
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
    )


def save_data_shift():
    _rows, location_id, _work_date = _absensi_choice()
    aksi = request.form.get("aksi") or "tambah"
    salah = ""
    if aksi == "ubah":
        raw = request.form.get("id") or ""
        if raw.isdigit():
            salah = save_shift(
                int(raw),
                request.form.get("nama") or "",
                request.form.get("mulai") or "",
                request.form.get("selesai") or "",
                request.form.get("tenggang") or "",
                request.form.get("hari") or "ini",
            )
    elif aksi == "hapus":
        raw = request.form.get("id") or ""
        if raw.isdigit():
            salah = delete_shift(int(raw))
    else:
        salah = add_shift(
            location_id,
            request.form.get("nama") or "",
            request.form.get("mulai") or "",
            request.form.get("selesai") or "",
            request.form.get("tenggang") or "",
            request.form.get("hari") or "ini",
        )
    redirect(url_for("data_shift", lokasi=location_id, salah=salah, simpan="" if salah else "1"))


def data_karyawan():
    tahap = request.args.get("tahap") or ""
    filters = officer_filters(request.args)
    people = employee_directory(
        request.args.get("q") or "",
        request.args.get("status") or "",
        tahap,
        filters,
        filters["sort"],
        filters["arah"],
    )
    waiting = employee_directory(tahap="menunggu")
    fresh = employee_directory(tahap="pengguna") + employee_directory(tahap="ditolak")
    return render_template(
        "data_karyawan.html",
        people=people,
        waiting=waiting,
        fresh=fresh,
        locations=attendance_locations(),
        positions=POSITIONS,
        employment=EMPLOYMENT,
        status_help=STATUS_HELP,
        query=request.args.get("q") or "",
        status=request.args.get("status") or "",
        filters=filters,
        sort=filters["sort"],
        arah=filters["arah"],
        tahap=tahap,
        tab=request.args.get("tab") or "daftar",
        saved=request.args.get("simpan") == "1",
        salah=request.args.get("salah") or "",
    )


def save_data_pengguna():
    salah, officer_id = create_user(request.form.get("nama") or "", request.form.get("username") or "")
    if salah or not officer_id:
        redirect(url_for("data_karyawan", tab="pengguna", salah=salah or "Pengguna belum tersimpan."))
    redirect(url_for("data_karyawan", tab="daftar-masuk", simpan="1"))


def save_data_daftar():
    raw = request.form.get("petugas") or ""
    location = request.form.get("lokasi") or ""
    salah = register_employee(
        int(raw) if raw.isdigit() else 0,
        int(location) if location.isdigit() else 0,
        request.form.get("masuk") or "",
        request.form.get("gaji") or "",
        request.form.get("status") or "",
    )
    redirect(url_for("data_karyawan", tab="setuju" if not salah else "daftar", salah=salah, simpan="" if salah else "1"))


def save_data_setuju(officer_id: str):
    salah = decide_onboarding(int(officer_id), request.form.get("keputusan") == "setuju")
    redirect(url_for("data_karyawan", tab="setuju", salah=salah, simpan="" if salah else "1"))


def _profile_snapshot(officer_id: int) -> dict:
    row = employee_row(officer_id)
    if not row:
        return {}
    return {
        "id": row["id"],
        "nama": row["nama"],
        "username": row["username"] or "",
        "lokasi": row["lokasi"] or "—",
        "lokasi_kedua": row["lokasi_kedua"] or "—",
        "posisi": row["posisi_label"] or "—",
        "status_label": row["status_label"],
        "account_code": row["account_code"],
        "new_joiner": bool(row["new_joiner"]),
        "ojt_menunggu": bool(row["ojt_menunggu"]),
        "joined": row["joined"] or "—",
        "resigned": row["resigned"] or "—",
        "masa_hari": row["masa_hari"],
        "masa_hari_label": row["masa_hari_label"],
        "masa_tahun_label": row["masa_tahun_label"],
        "gaji": rupiah(float(row["basic_salary"] or 0)),
        "bpjs_kes": rupiah(float(row["bpjs_kesehatan"] or 0)),
        "bpjs_tk": rupiah(float(row["bpjs_ketenagakerjaan"] or 0)),
    }


def _render_profile(officer_id: str, tab: str, saved: bool, salah: str, notice: dict, panel: bool) -> str:
    person = employee_profile(int(officer_id))
    if not person:
        abort(404)
    if tab not in ("kerja", "gaji", "rekening", "riwayat"):
        tab = "kerja"
    dates = status_date_fields(person)
    current = next((item for item in dates if item["code"] == person["status_choice"]), dates[0])
    source_deposit = file_deposit_for_username(person.get("username") or "")
    payroll_info = officer_payroll_info(person["id"], source_deposit)
    change_log = profile_log(person["id"]) if tab in ("gaji", "riwayat") else []
    pay_var_log = [
        row
        for row in change_log
        if (row.get("field_label") or "").startswith("Tunjangan Jabatan")
        or (row.get("field_label") or "") == "Deposit"
    ]
    return render_template(
        "data_profil_panel.html" if panel else "data_profil.html",
        person=person,
        locations=attendance_locations(),
        employment=EMPLOYMENT,
        tab=tab,
        panel=panel,
        status_dates=dates,
        when=current,
        log=change_log if tab == "riwayat" else [],
        pay_var_log=pay_var_log if tab == "gaji" else [],
        payroll_info=payroll_info,
        adjustment_log=payroll_info["log"] if tab in ("gaji", "riwayat") else [],
        saved=saved,
        salah=salah or "",
        snapshot=_profile_snapshot(person["id"]) if panel else {},
        mutasi={
            "dari": notice.get("dari") or "",
            "ke": notice.get("ke") or "",
            "salin": notice.get("salin") or "",
        },
    )


def _render_status_panel(officer_id: str, saved: bool, salah: str):
    person = employee_profile(int(officer_id))
    if person is None:
        abort(404)
    dates = status_date_fields(person)
    current = next((item for item in dates if item["code"] == person["status_choice"]), dates[0])
    return render_template(
        "data_status_panel.html",
        person=person,
        saved=saved,
        salah=salah,
        when=current,
        snapshot={"id": person["id"], "label": person["status_label"], "dates": dates},
    )


def _payroll_panel_mode(raw: str) -> str:
    mode = (raw or "catatan").strip().lower()
    if mode in ("pengurangan", "potongan"):
        return "potongan"
    if mode in ("penambahan", "tambahan"):
        return "tambahan"
    if mode in ("rincian", "detail"):
        return "rincian"
    if mode in ("selisih",):
        return "selisih"
    if mode in ("syarat", "catatan"):
        return "catatan"
    return "catatan"


def _officer_from_session(sid: str, username: str):
    """Load one officer (+ banding + attendance activity) from a payroll session source."""
    source = None
    if sid and SESSION_ID.match(sid):
        folder = (SESSIONS / sid).resolve()
        if folder.parent == SESSIONS.resolve() and folder.is_dir():
            candidate = folder / "source.xlsx"
            if candidate.exists():
                source = candidate
    if source is None:
        source = latest_source()
    if source is None or not source.exists():
        return None, [], None
    payroll = build(source)
    apply_payroll_conditions(payroll)
    key = (username or "").strip().lower()
    officer = next((row for row in payroll.officers if (row.username or "").strip().lower() == key), None)
    if officer is None:
        return None, payroll.banding, None
    line = next(
        (
            row
            for row in payroll.attendance
            if (row.username or "").strip().lower() == key
            and (not officer.lokasi or (row.lokasi or "").strip().lower() == (officer.lokasi or "").strip().lower())
        ),
        None,
    )
    if line is None:
        line = next(
            (row for row in payroll.attendance if (row.username or "").strip().lower() == key),
            None,
        )
    activity = officer.attendance_activity(
        payroll.banding,
        attendance_line=line,
        date_labels=list(payroll.attendance_dates or []),
    )
    return officer, payroll.banding, activity


def _render_selisih_panel(officer_id: str, file_deposit: float | None = None, sid: str = ""):
    person = employee_profile(int(officer_id))
    if person is None:
        abort(404)
    if file_deposit is None:
        file_deposit = file_deposit_for_username(person.get("username") or "")
    source_deposit = int(round(float(file_deposit or 0)))
    officer, banding, attendance = _officer_from_session(sid, person.get("username") or "")
    if officer is None:
        detail = {
            "paid_label": "—",
            "hitung_label": "—",
            "selisih_label": "—",
            "selisih": 0,
            "lines": [],
            "banding": [],
            "new_joiner": None,
        }
    else:
        detail = officer.selisih_breakdown(banding)
    return render_template(
        "data_payroll_selisih_panel.html",
        person=person,
        detail=detail,
        attendance=attendance,
        source_deposit=source_deposit,
        snapshot={
            "id": person["id"],
            "username": person.get("username") or "",
            "label": person["status_label"],
            "mode": "selisih",
            "deposit": source_deposit,
            "potongan": None,
            "tambahan": None,
            "note": "",
        },
    )


def _render_payroll_detail_panel(officer_id: str, file_deposit: float | None = None):
    person = employee_profile(int(officer_id))
    if person is None:
        abort(404)
    if file_deposit is None:
        file_deposit = file_deposit_for_username(person.get("username") or "")
    source_deposit = int(round(float(file_deposit or 0)))
    payroll_info = officer_payroll_info(person["id"], source_deposit)
    _officer, _banding, attendance = _officer_from_session(
        request.args.get("sid") or "",
        person.get("username") or "",
    )
    return render_template(
        "data_payroll_detail_panel.html",
        person=person,
        payroll_info=payroll_info,
        attendance=attendance,
        source_deposit=source_deposit,
        snapshot={
            "id": person["id"],
            "username": person.get("username") or "",
            "label": person["status_label"],
            "mode": "rincian",
            "deposit": payroll_info.get("effective_deposit"),
            "potongan": payroll_info["condition"].get("potongan"),
            "tambahan": payroll_info["condition"].get("tambahan"),
            "note": payroll_info.get("note") or "",
        },
    )


def _render_payroll_panel(
    officer_id: str,
    saved: bool,
    salah: str,
    file_deposit: float | None = None,
    mode: str = "catatan",
):
    person = employee_profile(int(officer_id))
    if person is None:
        abort(404)
    mode = _payroll_panel_mode(mode)
    if mode == "rincian":
        return _render_payroll_detail_panel(officer_id, file_deposit)
    if mode == "selisih":
        return _render_selisih_panel(officer_id, file_deposit, request.args.get("sid") or "")
    condition = payroll_condition_for(person.get("username") or "")
    source_deposit = file_deposit
    if source_deposit is None:
        source_deposit = file_deposit_for_username(person.get("username") or "")
    potongan_value = condition["potongan"] if condition.get("potongan") is not None else ""
    tambahan_value = condition["tambahan"] if condition.get("tambahan") is not None else ""
    labels = {"potongan": "Pengurangan payroll", "tambahan": "Penambahan payroll", "catatan": "Syarat payroll"}
    _officer, _banding, attendance = _officer_from_session(
        request.args.get("sid") or "",
        person.get("username") or "",
    )
    return render_template(
        "data_payroll_panel.html",
        person=person,
        condition=condition,
        mode=mode,
        mode_label=labels[mode],
        potongan_value=potongan_value,
        tambahan_value=tambahan_value,
        potongan_reasons=PAYROLL_POTONGAN_REASONS,
        tambahan_reasons=PAYROLL_TAMBAHAN_REASONS,
        source_deposit=int(round(float(source_deposit or 0))),
        adjustment_log=payroll_adjustment_log(person["id"]),
        attendance=attendance,
        saved=saved,
        salah=salah,
        snapshot={
            "id": person["id"],
            "username": person.get("username") or "",
            "label": person["status_label"],
            "mode": mode,
            "deposit": condition.get("deposit"),
            "potongan": condition.get("potongan"),
            "tambahan": condition.get("tambahan"),
            "note": condition.get("note") or "",
        },
    )


def data_profil(officer_id: str):
    if request.args.get("status") == "1":
        return _render_status_panel(officer_id, False, request.args.get("salah") or "")
    if request.args.get("syarat") == "1" or request.args.get("pay") in (
        "potongan",
        "tambahan",
        "pengurangan",
        "penambahan",
        "catatan",
        "syarat",
        "rincian",
        "detail",
        "selisih",
    ):
        raw = request.args.get("deposit") or ""
        file_deposit = float(raw) if raw.replace(".", "", 1).isdigit() else None
        mode = request.args.get("pay") or request.args.get("mode") or "catatan"
        return _render_payroll_panel(officer_id, False, request.args.get("salah") or "", file_deposit, mode)
    notice = {
        "dari": request.args.get("dari") or "",
        "ke": request.args.get("ke") or "",
        "salin": "1" if request.args.get("salin") == "1" else "",
    }
    return _render_profile(
        officer_id,
        request.args.get("tab") or "kerja",
        request.args.get("simpan") == "1",
        request.args.get("salah") or "",
        notice,
        request.args.get("panel") == "1",
    )


def save_data_profil(officer_id: str):
    bagian = request.form.get("bagian") or ""
    notice = {}
    if bagian == "syarat":
        mode = _payroll_panel_mode(request.form.get("pay_section") or "catatan")
        kwargs = {"note": request.form.get("catatan") or "", "section": mode}
        if mode == "potongan":
            kwargs["potongan_raw"] = request.form.get("potongan") or ""
            kwargs["potongan_reason"] = request.form.get("potongan_alasan") or ""
        elif mode == "tambahan":
            kwargs["tambahan_raw"] = request.form.get("tambahan") or ""
            kwargs["tambahan_reason"] = request.form.get("tambahan_alasan") or ""
        salah = save_payroll_syarat(int(officer_id), **kwargs)
        if not salah:
            for stamp in SESSIONS.glob("*/eligible.stamp"):
                stamp.unlink(missing_ok=True)
        if request.form.get("panel") == "1":
            return _render_payroll_panel(officer_id, not salah, salah or "", mode=mode)
        back_tab = request.form.get("tab") or "kerja"
        if back_tab not in ("kerja", "gaji", "rekening", "riwayat"):
            back_tab = "kerja"
        redirect(url_for("data_profil", officer_id=officer_id, tab=back_tab, salah=salah, simpan="" if salah else "1"))
    if bagian == "status":
        salah = save_account_status(int(officer_id), request.form.get("status") or "", request.form.get("tanggal") or "")
        if not salah:
            for stamp in SESSIONS.glob("*/eligible.stamp"):
                stamp.unlink(missing_ok=True)
        if request.form.get("panel") == "1":
            return _render_status_panel(officer_id, not salah, salah or "")
        redirect(url_for("data_profil", officer_id=officer_id, tab="kerja", salah=salah, simpan="" if salah else "1"))
    if bagian == "deposit":
        salah = save_officer_deposit(int(officer_id), request.form.get("deposit"))
        if not salah:
            for stamp in SESSIONS.glob("*/eligible.stamp"):
                stamp.unlink(missing_ok=True)
        tab = "gaji"
    elif bagian == "tunjangan":
        salah = save_officer_tunjangan(int(officer_id), request.form)
        tab = "gaji"
    elif bagian == "gaji":
        salah = save_compensation(int(officer_id), request.form)
        tab = "gaji"
    elif bagian == "rekening":
        salah = save_account(int(officer_id), request.form)
        tab = "rekening"
    elif bagian == "ojt":
        salah = decide_ojt(int(officer_id), request.form.get("keputusan") == "lulus")
        tab = "kerja"
    else:
        result = save_employee_profile(int(officer_id), request.form)
        if isinstance(result, tuple):
            salah, notice = result
        else:
            salah, notice = result, {}
        tab = "kerja"
    if request.form.get("panel") == "1":
        return _render_profile(officer_id, tab, not salah, salah or "", notice, True)
    redirect(
        url_for(
            "data_profil",
            officer_id=officer_id,
            tab=tab,
            salah=salah,
            simpan="" if salah else "1",
            dari=notice.get("dari") or "",
            ke=notice.get("ke") or "",
            salin=notice.get("salin") or "",
        )
    )


def static_file(rel: str) -> BytesResult:
    root = STATIC.resolve()
    target = (root / rel).resolve()
    if root not in target.parents or not target.is_file():
        abort(404)
    return BytesResult(target.read_bytes(), STATIC_TYPES.get(target.suffix.lower(), "application/octet-stream"))


ROUTES = [
    ("GET", re.compile(r"^/login$"), login_page),
    ("POST", re.compile(r"^/login$"), save_login),
    ("GET", re.compile(r"^/logout$"), logout),
    ("POST", re.compile(r"^/logout$"), logout),
    ("GET", re.compile(r"^/pengaturan/pengguna$"), users_page),
    ("POST", re.compile(r"^/pengaturan/pengguna$"), save_users),
    ("GET", re.compile(r"^/$"), home),
    ("GET", re.compile(r"^/kegiatan$"), kegiatan),
    ("POST", re.compile(r"^/kegiatan$"), save_kegiatan),
    ("POST", re.compile(r"^/kegiatan/(?P<activity_id>[0-9]+)/hapus$"), hapus_kegiatan),
    ("POST", re.compile(r"^/unggahan/(?P<sid>[0-9a-f]{32})/nama$"), simpan_alias),
    ("POST", re.compile(r"^/unggahan/(?P<sid>[0-9a-f]{32})/hapus$"), hapus_unggahan),
    ("GET", re.compile(r"^/istilah$"), istilah),
    ("GET", re.compile(r"^/arsip$"), arsip),
    ("GET", re.compile(r"^/arsip/(?P<month_id>[0-9]+)$"), arsip_bulan),
    ("GET", re.compile(r"^/banding-unggah$"), banding_unggah),
    ("POST", re.compile(r"^/analyze$"), analyze),
    ("GET", re.compile(r"^/data/lokasi$"), data_lokasi),
    ("POST", re.compile(r"^/data/lokasi$"), save_data_lokasi),
    ("GET", re.compile(r"^/data/lokasi/(?P<location_id>[0-9]+)$"), data_lokasi_profil),
    ("POST", re.compile(r"^/data/lokasi/(?P<location_id>[0-9]+)$"), save_data_lokasi_profil),
    ("GET", re.compile(r"^/penempatan$"), penempatan),
    ("POST", re.compile(r"^/penempatan$"), save_penempatan),
    ("GET", re.compile(r"^/data/shift$"), data_shift),
    ("POST", re.compile(r"^/data/shift$"), save_data_shift),
    ("GET", re.compile(r"^/data/karyawan$"), data_karyawan),
    ("POST", re.compile(r"^/data/karyawan/pengguna$"), save_data_pengguna),
    ("POST", re.compile(r"^/data/karyawan/daftar$"), save_data_daftar),
    ("POST", re.compile(r"^/data/karyawan/(?P<officer_id>[0-9]+)/setuju$"), save_data_setuju),
    ("GET", re.compile(r"^/data/karyawan/(?P<officer_id>[0-9]+)$"), data_profil),
    ("POST", re.compile(r"^/data/karyawan/(?P<officer_id>[0-9]+)$"), save_data_profil),
    ("GET", re.compile(r"^/absensi$"), absensi),
    ("POST", re.compile(r"^/absensi$"), save_absensi),
    ("GET", re.compile(r"^/absensi/shift$"), absensi_shift),
    ("POST", re.compile(r"^/absensi/shift$"), save_absensi_shift),
    ("GET", re.compile(r"^/absensi/petugas$"), absensi_petugas),
    ("POST", re.compile(r"^/absensi/petugas$"), save_absensi_petugas),
    ("GET", re.compile(r"^/absensi/banding$"), absensi_banding),
    ("POST", re.compile(r"^/absensi/banding/(?P<banding_id>[0-9]+)$"), putus_banding),
    ("GET", re.compile(r"^/kelola$"), kelola),
    ("POST", re.compile(r"^/kelola$"), save_kelola),
    ("GET", re.compile(r"^/kelola/petugas$"), kelola_petugas),
    ("POST", re.compile(r"^/kelola/petugas$"), save_petugas),
    ("GET", re.compile(r"^/kelola/lokasi/(?P<lokasi>.+)/petugas/(?P<username>.+)/profil$"), kelola_profil),
    ("POST", re.compile(r"^/kelola/lokasi/(?P<lokasi>.+)/petugas/(?P<username>.+)/profil$"), save_profil),
    ("GET", re.compile(r"^/kelola/lokasi/(?P<lokasi>.+)/petugas/(?P<username>.+)$"), kelola_absensi),
    ("GET", re.compile(r"^/kelola/lokasi/(?P<lokasi>.+)$"), kelola_lokasi),
    ("GET", re.compile(r"^/report/(?P<sid>[0-9a-f]{32})/lihat/final$"), view_final),
    ("POST", re.compile(r"^/report/(?P<sid>[0-9a-f]{32})/lihat/final$"), save_final),
    ("GET", re.compile(r"^/report/(?P<sid>[0-9a-f]{32})/lihat/(?P<book>hitung|transfer)$"), view_book),
    ("GET", re.compile(r"^/report/(?P<sid>[0-9a-f]{32})/download$"), download),
    ("GET", re.compile(r"^/report/(?P<sid>[0-9a-f]{32})/transfer$"), download_transfer),
    ("GET", re.compile(r"^/report/(?P<sid>[0-9a-f]{32})$"), report),
]


def dispatch(method: str, path: str):
    for allowed, pattern, handler in ROUTES:
        if method != allowed:
            continue
        match = pattern.match(path)
        if match:
            return handler(**match.groupdict())
    abort(404)


class AppHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def do_GET(self):
        self._handle()

    def do_POST(self):
        self._handle()

    def log_message(self, fmt: str, *args) -> None:
        print("%s - %s" % (self.address_string(), fmt % args))

    def _handle(self) -> None:
        header_token = _response_headers.set([])
        try:
            parsed = urlparse(self.path)
            path = unquote(parsed.path)
            length = int(self.headers.get("Content-Length") or 0)
            if length > MAX_BODY:
                self._send(413, b"File terlalu besar.", "text/plain; charset=utf-8")
                return
            body = self.rfile.read(length) if length else b""
            incoming = parse_request(
                self.command,
                parsed.query,
                self.headers.get("Content-Type", ""),
                body,
                self.headers.get("Cookie", ""),
            )
            incoming.user = current_user(incoming.cookies.get(SESSION_COOKIE))
            token = _current.set(incoming)
            actor_token = bind_actor(incoming.user)
            try:
                if path.startswith("/static/"):
                    result = static_file(path[len("/static/") :])
                else:
                    if path not in PUBLIC_PATHS and incoming.user is None:
                        next_url = path
                        if parsed.query:
                            next_url += "?" + parsed.query
                        redirect(url_for("login", next=next_url))
                    result = dispatch(self.command, path)
            finally:
                reset_actor(actor_token)
                _current.reset(token)
            self._emit(result)
        except Redirect as exc:
            self.send_response(302)
            self.send_header("Location", exc.location)
            self.send_header("Content-Length", "0")
            for key, value in list(_response_headers.get() or []) + list(exc.headers or []):
                self.send_header(key, value)
            self.end_headers()
        except Forbidden:
            self._send(403, "Akses ditolak.".encode(), "text/plain; charset=utf-8")
        except NotFound:
            self._send(404, "Halaman tidak ditemukan.".encode(), "text/plain; charset=utf-8")
        except Exception:
            traceback.print_exc()
            self._send(500, "Terjadi kesalahan.".encode(), "text/plain; charset=utf-8")
        finally:
            _response_headers.reset(header_token)

    def _emit(self, result) -> None:
        queued = _response_headers.get() or []
        cookie_headers = [(k, v) for k, v in queued if k.lower() == "set-cookie"]
        other = {k: v for k, v in queued if k.lower() != "set-cookie"}
        if isinstance(result, Download):
            data = result.path.read_bytes()
            headers = {"Content-Disposition": content_disposition(result.name), **other}
            self._send(200, data, XLSX_TYPE, headers, cookie_headers)
            return
        if isinstance(result, BytesResult):
            self._send(200, result.data, result.content_type, other, cookie_headers)
            return
        body = result.encode("utf-8") if isinstance(result, str) else result
        self._send(200, body, "text/html; charset=utf-8", other, cookie_headers)

    def _send(
        self,
        status: int,
        data: bytes,
        content_type: str,
        headers: dict | None = None,
        multi_headers: list[tuple[str, str]] | None = None,
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        for key, value in multi_headers or []:
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(data)


class AppServer(ThreadingHTTPServer):
    allow_reuse_address = True


def main() -> None:
    SESSIONS.mkdir(parents=True, exist_ok=True)
    try:
        seed_auth()
    except Exception:
        traceback.print_exc()
    server = AppServer((HOST, PORT), AppHandler)
    print(f"Buka http://{HOST}:{PORT}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer berhenti.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
