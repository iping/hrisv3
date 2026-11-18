"""Simple app login + user management (separate from petugas/officers)."""

from __future__ import annotations

import hashlib
import hmac
import secrets
from contextvars import ContextVar, Token
from datetime import datetime, timedelta, timezone

from db import connect

SESSION_COOKIE = "hris_session"
SESSION_DAYS = 14
PBKDF2_ROUNDS = 120_000
ROLES = (
    ("admin", "Admin"),
    ("staff", "Staff"),
)

_request_actor: ContextVar[dict | None] = ContextVar("request_actor", default=None)


def bind_actor(user: dict | None) -> Token:
    return _request_actor.set(user)


def reset_actor(token: Token) -> None:
    _request_actor.reset(token)


def current_actor() -> dict | None:
    return _request_actor.get()


def actor_snapshot(user: dict | None = None) -> tuple[int | None, str, str]:
    person = user if user is not None else current_actor()
    if not person:
        return None, "", ""
    username = (person.get("username") or "").strip()
    nama = (person.get("nama") or username).strip()
    return person.get("id"), username, nama


def format_actor_label(nama: str | None, username: str | None) -> str:
    nama = (nama or "").strip()
    username = (username or "").strip()
    if nama and username and nama.lower() != username.lower():
        return f"{nama} ({username})"
    return nama or username or "—"


def _now() -> datetime:
    return datetime.now(timezone.utc)


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", (password or "").encode("utf-8"), salt.encode("utf-8"), PBKDF2_ROUNDS)
    return f"pbkdf2_sha256${PBKDF2_ROUNDS}${salt}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        algo, rounds, salt, digest = (stored or "").split("$", 3)
    except ValueError:
        return False
    if algo != "pbkdf2_sha256":
        return False
    try:
        count = int(rounds)
    except ValueError:
        return False
    check = hashlib.pbkdf2_hmac("sha256", (password or "").encode("utf-8"), salt.encode("utf-8"), count).hex()
    return hmac.compare_digest(check, digest)


def ensure_auth_schema(db) -> None:
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS app_users (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            nama TEXT NOT NULL DEFAULT '',
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'staff',
            active BOOLEAN NOT NULL DEFAULT TRUE,
            created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
        );
        CREATE TABLE IF NOT EXISTS app_sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES app_users(id) ON DELETE CASCADE,
            created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            expires_at TIMESTAMPTZ NOT NULL,
            last_seen TIMESTAMPTZ NOT NULL DEFAULT now()
        );
        """
    )
    count = db.execute("SELECT COUNT(*) AS n FROM app_users").fetchone()["n"]
    if int(count or 0) == 0:
        db.execute(
            """
            INSERT INTO app_users (username, nama, password_hash, role, active)
            VALUES (?, ?, ?, 'admin', TRUE)
            """,
            ("admin", "Administrator", hash_password("admin123")),
        )


def seed_auth() -> None:
    with connect() as db:
        ensure_auth_schema(db)


def _clean_username(raw: str) -> str:
    return (raw or "").strip().lower()


def authenticate(username: str, password: str) -> dict | None:
    key = _clean_username(username)
    if not key or not password:
        return None
    with connect() as db:
        ensure_auth_schema(db)
        row = db.execute(
            """
            SELECT id, username, nama, password_hash, role, active
            FROM app_users
            WHERE lower(username) = lower(?)
            """,
            (key,),
        ).fetchone()
    if not row or not row["active"]:
        return None
    if not verify_password(password, row["password_hash"]):
        return None
    return {
        "id": row["id"],
        "username": row["username"],
        "nama": row["nama"] or row["username"],
        "role": row["role"] or "staff",
        "active": bool(row["active"]),
        "is_admin": (row["role"] or "") == "admin",
    }


def create_session(user_id: int) -> str:
    token = secrets.token_urlsafe(32)
    expires = _now() + timedelta(days=SESSION_DAYS)
    with connect() as db:
        ensure_auth_schema(db)
        db.execute(
            """
            INSERT INTO app_sessions (token, user_id, expires_at)
            VALUES (?, ?, ?)
            """,
            (token, user_id, expires),
        )
    return token


def destroy_session(token: str | None) -> None:
    if not token:
        return
    with connect() as db:
        ensure_auth_schema(db)
        db.execute("DELETE FROM app_sessions WHERE token = ?", (token,))


def current_user(token: str | None) -> dict | None:
    if not token:
        return None
    with connect() as db:
        ensure_auth_schema(db)
        row = db.execute(
            """
            SELECT s.token, s.expires_at, u.id, u.username, u.nama, u.role, u.active
            FROM app_sessions s
            JOIN app_users u ON u.id = s.user_id
            WHERE s.token = ?
            """,
            (token,),
        ).fetchone()
        if not row:
            return None
        expires = row["expires_at"]
        if expires is not None and getattr(expires, "tzinfo", None) is None:
            expires = expires.replace(tzinfo=timezone.utc)
        if expires is not None and expires < _now():
            db.execute("DELETE FROM app_sessions WHERE token = ?", (token,))
            return None
        if not row["active"]:
            return None
        db.execute("UPDATE app_sessions SET last_seen = now() WHERE token = ?", (token,))
    return {
        "id": row["id"],
        "username": row["username"],
        "nama": row["nama"] or row["username"],
        "role": row["role"] or "staff",
        "active": True,
        "is_admin": (row["role"] or "") == "admin",
    }


def list_users() -> list[dict]:
    with connect() as db:
        ensure_auth_schema(db)
        rows = db.execute(
            """
            SELECT id, username, nama, role, active, created_at, updated_at
            FROM app_users
            ORDER BY lower(username)
            """
        ).fetchall()
    for row in rows:
        row["role_label"] = dict(ROLES).get(row["role"] or "", row["role"] or "")
        row["active"] = bool(row["active"])
        row["is_admin"] = (row["role"] or "") == "admin"
    return rows


def get_user(user_id: int) -> dict | None:
    with connect() as db:
        ensure_auth_schema(db)
        row = db.execute(
            """
            SELECT id, username, nama, role, active, created_at, updated_at
            FROM app_users WHERE id = ?
            """,
            (user_id,),
        ).fetchone()
    if not row:
        return None
    row["role_label"] = dict(ROLES).get(row["role"] or "", row["role"] or "")
    row["active"] = bool(row["active"])
    row["is_admin"] = (row["role"] or "") == "admin"
    return row


def create_app_user(nama: str, username: str, password: str, role: str = "staff") -> str:
    nama = " ".join((nama or "").split())
    key = _clean_username(username)
    role = (role or "staff").strip().lower()
    if role not in dict(ROLES):
        role = "staff"
    if not nama:
        return "Nama wajib diisi."
    if not key or len(key) < 3:
        return "Username minimal 3 karakter."
    if " " in key:
        return "Username tanpa spasi."
    if len(password or "") < 6:
        return "Password minimal 6 karakter."
    with connect() as db:
        ensure_auth_schema(db)
        exists = db.execute(
            "SELECT id FROM app_users WHERE lower(username) = lower(?)",
            (key,),
        ).fetchone()
        if exists:
            return "Username sudah dipakai."
        db.execute(
            """
            INSERT INTO app_users (username, nama, password_hash, role, active)
            VALUES (?, ?, ?, ?, TRUE)
            """,
            (key, nama, hash_password(password), role),
        )
    return ""


def update_app_user(user_id: int, form, actor_id: int | None = None) -> str:
    person = get_user(user_id)
    if person is None:
        return "Pengguna tidak ditemukan."
    nama = " ".join((form.get("nama") or "").split())
    role = (form.get("role") or person["role"] or "staff").strip().lower()
    active_raw = form.get("active")
    active = True if active_raw in (None, "", "1", "true", "on", "ya") else False
    password = form.get("password") or ""
    if role not in dict(ROLES):
        return "Peran tidak valid."
    if not nama:
        return "Nama wajib diisi."
    if password and len(password) < 6:
        return "Password minimal 6 karakter."
    if actor_id and actor_id == user_id and not active:
        return "Tidak bisa menonaktifkan akun yang sedang dipakai."
    if actor_id and actor_id == user_id and role != "admin" and person["is_admin"]:
        # allow demotion only if another admin exists
        with connect() as db:
            others = db.execute(
                "SELECT COUNT(*) AS n FROM app_users WHERE role = 'admin' AND active AND id <> ?",
                (user_id,),
            ).fetchone()["n"]
        if int(others or 0) == 0:
            return "Harus ada minimal satu admin aktif."
    with connect() as db:
        ensure_auth_schema(db)
        if password:
            db.execute(
                """
                UPDATE app_users
                SET nama = ?, role = ?, active = ?, password_hash = ?, updated_at = now()
                WHERE id = ?
                """,
                (nama, role, active, hash_password(password), user_id),
            )
        else:
            db.execute(
                """
                UPDATE app_users
                SET nama = ?, role = ?, active = ?, updated_at = now()
                WHERE id = ?
                """,
                (nama, role, active, user_id),
            )
        if not active:
            db.execute("DELETE FROM app_sessions WHERE user_id = ?", (user_id,))
    return ""


def delete_app_user(user_id: int, actor_id: int | None = None) -> str:
    person = get_user(user_id)
    if person is None:
        return "Pengguna tidak ditemukan."
    if actor_id and actor_id == user_id:
        return "Tidak bisa menghapus akun sendiri."
    if person["is_admin"]:
        with connect() as db:
            others = db.execute(
                "SELECT COUNT(*) AS n FROM app_users WHERE role = 'admin' AND active AND id <> ?",
                (user_id,),
            ).fetchone()["n"]
        if int(others or 0) == 0:
            return "Harus ada minimal satu admin aktif."
    with connect() as db:
        ensure_auth_schema(db)
        db.execute("DELETE FROM app_users WHERE id = ?", (user_id,))
    return ""
