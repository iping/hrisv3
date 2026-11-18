"""PostgreSQL untuk unggahan, kegiatan, istilah, dan arsip bulan."""

from __future__ import annotations

import os
import re
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import psycopg2
from psycopg2.extras import RealDictCursor

ROOT = Path(__file__).resolve().parent
DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql://test@localhost:5432/hrisv3")
ADMIN_URL = os.environ.get("DATABASE_ADMIN_URL", "postgresql://test@localhost:5432/postgres")

MONTHS = ("Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des")

TERMS = [
    ("Gaji pokok", "Gaji sebelum tunjangan dan potongan. Di file disebut Basic Salary."),
    ("Gaji HRIS", "Gaji bersih di slip: gaji pokok ditambah tunjangan, dikurangi potongan. Di file disebut Paid Salary."),
    ("Hasil hitung", "Gaji yang dihitung ulang di sini. Lembur mengikuti durasi, denda alpa dari gaji pokok dibagi 25."),
    ("Selisih", "Hasil hitung dikurangi gaji HRIS. Angka ini dicek sebelum transfer."),
    ("Lembur", "Fee dari durasi: 1–4 jam Rp 50.000, 5–8 jam Rp 75.000, 9 jam atau lebih Rp 100.000."),
    ("Denda alpa", "Potongan karena tidak hadir tanpa izin. Gaji pokok dibagi 25, dikali jumlah status A."),
    ("Alpa", "Tidak hadir tanpa izin. Kode di absensi: A."),
    ("Terlambat", "Datang lewat dari jadwal. Kode di absensi: L."),
    ("Prorate", "Gaji yang dipotong atau ditambah karena masuk atau keluar di tengah periode."),
    ("Petugas baru", "Tanggal masuk di Officer Profile jatuh di dalam periode payroll ini. Tetap dihitung, supaya prorate bisa dicek."),
    ("Resign", "Petugas sudah keluar. Setelah ditandai, barisnya tidak masuk file transfer pada payroll berikutnya."),
    ("Akun testing", "Akun percobaan. Setelah ditandai, petugas ini tidak masuk file transfer pada payroll berikutnya."),
    ("Mutasi kantor", "Petugas pindah lokasi. Tandai kantor barunya, lalu tandai resign di kantor lama supaya gaji tidak dobel."),
    ("Banding", "Pengajuan perbaikan absensi. Disetujui, menunggu, atau ditolak."),
    ("Slip", "Rincian gaji satu petugas: pokok, tunjangan, lembur, dan potongan."),
    ("Transfer", "Daftar yang diunggah ke bank: rekening, nama penerima, nominal, dan remark."),
    ("Remark", "Catatan di mutasi bank, biasanya kode lokasi ditambah bulan."),
    ("Tunjangan", "Penambahan gaji. Di file, kolom ini nilainya positif."),
    ("Potongan", "Pengurangan gaji. Di file, kolom ini nilainya negatif."),
    ("Hadir", "Masuk sesuai shift. Kode di absensi: OT."),
    ("Izin", "Tidak masuk dengan izin. Kode di absensi: P."),
    ("Sakit", "Tidak masuk karena sakit. Kode di absensi: S."),
    ("Cuti", "Tidak masuk karena cuti. Kode di absensi: LEAVE."),
    ("Lokasi utama", "Lokasi utama petugas di Officer Profile."),
    ("Trial balance", "Ringkasan total: membandingkan angka dari file HRIS dengan hasil hitung di aplikasi."),
    ("Deposit", "Potongan yang sudah mengurangi gaji dihitung, biasanya dari saldo deposit petugas."),
    ("New Joiner", "Petugas yang absensi pertamanya jatuh di tengah periode payroll, jadi gajinya di-prorate."),
    ("Syarat", "Catatan, pengurangan, atau penambahan manual yang disimpan untuk payroll petugas ini."),
    ("Accounting", "Lembar transfer bank: rekening, nama penerima, nominal, dan remark."),
    ("Rekap", "Jumlah transfer digroup per remark, untuk cek total sebelum diunggah ke bank."),
    ("Tidak dihitung", "Petugas resign, blacklist, atau akun testing yang sengaja tidak masuk payroll bulan ini."),
    (
        "Normal Case",
        "Petugas yang angkanya sudah cocok dan tidak perlu dicek lagi: tanpa selisih, tanpa lembur, tanpa banding, tanpa potongan/deposit/denda, dan bukan New Joiner.",
    ),
]


class Result:
    def __init__(self, cursor):
        self.cursor = cursor

    def fetchone(self):
        row = self.cursor.fetchone()
        return dict(row) if row else None

    def fetchall(self):
        return [dict(row) for row in self.cursor.fetchall()]


class Conn:
    def __init__(self, raw):
        self.raw = raw

    def execute(self, sql: str, params=None) -> Result:
        cursor = self.raw.cursor(cursor_factory=RealDictCursor)
        cursor.execute(_adapt(sql, params), params)
        return Result(cursor)

    def executemany(self, sql: str, seq) -> None:
        rows = list(seq)
        if not rows:
            return
        cursor = self.raw.cursor()
        cursor.executemany(_adapt(sql, rows[0]), rows)

    def commit(self) -> None:
        self.raw.commit()


def _adapt(sql: str, params) -> str:
    if isinstance(params, dict):
        return re.sub(r":([A-Za-z_][A-Za-z0-9_]*)", r"%(\1)s", sql)
    return sql.replace("?", "%s")


def ensure_database() -> None:
    try:
        probe = psycopg2.connect(DATABASE_URL)
        probe.close()
        return
    except psycopg2.OperationalError as exc:
        if "does not exist" not in str(exc):
            raise
    admin = psycopg2.connect(ADMIN_URL)
    admin.autocommit = True
    name = DATABASE_URL.rsplit("/", 1)[-1].split("?")[0]
    with admin.cursor() as cursor:
        cursor.execute(f'CREATE DATABASE "{name}"')
    admin.close()


def ensure_schema(db: Conn) -> None:
    missing_payer = db.execute(
        """
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'attendance_benefits' AND column_name = 'penanggung'
        """
    ).fetchone() is None
    missing_lanjut = db.execute(
        """
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'attendance_officers' AND column_name = 'gaji_lanjut'
        """
    ).fetchone() is None
    missing_grace = db.execute(
        """
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'attendance_locations' AND column_name = 'grace_minutes'
        """
    ).fetchone() is None
    missing_pay_profile = db.execute(
        """
        SELECT 1 FROM information_schema.tables
        WHERE table_name = 'attendance_pay_profiles'
        """
    ).fetchone() is None
    missing_attendance = db.execute(
        """
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'attendance_officers' AND column_name = 'attendance_enabled'
        """
    ).fetchone() is None
    db.raw.cursor().execute(
        """
        CREATE TABLE IF NOT EXISTS months (
            id SERIAL PRIMARY KEY,
            period_start TEXT NOT NULL,
            period_end TEXT NOT NULL,
            source_name TEXT,
            source_mtime TEXT,
            saved_at TEXT,
            officer_count INTEGER,
            hris_total INTEGER,
            hitung_total INTEGER,
            baru_count INTEGER,
            resign_count INTEGER,
            appeal_count INTEGER,
            approved_count INTEGER,
            pending_count INTEGER,
            UNIQUE (period_start, period_end)
        );
        CREATE TABLE IF NOT EXISTS officers (
            person_key TEXT PRIMARY KEY,
            numeric_id TEXT,
            nama TEXT,
            username TEXT,
            lokasi TEXT,
            bank TEXT,
            rekening TEXT,
            phone TEXT,
            email TEXT,
            joined_date TEXT,
            resigned_date TEXT,
            activated TEXT,
            updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS payroll_lines (
            month_id INTEGER NOT NULL REFERENCES months(id) ON DELETE CASCADE,
            person_key TEXT NOT NULL,
            lokasi TEXT,
            nama TEXT,
            username TEXT,
            numeric_id TEXT,
            basic INTEGER,
            paid INTEGER,
            hitung INTEGER,
            selisih INTEGER,
            alpha INTEGER,
            status TEXT,
            joined_date TEXT,
            resigned_date TEXT,
            activated TEXT,
            PRIMARY KEY (month_id, person_key)
        );
        CREATE TABLE IF NOT EXISTS appeals (
            id SERIAL PRIMARY KEY,
            month_id INTEGER NOT NULL REFERENCES months(id) ON DELETE CASCADE,
            person_key TEXT,
            lokasi TEXT,
            nama TEXT,
            username TEXT,
            submitted_at TEXT,
            attendance_date TEXT,
            label TEXT,
            reason TEXT,
            detail TEXT,
            status TEXT,
            before_status TEXT,
            after_status TEXT,
            sequence INTEGER
        );
        CREATE TABLE IF NOT EXISTS uploads (
            id TEXT PRIMARY KEY,
            alias TEXT NOT NULL,
            uploaded_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            period_start TEXT,
            period_end TEXT,
            source_name TEXT
        );
        CREATE TABLE IF NOT EXISTS activities (
            id SERIAL PRIMARY KEY,
            upload_id TEXT REFERENCES uploads(id) ON DELETE CASCADE,
            kind TEXT NOT NULL,
            title TEXT NOT NULL,
            detail TEXT NOT NULL DEFAULT '',
            created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            updated_at TIMESTAMPTZ
        );
        CREATE TABLE IF NOT EXISTS terms (
            name TEXT PRIMARY KEY,
            description TEXT NOT NULL
        );
        """
    )
    db.raw.cursor().execute(
        """
        CREATE TABLE IF NOT EXISTS attendance_locations (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            active BOOLEAN NOT NULL DEFAULT TRUE,
            payable BOOLEAN NOT NULL DEFAULT TRUE,
            grace_minutes INTEGER NOT NULL DEFAULT 15
        );
        CREATE TABLE IF NOT EXISTS attendance_officers (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            nama TEXT NOT NULL,
            phone TEXT NOT NULL DEFAULT '',
            email TEXT NOT NULL DEFAULT '',
            bank TEXT NOT NULL DEFAULT '',
            rekening TEXT NOT NULL DEFAULT '',
            joined_date DATE,
            resigned_date DATE,
            testing BOOLEAN NOT NULL DEFAULT FALSE,
            gaji_pokok INTEGER NOT NULL DEFAULT 0,
            employment_status TEXT NOT NULL DEFAULT 'aktif',
            onboard_state TEXT NOT NULL DEFAULT 'disetujui',
            pending_location_id INTEGER,
            tunjangan_jabatan INTEGER NOT NULL DEFAULT 0,
            pemilik_rekening TEXT NOT NULL DEFAULT '',
            pemilik_sendiri BOOLEAN NOT NULL DEFAULT TRUE
        );
        CREATE TABLE IF NOT EXISTS attendance_placements (
            id SERIAL PRIMARY KEY,
            officer_id INTEGER NOT NULL REFERENCES attendance_officers(id) ON DELETE CASCADE,
            location_id INTEGER NOT NULL REFERENCES attendance_locations(id) ON DELETE CASCADE,
            start_date DATE NOT NULL,
            end_date DATE,
            UNIQUE (officer_id, location_id, start_date)
        );
        CREATE TABLE IF NOT EXISTS attendance_shifts (
            id SERIAL PRIMARY KEY,
            location_id INTEGER NOT NULL REFERENCES attendance_locations(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            start_time TIME NOT NULL,
            end_time TIME NOT NULL,
            grace_minutes INTEGER,
            ends_next_day BOOLEAN NOT NULL DEFAULT FALSE
        );
        ALTER TABLE attendance_shifts ADD COLUMN IF NOT EXISTS ends_next_day BOOLEAN NOT NULL DEFAULT FALSE;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS min_attend_minutes INTEGER;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS min_work_days INTEGER;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS min_off_days INTEGER;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS gender TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS marital_status TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS attendance_enabled BOOLEAN NOT NULL DEFAULT FALSE;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS activated_on DATE;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS status_on DATE;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS secondary_location_id INTEGER REFERENCES attendance_locations(id);
        CREATE TABLE IF NOT EXISTS attendance_user_profiles (
            officer_id INTEGER PRIMARY KEY REFERENCES attendance_officers(id) ON DELETE CASCADE,
            username TEXT NOT NULL,
            primary_location_id INTEGER REFERENCES attendance_locations(id),
            secondary_location_id INTEGER REFERENCES attendance_locations(id),
            basic_salary INTEGER NOT NULL DEFAULT 0,
            bpjs_kesehatan INTEGER NOT NULL DEFAULT 0,
            bpjs_ketenagakerjaan INTEGER NOT NULL DEFAULT 0,
            joined_date DATE,
            resigned_date DATE,
            updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
        );
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS posisi TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS max_attend_minutes INTEGER;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS ojt_months INTEGER;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS ojt_start DATE;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS ojt_decision TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS gaji_ojt INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS jabatan_ojt INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS gaji_lanjut INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS jabatan_lanjut INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE attendance_benefits ADD COLUMN IF NOT EXISTS struktur TEXT NOT NULL DEFAULT 'lanjut';
        ALTER TABLE attendance_benefits ADD COLUMN IF NOT EXISTS jenis TEXT NOT NULL DEFAULT 'lain';
        ALTER TABLE attendance_benefits ADD COLUMN IF NOT EXISTS penanggung TEXT NOT NULL DEFAULT 'perusahaan';
        ALTER TABLE attendance_benefits ALTER COLUMN nilai TYPE NUMERIC(14, 2) USING nilai::numeric;
        ALTER TABLE attendance_slots ADD COLUMN IF NOT EXISTS off_reason TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_locations ADD COLUMN IF NOT EXISTS grace_minutes INTEGER NOT NULL DEFAULT 15;
        ALTER TABLE attendance_locations ADD COLUMN IF NOT EXISTS min_attend_minutes INTEGER;
        ALTER TABLE attendance_locations ADD COLUMN IF NOT EXISTS min_work_days INTEGER;
        ALTER TABLE attendance_locations ADD COLUMN IF NOT EXISTS min_off_days INTEGER;
        ALTER TABLE attendance_locations ADD COLUMN IF NOT EXISTS latitude NUMERIC(10, 7);
        ALTER TABLE attendance_locations ADD COLUMN IF NOT EXISTS longitude NUMERIC(10, 7);
        ALTER TABLE attendance_locations ADD COLUMN IF NOT EXISTS gps_radius_meters INTEGER;
        ALTER TABLE attendance_shifts ALTER COLUMN grace_minutes DROP NOT NULL;
        ALTER TABLE attendance_shifts ALTER COLUMN grace_minutes DROP DEFAULT;
        CREATE TABLE IF NOT EXISTS attendance_slots (
            id SERIAL PRIMARY KEY,
            officer_id INTEGER NOT NULL REFERENCES attendance_officers(id) ON DELETE CASCADE,
            location_id INTEGER NOT NULL REFERENCES attendance_locations(id) ON DELETE CASCADE,
            work_date DATE NOT NULL,
            shift_id INTEGER REFERENCES attendance_shifts(id) ON DELETE SET NULL,
            slot_no SMALLINT NOT NULL DEFAULT 1,
            planned_code TEXT,
            clock_in TIME,
            clock_out TIME,
            code TEXT NOT NULL DEFAULT '',
            calculated_code TEXT NOT NULL DEFAULT '',
            code_source TEXT NOT NULL DEFAULT 'hitung',
            off_reason TEXT NOT NULL DEFAULT '',
            UNIQUE (officer_id, location_id, work_date, slot_no)
        );
        CREATE TABLE IF NOT EXISTS attendance_banding (
            id SERIAL PRIMARY KEY,
            slot_id INTEGER NOT NULL REFERENCES attendance_slots(id) ON DELETE CASCADE,
            requested_code TEXT NOT NULL,
            reason TEXT NOT NULL DEFAULT '',
            detail TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT 'pending',
            before_code TEXT NOT NULL DEFAULT '',
            after_code TEXT,
            parent_id INTEGER REFERENCES attendance_banding(id),
            sequence INTEGER NOT NULL DEFAULT 1,
            decided_note TEXT NOT NULL DEFAULT '',
            created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            decided_at TIMESTAMPTZ
        );
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS gaji_pokok INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS employment_status TEXT NOT NULL DEFAULT 'aktif';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS onboard_state TEXT NOT NULL DEFAULT 'disetujui';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS pending_location_id INTEGER;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS bank TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS rekening TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS tunjangan_jabatan INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS pemilik_rekening TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_officers ADD COLUMN IF NOT EXISTS pemilik_sendiri BOOLEAN NOT NULL DEFAULT TRUE;
        CREATE TABLE IF NOT EXISTS attendance_shift_months (
            id SERIAL PRIMARY KEY,
            officer_id INTEGER NOT NULL REFERENCES attendance_officers(id) ON DELETE CASCADE,
            location_id INTEGER NOT NULL REFERENCES attendance_locations(id) ON DELETE CASCADE,
            shift_id INTEGER NOT NULL REFERENCES attendance_shifts(id) ON DELETE CASCADE,
            year_month DATE NOT NULL,
            UNIQUE (officer_id, location_id, year_month)
        );
        CREATE TABLE IF NOT EXISTS attendance_pay_profiles (
            id SERIAL PRIMARY KEY,
            officer_id INTEGER NOT NULL REFERENCES attendance_officers(id) ON DELETE CASCADE,
            location_id INTEGER NOT NULL REFERENCES attendance_locations(id),
            active BOOLEAN NOT NULL DEFAULT FALSE,
            gaji_ojt INTEGER NOT NULL DEFAULT 0,
            jabatan_ojt INTEGER NOT NULL DEFAULT 0,
            gaji_lanjut INTEGER NOT NULL DEFAULT 0,
            jabatan_lanjut INTEGER NOT NULL DEFAULT 0,
            UNIQUE (officer_id, location_id)
        );
        ALTER TABLE attendance_benefits ADD COLUMN IF NOT EXISTS pay_profile_id INTEGER REFERENCES attendance_pay_profiles(id) ON DELETE CASCADE;
        ALTER TABLE attendance_payroll_conditions ADD COLUMN IF NOT EXISTS potongan INTEGER;
        ALTER TABLE attendance_payroll_conditions ADD COLUMN IF NOT EXISTS potongan_reason TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_payroll_conditions ADD COLUMN IF NOT EXISTS tambahan INTEGER;
        ALTER TABLE attendance_payroll_conditions ADD COLUMN IF NOT EXISTS tambahan_reason TEXT NOT NULL DEFAULT '';
        CREATE TABLE IF NOT EXISTS attendance_change_log (
            id SERIAL PRIMARY KEY,
            officer_id INTEGER NOT NULL REFERENCES attendance_officers(id) ON DELETE CASCADE,
            bagian TEXT NOT NULL,
            field_label TEXT NOT NULL DEFAULT '',
            before_text TEXT NOT NULL DEFAULT '',
            after_text TEXT NOT NULL DEFAULT '',
            created_at TIMESTAMPTZ NOT NULL DEFAULT now()
        );
        CREATE TABLE IF NOT EXISTS attendance_location_log (
            id SERIAL PRIMARY KEY,
            location_id INTEGER NOT NULL REFERENCES attendance_locations(id) ON DELETE CASCADE,
            bagian TEXT NOT NULL,
            field_label TEXT NOT NULL DEFAULT '',
            before_text TEXT NOT NULL DEFAULT '',
            after_text TEXT NOT NULL DEFAULT '',
            created_at TIMESTAMPTZ NOT NULL DEFAULT now()
        );
        CREATE TABLE IF NOT EXISTS attendance_payroll_conditions (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            officer_id INTEGER REFERENCES attendance_officers(id) ON DELETE SET NULL,
            deposit INTEGER,
            potongan INTEGER,
            potongan_reason TEXT NOT NULL DEFAULT '',
            tambahan INTEGER,
            tambahan_reason TEXT NOT NULL DEFAULT '',
            excluded BOOLEAN NOT NULL DEFAULT FALSE,
            note TEXT NOT NULL DEFAULT '',
            updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
        );
        CREATE TABLE IF NOT EXISTS attendance_payroll_adjustment_log (
            id SERIAL PRIMARY KEY,
            officer_id INTEGER REFERENCES attendance_officers(id) ON DELETE CASCADE,
            username TEXT NOT NULL DEFAULT '',
            kind TEXT NOT NULL,
            reason TEXT NOT NULL DEFAULT '',
            amount INTEGER NOT NULL DEFAULT 0,
            note TEXT NOT NULL DEFAULT '',
            created_at TIMESTAMPTZ NOT NULL DEFAULT now()
        );
        CREATE TABLE IF NOT EXISTS attendance_benefits (
            id SERIAL PRIMARY KEY,
            officer_id INTEGER NOT NULL REFERENCES attendance_officers(id) ON DELETE CASCADE,
            nama TEXT NOT NULL,
            ikut BOOLEAN NOT NULL DEFAULT TRUE,
            mode TEXT NOT NULL DEFAULT 'nominal',
            nilai INTEGER NOT NULL DEFAULT 0
        );
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
        ALTER TABLE activities ADD COLUMN IF NOT EXISTS actor_user_id INTEGER REFERENCES app_users(id) ON DELETE SET NULL;
        ALTER TABLE activities ADD COLUMN IF NOT EXISTS actor_username TEXT NOT NULL DEFAULT '';
        ALTER TABLE activities ADD COLUMN IF NOT EXISTS actor_nama TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_change_log ADD COLUMN IF NOT EXISTS actor_user_id INTEGER REFERENCES app_users(id) ON DELETE SET NULL;
        ALTER TABLE attendance_change_log ADD COLUMN IF NOT EXISTS actor_username TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_change_log ADD COLUMN IF NOT EXISTS actor_nama TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_payroll_adjustment_log ADD COLUMN IF NOT EXISTS actor_user_id INTEGER REFERENCES app_users(id) ON DELETE SET NULL;
        ALTER TABLE attendance_payroll_adjustment_log ADD COLUMN IF NOT EXISTS actor_username TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_payroll_adjustment_log ADD COLUMN IF NOT EXISTS actor_nama TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_location_log ADD COLUMN IF NOT EXISTS actor_user_id INTEGER REFERENCES app_users(id) ON DELETE SET NULL;
        ALTER TABLE attendance_location_log ADD COLUMN IF NOT EXISTS actor_username TEXT NOT NULL DEFAULT '';
        ALTER TABLE attendance_location_log ADD COLUMN IF NOT EXISTS actor_nama TEXT NOT NULL DEFAULT '';
        """
    )
    db.execute(
        """
        UPDATE attendance_benefits
        SET jenis = 'bpjs_kes', nama = 'BPJS Kesehatan'
        WHERE jenis = 'kesehatan'
        """
    )
    db.execute(
        """
        UPDATE attendance_benefits
        SET jenis = 'bpjs_tk', nama = 'BPJS Ketenagakerjaan'
        WHERE jenis = 'ketenagakerjaan'
        """
    )
    db.execute(
        """
        UPDATE attendance_benefits
        SET penanggung = 'deduction'
        WHERE jenis IN ('bpjs_tk', 'bpjs_kes') AND penanggung = 'overhead'
        """
    )
    folded = db.execute(
        """
        SELECT officer_id, struktur, COALESCE(SUM(nilai), 0) AS nilai
        FROM attendance_benefits
        WHERE jenis IN ('insentif_operasional', 'insentif_operational')
        GROUP BY officer_id, struktur
        """
    ).fetchall()
    for row in folded:
        current = db.execute(
            """
            SELECT id, nilai FROM attendance_benefits
            WHERE officer_id = ? AND struktur = ? AND jenis = 'insentif'
            ORDER BY id LIMIT 1
            """,
            (row["officer_id"], row["struktur"]),
        ).fetchone()
        if current:
            db.execute(
                "UPDATE attendance_benefits SET nama = 'Insentif', nilai = ? WHERE id = ?",
                (float(current["nilai"] or 0) + float(row["nilai"] or 0), current["id"]),
            )
        else:
            db.execute(
                """
                INSERT INTO attendance_benefits (officer_id, nama, ikut, mode, nilai, struktur, jenis, penanggung)
                VALUES (?, 'Insentif', TRUE, 'nominal', ?, ?, 'insentif', 'perusahaan')
                """,
                (row["officer_id"], row["nilai"], row["struktur"]),
            )
        db.execute(
            """
            DELETE FROM attendance_benefits
            WHERE officer_id = ? AND struktur = ? AND jenis IN ('insentif_operasional', 'insentif_operational')
            """,
            (row["officer_id"], row["struktur"]),
        )
    if missing_grace:
        db.execute("UPDATE attendance_shifts SET grace_minutes = NULL WHERE grace_minutes = 15")
    if missing_lanjut:
        db.execute(
            """
            UPDATE attendance_officers
            SET gaji_lanjut = gaji_pokok, jabatan_lanjut = tunjangan_jabatan
            """
        )
    if missing_payer:
        db.execute(
            """
            UPDATE attendance_benefits
            SET jenis = CASE
                WHEN lower(nama) LIKE '%kesehatan%' OR lower(nama) LIKE '%bpjs kes%' THEN 'kesehatan'
                WHEN lower(nama) LIKE '%ketenagakerjaan%' OR lower(nama) LIKE '%bpjs tk%'
                     OR lower(nama) LIKE '%bpjs ketenaga%' THEN 'ketenagakerjaan'
                WHEN lower(nama) LIKE '%jabatan%' THEN 'jabatan'
                ELSE 'lain'
            END
            """
        )
        db.execute(
            """
            INSERT INTO attendance_benefits (officer_id, nama, ikut, mode, nilai, struktur, jenis, penanggung)
            SELECT id, 'Tunjangan Jabatan', TRUE, 'nominal', jabatan_ojt, 'ojt', 'jabatan', 'perusahaan'
            FROM attendance_officers
            WHERE jabatan_ojt > 0
              AND NOT EXISTS (
                SELECT 1 FROM attendance_benefits b
                WHERE b.officer_id = attendance_officers.id
                  AND b.struktur = 'ojt' AND b.jenis = 'jabatan'
              )
            """
        )
        db.execute(
            """
            INSERT INTO attendance_benefits (officer_id, nama, ikut, mode, nilai, struktur, jenis, penanggung)
            SELECT id, 'Tunjangan Jabatan', TRUE, 'nominal', jabatan_lanjut, 'lanjut', 'jabatan', 'perusahaan'
            FROM attendance_officers
            WHERE jabatan_lanjut > 0
              AND NOT EXISTS (
                SELECT 1 FROM attendance_benefits b
                WHERE b.officer_id = attendance_officers.id
                  AND b.struktur = 'lanjut' AND b.jenis = 'jabatan'
              )
            """
        )
    if missing_pay_profile:
        db.execute(
            """
            INSERT INTO attendance_pay_profiles
                (officer_id, location_id, active, gaji_ojt, jabatan_ojt, gaji_lanjut, jabatan_lanjut)
            SELECT o.id, p.location_id, TRUE,
                   o.gaji_ojt, o.jabatan_ojt, o.gaji_lanjut, o.jabatan_lanjut
            FROM attendance_officers o
            JOIN LATERAL (
                SELECT location_id FROM attendance_placements
                WHERE officer_id = o.id
                ORDER BY start_date DESC
                LIMIT 1
            ) p ON TRUE
            WHERE NOT EXISTS (
                SELECT 1 FROM attendance_pay_profiles x
                WHERE x.officer_id = o.id AND x.location_id = p.location_id
            )
            """
        )
        db.execute(
            """
            UPDATE attendance_benefits b
            SET pay_profile_id = p.id
            FROM attendance_pay_profiles p
            WHERE b.pay_profile_id IS NULL
              AND p.officer_id = b.officer_id
              AND p.active
            """
        )
    if missing_attendance:
        db.execute(
            """
            UPDATE attendance_officers
            SET attendance_enabled = TRUE,
                activated_on = COALESCE(activated_on, joined_date)
            WHERE onboard_state = 'disetujui'
              AND employment_status IN ('aktif', 'ojt', 'lulus', 'mutasi', 'tunjangan', 'baru')
            """
        )
    db.execute(
        """
        INSERT INTO attendance_user_profiles (
            officer_id, username, primary_location_id, secondary_location_id,
            basic_salary, bpjs_kesehatan, bpjs_ketenagakerjaan, joined_date, resigned_date
        )
        SELECT o.id, o.username,
               COALESCE(p.location_id, o.pending_location_id),
               o.secondary_location_id,
               COALESCE(pay.gaji_ojt, o.gaji_ojt, 0)::int,
               COALESCE(kes.nilai, 0)::int,
               COALESCE(tk.nilai, 0)::int,
               o.joined_date, o.resigned_date
        FROM attendance_officers o
        LEFT JOIN LATERAL (
            SELECT location_id FROM attendance_placements
            WHERE officer_id = o.id
            ORDER BY (end_date IS NULL) DESC, start_date DESC
            LIMIT 1
        ) p ON TRUE
        LEFT JOIN LATERAL (
            SELECT gaji_ojt FROM attendance_pay_profiles
            WHERE officer_id = o.id AND active
            ORDER BY id DESC LIMIT 1
        ) pay ON TRUE
        LEFT JOIN LATERAL (
            SELECT b.nilai FROM attendance_benefits b
            JOIN attendance_pay_profiles pp ON pp.id = b.pay_profile_id
            WHERE pp.officer_id = o.id AND pp.active AND b.struktur = 'ojt' AND b.jenis = 'bpjs_kes'
            ORDER BY b.id LIMIT 1
        ) kes ON TRUE
        LEFT JOIN LATERAL (
            SELECT b.nilai FROM attendance_benefits b
            JOIN attendance_pay_profiles pp ON pp.id = b.pay_profile_id
            WHERE pp.officer_id = o.id AND pp.active AND b.struktur = 'ojt' AND b.jenis = 'bpjs_tk'
            ORDER BY b.id LIMIT 1
        ) tk ON TRUE
        WHERE NOT EXISTS (
            SELECT 1 FROM attendance_user_profiles u WHERE u.officer_id = o.id
        )
        """
    )
    db.commit()


@contextmanager
def connect():
    ensure_database()
    raw = psycopg2.connect(DATABASE_URL)
    raw.autocommit = False
    db = Conn(raw)
    try:
        raw.cursor().execute("SET TIME ZONE 'Asia/Jakarta'")
        ensure_schema(db)
        yield db
        raw.commit()
    except Exception:
        raw.rollback()
        raise
    finally:
        raw.close()


def stamp(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        value = datetime.fromisoformat(value)
    return f"{value.day} {MONTHS[value.month - 1]} {value.year}, {value.strftime('%H.%M')}"


def alias_label(alias: str, uploaded_at) -> str:
    when = stamp(uploaded_at)
    return f"{alias} · {when}" if when else alias


def remember_upload(sid: str, alias: str, period_start: str, period_end: str, source_name: str, uploaded_at=None) -> None:
    with connect() as db:
        db.execute(
            """
            INSERT INTO uploads (id, alias, uploaded_at, period_start, period_end, source_name)
            VALUES (?, ?, COALESCE(?::timestamptz, now()), ?, ?, ?)
            ON CONFLICT (id) DO UPDATE SET
                alias = EXCLUDED.alias,
                period_start = EXCLUDED.period_start,
                period_end = EXCLUDED.period_end,
                source_name = EXCLUDED.source_name
            """,
            (sid, alias, uploaded_at, period_start, period_end, source_name),
        )


def list_uploads() -> list[dict]:
    with connect() as db:
        rows = db.execute(
            "SELECT * FROM uploads ORDER BY uploaded_at DESC"
        ).fetchall()
    for row in rows:
        row["label"] = alias_label(row["alias"], row["uploaded_at"])
        row["when"] = stamp(row["uploaded_at"])
    return rows


def get_upload(sid: str) -> dict | None:
    with connect() as db:
        row = db.execute("SELECT * FROM uploads WHERE id = ?", (sid,)).fetchone()
    if row is None:
        return None
    row["label"] = alias_label(row["alias"], row["uploaded_at"])
    row["when"] = stamp(row["uploaded_at"])
    return row


def rename_upload(sid: str, alias: str) -> None:
    with connect() as db:
        db.execute("UPDATE uploads SET alias = ? WHERE id = ?", (alias.strip(), sid))


def delete_upload(sid: str) -> None:
    with connect() as db:
        db.execute("DELETE FROM uploads WHERE id = ?", (sid,))


def add_activity(
    kind: str,
    title: str,
    detail: str = "",
    upload_id: str | None = None,
    actor: dict | None = None,
) -> int:
    actor = actor or {}
    with connect() as db:
        row = db.execute(
            """
            INSERT INTO activities (
                upload_id, kind, title, detail,
                actor_user_id, actor_username, actor_nama
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            RETURNING id
            """,
            (
                upload_id,
                kind,
                title,
                detail or "",
                actor.get("id"),
                (actor.get("username") or "").strip(),
                (actor.get("nama") or actor.get("username") or "").strip(),
            ),
        ).fetchone()
    return row["id"]


def list_activities(upload_id: str = "") -> list[dict]:
    with connect() as db:
        if upload_id:
            rows = db.execute(
                """
                SELECT a.*, u.alias
                FROM activities a
                LEFT JOIN uploads u ON u.id = a.upload_id
                WHERE a.upload_id = ?
                ORDER BY a.created_at DESC, a.id DESC
                """,
                (upload_id,),
            ).fetchall()
        else:
            rows = db.execute(
                """
                SELECT a.*, u.alias
                FROM activities a
                LEFT JOIN uploads u ON u.id = a.upload_id
                ORDER BY a.created_at DESC, a.id DESC
                """
            ).fetchall()
    for row in rows:
        row["when"] = stamp(row["created_at"])
        nama = (row.get("actor_nama") or "").strip()
        username = (row.get("actor_username") or "").strip()
        if nama and username and nama.lower() != username.lower():
            row["actor_label"] = f"{nama} ({username})"
        else:
            row["actor_label"] = nama or username or "—"
    return rows


def get_activity(activity_id: int) -> dict | None:
    with connect() as db:
        return db.execute("SELECT * FROM activities WHERE id = ?", (activity_id,)).fetchone()


def update_activity(activity_id: int, title: str, detail: str, actor: dict | None = None) -> None:
    actor = actor or {}
    with connect() as db:
        if actor.get("id"):
            db.execute(
                """
                UPDATE activities
                SET title = ?, detail = ?, updated_at = now(),
                    actor_user_id = ?, actor_username = ?, actor_nama = ?
                WHERE id = ?
                """,
                (
                    title.strip(),
                    detail.strip(),
                    actor.get("id"),
                    (actor.get("username") or "").strip(),
                    (actor.get("nama") or actor.get("username") or "").strip(),
                    activity_id,
                ),
            )
        else:
            db.execute(
                "UPDATE activities SET title = ?, detail = ?, updated_at = now() WHERE id = ?",
                (title.strip(), detail.strip(), activity_id),
            )


def delete_activity(activity_id: int) -> None:
    with connect() as db:
        db.execute("DELETE FROM activities WHERE id = ?", (activity_id,))


def seed_terms(pairs: list[tuple[str, str]]) -> None:
    with connect() as db:
        db.executemany(
            """
            INSERT INTO terms (name, description) VALUES (?, ?)
            ON CONFLICT (name) DO UPDATE SET description = EXCLUDED.description
            """,
            pairs,
        )


def term_map() -> dict[str, str]:
    with connect() as db:
        rows = db.execute("SELECT name, description FROM terms").fetchall()
    return {row["name"].lower(): row["description"] for row in rows}


def all_terms() -> list[dict]:
    with connect() as db:
        return db.execute("SELECT name, description FROM terms ORDER BY name").fetchall()
