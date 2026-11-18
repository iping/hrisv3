"""Suntingan ringkasan akhir untuk satu sesi payroll.

Angka slip tetap dari hasil hitung. Yang disimpan di sini hanya perubahan
transaksi: rekening, nominal transfer, catatan, dan baris yang dikeluarkan.
File Accounting membaca catatan yang sama.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from calculate_payroll import text

FIELDS = ("bank", "rekening", "penerima", "remark", "amount", "catatan")


def closing_path(folder: Path) -> Path:
    return folder / "closing.json"


def load_closing(folder: Path) -> dict:
    path = closing_path(folder)
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    people = data.get("people") or {}
    return people if isinstance(people, dict) else {}


def save_closing(folder: Path, people: dict) -> None:
    path = closing_path(folder)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"people": people}, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_amount(raw: str):
    value = text(raw)
    if not value:
        return None
    negative = value.startswith("-")
    digits = re.sub(r"\D", "", value)
    if not digits:
        return None
    amount = int(digits)
    return -amount if negative else amount


def edits_from_form(form, originals: dict[str, dict]) -> dict:
    keys = form.getlist("key")
    banks = form.getlist("bank")
    accounts = form.getlist("rekening")
    receivers = form.getlist("penerima")
    amounts = form.getlist("amount")
    remarks = form.getlist("remark")
    notes = form.getlist("catatan")
    deleted = form.getlist("deleted")
    people = {}
    for index, key in enumerate(keys):
        origin = originals.get(key)
        if origin is None:
            continue
        change = {}
        if index < len(deleted) and deleted[index] == "1":
            change["deleted"] = True
        posted = {
            "bank": text(banks[index]) if index < len(banks) else "",
            "rekening": text(accounts[index]) if index < len(accounts) else "",
            "penerima": text(receivers[index]) if index < len(receivers) else "",
            "remark": text(remarks[index]) if index < len(remarks) else "",
            "catatan": text(notes[index]) if index < len(notes) else "",
        }
        amount = parse_amount(amounts[index]) if index < len(amounts) else None
        if amount is None:
            amount = int(origin["amount"])
        for field, value in posted.items():
            if value != text(origin.get(field)):
                change[field] = value
        if amount != int(origin["amount"]):
            change["amount"] = amount
        if change:
            people[key] = change
    return people
