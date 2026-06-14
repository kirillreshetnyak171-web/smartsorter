"""Локальный архив черновиков (папка пользователя)."""

from __future__ import annotations

import json
import os
import shutil
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from user_data import ensure_user_data_ready

MAX_ENTRIES = 50

DRAFT_FILE = "draft.csv"
ORIGINAL_FILE = "original.csv"
META_FILE = "meta.json"


@dataclass
class HistoryMeta:
    id: str
    created_at: str
    source_name: str
    source_path: str
    task_preview: str
    task_text: str
    template_index: int | None
    lang: str
    rows_before: int
    rows_after: int
    cells_changed: int
    report_summary: str
    report_details: str


@dataclass
class HistoryRecord:
    meta: HistoryMeta
    entry_dir: str


def _history_root() -> str:
    ensure_user_data_ready()
    from user_data import HISTORY_ROOT
    return HISTORY_ROOT


def _entries_dir() -> str:
    return os.path.join(_history_root(), "entries")


def _db_path() -> str:
    return os.path.join(_history_root(), "history.db")


def _ensure_dirs() -> None:
    os.makedirs(_entries_dir(), exist_ok=True)


def _connect() -> sqlite3.Connection:
    _ensure_dirs()
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS history (
            id TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            source_name TEXT NOT NULL,
            task_preview TEXT NOT NULL,
            entry_dir TEXT NOT NULL
        )
        """
    )
    conn.commit()
    return conn


def _new_entry_id() -> str:
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    return f"{stamp}_{uuid.uuid4().hex[:6]}"


def _meta_from_dict(data: dict) -> HistoryMeta:
    return HistoryMeta(
        id=data["id"],
        created_at=data["created_at"],
        source_name=data["source_name"],
        source_path=data.get("source_path", ""),
        task_preview=data.get("task_preview", ""),
        task_text=data.get("task_text", data.get("task_preview", "")),
        template_index=data.get("template_index"),
        lang=data.get("lang", "ru"),
        rows_before=int(data.get("rows_before", 0)),
        rows_after=int(data.get("rows_after", 0)),
        cells_changed=int(data.get("cells_changed", 0)),
        report_summary=data.get("report_summary", ""),
        report_details=data.get("report_details", ""),
    )


def _read_meta(entry_dir: str) -> HistoryMeta:
    path = os.path.join(entry_dir, META_FILE)
    with open(path, encoding="utf-8") as fh:
        return _meta_from_dict(json.load(fh))


def _prune_old(conn: sqlite3.Connection) -> None:
    rows = conn.execute(
        "SELECT id, entry_dir FROM history ORDER BY created_at DESC",
    ).fetchall()
    for row in rows[MAX_ENTRIES:]:
        entry_dir = row["entry_dir"]
        if os.path.isdir(entry_dir):
            shutil.rmtree(entry_dir, ignore_errors=True)
        conn.execute("DELETE FROM history WHERE id = ?", (row["id"],))
    conn.commit()


def save_history_entry(
    *,
    source_path: str,
    task: str,
    lang: str,
    original_df: pd.DataFrame,
    draft_df: pd.DataFrame,
    report_summary: str,
    report_details: str,
    cells_changed: int,
    template_index: int | None = None,
) -> str:
    entry_id = _new_entry_id()
    entry_dir = os.path.join(_entries_dir(), entry_id)
    os.makedirs(entry_dir, exist_ok=True)

    source_name = os.path.basename(source_path) if source_path else "—"
    task_text = task.strip()
    task_preview = " ".join(task_text.split())
    if len(task_preview) > 240:
        task_preview = task_preview[:237] + "…"

    meta = HistoryMeta(
        id=entry_id,
        created_at=datetime.now().isoformat(timespec="seconds"),
        source_name=source_name,
        source_path=source_path or "",
        task_preview=task_preview,
        task_text=task_text,
        template_index=template_index,
        lang=lang,
        rows_before=len(original_df),
        rows_after=len(draft_df),
        cells_changed=cells_changed,
        report_summary=report_summary,
        report_details=report_details,
    )

    draft_df.to_csv(os.path.join(entry_dir, DRAFT_FILE), index=False, encoding="utf-8")
    original_df.to_csv(
        os.path.join(entry_dir, ORIGINAL_FILE), index=False, encoding="utf-8",
    )
    with open(os.path.join(entry_dir, META_FILE), "w", encoding="utf-8") as fh:
        json.dump(meta.__dict__, fh, ensure_ascii=False, indent=2)

    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO history (id, created_at, source_name, task_preview, entry_dir)
            VALUES (?, ?, ?, ?, ?)
            """,
            (entry_id, meta.created_at, source_name, task_preview, entry_dir),
        )
        conn.commit()
        _prune_old(conn)
    finally:
        conn.close()
    return entry_id


def list_history_entries() -> list[HistoryRecord]:
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT entry_dir FROM history ORDER BY created_at DESC",
        ).fetchall()
    finally:
        conn.close()

    records: list[HistoryRecord] = []
    for row in rows:
        entry_dir = row["entry_dir"]
        if not os.path.isdir(entry_dir):
            continue
        try:
            meta = _read_meta(entry_dir)
            records.append(HistoryRecord(meta=meta, entry_dir=entry_dir))
        except (OSError, json.JSONDecodeError, KeyError):
            continue
    return records


def load_history_entry(entry_id: str) -> HistoryRecord | None:
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT entry_dir FROM history WHERE id = ?", (entry_id,),
        ).fetchone()
    finally:
        conn.close()
    if row is None:
        return None
    entry_dir = row["entry_dir"]
    if not os.path.isdir(entry_dir):
        return None
    try:
        return HistoryRecord(meta=_read_meta(entry_dir), entry_dir=entry_dir)
    except (OSError, json.JSONDecodeError, KeyError):
        return None


def load_history_dataframes(
    record: HistoryRecord,
) -> tuple[pd.DataFrame | None, pd.DataFrame]:
    draft_path = os.path.join(record.entry_dir, DRAFT_FILE)
    draft_df = pd.read_csv(draft_path)
    original_path = os.path.join(record.entry_dir, ORIGINAL_FILE)
    original_df = pd.read_csv(original_path) if os.path.exists(original_path) else None
    return original_df, draft_df


def delete_history_entry(entry_id: str) -> bool:
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT entry_dir FROM history WHERE id = ?", (entry_id,),
        ).fetchone()
        if row is None:
            return False
        entry_dir = row["entry_dir"]
        if os.path.isdir(entry_dir):
            shutil.rmtree(entry_dir, ignore_errors=True)
        conn.execute("DELETE FROM history WHERE id = ?", (entry_id,))
        conn.commit()
        return True
    finally:
        conn.close()
