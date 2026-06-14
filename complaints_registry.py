"""Номера обращений и статусы для сценария жалоб."""

from __future__ import annotations

import json
import os
from datetime import datetime

import pandas as pd

import user_data
from user_data import ensure_user_data_ready

COMPLAINT_NUMBER_COL = "Номер_обращения"
COMPLAINT_STATUS_COL = "Статус"
COMPLAINT_REGISTERED_COL = "Дата_регистрации"
COMPLAINT_PROCESSED_COL = "Дата_обработки"

STATUS_NEW = "Новая"
STATUS_REVIEW = "На проверке"
STATUS_READY = "Готово к отправке"
STATUS_SENT = "Отправлено"

_NUMBER_PREFIX = "ОБ"
_REGISTRY_COLS = (
    COMPLAINT_NUMBER_COL,
    COMPLAINT_STATUS_COL,
    COMPLAINT_REGISTERED_COL,
    COMPLAINT_PROCESSED_COL,
)


def _load_counter() -> dict:
    ensure_user_data_ready()
    if not os.path.exists(user_data.COUNTER_FILE):
        return {"year": datetime.now().year, "next": 1}
    try:
        with open(user_data.COUNTER_FILE, encoding="utf-8") as fh:
            data = json.load(fh)
        year = int(data.get("year", datetime.now().year))
        nxt = int(data.get("next", 1))
        return {"year": year, "next": max(1, nxt)}
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return {"year": datetime.now().year, "next": 1}


def _save_counter(state: dict) -> None:
    ensure_user_data_ready()
    with open(user_data.COUNTER_FILE, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2)


def allocate_complaint_numbers(count: int) -> list[str]:
    if count <= 0:
        return []
    state = _load_counter()
    year = datetime.now().year
    if state["year"] != year:
        state = {"year": year, "next": 1}
    numbers: list[str] = []
    for _ in range(count):
        numbers.append(f"{_NUMBER_PREFIX}-{year}-{state['next']:04d}")
        state["next"] += 1
    _save_counter(state)
    return numbers


def _is_empty_cell(value) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    return not str(value).strip()


def _normalize_registry_columns(df: pd.DataFrame) -> pd.DataFrame:
    """ИИ часто возвращает пустые статусы как float NaN — приводим к тексту."""
    out = df.copy()
    for col in _REGISTRY_COLS:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].map(
            lambda v: "" if _is_empty_cell(v) else str(v).strip(),
        ).astype("object")
    return out


def mark_complaints_sent(df: pd.DataFrame) -> pd.DataFrame:
    """Помечает все строки как «Отправлено» после раскладки по отделам."""
    if df is None or df.empty:
        return df
    out = _normalize_registry_columns(df)
    today = datetime.now().strftime("%d.%m.%Y")
    out[COMPLAINT_STATUS_COL] = STATUS_SENT
    out[COMPLAINT_PROCESSED_COL] = today
    return out


def enrich_complaints_draft(df: pd.DataFrame) -> pd.DataFrame:
    """Присваивает номера и статус «Новая» строкам без номера."""
    if df is None or df.empty:
        return df
    out = _normalize_registry_columns(df)

    today = datetime.now().strftime("%d.%m.%Y")
    empty_mask = out[COMPLAINT_NUMBER_COL].map(_is_empty_cell)
    need = int(empty_mask.sum())
    if need:
        numbers = allocate_complaint_numbers(need)
        out.loc[empty_mask, COMPLAINT_NUMBER_COL] = numbers

    status_mask = out[COMPLAINT_STATUS_COL].map(_is_empty_cell)
    out.loc[status_mask, COMPLAINT_STATUS_COL] = STATUS_NEW

    reg_mask = out[COMPLAINT_REGISTERED_COL].map(_is_empty_cell)
    out.loc[reg_mask, COMPLAINT_REGISTERED_COL] = today
    return out
