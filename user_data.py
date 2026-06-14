"""Папка данных пользователя (история, настройки, счётчики)."""

from __future__ import annotations

import os
import platform
import shutil
import sqlite3

APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "Умный помощник"

USER_DATA_ROOT = ""
CONFIG_FILE = ""
LANG_FILE = ""
THEME_FILE = ""
RULES_PATH_FILE = ""
HISTORY_ROOT = ""
COUNTER_FILE = ""
MIGRATION_FLAG_FILE = ""

_LEGACY_CONFIG_NAMES = ("config.txt", "language.txt", "theme.txt", "rules_path.txt")
_LEGACY_HISTORY_REL = os.path.join("data", "history")


def _default_user_data_root() -> str:
    system = platform.system()
    home = os.path.expanduser("~")
    if system == "Darwin":
        return os.path.join(home, "Library", "Application Support", APP_NAME)
    if system == "Windows":
        base = os.environ.get("APPDATA") or os.path.join(home, "AppData", "Roaming")
        return os.path.join(base, APP_NAME)
    xdg = os.environ.get("XDG_DATA_HOME") or os.path.join(home, ".local", "share")
    return os.path.join(xdg, "smartsorter")


def _bind_paths(root: str) -> None:
    global USER_DATA_ROOT, CONFIG_FILE, LANG_FILE, THEME_FILE, RULES_PATH_FILE
    global HISTORY_ROOT, COUNTER_FILE, MIGRATION_FLAG_FILE
    USER_DATA_ROOT = root
    CONFIG_FILE = os.path.join(root, "config.txt")
    LANG_FILE = os.path.join(root, "language.txt")
    THEME_FILE = os.path.join(root, "theme.txt")
    RULES_PATH_FILE = os.path.join(root, "rules_path.txt")
    HISTORY_ROOT = os.path.join(root, "history")
    COUNTER_FILE = os.path.join(root, "complaint_counter.json")
    MIGRATION_FLAG_FILE = os.path.join(root, ".migrated_from_app_dir")


def _copy_file_if_missing(src: str, dst: str) -> None:
    if os.path.exists(src) and not os.path.exists(dst):
        shutil.copy2(src, dst)


def _migrate_history_db(old_root: str, new_root: str) -> None:
    old_db = os.path.join(old_root, "history.db")
    new_db = os.path.join(new_root, "history.db")
    if not os.path.exists(old_db):
        return
    if os.path.exists(new_db):
        return
    shutil.copy2(old_db, new_db)
    conn = sqlite3.connect(new_db)
    try:
        rows = conn.execute("SELECT id, entry_dir FROM history").fetchall()
        for entry_id, entry_dir in rows:
            if not entry_dir or not entry_dir.startswith(old_root):
                continue
            suffix = os.path.relpath(entry_dir, old_root)
            new_entry_dir = os.path.join(new_root, suffix)
            conn.execute(
                "UPDATE history SET entry_dir = ? WHERE id = ?",
                (new_entry_dir, entry_id),
            )
        conn.commit()
    finally:
        conn.close()


def _migrate_legacy_data() -> None:
    legacy_history = os.path.join(APP_DIR, _LEGACY_HISTORY_REL)
    new_history = HISTORY_ROOT
    os.makedirs(USER_DATA_ROOT, exist_ok=True)
    os.makedirs(new_history, exist_ok=True)

    for name in _LEGACY_CONFIG_NAMES:
        _copy_file_if_missing(os.path.join(APP_DIR, name), os.path.join(USER_DATA_ROOT, name))

    if os.path.isdir(legacy_history):
        if not os.listdir(new_history):
            shutil.copytree(legacy_history, new_history, dirs_exist_ok=True)
        else:
            old_entries = os.path.join(legacy_history, "entries")
            new_entries = os.path.join(new_history, "entries")
            if os.path.isdir(old_entries):
                os.makedirs(new_entries, exist_ok=True)
                for name in os.listdir(old_entries):
                    src = os.path.join(old_entries, name)
                    dst = os.path.join(new_entries, name)
                    if os.path.isdir(src) and not os.path.exists(dst):
                        shutil.copytree(src, dst)
        _migrate_history_db(legacy_history, new_history)

    with open(MIGRATION_FLAG_FILE, "w", encoding="utf-8") as fh:
        fh.write("ok\n")


def ensure_user_data_ready() -> str:
    """Создаёт папку пользователя и один раз переносит старые данные из папки приложения."""
    if USER_DATA_ROOT:
        return USER_DATA_ROOT
    root = _default_user_data_root()
    _bind_paths(root)
    os.makedirs(HISTORY_ROOT, exist_ok=True)
    os.makedirs(os.path.join(HISTORY_ROOT, "entries"), exist_ok=True)
    if not os.path.exists(MIGRATION_FLAG_FILE):
        _migrate_legacy_data()
    return USER_DATA_ROOT
