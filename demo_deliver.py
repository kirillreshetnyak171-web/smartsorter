#!/usr/bin/env python3
"""
Демо-доставка «писем» по отделам в папки на рабочем столе.

Как проверить:
  1. В app.py: классифицируйте жалобы → «В папку» (получите routing_YYYY-MM-DD_HHMMSS/)
  2. Запустите:
       python3 demo_deliver.py ~/Desktop/routing_2026-06-06_123456

Или сразу из Excel с классификацией:
       python3 demo_deliver.py --xlsx routing.xlsx

Папки отделов: ~/Desktop/SmartSorter_Почта_отделов/
Конфиг: department_folders.demo.txt (рядом с этим скриптом)
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

# Импорт логики из основного приложения
from app import (
    ROUTING_DEPT_NAMES,
    build_complaints_readable_text,
    build_department_email,
    export_routing_folder,
    find_routing_column,
    load_allowed_departments,
    read_table,
)

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = SCRIPT_DIR / "department_folders.demo.txt"

SPECIAL_FILE_MAP = {
    "на проверку": "На проверку",
    "needs review": "На проверку",
    "prüfung": "На проверку",
    "проверить_отдел": "На проверку",
    "check_department": "На проверку",
    "abteilung_prüfen": "На проверку",
    "сводка": "Сводка",
    "summary": "Сводка",
    "zusammenfassung": "Сводка",
}


@dataclass
class DepartmentContact:
    name: str
    folder: Path
    email: str = ""


def parse_department_contacts(config_path: Path) -> dict[str, DepartmentContact]:
    """Читает department_folders.demo.txt → словарь отдел → контакт."""
    text = config_path.read_text(encoding="utf-8")
    contacts: dict[str, DepartmentContact] = {}
    section = ""
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        upper = stripped.upper()
        if "КОНТАКТЫ ОТДЕЛОВ" in upper:
            section = "dept"
            continue
        if "СПЕЦИАЛЬНЫЕ ПАПКИ" in upper:
            section = "special"
            continue
        if not stripped.startswith("- "):
            continue
        body = stripped[2:].strip()
        parts = [p.strip() for p in body.split("|")]
        if len(parts) < 2:
            continue
        name, folder = parts[0], parts[1]
        email = parts[2] if len(parts) > 2 else ""
        contacts[name.casefold()] = DepartmentContact(
            name=name,
            folder=Path(folder).expanduser(),
            email=email,
        )
    return contacts


def _norm_key(name: str) -> str:
    return re.sub(r"^00_", "", name.strip()).casefold()


def _match_contact(file_stem: str, contacts: dict[str, DepartmentContact]) -> DepartmentContact | None:
    key = _norm_key(file_stem)
    if key in contacts:
        return contacts[key]
    for special_key, contact_name in SPECIAL_FILE_MAP.items():
        if special_key in key:
            return contacts.get(contact_name.casefold())
    for ck, contact in contacts.items():
        if ck in key or key in ck:
            return contact
    return None


def _email_labels() -> dict[str, str]:
    return {
        "email_body": (
            "Тема: Жалобы — {department} — {date}\n\n"
            "Здравствуйте!\n\n"
            "Во вложении {count} обращение(й) для отдела «{department}».\n"
            "Высокий приоритет: {high}\n"
            "Требует проверки: {review}\n\n"
            "{attach_hint}\n\n"
            "С уважением"
        ),
        "email_attach": (
            "В этой папке: Excel с таблицей и файл «жалобы_прочитать_…txt» — "
            "откройте его, чтобы прочитать тексты обращений."
        ),
    }


def _write_readable_complaints(
    dept_df: pd.DataFrame,
    folder: Path,
    stamp: str,
    title: str,
) -> Path:
    """Создаёт .txt с текстами жалоб для чтения без Excel."""
    path = folder / f"жалобы_прочитать_{stamp}.txt"
    path.write_text(
        build_complaints_readable_text(dept_df, title),
        encoding="utf-8",
    )
    return path


def deliver_from_export_folder(
    export_folder: Path,
    contacts: dict[str, DepartmentContact],
    df: pd.DataFrame | None = None,
) -> list[str]:
    """Копирует файлы из routing_* в папки отделов + пишет письмо.txt."""
    if not export_folder.is_dir():
        raise FileNotFoundError(f"Папка не найдена: {export_folder}")

    log: list[str] = []
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")

    for xlsx in sorted(export_folder.glob("*.xlsx")):
        stem = xlsx.stem
        contact = _match_contact(stem, contacts)
        if contact is None:
            log.append(f"⚠ Пропуск (нет в конфиге): {xlsx.name}")
            continue

        contact.folder.mkdir(parents=True, exist_ok=True)
        dest_xlsx = contact.folder / f"{stamp}_{xlsx.name}"
        shutil.copy2(xlsx, dest_xlsx)

        # Тексты жалоб — из готового _прочитать.txt или из Excel
        readable_src = export_folder / f"{stem}_прочитать.txt"
        if readable_src.exists():
            dest_readable = contact.folder / f"жалобы_прочитать_{stamp}.txt"
            shutil.copy2(readable_src, dest_readable)
        elif contact.name not in ("Сводка",):
            try:
                dept_df = read_table(str(xlsx))
                dest_readable = _write_readable_complaints(
                    dept_df,
                    contact.folder,
                    stamp,
                    f"Жалобы — {contact.name}",
                )
            except (OSError, ValueError, pd.errors.ParserError):
                dest_readable = None
        else:
            dest_readable = None

        letter_name = f"письмо_{stamp}.txt"
        letter_path = contact.folder / letter_name

        email_df = None
        if contact.name not in ("На проверку", "Сводка"):
            try:
                email_df = read_table(str(xlsx))
            except (OSError, ValueError, pd.errors.ParserError):
                email_df = df
            if email_df is not None:
                dept_col = find_routing_column(email_df, ROUTING_DEPT_NAMES)
                if dept_col and contact.name in email_df[dept_col].astype(
                    str,
                ).str.strip().values:
                    body = build_department_email(
                        email_df, contact.name, _email_labels(),
                    )
                else:
                    body = _simple_letter(contact, dest_xlsx.name, dest_readable)
            else:
                body = _simple_letter(contact, dest_xlsx.name, dest_readable)
        else:
            body = _simple_letter(contact, dest_xlsx.name, dest_readable)

        header = f"Кому: {contact.email or '(email не указан)'}\n"
        header += f"Папка: {contact.folder}\n"
        header += f"Таблица: {dest_xlsx.name}\n"
        if dest_readable:
            header += f"Прочитать тексты: {dest_readable.name}\n"
        header += "-" * 50 + "\n\n"
        letter_path.write_text(header + body, encoding="utf-8")

        extra = f", {dest_readable.name}" if dest_readable else ""
        log.append(f"✓ {contact.name} → {contact.folder.name}/ ({dest_xlsx.name}{extra})")

    return log


def _simple_letter(
    contact: DepartmentContact,
    attachment: str,
    readable: Path | None = None,
) -> str:
    date = datetime.now().strftime("%d.%m.%Y")
    read_hint = ""
    if readable:
        read_hint = (
            f"\nЧтобы прочитать, что произошло — откройте «{readable.name}» "
            f"(тексты жалоб простым языком).\n"
        )
    return (
        f"Тема: Жалобы — {contact.name} — {date}\n\n"
        f"Здравствуйте!\n\n"
        f"Во вложении файл «{attachment}» для отдела «{contact.name}»."
        f"{read_hint}\n"
        f"(Демо: вместо реальной почты файлы лежат в папке отдела.)\n\n"
        f"С уважением"
    )


def deliver_from_xlsx(xlsx_path: Path, contacts: dict[str, DepartmentContact]) -> list[str]:
    """Экспортирует routing во временную папку и доставляет по отделам."""
    df = read_table(str(xlsx_path))
    temp_parent = xlsx_path.parent / f"_demo_routing_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    temp_parent.mkdir(exist_ok=True)
    labels = {
        "sheet_summary": "Сводка",
        "sheet_review": "На проверку",
        "sheet_unknown_dept": "Проверить_отдел",
        "dept": "Отдел",
        "total": "Всего",
        "high": "Высокий",
        "review": "Проверка",
        "unknown": "Не указан",
        "log_title": "Лог",
        "log_date": "Дата",
        "log_total": "Всего",
        "log_review": "Проверка",
        "log_unknown": "Неизвестный",
        "log_by_dept": "По отделам",
        "log_unknown_list": "Неизвестные",
    }
    allowed = load_allowed_departments()
    folder, _ = export_routing_folder(df, str(temp_parent), labels, allowed or None)
    log = deliver_from_export_folder(Path(folder), contacts, df)
    log.insert(0, f"Создана промежуточная папка: {folder}")
    return log


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Демо: доставить файлы по отделам в папки на рабочем столе",
    )
    parser.add_argument(
        "path",
        nargs="?",
        help="Папка routing_* из приложения или .xlsx с классификацией",
    )
    parser.add_argument(
        "--config",
        default=str(DEFAULT_CONFIG),
        help="Файл с контактами отделов (по умолчанию department_folders.demo.txt)",
    )
    parser.add_argument(
        "--xlsx",
        action="store_true",
        help="Путь — это Excel, сначала разбить по отделам",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Показать отделы и папки из конфига",
    )
    args = parser.parse_args()

    config_path = Path(args.config).expanduser()
    if not config_path.exists():
        print(f"Нет конфига: {config_path}", file=sys.stderr)
        return 1

    contacts = parse_department_contacts(config_path)
    if args.list:
        print("Отделы и папки (демо):\n")
        for c in contacts.values():
            email = f"  ({c.email})" if c.email else ""
            print(f"  {c.name}{email}")
            print(f"    → {c.folder}\n")
        return 0

    if not args.path:
        parser.print_help()
        print("\nПример:")
        print("  python3 demo_deliver.py ~/Desktop/routing_2026-06-06_143000")
        return 1

    target = Path(args.path).expanduser()
    print(f"Конфиг: {config_path}")
    print(f"Источник: {target}\n")

    try:
        if args.xlsx or target.suffix.lower() == ".xlsx":
            lines = deliver_from_xlsx(target, contacts)
        else:
            df = None
            log_file = target / "log.txt"
            if log_file.exists():
                pass
            # попробуем собрать df из всех xlsx отделов для текста письма
            dept_files = [f for f in target.glob("*.xlsx") if not f.stem.startswith("00_")]
            if dept_files:
                frames = [read_table(str(f)) for f in dept_files]
                if frames:
                    df = pd.concat(frames, ignore_index=True)
            lines = deliver_from_export_folder(target, contacts, df)
    except Exception as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    if not lines:
        print("Нечего доставлять — положите в папку .xlsx файлы по отделам.")
        return 1

    for line in lines:
        print(line)
    print(f"\nГотово. Откройте: ~/Desktop/SmartSorter_Почта_отделов/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
