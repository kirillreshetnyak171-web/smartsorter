#!/bin/bash
# Запуск SmartSorter двойным кликом (macOS).
# ПКМ → Открыть при первом запуске, если macOS блокирует файл.

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR" || exit 1

if [ -x "$DIR/.venv/bin/python3" ]; then
    PYTHON="$DIR/.venv/bin/python3"
elif command -v python3 >/dev/null 2>&1; then
    PYTHON="python3"
else
    osascript -e 'display dialog "Python 3 не найден.\n\nУстановите с python.org и выполните в Терминале:\npip install -r requirements.txt" buttons {"OK"} default button 1 with title "SmartSorter"' 2>/dev/null
    exit 1
fi

if ! "$PYTHON" -c "import PyQt6" 2>/dev/null; then
    osascript -e 'display dialog "Зависимости не установлены.\n\nОткройте Терминал в папке проекта и выполните:\npython3 -m pip install -r requirements.txt" buttons {"OK"} default button 1 with title "SmartSorter"' 2>/dev/null
    exit 1
fi

exec "$PYTHON" "$DIR/app.py"
