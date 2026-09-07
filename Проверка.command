#!/bin/bash
# Двойной клик по этому файлу проверяет, всё ли готово к работе.
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "На этом компьютере не установлен Python."
  echo "Установите инструменты разработчика командой:  xcode-select --install"
  echo "После установки запустите проверку снова."
  echo
  echo "Нажмите Enter, чтобы закрыть окно."
  read -r
  exit 1
fi

python3 scripts/check.py
echo
echo "Нажмите Enter, чтобы закрыть окно."
read -r
