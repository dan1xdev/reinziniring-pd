#!/bin/zsh
cd "$(dirname "$0")" || exit 1
if [[ -x .venv/bin/python ]]; then
  ui_python=.venv/bin/python
elif [[ -x ../../work/registry-venv/bin/python ]]; then
  ui_python=../../work/registry-venv/bin/python
else
  ui_python=python3
fi
if ! "$ui_python" -c 'import openpyxl, defusedxml, keyring' 2>/dev/null; then
  echo 'Сначала установите зависимости по инструкции UI_README.txt.'
  read 'ui_reply?Нажмите Enter, чтобы закрыть окно.'
  exit 1
fi
"$ui_python" -m ou_processor.web
