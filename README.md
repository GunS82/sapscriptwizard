# sapscriptwizard

Библиотека для автоматизации SAP GUI Scripting на Python (Windows).

## Установка

```bash
pip install .
```

## Использование

```python
from sapscriptwizard import Sapscript
sap = Sapscript()
# ... дальнейшая работа с SAP GUI
```

## Требования
- Python 3.8+
- pywin32
- pandas
- polars
- Pillow

## Структура пакета
- sapscriptwizard.py
- window.py
- shell_table.py
- gui_tree.py
- element_finder.py
- locator_helpers.py
- types_/
- utils/
- parallel/

## История запусков

```python
from sapscriptwizard import Sapscript
from sapscriptwizard.run_history import RunHistory

history = RunHistory("run_history.db")
sap = Sapscript(run_history=history)

run_id = sap.start_run("example_script")
try:
    # основной код работы с SAP
    sap.finish_run(run_id, "success")
except Exception as exc:
    sap.handle_exception_with_screenshot(exc)
    sap.finish_run(run_id, "error", error_message=str(exc))
```

Историю запусков можно посмотреть командой:

```bash
python run_history_cli.py --limit 5
```

## KeyValue Clipboard Helper

В файле `temp.py` находится простое приложение на Tkinter. Оно читает пары `ключ-значение` из `key_value_data.json` и копирует значения в буфер обмена. Запустить приложение можно командой:

```bash
python temp.py
```

Используйте сочетание `Ctrl+Shift+C` или иконку в трее, чтобы показать или скрыть окно.


## Лицензия
MIT