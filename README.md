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

### Подключение семантических локаторов

```python
from sapscriptwizard import Window
from sapscriptwizard_semantic.element_finder import SapElementFinder

win = Window(app, 0, conn_handle, 0, sess_handle,
             element_finder=SapElementFinder(sess_handle))
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
- sapscriptwizard_semantic/ (optional plugin with `SapElementFinder`)
- types_/
- utils/
- parallel/
## KeyValue Clipboard Helper

В файле `temp.py` находится простое приложение на Tkinter. Оно читает пары `ключ-значение` из `key_value_data.json` и копирует значения в буфер обмена. Запустить приложение можно командой:

```bash
python temp.py
```

Используйте сочетание `Ctrl+Shift+C` или иконку в трее, чтобы показать или скрыть окно.


## Лицензия
MIT
