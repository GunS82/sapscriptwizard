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

## Лицензия
MIT