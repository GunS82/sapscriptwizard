# SAP Script Wizard

Двухслойная библиотека для автоматизации SAP GUI Scripting на Python.
Нижний уровень предоставляет изоморфный к VBS API, верхний уровень —
"рецепты" для типовых транзакций.

## Возможности

* Управление сессиями через `SapSession` и окно `Window`.
* Надежные ретраи и обработка ошибок с кодами (`SapErrorCode`).
* Инспекция экрана в JSON через `Window.dump_gui_structure()`.
* Семантические локаторы (`Window.finder.find()` и `helpers.explain`).
* Готовые контролы: `ShellTable`, `GuiTree`, `GuiAbapEditor`, `GuiTextEditor`.
* Рецепты для транзакций SE80, WE19, SPROXY.

## Установка

```bash
pip install .
```

## Быстрый старт

```python
from sapscriptwizard import SapSession

session = SapSession.from_gui()
window = session.window()
window.write("usr/txtUser", "DEMO")
window.press("usr/btnExecute")
print(window.dump_gui_structure())
```

Дополнительные примеры доступны в каталоге `examples/` и в `docs/cookbook/`.

## Требования

* Python 3.10+
* SAP GUI Scripting (Windows) + `pywin32`

Для разработки используйте extras:

```bash
pip install .[dev]
```

## Тесты и линтеры

```bash
pytest
ruff check
mypy
```

## Совместимость

Публичный API 0.x вынесен в пакет `sapscriptwizard.compat` и помечен
`DeprecationWarning`. Подробности в [MIGRATION.md](MIGRATION.md).

## Лицензия

MIT
