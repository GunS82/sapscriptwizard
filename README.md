# sapscriptwizard

Двухслойная библиотека для автоматизации SAP GUI Scripting на Python. Нижний слой
повторяет VBS-подход через объекты :class:`SapSession` и :class:`Window`. Верхний
слой предоставляет «рецепты» для типовых транзакций и вспомогательные контролы.

## Установка

```bash
pip install .
```

## Быстрый старт

```python
from sapscriptwizard import SapSession

session = SapSession()
session.connect()
window = session.window()
window.write("wnd[0]/usr/txtRSYST-MANDT", "100")
```

## Основные модули

- `sapscriptwizard.core` — подключение к SAP GUI, управление окнами и ожидания.
- `sapscriptwizard.gui` — контролы (таблицы, деревья, редакторы) и семантические локаторы.
- `sapscriptwizard.features` — рецепты для SE80, WE19 и SPROXY.
- `sapscriptwizard.helpers` — объяснения элементов и подсказки локаторов.
- `sapscriptwizard.compat` — слой совместимости с устаревшим API (`DeprecationWarning`).

## Дополнительно

- `MIGRATION.md` содержит карту перехода со старого API.
- `docs/vbs_to_python.md` описывает перенос VBS-скриптов.
- `examples/` хранит готовые скрипты для запуска.

## Требования

- Python 3.10+
- Windows c включённым SAP GUI Scripting
- `pywin32`

## Лицензия

MIT
