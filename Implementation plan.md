Коротко: внедряем двухслойную архитектуру. Низкий слой VBS-like API. Высокий слой Page Objects/«рецепты». Добавляем инспекцию экрана и семантические локаторы. Ниже подробный план.

0) Цели и границы

Цели: понятный API для людей и LLM. Изоморфность VBS на низком уровне. Рецепты для SE80/WE19/SPROXY. Инспекция экрана в JSON.

Не цели: автоматическая трансляция VBS→Python внутри библиотеки.


1) Версионирование и миграция

Семантическая версия: 1.0.0.

Ветка v0 в заморозке. Новая ветка main.

Слой совместимости compat/ с DeprecationWarnings. Срок удаления старого API: следующий мажор.

MIGRATION.md: таблица «старый → новый».


2) Структура репозитория

sapscriptwizard/
  __init__.py
  core/
    session.py
    window.py
    com_gateway.py
    waits.py
    errors.py
    types.py
  gui/
    controls/
      table.py
      tree.py
      abap_editor.py
      text_editor.py
    finders/
      locator_model.py
      semantic_finder.py
  features/
    se80.py
    we19.py
    sproxy.py
  helpers/
    explain.py
  cli/
    sapwiz.py
  compat/               # адаптеры старого API
docs/
  vbs_to_python.md
  cookbook/*.md
examples/
tests/
pyproject.toml

3) Этапы внедрения

Этап 1. Инфраструктура проекта

Настроить packaging (pyproject.toml), Ruff, MyPy/Pyright, pytest.

Логи: logging с форматами JSON-строки, уровни INFO/DEBUG.

Pre-commit: ruff, isort, mypy.

GitHub Actions: линтеры и юнит-тесты. Интеграционные тесты запускать локально вручную.


Критерии приёмки

Успешный прогон линтеров и тестов.

Сборка колеса и sdist.


Этап 2. Базовый «ядро» (core)

core.session.SapSession: подключение к GUI Scripting, управление жизненным циклом, retry-политики.

core.window.Window:

find_by_id(id), read_text(id), write(id, text), press(id), send_vkey(code).

Исключения из core.errors с кодами: E_NOT_FOUND, E_DISABLED, E_TIMEOUT, E_COM_FAIL.


core.waits: wait_visible(id), wait_ready(), wait_status(text|code).

core.com_gateway: единая точка работы с COM, timeouts, повтор.


Критерии приёмки

Юнит-тесты на прокси-объекты COM (моки).

Докстроки и типы. Примеры в examples/core_minimal.py.


Этап 3. Инспекция экрана

Window.dump_gui_structure(max_depth, max_children) -> dict.

JSON-схема:

window_id, title, status_bar, containers[], elements[].

Элемент: id, type, tech_name, label, value, enabled, visible, children[].


helpers.explain.explain_id(id) -> ElementInfo.

helpers.explain.suggest_locator(ElementInfo) -> LocatorHint
(типы: HLabel, VLabel, ByTitle, ByColRow).


Критерии приёмки

Снимок структуры на типовых окнах (моки). Валидация схемы.

Примеры в examples/dump_and_suggest.py.


Этап 4. Контролы GUI (обёртки)

gui.controls.table.ShellTable:

Чтение заголовков и ячеек.

Навигация по строкам.

to_dicts(), to_csv(path).


gui.controls.tree.GuiTree:

expand_node(path|text), select_node(path|text), get_children(node).


gui.controls.abap_editor.GuiAbapEditor:

get_all_text(), set_all_text(text), find(text), goto_line(n).


gui.controls.text_editor.GuiTextEdit:

get_content(), set_content(text), load_from_file(path).



Критерии приёмки

Контрактные тесты на API против моков.

Пример чтения ALV, дерева, редакторов.


Этап 5. Семантический поиск

gui.finders.locator_model:

Типы локаторов: ById, ByLabel, ByTitle, ByRole, ByColumn("Статус").


gui.finders.semantic_finder.SemanticFinder:

find(query: str) -> ById с грамматикой:
@"Метка", = "Заголовок кнопки", col:"Статус" row:5.

Резолв к find_by_id.


Связка с dump_gui_structure для индексации меток.


Критерии приёмки

Набор тестов для 20 запросов-шаблонов.

Пример examples/semantic_select.py.


Этап 6. Низкоуровневый VBS-like API

На уровне Window: стабильные синонимы:

set_text(id, text) ↔ write(id, text).

press(id), select(id, key), set_checked(id, bool).


Чёткие ошибки с кодами. Поля исключений: id, locator, hint.


Критерии приёмки

Таблица соответствий в docs/vbs_to_python.md.

Тест-кейсы «1-в-1» перепись VBS строк.


Этап 7. Высокоуровневые «рецепты»

features.se80.SE80:

open_program(name) -> SE80.

read_source() -> str.

where_used(kind) -> list[Ref].


features.we19.WE19:

upload_xml(path) -> WE19.

fill(fields: dict) -> WE19.

execute() -> RunResult.


features.sproxy.SProxy:

open_service(name), import_wsdl(path|url), generate_objects().



Критерии приёмки

Интеграционные сценарии в виде псевдо-тестов с мок-скриншотами.

Cookbook: шаг-за-шагом примеры.


Этап 8. CLI

cli/sapwiz.py (Typer/argparse):

sapwiz analyze se80 --program Z_PROG.

sapwiz we19 --xml in.xml --type ORDERS.

sapwiz dump --json out.json.



Критерии приёмки

Помощь --help понятна. Команды работают на моках.


Этап 9. Документация и примеры

docs/vbs_to_python.md: 30+ канонических пар «VBS → Python».

docs/cookbook/: SE80, WE19, SPROXY.

Диаграммы: архитектура слоёв, жизненный цикл окна.

Примеры в examples/ с минимальным кодом.


Критерии приёмки

Полный проход по чек-листу LLM-readiness: названия методов, сигнатуры, докстроки, стабильные паттерны.


Этап 10. Тестирование

Юнит: 80%+ покрытие core/finder.

Контрактные тесты контролов.

Snapshot-тесты JSON структуры.

«Record-replay» слой: фикстуры с сериализацией COM-вызовов.

Ручные прогоны против реального SAP GUI на dev-машине.


Критерии приёмки

Отчёт о покрытии.

Стабильность snapshot-тестов.


Этап 11. Слой совместимости

compat/ модули и адаптер-классы. DeprecationWarnings с ссылкой на MIGRATION.md.

Тесты, что старые вызовы не падают.


Критерии приёмки

Список старых публичных точек закрыт адаптерами.


Этап 12. Релиз

CHANGELOG.md с breaking changes.

Теги Git. Публикация на PyPI.

Выпуск «What’s new» с примерами.


4) Политики ошибок и ретраев

Исключения:

SapError(code, message, details).

Коды: E_NOT_FOUND, E_DISABLED, E_TIMEOUT, E_COM_FAIL, E_UNSUPPORTED, E_AMBIGUOUS.


Ретраи: экспоненциальная задержка, max_attempts в com_gateway.

Тайм-ауты: глобально в session, точечные аргументы в вызовах.


5) Локаторы и грамматика запросов

Поддержка:

@"Метка" — по вертикальной/горизонтальной метке.

="Текст кнопки" — по заголовку.

id:"wnd[0]/usr/..." — прямой ID.

col:"Статус" row:5 — ячейка таблицы.


Разрешение конфликтов: детерминированная сортировка, эвристики близости.


6) Контракты API (срез)

Window.dump_gui_structure(max_depth:int=2, max_children:int=100) -> dict

Window.finder.find(query:str) -> str  (ID)

ShellTable.to_dicts(limit_rows:int|None=None) -> list[dict]

GuiAbapEditor.get_all_text() -> str

WE19.upload_xml(path:str) -> WE19


7) MIGRATION.md (ядро)

Старое session.findById(...).text="X" → win.write(id, "X") или win.set_text(id, "X").

Старое pressButton → win.press(id).

Старое чтение ALV вручную → ShellTable.to_dicts().


8) Риски и снижения

Зависимость от раскладки экранов: снижать через dump_gui_structure и семантические локаторы.

Хрупкость ID: предлагать suggest_locator.

COM-нестабильность: централизованные ретраи в com_gateway.


9) Контроль качества

Чек-лист API стабильности: имена, порядок аргументов, типы.

Чек-лист читаемости: примеры покрывают три ключевых сценария.

Чек-лист LLM: наличие «скелетонов» для автодополнения и очевидных названий.


10) Минимальный инкремент «готово»

Реализованы: core, dump, finders, ShellTable, GuiTree, AbapEditor, TextEditor.

Есть: SE80.read_source, WE19.upload_xml/execute.

Есть: vbs_to_python.md, 10 примеров.

Тесты: юнит + snapshot.


Готов к реализации по этапам.
