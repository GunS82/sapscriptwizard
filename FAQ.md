# Документация по библиотеке `sapscriptwizard`

## Оглавление

1.  [Обзор](#1-обзор)
2.  [Установка](#2-установка)
3.  [Основные компоненты](#3-основные-компоненты)
4.  [Начало работы (Quick Start)](#4-начало-работы-quick-start)
5.  [Ключевые концепции](#5-ключевые-концепции)
    *   [SAP GUI Scripting API](#sap-gui-scripting-api)
    *   [Соединения и сессии](#соединения-и-сессии)
    *   [Идентификаторы элементов (Element ID)](#идентификаторы-элементов-element-id)
    *   [Семантические локаторы](#семантические-локаторы)
6.  [API Reference](#6-api-reference)
    *   [`sapscriptwizard.Sapscript`](#class-sapscriptwizardsapscript)
    *   [`sapscriptwizard.Window`](#class-sapscriptwizardwindow)
    *   [`sapscriptwizard.element_finder.SapElementFinder`](#class-sapscriptwizardelement_findersapelementfinder)
    *   [`sapscriptwizard.locator_helpers`](#module-sapscriptwizardlocator_helpers)
    *   [`sapscriptwizard.shell_table.ShellTable`](#class-sapscriptwizardshell_tableshelltable)
    *   [`sapscriptwizard.gui_tree.GuiTree`](#class-sapscriptwizardgui_treeguitree)
    *   [`sapscriptwizard.parallel.api.run_parallel`](#function-sapscriptwizardparallelapirun_parallel)
    *   [`sapscriptwizard.parallel.runner.SapParallelRunner`](#class-sapscriptwizardparallelrunnersapparallelrunner)
    *   [`sapscriptwizard.types_.exceptions`](#module-sapscriptwizardtypes_exceptions)
    *   [`sapscriptwizard.types_.types`](#module-sapscriptwizardtypes_types)
    *   [`sapscriptwizard.utils.sap_config.SapLogonConfig`](#class-sapscriptwizardutilssap_configsaplogonconfig)
    *   [`sapscriptwizard.utils.utils`](#module-sapscriptwizardutilsutils)
7.  [Работа с семантическими локаторами](#7-работа-с-семантическими-локаторами)
8.  [Параллельное выполнение сценариев](#8-параллельное-выполнение-сценариев)
9.  [Обработка исключений](#9-обработка-исключений)
10. [Логирование](#10-логирование)
11. [Продвинутые темы и советы](#11-продвинутые-темы-и-советы)
    *   [Кэширование элементов](#кэширование-элементов)
    *   [Отладка и анализ GUI](#отладка-и-анализ-gui)
    *   [Работа с COM-объектами](#работа-с-com-объектами)

---

## 1. Обзор

`sapscriptwizard` - это библиотека Python, предназначенная для упрощения автоматизации задач в SAP GUI для Windows. Она предоставляет высокоуровневый API для взаимодействия с SAP GUI Scripting Engine, позволяя разработчикам программно управлять сессиями SAP, окнами, элементами управления и извлекать данные.

**Ключевые возможности:**

*   Запуск и завершение работы SAP Logon.
*   Подключение к существующим соединениям и сессиям SAP.
*   Открытие новых сессий (окон).
*   Управление окнами: максимизация, восстановление, закрытие.
*   Взаимодействие с элементами GUI по их ID: нажатие кнопок, ввод текста, чтение значений, установка чекбоксов и т.д.
*   **Семантические локаторы:** Мощный механизм поиска элементов по их текстовым меткам, содержимому или взаимному расположению, что делает скрипты более читаемыми и устойчивыми к изменениям ID.
*   Работа со специфичными элементами SAP:
    *   **Таблицы (`ShellTable` / `GuiGridView`):** Чтение данных в DataFrame (Polars, Pandas), экспорт в CSV, прокрутка для загрузки всех строк.
    *   **Деревья (`GuiTree`):** Разворачивание/сворачивание узлов, выбор узлов, получение информации об узлах и их дочерних элементах.
*   Чтение и проверка сообщений в строке состояния.
*   Запуск транзакций (включая "надёжный" запуск с проверкой ошибок).
*   Итерация по элементам GUI с использованием шаблонов ID.
*   Получение и установка свойств элементов.
*   Создание "слепков" (snapshots) иерархии GUI для отладки и анализа.
*   Обработка неожиданных всплывающих окон.
*   **Параллельное выполнение:** Возможность запускать сценарии в нескольких сессиях SAP одновременно для ускорения обработки больших объемов данных.
*   Вспомогательные утилиты для работы с конфигурацией SAP Logon (`saplogon.ini`).
*   Автоматическое создание скриншотов при ошибках.
*   Управление историей ввода SAP GUI.

Библиотека построена на основе `pywin32` для взаимодействия с COM-объектами SAP GUI Scripting Engine.

## 2. Установка

Для установки библиотеки используйте pip:

```bash
pip install -r requirements.txt
# или, если setup.py настроен для публикации:
# pip install sapscriptwizard
```

Предполагается, что файл `requirements.txt` содержит:

```
pywin32
pandas
polars
Pillow
PyYAML # Для сохранения snapshot в YAML
```

**Зависимости:**

*   Python 3.8+
*   `pywin32`: Для COM-взаимодействия.
*   `pandas`: Для конвертации табличных данных (опционально, если используется `to_pandas_dataframe`).
*   `polars`: Для основной работы с табличными данными.
*   `Pillow`: Для создания скриншотов.
*   `PyYAML` (опционально): Для сохранения "слепков" GUI в формате YAML.

**Требования к системе:**

*   Windows OS.
*   Установленный SAP GUI для Windows.
*   Включенный SAP GUI Scripting на стороне сервера и клиента.

## 3. Основные компоненты

Библиотека состоит из следующих ключевых модулей и классов:

*   **`sapscriptwizard.Sapscript`**: Основной класс для инициализации и управления SAP GUI.
*   **`sapscriptwizard.Window`**: Представляет окно сессии SAP и предоставляет методы для взаимодействия с его элементами.
*   **`sapscriptwizard.element_finder.SapElementFinder`**: Внутренний класс, используемый `Window` для поиска элементов с помощью семантических локаторов.
*   **`sapscriptwizard.locator_helpers`**: Содержит dataclass'ы для представления позиций элементов, информации об элементах и стратегий локаторов.
*   **`sapscriptwizard.shell_table.ShellTable`**: Класс для работы с таблицами SAP (ALV Grid, `GuiGridView`).
*   **`sapscriptwizard.gui_tree.GuiTree`**: Класс для работы с древовидными структурами SAP (`GuiShell` типа "Tree").
*   **`sapscriptwizard.parallel.api.run_parallel`**: Функция для запуска параллельного выполнения сценариев.
*   **`sapscriptwizard.parallel.runner.SapParallelRunner`**: Класс, управляющий пулом процессов для параллельной работы с SAP.
*   **`sapscriptwizard.types_.exceptions`**: Пользовательские исключения библиотеки.
*   **`sapscriptwizard.types_.types`**: Определения типов, например, `NavigateAction`.
*   **`sapscriptwizard.utils`**: Вспомогательные утилиты.

## 4. Начало работы (Quick Start)

```python
import logging
from pathlib import Path
from sapscriptwizard import Sapscript, exceptions, NavigateAction

# Настройка логирования (рекомендуется)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')
log = logging.getLogger(__name__)

# Параметры для подключения (замените на свои)
SAP_SID = "S4H"
SAP_CLIENT = "100"
SAP_USER = "MYUSER"
SAP_PASSWORD = "MYPASSWORD"
SAP_LANGUAGE = "EN" # или "RU"

def main():
    sap = Sapscript()
    sap.enable_screenshots_on_error()
    sap.set_screenshot_directory(Path("./sap_errors"))

    try:
        # Запуск SAP Logon, если он еще не запущен
        # Sapscript.start_saplogon() # Раскомментируйте, если нужно

        # Вариант 1: Запуск нового экземпляра SAP и подключение
        # sap.launch_sap(
        #     sid=SAP_SID,
        #     client=SAP_CLIENT,
        #     user=SAP_USER,
        #     password=SAP_PASSWORD,
        #     language=SAP_LANGUAGE,
        #     root_sap_dir=Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui") # Укажите ваш путь
        # )
        # window = sap.attach_window(connection_index=0, session_index=0)

        # Вариант 2: Подключение к уже запущенной сессии (если вы знаете индексы)
        # Предположим, сессия уже открыта
        log.info("Попытка подключения к существующей сессии...")
        all_connections = sap.get_all_connections_info()
        if not all_connections or not all_connections[0]['sessions']:
            log.error("Нет активных сессий SAP для подключения. Запустите SAP и войдите в систему.")
            return

        target_conn_idx = all_connections[0]['index']
        target_sess_idx = all_connections[0]['sessions'][0]['index']
        window = sap.attach_window(connection_index=target_conn_idx, session_index=target_sess_idx)
        log.info(f"Успешно подключено к окну: {window}")

        # Базовые действия с окном
        window.maximize()

        # Запуск транзакции
        window.start_transaction("SE16") # или window.start_transaction_robust("SE16")

        # Ожидание появления элементов и взаимодействие с использованием ID
        # (ID могут отличаться в вашей системе, используйте SAP GUI Scripting Tracker)
        table_name_field_id = "wnd[0]/usr/ctxtDATABROWSE-TABLENAME"
        if window.exists(table_name_field_id):
            window.write(table_name_field_id, "MARA")
            window.press("wnd[0]/tbar[1]/btn[7]") # Кнопка "Execute" (F8)
        else:
            log.warning(f"Поле {table_name_field_id} не найдено.")
            # Можно добавить обработку всплывающих окон, если они ожидаются
            # window.handle_unexpected_popup(press_button_id="tbar[0]/btn[0]") # Пример: нажать Enter

        # Взаимодействие с использованием семантических локаторов
        log.info("Поиск поля 'Имя таблицы' с помощью семантического локатора...")
        # Пример для русскоязычного интерфейса, если метка "Имя таблицы"
        #mara_field_id = window.find_element_id_by_locator("Имя таблицы")
        # Для англоязычного:
        mara_field_id_loc = window.find_element_id_by_locator("Table Name")

        if mara_field_id_loc:
            log.info(f"Найдено поле 'Table Name' по локатору, ID: {mara_field_id_loc}")
            window.write(mara_field_id_loc, "T001")
            # Нажать кнопку "Execute" по ее тексту/tooltip (если он "Execute")
            window.press_by_locator("=Execute")
        else:
            log.warning("Поле 'Table Name' не найдено с помощью семантического локатора.")


        # Пример чтения таблицы (если открыта таблица T001)
        # ID таблицы может быть, например, "wnd[0]/usr/cntlGRID1/shellcont/shell"
        # Используйте Scripting Tracker для определения ID вашей таблицы
        # shell_table_id = "wnd[0]/usr/cntlGRID1/shellcont/shell" # Замените на ваш ID
        # if window.exists(shell_table_id):
        #     shell_table = window.read_shell_table(shell_table_id)
        #     log.info(f"Прочитана таблица с {shell_table.rows} строками и {shell_table.columns} колонками.")
        #     log.info(f"Первые 5 строк:\n{shell_table.data.head()}")
        #     shell_table.to_csv("t001_export.csv")
        #     log.info("Таблица сохранена в t001_export.csv")

        # Возврат на главный экран
        window.navigate(NavigateAction.back)
        window.navigate(NavigateAction.back) # Дважды, если были в результатах SE16

        log.info("Сценарий выполнен успешно.")

    except exceptions.SapGuiComException as e:
        log.error(f"Ошибка SAP GUI COM: {e}", exc_info=True)
        sap.handle_exception_with_screenshot(e, filename_prefix="sap_com_error")
    except exceptions.ActionException as e:
        log.error(f"Ошибка действия: {e}", exc_info=True)
        sap.handle_exception_with_screenshot(e, filename_prefix="action_error")
    except FileNotFoundError as e:
        log.error(f"Файл не найден: {e}", exc_info=True)
    except Exception as e:
        log.error(f"Непредвиденная ошибка: {e}", exc_info=True)
        sap.handle_exception_with_screenshot(e, filename_prefix="unexpected_error")
    finally:
        # Завершение работы SAP (если был запущен через launch_sap с quit_auto=True,
        # это произойдет автоматически при выходе из скрипта).
        # Если подключались к существующей сессии, закрывать SAP обычно не нужно.
        # sap.quit() # Раскомментируйте, если нужно принудительно закрыть SAP
        log.info("Завершение работы.")

if __name__ == "__main__":
    main()
```

## 5. Ключевые концепции

### SAP GUI Scripting API
SAP GUI Scripting - это интерфейс автоматизации, предоставляемый SAP, который позволяет внешним приложениям управлять SAP GUI. `sapscriptwizard` использует этот API через COM-объекты Windows. Для работы библиотеки необходимо, чтобы SAP GUI Scripting был разрешен как на сервере SAP, так и в настройках клиента SAP Logon.

### Соединения и сессии
*   **Соединение (Connection):** Представляет подключение к определенной системе SAP (SID). Одно приложение SAP Logon может управлять несколькими соединениями. Индексируются с 0.
*   **Сессия (Session):** Внутри одного соединения может быть открыто несколько сессий (окон). Каждая сессия работает независимо. Стандартное ограничение SAP - 6 сессий на одно соединение. Индексируются с 0 внутри каждого соединения.

### Идентификаторы элементов (Element ID)
Каждый элемент в SAP GUI (поле ввода, кнопка, таблица и т.д.) имеет уникальный иерархический ID, например, `wnd[0]/usr/txtRSYST-BNAME` или `wnd[0]/tbar[0]/btn[3]`. Эти ID используются для прямого взаимодействия с элементами. Они могут быть определены с помощью инструмента "Scripting Tracker", доступного в SAP GUI (Alt+F12 -> Script Recording and Playback -> кнопка "Record" создает VBS-скрипт, где видны ID).

### Семантические локаторы
Поскольку ID элементов могут меняться (хотя и редко в стандартных транзакциях), `sapscriptwizard` вводит понятие семантических локаторов. Это строки, описывающие элемент по его визуальным признакам или расположению относительно других элементов, вместо использования жестко заданных ID. Это делает скрипты более читаемыми и устойчивыми к незначительным изменениям интерфейса. Подробнее см. раздел [Работа с семантическими локаторами](#7-работа-с-семантическими-локаторами).

## 6. API Reference

### Class `sapscriptwizard.Sapscript`
Основной класс для инициализации и управления SAP GUI.

```python
class Sapscript:
    def __init__(self, default_window_title: str = "SAP Easy Access")
```
*   **`default_window_title`**: Заголовок окна SAP по умолчанию, используемый для проверок при запуске.

**Методы:**

*   **`_ensure_com_objects()` -> `None`**: (Внутренний) Гарантирует инициализацию COM-объектов SAP GUI. Вызывает `SapGuiComException` при неудаче.
*   **`launch_sap(sid: str, client: str, user: str, password: str, *, root_sap_dir: Path = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui"), maximise: bool = True, language: str = "en", quit_auto: bool = True)` -> `None`**:
    Запускает SAP с использованием `sapshcut.exe`.
    *   `sid`, `client`, `user`, `password`: Учетные данные.
    *   `root_sap_dir`: Путь к директории установки SAP GUI.
    *   `maximise`: Максимизировать окно после запуска.
    *   `language`: Язык входа (например, "EN", "RU").
    *   `quit_auto`: Если `True`, регистрирует обработчик `atexit` для попытки корректного завершения работы SAP при выходе из скрипта.
    *   Вызывает: `FileNotFoundError`, `WindowDidNotAppearException`, `SapGuiComException`.
*   **`quit()` -> `None`**:
    Пытается корректно завершить сессию (System -> Log Off) и затем принудительно завершает процесс `saplogon.exe`.
*   **`attach_window(connection_index: int, session_index: int)` -> `Window`**:
    Подключается к указанной сессии SAP.
    *   `connection_index`: 0-based индекс соединения.
    *   `session_index`: 0-based индекс сессии внутри соединения.
    *   Возвращает: Объект `Window`.
    *   Вызывает: `AttributeError`, `SapGuiComException`, `AttachException`.
*   **`open_new_window(window_to_handle_opening: Window)` -> `None`**:
    Открывает новую сессию SAP, используя существующий объект `Window` для инициации команды.
    *   Вызывает: `ActionException`, `WindowDidNotAppearException`.
*   **`start_saplogon(saplogon_path: Union[str, Path] = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"))` -> `bool`**: (Статический метод)
    Запускает `saplogon.exe`, если объект SAP GUI scripting не найден.
    *   Возвращает: `True`, если процесс запущен или уже работал, `False` в случае ошибки.
*   **`get_connection_count()` -> `int`**:
    Возвращает количество открытых соединений SAP.
    *   Вызывает: `SapGuiComException`.
*   **`get_connection_info(connection_index: int)` -> `Optional[Dict[str, Any]]`**:
    Возвращает базовую информацию о соединении (например, описание).
*   **`get_active_session_indices(connection_index: int)` -> `List[int]`**:
    Возвращает список активных (доступных) индексов сессий для данного соединения.
*   **`get_session_info(connection_index: int, session_index: int)` -> `Optional[Dict[str, Any]]`**:
    Возвращает детальную информацию о сессии (SID, пользователь, клиент, транзакция и т.д.).
*   **`get_all_connections_info()` -> `List[Dict[str, Any]]`**:
    Сканирует и возвращает структурированную информацию обо всех активных соединениях и их сессиях.
*   **`find_session_by_sid_user(sid: str, user: str)` -> `Optional[Window]`**:
    Находит первую активную сессию, соответствующую указанным SID и пользователю.
*   **`enable_screenshots_on_error()` -> `None`**: Включает автоматическое создание скриншотов при ошибках.
*   **`disable_screenshots_on_error()` -> `None`**: Отключает автоматическое создание скриншотов.
*   **`set_screenshot_directory(directory: Union[str, Path])` -> `None`**: Устанавливает директорию для сохранения скриншотов.
*   **`handle_exception_with_screenshot(exception: Exception, filename_prefix: str = "pysap_error")` -> `None`**:
    Обрабатывает исключение: логирует его и создает скриншот (если включено).
*   **`disable_history()` -> `bool`**: Отключает историю ввода в SAP GUI (`HistoryEnabled=False`).
*   **`enable_history()` -> `bool`**: Включает историю ввода в SAP GUI (`HistoryEnabled=True`).

---

### Class `sapscriptwizard.Window`
Представляет окно сессии SAP и методы для взаимодействия с ним.

```python
class Window:
    def __init__(
        self,
        application: win32com.client.CDispatch,
        connection: int,
        connection_handle: win32com.client.CDispatch,
        session: int,
        session_handle: win32com.client.CDispatch,
    )
```
*   `application`: COM-объект `ScriptingEngine`.
*   `connection`, `session`: Индексы соединения и сессии.
*   `connection_handle`, `session_handle`: COM-объекты для соединения и сессии.

**Методы:**

*   **`maximize()` / `restore()` / `close_window()` -> `None`**: Управление состоянием окна.
*   **`navigate(action: NavigateAction)` -> `None`**: Навигация (Enter, Back, End, Cancel, Save). `NavigateAction` - это enum из `types_.types`.
*   **`start_transaction(transaction: str)` -> `None`**: Запускает транзакцию.
*   **`press(element: str)` -> `None`**: Нажимает элемент (кнопку).
*   **`select(element: str)` -> `None`**: Выбирает элемент (например, пункт меню, радиокнопку).
*   **`is_selected(element: str)` -> `bool`**: Проверяет, выбран ли элемент (чекбокс, радиокнопка).
*   **`set_checkbox(element: str, selected: bool)` -> `None`**: Устанавливает состояние чекбокса.
*   **`write(element: str, text: str)` -> `None`**: Вводит текст в поле.
*   **`read(element: str)` -> `str`**: Читает текстовое содержимое элемента.
*   **`visualize(element: str, seconds: int = 1)` -> `None`**: Подсвечивает элемент красной рамкой.
*   **`exists(element: str)` -> `bool`**: Проверяет существование элемента по ID.
*   **`send_v_key(element: str = "wnd[0]", *, focus_element: Optional[str] = None, value: int = 0)` -> `None`**: Отправляет виртуальную клавишу (VKey) элементу.
*   **`read_html_viewer(element: str)` -> `str`**: Читает HTML-содержимое из элемента `GuiHTMLViewer`.
*   **`read_shell_table(element: str, load_table: bool = True)` -> `ShellTable`**:
    Читает данные из таблицы (`GuiGridView`) и возвращает объект `ShellTable`.
    *   `load_table`: Если `True`, пытается прокрутить таблицу для загрузки всех строк.
*   **`get_tree(element_id: str)` -> `GuiTree`**:
    Возвращает объект `GuiTree` для взаимодействия с древовидным элементом.
    *   Вызывает: `ElementNotFoundException`, `InvalidElementTypeException`, `SapGuiComException`.
*   **`get_status_message(window_id: str = "wnd[0]")` -> `Optional[Tuple[str, str, str, str]]`**:
    Читает сообщение из строки состояния. Возвращает кортеж `(type, id, number, text)` или `None`.
*   **`assert_status_bar(...)` -> `bool`**:
    Проверяет сообщение в строке состояния на соответствие ожиданиям. Имеет множество параметров для задания ожидаемых значений типа, ID, номера, текста сообщения, а также таймауты.
    *   Вызывает: `StatusBarAssertionError`, `StatusBarException`, `ValueError`.
*   **`select_menu_item_by_name(menu_path: List[str], window_id: str = "wnd[0]")` -> `None`**:
    Выбирает пункт меню, перемещаясь по именам пунктов. `menu_path` - список строк.
    *   Вызывает: `MenuNotFoundException`, `ActionException`.
*   **`start_transaction_robust(transaction: str, check_errors: bool = True)` -> `None`**:
    Запускает транзакцию с префиксом `/n` и опционально проверяет строку состояния на стандартные ошибки (не найдено, нет авторизации, заблокировано).
    *   Вызывает: `TransactionNotFoundError`, `AuthorizationError`, `ActionBlockedError`, `ActionException`.
*   **`iterate_elements_by_template(root_element_id: str, id_template: str, start_index: int, max_index: int = 50)` -> `Generator[Tuple[int, win32com.client.CDispatch], None, None]`**:
    Итерирует по элементам GUI, ID которых соответствует шаблону (например, `usr/tabsTABSTRIP/tabpTAB{index}`).
*   **`print_all_elements(root_element_id: str = "wnd[0]")` -> `None`**:
    Выводит ID и типы всех прямых дочерних элементов указанного корневого элемента.
*   **`scroll_element(element_id: str, position: int)` -> `None`**:
    Прокручивает вертикальный скроллбар элемента к указанной позиции.
*   **`get_element_property(element_id: str, property_name: str)` -> `Any`**:
    Получает значение указанного свойства элемента.
    *   Вызывает: `ElementNotFoundException`, `PropertyNotFoundException`.
*   **`set_element_property(element_id: str, property_name: str, value: Any)` -> `None`**:
    Устанавливает значение свойства элемента (использовать с осторожностью).
*   **`dump_element_state(element_id: str, recursive: bool = True, max_depth: int = 3, print_output: bool = True)` -> `Optional[Dict[str, Any]]`**:
    Собирает и опционально выводит состояние элемента (свойства, дочерние элементы).
*   **`save_gui_snapshot(filepath: Union[str, Path], root_element_id: str = "wnd[0]", max_depth: Optional[int] = None, properties_to_include: Optional[List[str]] = None, properties_to_exclude: Optional[List[str]] = None, output_format: str = 'json', include_children: bool = True)` -> `None`**:
    Создает "слепок" иерархии GUI и сохраняет его в файл JSON или YAML.
    *   `properties_to_include`: Список свойств для включения. Если `None` - все доступные.
    *   `properties_to_exclude`: Список свойств для исключения.
*   **`save_gui_snapshot_from_schema(filepath: Union[str, Path], object_schema: Dict[str, Any], ...)` -> `None`**:
    Аналогично `save_gui_snapshot`, но использует предоставленную схему объектов (например, из `sap_gui_objects.json`) для определения, какие свойства пытаться прочитать.
*   **`handle_unexpected_popup(...)` -> `bool`**:
    Обнаруживает и пытается обработать простые неожиданные всплывающие окна.
    *   `popup_ids`: Список ID окон для проверки (по умолчанию `["wnd[1]", "wnd[2]"]`).
    *   `press_no_button_id`: ID кнопки "Нет" (или отмены), имеет высший приоритет.
    *   `press_button_id`: ID основной кнопки (ОК, Да).
    *   `action_vkey`: VKey для отправки, если кнопки не найдены.
    *   Возвращает: `True`, если окно обработано.
*   **Методы с семантическими локаторами:**
    *   **`find_element_id_by_locator(locator_str: str, target_element_types: Optional[List[str]] = None)` -> `Optional[str]`**:
        Находит ID элемента по семантическому локатору. Использует `SapElementFinder`.
    *   **`press_by_locator(locator_str: str, target_element_types: Optional[List[str]] = ...)` -> `None`**
    *   **`write_by_locator(locator_str: str, text: str, target_element_types: Optional[List[str]] = ...)` -> `None`**
    *   **`read_by_locator(locator_str: str, target_element_types: Optional[List[str]] = ...)` -> `str`**
    *   **`select_by_locator(locator_str: str, target_element_types: Optional[List[str]] = ...)` -> `None`**
    *   **`is_selected_by_locator(locator_str: str, target_element_types: Optional[List[str]] = ...)` -> `bool`**
    *   **`set_checkbox_by_locator(locator_str: str, selected: bool, target_element_types: Optional[List[str]] = ...)` -> `None`**
    *   **`visualize_by_locator(locator_str: str, seconds: int = 1, target_element_types: Optional[List[str]] = None)` -> `None`**
    *   **`exists_by_locator(locator_str: str, target_element_types: Optional[List[str]] = None)` -> `bool`**
    *   Для этих методов `target_element_types` определяет, среди каких типов элементов искать целевой объект. Если `None`, используются типы по умолчанию из `SapElementFinder.DEFAULT_TARGET_TYPES`.

---

### Class `sapscriptwizard.element_finder.SapElementFinder`
Этот класс отвечает за поиск элементов SAP GUI с использованием семантических локаторов. Обычно он используется внутренне классом `Window` и не требует прямого вызова разработчиком, но понимание его работы полезно.

```python
class SapElementFinder:
    def __init__(self, session_handle: win32com.client.CDispatch)
```
*   `session_handle`: COM-объект текущей сессии SAP.

**Ключевые аспекты:**

*   **Кэширование:** `SapElementFinder` кэширует информацию об элементах (ID, тип, текст, позиция) текущего активного окна для ускорения повторных поисков. Кэш автоматически обновляется при смене активного окна SAP.
    *   `_check_and_refresh_cache()`: Проверяет и обновляет кэш.
    *   `_scan_window_elements()`: Сканирует активное окно и заполняет кэш.
*   **Парсинг локаторов:**
    *   `_parse_locator(locator_str: str)`: Преобразует строку локатора в объект стратегии поиска (`ContentLocator`, `HLabelLocator` и т.д.).
*   **Поиск элементов:**
    *   `find_element(locator_str: str, target_element_types: Optional[List[str]] = None)`: Основной метод поиска.
        *   `locator_str`: Строка семантического локатора (см. раздел [Работа с семантическими локаторами](#7-работа-с-семантическими-локаторами)).
        *   `target_element_types`: Список типов SAP GUI элементов (например, `["GuiTextField", "GuiButton"]`), среди которых будет искаться целевой элемент. Если `None`, используются `DEFAULT_TARGET_TYPES` (поля ввода, кнопки и т.д.).
        *   Возвращает: ID найденного элемента или `None`.
*   **`DEFAULT_TARGET_TYPES`**: Список типов элементов по умолчанию, которые считаются целью локаторов (например, `GuiTextField`, `GuiButton`).
*   **`LABEL_ELEMENT_TYPES`**: Список типов элементов, которые могут выступать в роли меток (например, `GuiLabel`).

---

### Module `sapscriptwizard.locator_helpers`
Содержит dataclass'ы, используемые `SapElementFinder`.

*   **`@dataclass(frozen=True) Position`**:
    Хранит позицию (left, top, width, height) и размеры элемента. Включает вычисляемые свойства (right, bottom, center_x, center_y) и методы для проверки взаимного расположения (`is_horizontally_aligned_with`, `is_right_of` и т.д.).
*   **`@dataclass ElementInfo`**:
    Хранит информацию о найденном элементе GUI: `element_id`, `element_type`, `text`, `tooltip`, `position`, `name`, `changeable`.
*   **`@dataclass LocatorStrategy`**: Базовый класс для стратегий локаторов.
    *   **`ContentLocator(value: str)`**: Поиск по точному совпадению текста или tooltip элемента. Локатор: `"=Текст кнопки"`
    *   **`HLabelLocator(label: str)`**: Поиск элемента справа от горизонтальной метки. Локатор: `"Метка"`
    *   **`VLabelLocator(label: str)`**: Поиск элемента под вертикальной меткой. Локатор: `"@ Метка"`
    *   **`HLabelVLabelLocator(h_label: str, v_label: str)`**: Поиск элемента на пересечении горизонтальной и вертикальной метки (справа от `h_label` и ниже `v_label`). Локатор: `"Гориз. метка @ Верт. метка"`
    *   **`HLabelHLabelLocator(left_label: str, right_label: str)`**: Поиск элемента между двумя горизонтальными метками/элементами. Локатор: `"Левая метка >> Правая метка"`

---

### Class `sapscriptwizard.shell_table.ShellTable`
Представляет таблицу SAP (`GuiShell` типа `GridView` или `GuiGridView`).

```python
class ShellTable:
    def __init__(self, session_handle: win32com.client.CDispatch, element: str, load_table: bool = True)
```
*   `session_handle`: COM-объект сессии SAP.
*   `element`: ID элемента таблицы.
*   `load_table`: Если `True`, пытается прокрутить таблицу для загрузки всех строк при инициализации.
*   Вызывает: `ElementNotFoundException`, `InvalidElementTypeException`, `ActionException`.

**Атрибуты:**

*   **`data`**: `polars.DataFrame`, содержащий данные таблицы.
*   **`rows`**: Количество строк.
*   **`columns`**: Количество колонок.

**Методы:**

*   **`_read_shell_table(load_table: bool = True)` -> `pl.DataFrame`**: (Внутренний) Читает данные из COM-объекта таблицы.
*   **`to_polars_dataframe()` -> `pl.DataFrame`**: Возвращает копию данных таблицы как Polars DataFrame.
*   **`to_pandas_dataframe()` -> `pandas.DataFrame`**: Конвертирует и возвращает данные таблицы как Pandas DataFrame.
*   **`to_dict(as_series: bool = False)` -> `Dict[str, Any]`**: Возвращает данные таблицы как словарь.
*   **`to_dicts()` -> `List[Dict[str, Any]]`**: Возвращает данные таблицы как список словарей (один словарь на строку).
*   **`to_csv(file_path: Union[str, Path], separator: str = ';', include_header: bool = True, **kwargs)` -> `None`**:
    Сохраняет данные таблицы в CSV файл.
    *   `**kwargs`: Дополнительные аргументы для `polars.DataFrame.write_csv()`.
*   **`get_column_names()` -> `List[str]`**: Возвращает список имен колонок.
*   **`cell(row: int, column: Union[str, int])` -> `Any`**: Возвращает значение ячейки.
*   **`load(move_by: int = 20, move_by_table_end: int = 2)` -> `None`**:
    Прокручивает таблицу для загрузки всех строк (SAP обычно загружает только видимые).
*   **`press_button(button: str)` -> `None`**: Нажимает кнопку на панели инструментов таблицы.
*   **`select_rows(indexes: List[int])` -> `None`**: Выбирает строки по их 0-based индексам.
*   **`change_checkbox(row: int, column_id: str, flag: bool)` -> `None`**: Устанавливает состояние чекбокса в ячейке.

---

### Class `sapscriptwizard.gui_tree.GuiTree`
Представляет древовидный элемент SAP (`GuiShell` типа `Tree` или `GuiTreeControl`).

```python
class GuiTree:
    def __init__(self, session_handle: win32com.client.CDispatch, element_id: str)
```
*   `session_handle`: COM-объект сессии SAP.
*   `element_id`: ID элемента дерева.
*   Вызывает: `ElementNotFoundException`, `InvalidElementTypeException`, `SapGuiComException`.

**Методы:**

*   **`expand_node(node_key: str)` -> `None`**: Разворачивает узел.
*   **`collapse_node(node_key: str)` -> `None`**: Сворачивает узел.
*   **`select_node(node_key: str, ensure_visible_first: bool = False, top_node_key_if_needed: Optional[str] = None)` -> `None`**:
    Выбирает узел.
    *   `ensure_visible_first`: Если `True`, пытается сделать узел видимым (`TopNode`), прежде чем выбрать.
    *   `top_node_key_if_needed`: Ключ узла для установки как `TopNode`.
*   **`selected_node` (property) -> `Optional[str]`**: Возвращает ключ текущего выбранного узла.
*   **`top_node` (property) -> `Optional[str]`**: Возвращает ключ самого верхнего видимого узла.
*   **`set_top_node(node_key: str)` -> `None`**: Устанавливает самый верхний видимый узел.
*   **`get_node_text(node_key: str)` -> `str`**: Возвращает отображаемый текст узла.
*   **`get_all_node_keys()` -> `List[str]`**: Возвращает список ключей всех загруженных узлов.
*   **`get_column_names()` -> `List[str]`**: Возвращает технические имена колонок (для деревьев с колонками).
*   **`get_item_text(node_key: str, column_name: str)` -> `str`**: Возвращает текст ячейки в строке узла (для деревьев-списков/колоночных).
*   **`double_click_node(node_key: str)` -> `None`**: Выполняет двойной клик по узлу.
*   **`get_node_children_info(parent_node_key: str, auto_expand: bool = True)` -> `List[Tuple[str, str]]`**:
    Возвращает информацию (ключ, текст) о прямых дочерних элементах узла.
    *   `auto_expand`: Если `True`, пытается развернуть родительский узел.
*   **`find_node_key_by_text(target_text: str, case_sensitive: bool = False, search_depth: Optional[int] = None)` -> `Optional[str]`**:
    Находит ключ узла по его отображаемому тексту. `search_depth` пока не полностью реализован для поиска по всей глубине неразвернутых узлов.

---

### Function `sapscriptwizard.parallel.api.run_parallel`
Основная функция для запуска сценариев SAP параллельно или последовательно.

```python
def run_parallel(
    enabled: bool,
    num_processes: int,
    worker_function: Callable[[Window, List[Any]], Any],
    input_data_list: Optional[List[Any]] = None,
    input_data_file: Optional[str] = None,
    interactive: bool = False,
    **runner_kwargs: Any
) -> Optional[Any]
```
*   `enabled`: Если `True`, запускает параллельное выполнение. Иначе - последовательное.
*   `num_processes`: Желаемое количество параллельных процессов (если `enabled=True` и `mode='new'`).
*   `worker_function`: Пользовательская функция, принимающая объект `Window` и список данных для обработки. Может возвращать результат в последовательном режиме.
*   `input_data_list`: Список входных данных.
*   `input_data_file`: Путь к файлу с входными данными (каждая строка - один элемент).
*   `interactive`: Если `True`, запрашивает у пользователя выбор соединения, режим (`new`/`existing`) и сессии. Иначе используются значения по умолчанию.
*   `**runner_kwargs`: Дополнительные аргументы для `SapParallelRunner` (например, `popup_check_delay`).
*   Возвращает: Результат `worker_function` в последовательном режиме, `None` в параллельном.
*   Вызывает: `ValueError`, `FileNotFoundError`, `AttachException`, `SystemExit`.

Эта функция выполняет сканирование доступных сессий, взаимодействует с пользователем (если `interactive=True`) для выбора целевых сессий или режима создания новых, а затем либо выполняет `worker_function` в текущем процессе (если `enabled=False`), либо создает и запускает `SapParallelRunner`.

---

### Class `sapscriptwizard.parallel.runner.SapParallelRunner`
Управляет параллельным выполнением `worker_function` в нескольких сессиях SAP.

```python
class SapParallelRunner:
    def __init__(self,
                 num_processes: int,
                 worker_function: WorkerFunctionType,
                 target_connection_index: int,
                 mode: str, # 'new' or 'existing'
                 target_session_indices: Optional[List[int]],
                 input_data_file: Optional[str] = None,
                 input_data_list: Optional[List[Any]] = None,
                 session_attach_interval: int = 5,
                 popup_check_delay: int = 10,
                 wait_before_launch: int = 15)
```
*   `num_processes`: Эффективное количество параллельных процессов.
*   `worker_function`: Функция для выполнения в каждой сессии.
*   `target_connection_index`: Индекс соединения SAP.
*   `mode`: `'new'` (создать новые сессии) или `'existing'` (использовать существующие).
*   `target_session_indices`: Список индексов сессий для использования (если `mode='existing'`).
*   `input_data_file` / `input_data_list`: Источники данных.
*   `session_attach_interval`: Задержка (в сек) между попытками подключения рабочего процесса к сессии.
*   `popup_check_delay`: Задержка (в сек) после открытия нового окна (для `mode='new'`).
*   `wait_before_launch`: Задержка (в сек) перед запуском рабочих процессов (для `mode='new'`).

**Методы:**

*   **`run()` -> `None`**: Основной метод, запускающий весь процесс:
    1.  Чтение данных (`_read_data`).
    2.  Если `mode='new'`, открытие новых сессий (`_open_sessions`), при этом определяется `_actual_session_indices_to_use`. Учитывается лимит в 6 сессий.
    3.  Если `mode='existing'`, используются `target_session_indices` как `_actual_session_indices_to_use`.
    4.  Разделение данных на части и подготовка временных файлов (`_split_list`, `_prepare_data_files`).
    5.  Запуск рабочих процессов (`_launch_workers`). Каждый процесс нацелен на определенный `connection_index` и `session_index` из `_actual_session_indices_to_use`.
    6.  Ожидание завершения всех рабочих процессов (`_wait_for_workers`).
    7.  Очистка временных файлов (`_cleanup_temp_files`).

*   **`_worker_process_target(worker_function, file_path, connection_index, session_index)`**: (Статический метод)
    Функция, выполняемая каждым дочерним процессом.
    1.  Читает свою порцию данных из `file_path`.
    2.  Создает экземпляр `Sapscript`.
    3.  Пытается подключиться к *назначенному* `connection_index` и `session_index` с несколькими попытками.
    4.  Вызывает пользовательскую `worker_function`, передавая ей объект `Window` и данные.
    5.  Обрабатывает исключения, включая создание скриншотов.

---

### Module `sapscriptwizard.types_.exceptions`
Определяет пользовательские исключения, специфичные для библиотеки.

*   `WindowDidNotAppearException(Exception)`: Окно SAP не появилось.
*   `AttachException(Exception)`: Ошибка подключения к соединению или сессии.
*   `ActionException(Exception)`: Ошибка выполнения действия (клик, выбор и т.д.).
*   `SapGuiComException(Exception)`: Общая ошибка взаимодействия с COM-объектом SAP GUI.
*   `ElementNotFoundException(SapGuiComException)`: Элемент GUI не найден по ID.
*   `PropertyNotFoundException(SapGuiComException)`: У элемента нет запрошенного свойства.
*   `InvalidElementTypeException(SapGuiComException)`: Элемент имеет неверный тип для операции.
*   `MenuNotFoundException(ElementNotFoundException)`: Пункт меню не найден по имени.
*   `StatusBarException(SapGuiComException)`: Ошибка чтения строки состояния.
*   `TransactionException(Exception)`: Базовое исключение для ошибок транзакций.
*   `TransactionNotFoundError(TransactionException)`: Код транзакции не существует.
*   `AuthorizationError(TransactionException)`: Пользователь не авторизован.
*   `ActionBlockedError(TransactionException)`: Действие заблокировано (например, объект заблокирован).
*   `SapLogonConfigError(Exception)`: Ошибка, связанная с конфигурацией `saplogon.ini`.
*   `StatusBarAssertionError(SapGuiComException)`: Содержимое строки состояния не соответствует ожиданиям.

---

### Module `sapscriptwizard.types_.types`

*   **`enum NavigateAction(Enum)`**:
    Перечисление для метода `Window.navigate()`: `enter`, `back`, `end`, `cancel`, `save`.

---

### Class `sapscriptwizard.utils.sap_config.SapLogonConfig`
Утилита для чтения информации из файлов `saplogon.ini`.

```python
class SapLogonConfig:
    # Является синглтоном для хранения путей к ini-файлам
    def set_ini_files(self, *file_paths: Union[str, Path]) -> None
    def get_connect_name_by_sid(self, sid: str, first_only: bool = True) -> Optional[Union[str, List[str]]]
```
*   **`set_ini_files(*file_paths)`**: Устанавливает пути к одному или нескольким файлам `saplogon.ini`.
*   **`get_connect_name_by_sid(sid, first_only=True)`**:
    Находит имя (описание) соединения по его SID.
    *   `first_only`: Если `True`, возвращает первое найденное имя, иначе список всех найденных имен.
    *   Вызывает `SapLogonConfigError`, если ini-файлы не установлены или SID не найден.

---

### Module `sapscriptwizard.utils.utils`

*   **`kill_process(process: str)`**: Принудительно завершает процесс по его имени.
*   **`wait_for_window_title(title: str, timeout_loops: int = 10)`**:
    Ожидает появления окна с указанным заголовком.
    *   Вызывает: `WindowDidNotAppearException`.

---

## 7. Работа с семантическими локаторами

Семантические локаторы позволяют находить элементы GUI без использования их жестко заданных ID, опираясь на их текстовое содержимое или взаимное расположение. Это делает скрипты более надежными и читаемыми.

**Поддерживаемые типы локаторов:**

1.  **По содержимому (Content Locator):**
    *   Синтаксис: `=ТекстЭлемента`
    *   Ищет элемент (любого типа из кэша), у которого свойство `Text` или `Tooltip` точно совпадает с `ТекстЭлемента`.
    *   Пример: `window.press_by_locator("=Сохранить")`

2.  **По горизонтальной метке (HLabel Locator):**
    *   Синтаксис: `ТекстМетки`
    *   Ищет элемент-метку (типа `GuiLabel` или `GuiTextField`/`GuiCTextField`, используемый как метка) с текстом `ТекстМетки`. Затем ищет целевой элемент (например, поле ввода) справа от этой метки, выровненный с ней по горизонтали и находящийся как можно ближе.
    *   Пример: `window.write_by_locator("Пользователь", "MYUSER")`

3.  **По вертикальной метке (VLabel Locator):**
    *   Синтаксис: `@ТекстМетки`
    *   Аналогично HLabel, но ищет целевой элемент под меткой, выровненный по вертикали.
    *   Пример: `window.write_by_locator("@Пароль", "secret")`

4.  **По горизонтальной и вертикальной меткам (HLabelVLabel Locator):**
    *   Синтаксис: `ГоризонтальнаяМетка @ ВертикальнаяМетка`
    *   Ищет элемент, который находится правее `ГоризонтальнаяМетка` и ниже `ВертикальнаяМетка`, и выровнен с ними соответственно.
    *   Пример: `window.write_by_locator("Область данных @ Значение", "123")` (если "Область данных" слева, а "Значение" сверху от поля)

5.  **По двум горизонтальным меткам/элементам (HLabelHLabel Locator):**
    *   Синтаксис: `ЛевыйЭлемент >> ПравыйЭлемент`
    *   Ищет `ЛевыйЭлемент` (может быть меткой или полем с текстом), затем ищет `ПравыйЭлемент` (поле ввода, кнопка и т.д. с указанным текстом/tooltip), который находится справа от `ЛевыйЭлемент` и выровнен с ним по горизонтали. Используется, когда сама метка искомого поля не уникальна, но комбинация с соседним элементом уникальна.
    *   Пример: `window.write_by_locator("От >> До", "31.12.2023")` (если поле "До" находится справа от поля или метки "От")

**Использование:**

Методы `Window`, имеющие суффикс `_by_locator` (например, `press_by_locator`, `write_by_locator`), принимают строку локатора вместо ID элемента.

```python
# Найти поле ввода справа от метки "Имя пользователя" и ввести текст
window.write_by_locator("Имя пользователя", "testuser")

# Найти кнопку с текстом "Выполнить" и нажать ее
window.press_by_locator("=Выполнить")

# Найти поле ввода под меткой "Код компании"
company_code = window.read_by_locator("@Код компании")
```

**Важно:**
*   Поиск по локаторам выполняется в текущем активном окне SAP.
*   Эффективность и точность зависят от уникальности меток и структуры экрана.
*   Для отладки локаторов можно использовать `window.visualize_by_locator("Ваш локатор")`.
*   Параметр `target_element_types` в методах `_by_locator` позволяет сузить поиск до определенных типов элементов (например, искать только `GuiTextField`).

## 8. Параллельное выполнение сценариев

Модуль `sapscriptwizard.parallel` предоставляет возможность запускать задачи автоматизации SAP в нескольких сессиях одновременно. Это может значительно ускорить обработку больших объемов данных или выполнение множества независимых операций.

**Основные компоненты:**

*   **`run_parallel()` (из `parallel.api`)**: Точка входа для запуска параллельного выполнения.
*   **`SapParallelRunner` (из `parallel.runner`)**: Класс, управляющий созданием процессов и распределением задач.
*   **Рабочая функция (Worker Function)**: Пользовательская функция, которая будет выполняться в каждом параллельном процессе. Она должна принимать два аргумента:
    1.  `window: Window`: Объект `Window`, подключенный к конкретной сессии SAP.
    2.  `data_chunk: List[Any]`: Порция данных, назначенная этому процессу.

**Пример рабочей функции:**

```python
from sapscriptwizard import Window # Внутри файла с рабочей функцией
import logging

logger = logging.getLogger(__name__) # Логгер для рабочей функции

def my_sap_worker(window: Window, data_items: list):
    logger.info(f"Процесс {window.session} начал работу с {len(data_items)} элементами.")
    try:
        for item in data_items:
            # Пример: Запуск транзакции и ввод данных
            window.start_transaction("MM03")
            window.write_by_locator("Материал", str(item)) # Предполагаем, что локатор "Материал" существует
            window.press_by_locator("=Основные данные") # Пример кнопки
            # ... другие действия ...
            logger.info(f"Обработан элемент: {item} в сессии {window.session}")
            # Важно: обрабатывайте исключения внутри рабочей функции, если это необходимо
            # для продолжения обработки других элементов в этом же процессе.
    except Exception as e:
        logger.error(f"Ошибка в сессии {window.session} при обработке элемента {item if 'item' in locals() else 'N/A'}: {e}")
        # Скриншот будет сделан автоматически, если включен в Sapscript и ошибка не перехвачена здесь
        raise # Перебрасываем, чтобы SapParallelRunner мог залогировать и сделать скриншот
```

**Запуск параллельного выполнения:**

```python
from sapscriptwizard.parallel.api import run_parallel
# from my_workers import my_sap_worker # Предполагается, что my_sap_worker в этом файле

# ... (код my_sap_worker выше) ...

if __name__ == "__main__":
    # Настройка основного логирования
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')

    # Данные для обработки
    my_data = [f"MATERIAL_{i:03}" for i in range(10)] # 10 материалов

    try:
        run_parallel(
            enabled=True,                     # Включить параллельный режим
            num_processes=3,                  # Желаемое количество новых сессий (если mode='new')
            worker_function=my_sap_worker,    # Ваша рабочая функция
            input_data_list=my_data,          # Передача данных списком
            # input_data_file="path/to/my_data.txt", # или из файла
            interactive=True,                 # Запросить у пользователя выбор сессий/режима
            # **runner_kwargs:
            popup_check_delay=5,              # Задержка после открытия нового окна
            wait_before_launch=10             # Задержка перед стартом рабочих процессов
        )
        logging.info("Параллельная обработка завершена.")
    except Exception as e:
        logging.error(f"Ошибка при запуске параллельной обработки: {e}", exc_info=True)

```

**Режимы работы `run_parallel`:**

1.  **`mode='new'` (по умолчанию для неинтерактивного режима, выбирается в интерактивном):**
    *   `SapParallelRunner` попытается открыть `num_processes` новых сессий SAP в указанном `target_connection_index`.
    *   Учитывается лимит SAP в 6 сессий на соединение. Если запрошено больше, чем можно открыть, количество процессов будет уменьшено.
    *   Рабочие процессы будут подключены к этим вновь открытым сессиям.

2.  **`mode='existing'` (выбирается в интерактивном режиме):**
    *   Пользователю будет предложено выбрать существующие активные сессии из указанного `target_connection_index`.
    *   `num_processes` будет равно количеству выбранных пользователем сессий.
    *   Рабочие процессы будут подключены к этим уже существующим сессиям.

**Важные моменты:**

*   **Изоляция процессов:** Каждый рабочий процесс выполняется независимо. Ошибки в одном процессе (если не фатальные для SAP GUI) не должны влиять на другие.
*   **Распределение данных:** `SapParallelRunner` автоматически делит `input_data_list` или содержимое `input_data_file` на примерно равные части для каждого процесса.
*   **Логирование:** Рекомендуется настраивать логирование как в основном скрипте, так и внутри рабочей функции для отладки. Логи из разных процессов могут перемешиваться, если не использовать специальные обработчики.
*   **Обработка ошибок:** Ошибки в рабочих процессах логируются, и `Sapscript` внутри процесса попытается сделать скриншот.
*   **COM Инициализация:** Каждый процесс инициализирует COM отдельно. `pythoncom.CoInitialize()` вызывается автоматически `pywin32` при первом использовании COM в потоке/процессе.

## 9. Обработка исключений

Библиотека использует ряд пользовательских исключений (см. `sapscriptwizard.types_.exceptions`) для обозначения различных проблем. Рекомендуется оборачивать код взаимодействия с библиотекой в блоки `try...except` для корректной обработки ошибок.

```python
from sapscriptwizard import Sapscript, exceptions

sap = Sapscript()
sap.enable_screenshots_on_error() # Включить скриншоты при ошибках

try:
    window = sap.attach_window(0, 0)
    window.write("несуществующий_id", "текст")
except exceptions.ElementNotFoundException as e:
    print(f"Элемент не найден: {e}")
    sap.handle_exception_with_screenshot(e)
except exceptions.AttachException as e:
    print(f"Ошибка подключения: {e}")
    sap.handle_exception_with_screenshot(e)
except exceptions.SapGuiComException as e: # Более общее исключение SAP
    print(f"Общая ошибка SAP GUI: {e}")
    sap.handle_exception_with_screenshot(e)
except Exception as e: # Любые другие ошибки
    print(f"Непредвиденная ошибка: {e}")
    sap.handle_exception_with_screenshot(e)
```

Метод `Sapscript.handle_exception_with_screenshot(e)` логирует исключение и, если включено, сохраняет скриншот текущего экрана в директорию, указанную через `set_screenshot_directory()`.

## 10. Логирование
Библиотека `sapscriptwizard` использует стандартный модуль `logging` Python. Для просмотра сообщений от библиотеки необходимо настроить обработчик логирования в вашем основном скрипте.

Пример базовой настройки:
```python
import logging

logging.basicConfig(
    level=logging.INFO,  # Уровни: DEBUG, INFO, WARNING, ERROR, CRITICAL
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Если нужно более детальное логирование от конкретного модуля библиотеки:
# logging.getLogger('sapscriptwizard.window').setLevel(logging.DEBUG)
# logging.getLogger('sapscriptwizard.element_finder').setLevel(logging.DEBUG)

# ... остальной ваш код ...
```
Логи помогут отслеживать выполнение операций, диагностировать проблемы и понимать поведение семантических локаторов и кэша элементов.

## 11. Продвинутые темы и советы

### Кэширование элементов
Класс `SapElementFinder` (используемый `Window` для семантических локаторов) кэширует информацию обо всех элементах текущего активного окна SAP.
*   **Преимущества:** Значительно ускоряет последовательные вызовы методов поиска (`find_element_id_by_locator`, `press_by_locator` и т.д.) в пределах одного и того же экрана SAP, так как не требует повторного сканирования всех элементов.
*   **Обновление кэша:** Кэш автоматически инвалидируется и обновляется при изменении активного окна SAP (например, при переходе на новый экран, открытии диалогового окна). Это определяется по изменению `Id` объекта `session.ActiveWindow`.
*   **Ручное обновление:** Обычно не требуется. Если есть подозрение, что кэш устарел, а окно не изменилось (редкий случай, возможно, при динамическом обновлении части экрана без смены `ActiveWindow.Id`), можно попробовать создать новый экземпляр `Window` или временно взаимодействовать через ID, если это критично. Однако, в большинстве случаев автоматическое обновление работает корректно.

### Отладка и анализ GUI
*   **SAP GUI Scripting Tracker:** Ваш главный инструмент. (Alt+F12 в SAP GUI -> Script Recording and Playback). Запись скрипта покажет ID элементов, их типы и вызываемые методы.
*   **`Window.print_all_elements(root_element_id)`**: Выводит ID и типы прямых дочерних элементов указанного контейнера (например, `wnd[0]`). Полезно для понимания структуры текущего экрана.
*   **`Window.dump_element_state(element_id, recursive=True, max_depth=N)`**: Выводит в консоль подробную информацию о свойствах элемента и его дочерних элементах (рекурсивно до `max_depth`).
*   **`Window.save_gui_snapshot(filepath, ...)`**: Сохраняет полный "слепок" иерархии GUI в JSON или YAML файл. Это может быть очень полезно для анализа структуры сложных экранов вне SAP, сравнения состояний GUI или предоставления информации для отладки.
    *   `save_gui_snapshot_from_schema` позволяет использовать предопределенную схему свойств, что может быть полезно для стандартизированного сбора данных или если `dir()` на COM-объектах работает непредсказуемо.
*   **`Window.visualize(element_id)` / `Window.visualize_by_locator(locator_str)`**: Подсвечивает элемент красной рамкой, позволяя визуально проверить, тот ли элемент был найден.
*   **Логирование:** Включение `DEBUG` уровня логирования для модулей `sapscriptwizard` даст много информации о процессе поиска элементов и других внутренних операциях.

### Работа с COM-объектами
*   Библиотека абстрагирует прямое взаимодействие с COM, но понимание основ полезно.
*   `_ensure_com_objects()` в `Sapscript` отвечает за получение корневых COM-объектов (`SAPGUI`, `ScriptingEngine`).
*   Каждый элемент SAP GUI является COM-объектом со своими свойствами и методами, которые можно найти в документации SAP Scripting API или с помощью Scripting Tracker.
*   Будьте осторожны при прямом вызове методов COM-объектов, которых нет в API `sapscriptwizard`, так как это может привести к неожиданному поведению или ошибкам.

