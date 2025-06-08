"""
.. include:: ../README.md
"""

from .sapscriptwizard import Sapscript
from .window import Window
from .shell_table import ShellTable
# --- НОВЫЙ КОД ---
from .gui_tree import GuiTree # Добавляем импорт нового класса
# --- КОНЕЦ НОВОГО КОДА ---
from .run_history import RunHistory
from .types_.types import NavigateAction
from .types_ import exceptions
from .parallel import run_parallel, SapParallelRunner
