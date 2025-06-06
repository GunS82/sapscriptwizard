from typing import Self, Any, ClassVar, overload, Union, Dict, List # Добавлено Union
# --- НОВЫЙ КОД ---
from pathlib import Path # Добавлено для to_csv
# --- КОНЕЦ НОВОГО КОДА ---

import win32com.client
from win32com.universal import com_error
import polars as pl
import pandas

from sapscriptwizard.types_ import exceptions


class ShellTable:
    """
    A class representing a shell table (typically GuiGridView / ALV Grid).
    Provides methods to read data and interact with the table.
    """

    def __init__(self, session_handle: win32com.client.CDispatch, element: str, load_table: bool = True) -> None:
        """
        Args:
            session_handle (win32com.client.CDispatch): SAP session handle
            element (str): SAP table element ID (usually a GuiShell of subtype GridView)
            load_table (bool): Loads table by scrolling if True, default True

        Raises:
            ActionException: error reading table data or wrong element type.
        """
        self.table_element = element
        self._session_handle = session_handle
        try:
            # --- ДОБАВЛЕНИЕ: Проверка типа элемента ---
            self._com_object = self._session_handle.findById(self.table_element)
            element_type = getattr(self._com_object, "Type", "")
            # Чаще всего это GuiShell, но GuiGridView более специфичен
            if element_type not in ["GuiShell", "GuiGridView"]:
                 # Можно добавить другие совместимые типы, если они известны
                 raise exceptions.InvalidElementTypeException(
                     f"Element '{element}' is type '{element_type}', expected GuiShell/GuiGridView for ShellTable.")
            # --- КОНЕЦ ДОБАВЛЕНИЯ ---
        except Exception as e:
             raise exceptions.ElementNotFoundException(f"Error finding element '{element}': {e}")

        self.data = self._read_shell_table(load_table)
        self.rows = self.data.shape[0]
        self.columns = self.data.shape[1]


    def __repr__(self) -> str:
        # Limit representation size for large tables
        with pl.Config(tbl_rows=10, tbl_cols=8):
             return repr(self.data)

    def __str__(self) -> str:
        # Limit string size for large tables
        with pl.Config(tbl_rows=20, tbl_cols=8):
             return str(self.data)

    # --- ИЗМЕНЕНИЕ: __eq__ должен сравнивать данные, а не объекты ---
    def __eq__(self, other: object) -> bool:
        if isinstance(other, ShellTable):
             # Use polars DataFrame equality check
             try:
                  return self.data.equals(other.data)
             except Exception: # Catch potential errors during comparison
                  return False
        elif isinstance(other, pl.DataFrame):
             try:
                  return self.data.equals(other)
             except Exception:
                  return False
        # Comparing with pandas DataFrame might be needed, but requires conversion
        # elif isinstance(other, pandas.DataFrame):
        #     try:
        #         return self.to_pandas_dataframe().equals(other)
        #     except Exception:
        #         return False
        return False
    # --- КОНЕЦ ИЗМЕНЕНИЯ ---

    def __hash__(self) -> hash:
        # Hashing a DataFrame is problematic. Hash based on shape and element ID for basic identity.
        # Note: This might not be ideal if data changes. Consider if hashability is truly needed.
        return hash((self._session_handle.SessionInfo.SystemSessionId, self.table_element, self.data.shape))

    # ... (существующие методы __getitem__, __iter__, _read_shell_table, etc.) ...
    # Копируем существующие методы для контекста
    def __getitem__(self, item) -> Union[Dict[str, Any], List[Dict[str, Any]]]: # Use Union
        if isinstance(item, int):
            # Handle negative index? Polars might do it automatically.
            if item >= self.rows or item < -self.rows :
                raise IndexError(f"Row index {item} out of bounds for table with {self.rows} rows.")
            return self.data.row(item, named=True)
        elif isinstance(item, slice):
            if item.step is not None and item.step != 1:
                raise NotImplementedError("Step slicing is not supported")
            # Polars slice: start, length. Need to calculate length carefully.
            start = item.start if item.start is not None else 0
            stop = item.stop if item.stop is not None else self.rows
            if start < 0: start += self.rows
            if stop < 0: stop += self.rows
            start = max(0, min(start, self.rows))
            stop = max(start, min(stop, self.rows))
            length = stop - start

            sl = self.data.slice(start, length)
            return sl.to_dicts()
        else:
            raise TypeError("Table indices must be integers or slices")

    def __iter__(self) -> Self:
        # Reset index for iteration if needed, or rely on iterator class state
        self._iterator_index = 0
        return self # Or return ShellTableRowIterator(self.data) if using separate iterator class

    def __next__(self) -> Dict[str, Any]:
        # Simple iterator implementation directly in the class
        if self._iterator_index >= self.rows:
            raise StopIteration
        value = self.data.row(self._iterator_index, named=True)
        self._iterator_index += 1
        return value

    def _read_shell_table(self, load_table: bool = True) -> pl.DataFrame:
        """ Reads table data from the GuiShell/GuiGridView element. """
        try:
            # Use the stored COM object
            shell = self._com_object # self._session_handle.findById(self.table_element)

            columns = shell.ColumnOrder
            # Check if columns is a tuple/list or needs conversion
            if not isinstance(columns, (list, tuple)):
                 # Handle cases where ColumnOrder might return a COM collection
                 try:
                      columns = [columns.Item(i) for i in range(columns.Count)]
                 except Exception:
                     raise exceptions.ActionException(f"Could not interpret ColumnOrder for {self.table_element}")

            if not columns:
                 print(f"Warning: No columns found for table {self.table_element}. Returning empty DataFrame.")
                 return pl.DataFrame()

            # Get row count *after* potential loading
            if load_table:
                self.load() # Use the instance's load method
                # Re-fetch row count after loading
                rows_count = shell.RowCount
            else:
                 rows_count = shell.RowCount

            if rows_count == 0:
                return pl.DataFrame()

            # Efficient data reading (consider batching if very large)
            data = []
            for i in range(rows_count):
                row_data = {}
                for column in columns:
                    try:
                        row_data[column] = shell.GetCellValue(i, column)
                    except Exception as cell_ex:
                         # Log error for specific cell, maybe put None or error string
                         print(f"Warning: Error reading cell ({i}, {column}): {cell_ex}")
                         row_data[column] = None # Or some error marker
                data.append(row_data)

            # Handle potential empty data list
            if not data:
                 return pl.DataFrame({col:[] for col in columns})

            return pl.DataFrame(data, schema={col:pl.Utf8 for col in columns}) # Assume string initially

        except Exception as ex:
            raise exceptions.ActionException(f"Error reading element {self.table_element}: {ex}")

    def to_polars_dataframe(self) -> pl.DataFrame:
        """ Get table data as a polars DataFrame """
        return self.data.clone() # Return a copy

    def to_pandas_dataframe(self) -> pandas.DataFrame:
        """ Get table data as a pandas DataFrame """
        return self.data.to_pandas()

    def to_dict(self, as_series: bool = False) -> Dict[str, Any]:
        """ Get table data as a dictionary """
        # Note: Polars' to_dict(as_series=False) returns Dict[str, List[Any]]
        return self.data.to_dict(as_series=as_series)

    def to_dicts(self) -> List[Dict[str, Any]]:
        """ Get table data as a list of dictionaries (one dict per row) """
        return self.data.to_dicts()

    def get_column_names(self) -> List[str]:
        """ Get column names """
        return self.data.columns

    @overload
    def cell(self, row: int, column: int) -> Any: ...

    @overload
    def cell(self, row: int, column: str) -> Any: ...

    def cell(self, row: int, column: Union[str, int]) -> Any:
        """ Get cell value from the DataFrame """
        try:
            return self.data.item(row, column)
        except IndexError:
             raise IndexError(f"Row index {row} out of bounds.")
        except pl.ColumnNotFoundError:
             raise KeyError(f"Column '{column}' not found.")
        except Exception as e:
             raise RuntimeError(f"Error accessing cell ({row}, {column}): {e}")

    def load(self, move_by: int = 20, move_by_table_end: int = 2) -> None:
        """ Skims through the table to load all data, as SAP only loads visible data """
        row_position = 0
        try:
            shell = self._com_object # Use stored object

            # Scroll down quickly first
            while True:
                try:
                    # Setting currentCellRow might be enough to trigger loading
                    shell.currentCellRow = row_position
                    # Optional: Verify if data actually loaded or add a small sleep
                    # sleep(0.05)
                    if row_position >= shell.RowCount - shell.VisibleRowCount: # Stop near the end
                         break
                except com_error as ce:
                    # Check if error code indicates "out of bounds" or similar benign error at the end
                    # HRESULT for "invalid row number" might vary. Check specific error.
                    # For now, just break on any COM error during fast scroll.
                    # print(f"COM error during fast scroll at row {row_position}: {ce}. Assuming end reached.")
                    break
                except Exception as e:
                     print(f"Warning: Unexpected error during fast scroll at row {row_position}: {e}. Stopping scroll.")
                     break # Stop on unexpected errors
                row_position += move_by

            # Scroll slowly near the end to catch final rows
            # Start from slightly before the last known position
            row_position = max(0, row_position - move_by)
            final_row_count = shell.RowCount # Get current count before slow scroll
            while row_position < final_row_count: # Scroll until the current end
                 try:
                     shell.currentCellRow = row_position
                     # sleep(0.05)
                     # Update final_row_count in case more rows load during slow scroll
                     new_row_count = shell.RowCount
                     if new_row_count > final_row_count:
                         final_row_count = new_row_count
                 except com_error:
                     # Benign error expected when trying to set currentCellRow past the actual end
                     # print(f"COM error during slow scroll at row {row_position}. Reached end.")
                     break
                 except Exception as e:
                      print(f"Warning: Unexpected error during slow scroll at row {row_position}: {e}. Stopping scroll.")
                      break
                 row_position += move_by_table_end

            # One final scroll to the very last row might be needed
            try:
                 shell.currentCellRow = max(0, shell.RowCount - 1)
            except Exception:
                 pass # Ignore errors setting to the very last row

        except Exception as e:
            # Don't raise ActionException, just print warning, as reading might still work partially
            print(f"Warning: Error occurred during table load for {self.table_element}: {e}")


    def press_button(self, button: str) -> None:
        """ Presses button that is within the shell/table element's toolbar """
        try:
            # Button might be relative to the shell object
            self._com_object.pressButton(button)
        except Exception as e:
            raise exceptions.ActionException(f"Error pressing shell button '{button}' in {self.table_element}: {e}")

    def select_rows(self, indexes: List[int]) -> None:
        """ Selects rows in the shell table by their 0-based indexes """
        try:
            # Convert list of integers to comma-separated string
            value = ",".join(map(str, indexes))
            self._com_object.selectedRows = value
        except Exception as e:
            raise exceptions.ActionException(
                f"Error selecting rows with indexes {indexes} in {self.table_element}: {e}"
            )

    def change_checkbox(self, row: int, column_id: str, flag: bool) -> None: # Changed signature
        """ Sets checkbox in a specific cell of the shell table """
        try:
            # Ensure flag is boolean
            self._com_object.modifyCell(row, column_id, bool(flag)) # Use modifyCell for checkboxes
            # Note: Some tables might use `changeCheckbox` on a different object or need different args.
            # This assumes `modifyCell` works for checkbox columns in this specific shell type.
        except Exception as e:
            raise exceptions.ActionException(f"Error setting checkbox in cell ({row}, {column_id}) for {self.table_element}: {e}")

    # --- НОВЫЙ КОД: Сохранение в CSV ---
    def to_csv(self, file_path: Union[str, Path], separator: str = ';', include_header: bool = True, **kwargs) -> None:
        """
        Saves the table data to a CSV file using polars.

        Args:
            file_path (Union[str, Path]): The path to the output CSV file.
            separator (str): The delimiter to use in the CSV file. Defaults to ';'.
            include_header (bool): Whether to write the header row. Defaults to True.
            **kwargs: Additional keyword arguments passed directly to polars.DataFrame.write_csv().
                      See Polars documentation for options like encoding, date_format, etc.
        """
        try:
            # Ensure directory exists if file_path includes directories
            path_obj = Path(file_path)
            path_obj.parent.mkdir(parents=True, exist_ok=True)

            print(f"Saving table data to CSV: {path_obj.resolve()}")
            self.data.write_csv(
                file=path_obj,
                separator=separator,
                include_header=include_header,
                **kwargs
            )
            print("Save to CSV complete.")
        except Exception as e:
            raise exceptions.ActionException(f"Error saving table {self.table_element} to CSV '{file_path}': {e}")
    # --- КОНЕЦ НОВОГО КОДА ---
    # --- Существующий итератор (без изменений) ---
class ShellTableRowIterator:
    """ Iterator for shell table rows """
    def __init__(self, data: pl.DataFrame) -> None:
        self.data = data
        self.index = 0

    def __iter__(self) -> Self:
        return self

    def __next__(self) -> Dict[str, Any]:
        if self.index >= self.data.shape[0]:
            raise StopIteration
        value = self.data.row(self.index, named=True)
        self.index += 1
        return value