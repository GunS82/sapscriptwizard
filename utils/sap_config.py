# --- НОВЫЙ ФАЙЛ: pysapscript\utils\sap_config.py ---
"""Utilities for reading saplogon.ini configuration files."""
import configparser
import os
from pathlib import Path
from typing import Union, List, Optional

from sapscriptwizard.types_.exceptions import SapLogonConfigError

DESCRIPTION_SECTION = "Description"
SID_SECTION = "MSSysName"

class SapLogonConfig:
    """
    Utility class to read information from saplogon.ini files.
    Based on SAPLogonINI from pysapgui.
    """
    _instance = None
    _ini_files: List[Path] = []

    def __new__(cls, *args, **kwargs):
        # Singleton pattern to hold ini file paths globally if needed
        if not cls._instance:
            cls._instance = super(SapLogonConfig, cls).__new__(cls)
        return cls._instance

    def set_ini_files(self, *file_paths: Union[str, Path]) -> None:
        """
        Sets the paths to the saplogon.ini files to be searched.

        Args:
            *file_paths (Union[str, Path]): One or more paths to saplogon.ini files.
        """
        self._ini_files = []
        for file_path in file_paths:
            path = Path(file_path)
            if path.exists() and path.is_file():
                self._ini_files.append(path.resolve())
            else:
                # Optionally raise an error or log a warning
                print(f"Warning: saplogon.ini file not found or is not a file: {path}")

    def get_connect_name_by_sid(self, sid: str, first_only: bool = True) -> Optional[Union[str, List[str]]]:
        """
        Finds the connection description (name) based on the System ID (SID).

        Args:
            sid (str): The System ID (e.g., "SQ4").
            first_only (bool): If True, returns the first match found.
                               If False, returns a list of all matches.

        Returns:
            Optional[Union[str, List[str]]]: The connection name(s) or None if not found.

        Raises:
            SapLogonConfigError: If no ini files were set or if SID is not found.
        """
        if not self._ini_files:
            raise SapLogonConfigError("No saplogon.ini files have been set using set_ini_files().")

        found_names: List[str] = []
        sid_upper = sid.upper()

        for file_path in self._ini_files:
            config = configparser.ConfigParser()
            try:
                # Use 'utf-8' or 'latin-1' encoding, common for saplogon.ini
                config.read(file_path, encoding='latin-1')

                if SID_SECTION not in config or DESCRIPTION_SECTION not in config:
                    continue # Skip files without the required sections

                items = dict(config[SID_SECTION].items())
                item_id = None
                for key, value in items.items():
                    if value.upper() == sid_upper:
                        item_id = key
                        break

                if item_id and item_id in config[DESCRIPTION_SECTION]:
                    conn_name = config[DESCRIPTION_SECTION][item_id]
                    if first_only:
                        return conn_name
                    else:
                        if conn_name not in found_names: # Avoid duplicates from same file
                             found_names.append(conn_name)

            except configparser.Error as e:
                print(f"Warning: Could not parse {file_path}: {e}")
            except Exception as e:
                 print(f"Warning: An unexpected error occurred reading {file_path}: {e}")


        if not found_names:
            # Raise error only after checking all files
            raise SapLogonConfigError(f"System ID '{sid}' not found in configured saplogon.ini files.")
        else:
            return found_names if not first_only else found_names[0]

# --- КОНЕЦ НОВОГО ФАЙЛА ---