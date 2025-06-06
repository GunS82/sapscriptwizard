# File: sapscriptwizard/sapscriptwizard.py

import time
import atexit
import logging # Added for logging
from pathlib import Path
from subprocess import Popen
from datetime import datetime
from typing import Optional, Dict, Any, List, Union # Ensure necessary types are imported

import win32com.client
try:
    from PIL import ImageGrab # Pillow library for screenshots
except ImportError:
    ImageGrab = None # Handle case where Pillow is not installed

# Adjust imports based on your structure
from . import window # Use relative import if window.py is in the same directory/package
from .utils import utils
from .types_ import exceptions

# --- Logger for this module ---
log = logging.getLogger(__name__)


class Sapscript:
    """
    Main class for interacting with SAP GUI Scripting API.
    Handles launching SAP, attaching to sessions, managing windows,
    and providing access to GUI elements.
    """
    def __init__(self, default_window_title: str = "SAP Easy Access") -> None:
        """
        Initializes the Sapscript object.

        Args:
            default_window_title (str): Default SAP window title used for checks.
        """
        self._sap_gui_auto: Optional[win32com.client.CDispatch] = None
        self._application: Optional[win32com.client.CDispatch] = None
        self.default_window_title = default_window_title
        # --- Screenshot Attributes ---
        self._screenshots_on_error_enabled: bool = True # Enabled by default
        self._screenshot_directory: Optional[Path] = None # Default to current dir
        # --- History Attribute (managed by methods) ---
        # self._manage_history = False # Example if we wanted auto-management

        # Initial check for COM objects can be deferred until first use via _ensure_com_objects
        log.debug("Sapscript object initialized.")


    def __repr__(self) -> str:
        return f"Sapscript(default_window_title={self.default_window_title})"

    def __str__(self) -> str:
        return f"Sapscript(default_window_title={self.default_window_title})"

    def _ensure_com_objects(self) -> None:
        """
        Ensures the basic COM objects for SAP GUI (_sap_gui_auto, _application)
        are initialized. Raises SapGuiComException if initialization fails.
        """
        if not isinstance(self._sap_gui_auto, win32com.client.CDispatch):
            try:
                log.debug("Attempting to GetObject('SAPGUI')...")
                self._sap_gui_auto = win32com.client.GetObject("SAPGUI")
                log.debug("GetObject('SAPGUI') successful.")
            except Exception as e:
                log.error(f"Failed to GetObject('SAPGUI'): {e}")
                raise exceptions.SapGuiComException(f"Could not get SAPGUI object. Is SAP Logon running? Error: {e}")

        if not isinstance(self._application, win32com.client.CDispatch):
            # Ensure _sap_gui_auto exists before trying to get ScriptingEngine
            if not isinstance(self._sap_gui_auto, win32com.client.CDispatch):
                 # This case should ideally not happen if the first part succeeded, but defensively check.
                 log.error("Cannot get Scripting Engine because SAPGUI object is missing.")
                 raise exceptions.SapGuiComException("Cannot get Scripting Engine, SAPGUI object is missing.")
            try:
                log.debug("Attempting to GetScriptingEngine...")
                self._application = self._sap_gui_auto.GetScriptingEngine
                log.debug("GetScriptingEngine successful.")
            except Exception as e:
                 log.error(f"Failed to GetScriptingEngine: {e}")
                 raise exceptions.SapGuiComException(f"Could not get Scripting Engine. Is GUI Scripting enabled? Error: {e}")

        # Final check
        if not self._application:
            # Should not be reachable if above logic is correct, but as a safeguard
            raise exceptions.SapGuiComException("Failed to initialize SAP GUI Scripting Engine.")


    def launch_sap(self,
                   sid: str,
                   client: str,
                   user: str,
                   password: str,
                   *,
                   root_sap_dir: Path = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui"),
                   maximise: bool = True,
                   language: str = "en",
                   quit_auto: bool = True) -> None:
        """
        Launches SAP using sapshcut.exe and waits for it to load.

        Args:
            sid: SAP system ID.
            client: SAP client.
            user: SAP user.
            password: SAP password.
            root_sap_dir: Path to the SAP GUI installation directory.
            maximise: Maximise window after start if True.
            language: SAP language (e.g., "EN", "DE").
            quit_auto: Register atexit handler to attempt graceful quit if True.

        Raises:
            FileNotFoundError: If sapshcut.exe is not found.
            WindowDidNotAppearException: If the SAP window doesn't appear after launch.
            SapGuiComException: If COM objects cannot be initialized after launch.
        """
        log.info(f"Launching SAP: SID={sid}, Client={client}, User={user}, Lang={language}")
        self._launch(
            root_sap_dir,
            sid,
            client,
            user,
            password,
            maximise,
            language
        )

        # Wait a bit longer for SAP to potentially initialize scripting fully
        time.sleep(7)
        try:
             self._ensure_com_objects() # Ensure COM objects are ready after launch
             log.info("SAP launched and scripting engine accessed successfully.")
        except exceptions.SapGuiComException as e:
             log.error(f"Error ensuring COM objects after launch: {e}")
             # Decide if this should be fatal or just a warning
             raise # Re-raise as it's likely unusable

        if quit_auto:
            log.info("Registering automatic quit handler.")
            atexit.register(self.quit)


    def _launch(self, working_dir: Path, sid: str, client: str,
                user: str, password: str, maximise: bool, language: str) -> None:
        """Internal helper to launch sapshcut.exe."""
        working_dir = working_dir.resolve()
        sap_executable = working_dir / "sapshcut.exe"

        if not sap_executable.is_file():
            log.error(f"sapshcut.exe not found at: {sap_executable}")
            raise FileNotFoundError(f"sapshcut.exe not found at: {sap_executable}")

        maximise_sap = "-max" if maximise else ""
        language_sap = f"-language={language.upper()}" if language else "" # Use uppercase for language

        # Construct command line arguments carefully
        # Base command includes credentials
        cmd_parts = [
            f"-system={sid}",
            f"-client={client}",
            f"-user={user}",
            f'-pw="{password}"' # Enclose password in quotes if it might contain spaces/special chars
        ]
        # Add optional parts only if they have values
        if maximise_sap: cmd_parts.append(maximise_sap)
        if language_sap: cmd_parts.append(language_sap)

        command_str = " ".join(cmd_parts)
        full_cmd = [str(sap_executable)] + cmd_parts # Pass as list to Popen

        log.debug(f"Executing SAP launch command: {full_cmd}") # Be careful logging passwords

        tryouts = 2
        while tryouts > 0:
            try:
                # Use Popen for better process handling (though not fully utilized here)
                process = Popen(full_cmd)
                log.info(f"Launched sapshcut.exe (PID likely {process.pid}). Waiting for window...")
                # Wait for a window with the *default* title to appear.
                # This might need adjustment if the initial screen title is different.
                utils.wait_for_window_title(self.default_window_title, timeout_loops=20) # Increase timeout
                log.info("SAP window detected.")
                break # Success

            except exceptions.WindowDidNotAppearException:
                log.warning(f"SAP window did not appear after launch attempt (Try {3-tryouts}/2).")
                tryouts -= 1
                # Attempt to kill residual processes before retrying
                utils.kill_process("sapshcut.exe")
                utils.kill_process("saplogon.exe")
                time.sleep(3) # Wait after killing

            except Exception as launch_err:
                 log.exception(f"Unexpected error during Popen or wait_for_window_title: {launch_err}")
                 raise launch_err # Re-raise unexpected errors

        else: # Loop finished without break (all tryouts failed)
            log.error("Failed to launch SAP - Window did not appear after multiple attempts.")
            raise exceptions.WindowDidNotAppearException(
                "Failed to launch SAP - Window did not appear after multiple attempts."
            )


    def quit(self) -> None:
        """
        Attempts to gracefully log off the first session (0, 0) and then kills saplogon.exe.
        Use with caution, especially the process killing part.
        """
        log.info("Attempting graceful SAP quit via System -> Log Off...")
        try:
            # Ensure COM objects are available before trying to quit
            self._ensure_com_objects()
            main_window = self.attach_window(0, 0) # Attach to the main session
            # System -> Log Off menu path (may vary slightly by version/language)
            # Using select_menu_item_by_name might be more robust if available/implemented
            menu_item_logoff = "wnd[0]/mbar/menu[0]/menu[11]" # Common path for System->Log Off
            log.debug(f"Attempting to select menu item: {menu_item_logoff}")
            main_window.select(menu_item_logoff)
            time.sleep(1) # Wait for potential confirmation popup

            # Check for common confirmation popup (wnd[1])
            popup_yes_button = "wnd[1]/usr/btnSPOP-OPTION1" # 'Yes' button ID
            if main_window.exists(popup_yes_button):
                log.info("Logoff confirmation popup detected. Clicking 'Yes'.")
                main_window.press(popup_yes_button)
                time.sleep(2) # Wait for logoff process
            else:
                log.info("No standard logoff confirmation popup detected.")

        except exceptions.AttachException:
            log.warning("Could not attach to session (0, 0) during quit. SAP might already be closed.")
        except exceptions.ActionException as e:
            log.warning(f"Error performing logoff action during quit: {e}. Will proceed to kill process.")
        except exceptions.SapGuiComException as e:
             log.warning(f"COM error during graceful quit attempt: {e}. Will proceed to kill process.")
        except Exception as e:
            log.exception(f"Unexpected error during graceful SAP quit: {e}. Will proceed to kill process.")
        finally:
            # Always attempt to kill the process as a fallback or final step
            log.warning("Killing saplogon.exe process as final step or fallback.")
            utils.kill_process("saplogon.exe")


    def attach_window(self, connection_index: int, session_index: int) -> window.Window:
        """
        Attaches to a specific SAP GUI session.

        Args:
            connection_index: Zero-based index of the connection.
            session_index: Zero-based index of the session within the connection.

        Returns:
            A Window object representing the attached session.

        Raises:
            AttributeError: If connection_index or session_index are not integers.
            SapGuiComException: If COM objects cannot be initialized.
            AttachException: If the specified connection or session cannot be accessed.
        """
        if not isinstance(connection_index, int):
            raise AttributeError("Connection index must be an integer!")
        if not isinstance(session_index, int):
            raise AttributeError("Session index must be an integer!")

        log.info(f"Attempting to attach to Connection {connection_index}, Session {session_index}...")
        self._ensure_com_objects() # Ensure _application is ready

        try:
            # Get the connection handle (COM object)
            connection_handle = self._application.Children(connection_index)
            log.debug(f"Accessed Connection handle for index {connection_index}.")
        except Exception as e:
            log.error(f"Failed to access Connection handle for index {connection_index}: {e}")
            raise exceptions.AttachException(f"Could not attach connection {connection_index}: {e}")

        try:
            # Get the session handle (COM object) from the connection
            session_handle = connection_handle.Children(session_index)
            log.debug(f"Accessed Session handle for index {session_index} on Connection {connection_index}.")
        except Exception as e:
            log.error(f"Failed to access Session handle for index {session_index} on Connection {connection_index}: {e}")
            # Check if connection handle is valid before raising session error
            if not connection_handle: # Should not happen if first try succeeded, but check
                 raise exceptions.AttachException(f"Cannot attach session {session_index}, parent connection {connection_index} is invalid.")
            raise exceptions.AttachException(f"Could not attach session {session_index} on Connection {connection_index}: {e}")

        # Create and return the Window object --- THIS IS THE CORRECTED PART ---
        win = window.Window(
            application=self._application,
            connection=connection_index,        # Use 'connection' keyword
            connection_handle=connection_handle,
            session=session_index,              # Use 'session' keyword
            session_handle=session_handle,
        )
        # --- END OF CORRECTION ---
        log.info(f"Successfully attached to Connection {connection_index}, Session {session_index}.")
        return win


    def open_new_window(self, window_to_handle_opening: window.Window) -> None:
        """
        Opens a new SAP GUI session (window) using an existing session.

        Args:
            window_to_handle_opening: An existing, idle Window object that will be used
                                      to trigger the creation of the new session.

        Raises:
            ActionException: If the command to create a session fails.
            WindowDidNotAppearException: If a new window with the default title
                                         doesn't appear within the timeout.
        """
        if not isinstance(window_to_handle_opening, window.Window):
            raise TypeError("window_to_handle_opening must be a valid Window object.")

        log.info(f"Requesting new session via Connection {window_to_handle_opening.connection}, Session {window_to_handle_opening.session}...")
        try:
            # Use the session handle of the provided window object
            window_to_handle_opening.session_handle.createSession()
            log.info("createSession command sent. Waiting for new window to appear...")
            # Wait for *any* window with the default title. This might be fragile if
            # multiple windows are opening simultaneously or have different titles.
            utils.wait_for_window_title(self.default_window_title, timeout_loops=15) # Increased timeout
            log.info("Detected a window with the default title after createSession call.")
        except exceptions.WindowDidNotAppearException:
            log.error("New SAP window did not appear after createSession call within timeout.")
            raise # Re-raise the specific exception
        except AttributeError:
             # If session_handle doesn't have createSession (shouldn't happen with valid Window obj)
             log.error("Provided Window object's session_handle does not have createSession method.")
             raise exceptions.ActionException("Cannot create session: Invalid session handle in provided Window object.")
        except Exception as e:
            # Catch other potential COM errors during createSession
            log.exception(f"Error sending createSession command: {e}")
            raise exceptions.ActionException(f"Error sending createSession command: {e}")


    @staticmethod
    def start_saplogon(saplogon_path: Union[str, Path] = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")) -> bool:
        """
        Starts the saplogon.exe process if SAP GUI scripting object is not found.

        Args:
            saplogon_path: The full path to saplogon.exe.

        Returns:
            True if the process was potentially started or already running,
            False if saplogon.exe was not found or failed to start.

        Note:
            This is a basic check. It doesn't guarantee SAP Logon is fully functional.
        """
        saplogon_path = Path(saplogon_path).resolve()
        log.info(f"Checking if SAP Logon needs starting (Path: {saplogon_path})...")

        # Check if scripting object is already available
        try:
            win32com.client.GetObject("SAPGUI")
            log.info("SAP GUI Scripting object already available. Assuming SAP Logon is running.")
            return True # Already running or accessible
        except Exception:
            log.info("SAP GUI Scripting object not found. Attempting to start SAP Logon...")
            pass # Object not found, proceed to start

        # Check if executable exists
        if not saplogon_path.is_file():
            log.error(f"saplogon.exe not found at the specified path: {saplogon_path}")
            return False

        # Attempt to start the process
        try:
            log.info(f"Starting SAP Logon process: {saplogon_path}")
            Popen([str(saplogon_path)])
            log.info("SAP Logon process launched. Waiting a few seconds for initialization...")
            time.sleep(7) # Increased wait time for SAP Logon to initialize

            # Verify again if the object is now available
            try:
                win32com.client.GetObject("SAPGUI")
                log.info("SAP GUI Scripting object now available after starting SAP Logon.")
                return True
            except Exception:
                log.error("Failed to get SAP GUI Scripting object even after attempting to start SAP Logon.")
                return False # Failed to become available
        except Exception as e:
            log.exception(f"Error occurred while trying to start saplogon.exe: {e}")
            return False

    # --- Methods for Getting Connection/Session Info ---

    def get_connection_count(self) -> int:
        """Gets the number of currently open SAP connections (systems)."""
        self._ensure_com_objects()
        try:
            count = self._application.Children.Count
            log.debug(f"Found {count} connection(s).")
            return count
        except Exception as e:
            log.error(f"Could not get connection count: {e}")
            # Consider raising or returning 0/None based on desired behavior
            raise exceptions.SapGuiComException(f"Could not get connection count: {e}")

    def get_connection_info(self, connection_index: int) -> Optional[Dict[str, Any]]:
        """Gets basic information about a specific SAP connection."""
        self._ensure_com_objects()
        try:
            connection_handle = self._application.Children(connection_index)
            info = {
                "Description": getattr(connection_handle, "Description", "N/A"),
                # Add other useful top-level connection properties if needed
                # e.g., IsBusy, IsConnected? (availability may vary)
            }
            log.debug(f"Retrieved info for Connection {connection_index}: {info}")
            return info
        except Exception as e:
            # Index out of bounds or other COM error
            log.warning(f"Could not get info for Connection {connection_index}: {e}")
            return None

    def get_active_session_indices(self, connection_index: int) -> List[int]:
        """
        Gets a list of active (accessible) session indices for a given connection.
        This iterates through potential indices as the COM collection might be sparse.
        """
        self._ensure_com_objects()
        indices = []
        log.debug(f"Scanning active session indices for Connection {connection_index}...")
        try:
            connection_handle = self._application.Children(connection_index)
            # We need to probe indices as Children.Count might not reflect closed sessions accurately.
            # Probe a reasonable range (e.g., 0 to 10, as max is usually 6).
            max_sessions_to_probe = 10
            for i in range(max_sessions_to_probe):
                 try:
                     # Attempt to access the session object by index
                     session_handle = connection_handle.Children(i)
                     # Perform a lightweight check to see if it's responsive
                     if session_handle and hasattr(session_handle, 'Info'):
                         # Accessing a property can confirm it's somewhat alive
                          _ = session_handle.Info.SystemSessionId
                          indices.append(i)
                          # log.debug(f"Found active session at index {i} on Connection {connection_index}.")
                 except Exception:
                     # This index is not available or session is closed/invalid
                     # log.debug(f"No active session found at index {i} on Connection {connection_index}.")
                     continue
            log.info(f"Found active session indices for Connection {connection_index}: {indices}")
            return indices
        except Exception as e:
             # Error accessing the connection itself
             log.error(f"Could not get active session indices for Connection {connection_index} due to connection error: {e}")
             return [] # Return empty list on error


    def get_session_info(self, connection_index: int, session_index: int) -> Optional[Dict[str, Any]]:
        """
        Gets detailed information about a specific SAP session.

        Returns dict with info or None if session not found/accessible/logged in.
        Adds 'index' key to the returned dictionary.
        """
        self._ensure_com_objects()
        log.debug(f"Getting detailed info for Conn {connection_index}, Session {session_index}...")
        try:
            connection_handle = self._application.Children(connection_index)
            session_handle = connection_handle.Children(session_index)
            info = session_handle.Info # Get the GuiSessionInfo object

            # Extract properties safely using getattr
            session_info = {
                "index": session_index, # Add the index for reference
                "SystemName": getattr(info, "SystemName", None),
                "Client": getattr(info, "Client", None),
                "User": getattr(info, "User", None),
                "Language": getattr(info, "Language", None),
                "Transaction": getattr(info, "Transaction", None),
                "WindowHandle": getattr(info, "WindowHandle", None),
                "ApplicationServer": getattr(info, "ApplicationServer", None),
                "SystemNumber": getattr(info, "SystemNumber", None),
                "SystemSessionId": getattr(info, "SystemSessionId", None),
                # Add other info properties as needed (e.g., IsLowSpeedConnection)
            }

            # Basic check: If user is empty, it's likely not fully logged in
            if not session_info["User"]:
                 log.warning(f"Session Conn:{connection_index}, Sess:{session_index} appears not fully logged in (User is empty).")
                 # Decide whether to return None or raise AuthorizationError
                 # Returning None might be safer for scanning purposes.
                 # raise exceptions.AuthorizationError(...)
                 return None

            log.debug(f"Successfully retrieved info for Conn {connection_index}, Session {session_index}.")
            return session_info

        except Exception as e:
            # Index out of bounds, session closed, COM error, etc.
            log.warning(f"Could not get info for Conn {connection_index}, Session {session_index}: {e}")
            return None


    def get_all_connections_info(self) -> List[Dict[str, Any]]:
        """
        Scans and returns structured information about all active connections
        and their sessions.
        """
        log.info("Scanning all active SAP GUI connections and sessions...")
        all_info = []
        self._ensure_com_objects() # Ensure we can access _application

        try:
            connection_count = self.get_connection_count()
            for conn_idx in range(connection_count):
                conn_details = self.get_connection_info(conn_idx)
                if conn_details:
                    connection_data = {
                        "index": conn_idx,
                        "description": conn_details.get("Description", "N/A"),
                        "sessions": []
                    }
                    # Get active sessions for this connection
                    active_session_indices = self.get_active_session_indices(conn_idx)
                    for sess_idx in active_session_indices:
                        session_details = self.get_session_info(conn_idx, sess_idx)
                        if session_details: # Add only if info could be retrieved (e.g., logged in)
                            connection_data["sessions"].append(session_details)
                    all_info.append(connection_data)
                else:
                     log.warning(f"Skipping connection index {conn_idx} as basic info couldn't be retrieved.")

        except Exception as e:
            log.exception(f"Error occurred while scanning all connections: {e}")
            # Return potentially partial data or empty list? Returning partial for now.

        log.info(f"Scan complete. Found info for {len(all_info)} connection(s).")
        return all_info


    def find_session_by_sid_user(self, sid: str, user: str) -> Optional[window.Window]:
        """Finds the first active session matching the given SID and User."""
        log.info(f"Searching for session with SID='{sid}', User='{user}'...")
        self._ensure_com_objects()
        sid_upper = sid.upper()
        user_upper = user.upper()

        try:
            for conn_idx in range(self.get_connection_count()):
                active_sessions = self.get_active_session_indices(conn_idx)
                for sess_idx in active_sessions:
                    session_info = self.get_session_info(conn_idx, sess_idx)
                    if session_info:
                        current_sid = session_info.get("SystemName", "")
                        current_user = session_info.get("User", "")
                        if current_sid and current_user and \
                           current_sid.upper() == sid_upper and \
                           current_user.upper() == user_upper:
                            log.info(f"Found matching session: Conn {conn_idx}, Session {sess_idx}.")
                            # Attach and return the window object
                            return self.attach_window(conn_idx, sess_idx)
        except Exception as e:
            log.exception(f"Error searching for session by SID/User: {e}")

        log.info(f"Session with SID='{sid}', User='{user}' not found.")
        return None


    # --- Screenshot Methods ---

    def enable_screenshots_on_error(self) -> None:
        """Enables automatic screenshot capture on exceptions handled by handle_exception_with_screenshot."""
        if ImageGrab is None:
            log.warning("Pillow library not found. Screenshots on error cannot be enabled. Install with: pip install Pillow")
            self._screenshots_on_error_enabled = False
        else:
            self._screenshots_on_error_enabled = True
            log.info("Screenshots on error enabled.")

    def disable_screenshots_on_error(self) -> None:
        """Disables automatic screenshot capture."""
        self._screenshots_on_error_enabled = False
        log.info("Screenshots on error disabled.")

    def set_screenshot_directory(self, directory: Union[str, Path]) -> None:
        """Sets the directory where screenshots will be saved."""
        try:
            path_obj = Path(directory).resolve()
            # Attempt to create directory if it doesn't exist
            path_obj.mkdir(parents=True, exist_ok=True)
            self._screenshot_directory = path_obj
            log.info(f"Screenshot directory set to: {self._screenshot_directory}")
        except Exception as e:
            log.error(f"Error setting screenshot directory '{directory}': {e}. Screenshots will be saved to current directory.", exc_info=True)
            self._screenshot_directory = None

    def _take_screenshot(self, filename_prefix: str = "pysap_error") -> Optional[Path]:
        """Internal method to capture and save a screenshot."""
        if ImageGrab is None:
            log.warning("Cannot take screenshot: Pillow library not installed.")
            return None
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            filename = f"{filename_prefix}_{timestamp}.png"
            save_dir = self._screenshot_directory if self._screenshot_directory else Path(".")
            filepath = save_dir.joinpath(filename)
            log.info(f"Attempting to save screenshot to: {filepath}")
            screenshot = ImageGrab.grab()
            screenshot.save(filepath)
            log.info(f"Screenshot saved successfully: {filepath}")
            return filepath
        except Exception as e:
            log.exception(f"Error taking or saving screenshot: {e}")
            return None

    def handle_exception_with_screenshot(self,
                                         exception: Exception, # Keep argument name generic
                                         filename_prefix: str = "pysap_error") -> None:
        """
        Handles an exception by logging it and taking a screenshot if enabled.
        Typically called from an except block.
        """
        # Log the exception regardless of screenshot setting
        log.error(f"Handling exception: {type(exception).__name__}: {exception}", exc_info=True) # Add stack trace to log

        if self._screenshots_on_error_enabled:
            log.info("Taking screenshot because screenshots_on_error is enabled.")
            self._take_screenshot(filename_prefix=filename_prefix)
        else:
            log.info("Screenshot on error is disabled, skipping capture.")

    # --- HistoryEnabled Methods ---

    def disable_history(self) -> bool:
        """Disables the SAP GUI input history (sets HistoryEnabled=False)."""
        log.info("Attempting to disable SAP GUI input history...")
        try:
            self._ensure_com_objects() # Ensure _application exists
            self._application.HistoryEnabled = False
            log.info("SAP GUI input history disabled (HistoryEnabled=False).")
            return True
        except exceptions.SapGuiComException as e:
             log.error(f"Failed to disable history (COM object issue): {e}")
             return False
        except Exception as e:
            # Catch potential errors setting the property (e.g., read-only?)
            log.error(f"Error setting HistoryEnabled to False: {e}", exc_info=True)
            return False

    def enable_history(self) -> bool:
        """Enables the SAP GUI input history (sets HistoryEnabled=True)."""
        log.info("Attempting to enable SAP GUI input history...")
        try:
            self._ensure_com_objects()
            self._application.HistoryEnabled = True
            log.info("SAP GUI input history enabled (HistoryEnabled=True).")
            return True
        except exceptions.SapGuiComException as e:
             log.error(f"Failed to enable history (COM object issue): {e}")
             return False
        except Exception as e:
            log.error(f"Error setting HistoryEnabled to True: {e}", exc_info=True)
            return False