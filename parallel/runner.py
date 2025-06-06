# File: pysapscript/parallel/runner.py
"""Worker process used by run_parallel to drive SAP sessions."""

import multiprocessing
import time
import os
import tempfile
import logging
import traceback # For full error stack trace
from typing import Callable, List, Optional, Any, Dict

# Import win32com for diagnostics, handle potential import error
try:
    import win32com.client
except ImportError:
    win32com = None # Set to None if library is not available

# Adjust imports based on your project structure
from ..sapscriptwizard import Sapscript
from ..window import Window
from ..types_.exceptions import AttachException, SapGuiComException
# from ..utils import utils # Keep commented if kill_process is not used

# Set up logger for this module
log = logging.getLogger(__name__)

# Define the type for the worker function
WorkerFunctionType = Callable[[Window, List[Any]], None]

class SapParallelRunner:
    """
    Manages the parallel execution of a worker function across multiple SAP GUI sessions.
    Supports modes for using existing sessions or creating new ones.
    """
    def __init__(self,
                 num_processes: int, # This is now the *effective* number of processes
                 worker_function: WorkerFunctionType,
                 # --- New parameters from api.py ---
                 target_connection_index: int,
                 mode: str, # 'new' or 'existing'
                 target_session_indices: Optional[List[int]], # Indices if mode='existing'
                 # --- Data handling ---
                 input_data_file: Optional[str] = None,
                 input_data_list: Optional[List[Any]] = None,
                 # --- Timing parameters ---
                 session_attach_interval: int = 5,  # Delay between attach attempts in worker
                 popup_check_delay: int = 10,       # Delay after opening each new window
                 wait_before_launch: int = 15):     # Delay before starting any worker processes
        """
        Initializes the SapParallelRunner.

        Args:
            num_processes: The effective number of parallel processes to run.
            worker_function: The function to execute in each session.
            target_connection_index: The index of the SAP connection to use.
            mode: 'new' to create new sessions, 'existing' to use specified ones.
            target_session_indices: List of session indices to use if mode='existing'.
            input_data_file: Path to data file.
            input_data_list: List of data items.
            session_attach_interval: Delay (s) between worker attach attempts.
            popup_check_delay: Delay (s) after opening a new window (if mode='new').
            wait_before_launch: Delay (s) before starting worker processes (if mode='new').
        """
        # Basic validation
        if mode not in ['new', 'existing']:
            raise ValueError("Invalid mode specified. Must be 'new' or 'existing'.")
        if mode == 'existing' and not target_session_indices:
            raise ValueError("target_session_indices must be provided when mode is 'existing'.")
        if mode == 'existing' and len(target_session_indices) != num_processes:
             log.warning(f"Runner created for 'existing' mode with num_processes={num_processes} but {len(target_session_indices)} indices provided. Effective processes will be {len(target_session_indices)}.")
             num_processes = len(target_session_indices) # Adjust effective number
        if num_processes <= 0:
             # This case should be handled by api.py, but check defensively
             raise ValueError("num_processes must be greater than 0.")
        # Data source validation
        if not (input_data_file or input_data_list is not None):
            raise ValueError("Either input_data_file or input_data_list must be provided.")
        if input_data_file and input_data_list is not None:
            raise ValueError("Cannot use both input_data_file and input_data_list.")

        # Store parameters
        self.effective_num_processes = num_processes
        self.worker_function = worker_function
        self.target_connection_index = target_connection_index
        self.mode = mode
        self.target_session_indices = target_session_indices # Only used if mode='existing'
        self.input_data_file = input_data_file
        self.input_data_list = input_data_list
        self.session_attach_interval = session_attach_interval
        self.popup_check_delay = popup_check_delay
        self.wait_before_launch = wait_before_launch

        # Internal state
        self._sap_main: Optional[Sapscript] = None # Instance for the main process (session opening)
        self._all_data: List[Any] = []
        self._temp_files: List[Optional[str]] = [] # Temp file paths for data chunks
        self._processes: List[multiprocessing.Process] = []
        # --- Stores the actual session indices workers will connect to ---
        self._actual_session_indices_to_use: List[int] = []

    def run(self):
        """Executes the parallel processing workflow."""
        log.info(f"Starting parallel run: Conn={self.target_connection_index}, Mode='{self.mode}', Processes={self.effective_num_processes}")
        try:
            self._read_data()

            # --- Session Handling based on Mode ---
            if self.mode == 'new':
                log.info("Mode is 'new', attempting to open sessions...")
                self._open_sessions() # This will now populate _actual_session_indices_to_use
                # Adjust effective_num_processes if fewer sessions were opened than requested due to limit
                if len(self._actual_session_indices_to_use) != self.effective_num_processes:
                     log.warning(f"Number of actually available new sessions ({len(self._actual_session_indices_to_use)}) differs from requested ({self.effective_num_processes}). Adjusting process count.")
                     self.effective_num_processes = len(self._actual_session_indices_to_use)
                     if self.effective_num_processes == 0:
                          log.error("No new sessions could be opened or identified. Cannot proceed.")
                          return # Exit run method if no sessions available

            elif self.mode == 'existing':
                log.info(f"Mode is 'existing', using pre-defined sessions: {self.target_session_indices}")
                if not self.target_session_indices: # Should be caught in init, but double check
                    log.error("Mode is 'existing' but no target_session_indices provided.")
                    return
                self._actual_session_indices_to_use = self.target_session_indices
                # Ensure main Sapscript object is created for potential future use, but don't open windows
                try:
                    self._sap_main = Sapscript()
                except Exception as e:
                    log.warning(f"Could not create main Sapscript instance in 'existing' mode (may not be needed): {e}")


            # --- Data Splitting and File Prep (uses effective_num_processes) ---
            chunks = self._split_list(self._all_data, self.effective_num_processes)
            self._prepare_data_files(chunks)

            # --- Delay Before Launching Workers (especially important for 'new' mode) ---
            if self.mode == 'new':
                log.info(f"Waiting {self.wait_before_launch} seconds before launching workers...")
                time.sleep(self.wait_before_launch)
            else: # 'existing' mode, less critical delay, but a small one might not hurt
                 time.sleep(1)

            # --- Launch and Wait ---
            self._launch_workers()
            self._wait_for_workers()

        except Exception as e:
             log.exception(f"An error occurred during the parallel run setup or execution: {e}")
        finally:
            log.info("Parallel run workflow finished. Cleaning up temporary files...")
            self._cleanup_temp_files()
            # No automatic SAP closing
            log.info("Cleanup complete. SAP sessions remain open.")

    def _read_data(self):
        """Reads input data from the specified source."""
        # (No changes needed from previous version - keeping for completeness)
        if self.input_data_list is not None:
            self._all_data = list(self.input_data_list)
            log.info(f"Read {len(self._all_data)} items from input_data_list.")
        elif self.input_data_file:
            log.info(f"Reading data from file: {self.input_data_file}")
            try:
                with open(self.input_data_file, 'r', encoding='utf-8') as f:
                    self._all_data = [line.strip() for line in f]
                log.info(f"Read {len(self._all_data)} lines from {self.input_data_file}.")
            except FileNotFoundError:
                log.error(f"Input data file not found: {self.input_data_file}")
                raise
            except Exception as e:
                log.error(f"Error reading input data file {self.input_data_file}: {e}")
                raise
        else:
             self._all_data = []
             log.warning("No input data source specified.")

    def _split_list(self, lst: List[Any], n: int) -> List[List[Any]]:
        """Splits a list into n roughly equal chunks."""
        # (No changes needed from previous version)
        if n <= 0: return []
        if not lst: return [[] for _ in range(n)]
        k, m = divmod(len(lst), n)
        chunks = [lst[i * k + min(i, m):(i + 1) * k + min(i + 1, m)] for i in range(n)]
        while len(chunks) < n: chunks.append([])
        log.info(f"Split data into {len(chunks)} chunks. Sizes: {[len(c) for c in chunks]}")
        return chunks

    def _prepare_data_files(self, data_chunks: List[List[Any]]):
        """Creates temporary files for each non-empty data chunk."""
        # (No changes needed from previous version)
        self._temp_files = []
        log.info("Preparing temporary data files for workers...")
        for i, chunk in enumerate(data_chunks):
            if chunk:
                try:
                    fd, path = tempfile.mkstemp(suffix=f"_sapdata_p{i}.txt", prefix="pysap_", text=True)
                    with os.fdopen(fd, 'w', encoding='utf-8') as f:
                        for item in chunk: f.write(str(item) + '\n')
                    self._temp_files.append(path)
                except Exception as e:
                    log.error(f"Failed to create temp file for chunk {i}: {e}")
                    self._temp_files.append(None)
                    raise
            else:
                self._temp_files.append(None)
        log.info(f"Prepared {len(self._temp_files)} temp file paths (None indicates empty/no chunk).")

    def _open_sessions(self):
        """
        Opens new SAP GUI sessions on the target connection (only if mode='new').
        Respects the 6-session limit and identifies the indices of newly opened sessions.
        Populates `self._actual_session_indices_to_use`.
        """
        if self.mode != 'new':
            log.error("_open_sessions called unexpectedly when mode is not 'new'.")
            return # Should not happen if called correctly from run()

        log.info(f"Attempting to open {self.effective_num_processes} new session(s) on Connection {self.target_connection_index}...")
        initial_indices: List[int] = []
        opened_indices: List[int] = []

        try:
            self._sap_main = Sapscript()
            # --- Get initial state ---
            initial_indices = self._sap_main.get_active_session_indices(self.target_connection_index)
            log.info(f"Connection {self.target_connection_index}: Initial active session indices: {initial_indices}")
            current_session_count = len(initial_indices)
            if not initial_indices:
                 # We need at least one session to originate the 'open new' command
                 raise AttachException(f"Cannot open new sessions: No existing sessions found on Connection {self.target_connection_index} to act as base.")

            # --- Determine base session and attach ---
            # Use the lowest index as the base for opening new ones
            base_session_index = min(initial_indices)
            log.info(f"Using session {base_session_index} on Connection {self.target_connection_index} as base to open new sessions.")
            base_window = self._sap_main.attach_window(self.target_connection_index, base_session_index)

            # --- Loop to open required number of sessions ---
            sessions_opened_count = 0
            for i in range(self.effective_num_processes):
                current_session_count = len(self._sap_main.get_active_session_indices(self.target_connection_index))
                log.info(f"Checking session limit: Currently {current_session_count} sessions on Conn {self.target_connection_index}.")
                if current_session_count >= 6:
                    log.warning(f"Reached 6-session limit on Connection {self.target_connection_index}. Cannot open more sessions.")
                    break # Stop opening

                log.info(f"Opening new session #{i+1} (attempting)...")
                try:
                    self._sap_main.open_new_window(base_window)
                    log.info(f"Open command sent for session #{i+1}. Waiting {self.popup_check_delay}s...")
                    time.sleep(self.popup_check_delay)
                    sessions_opened_count += 1
                except Exception as e:
                    log.error(f"Failed to send open command for new session #{i+1}: {e}")
                    # Decide if we should stop or continue trying others
                    break # Stop opening on error

            # --- Identify newly opened sessions ---
            log.info(f"Finished attempting to open sessions. Opened count: {sessions_opened_count}.")
            final_indices = self._sap_main.get_active_session_indices(self.target_connection_index)
            log.info(f"Connection {self.target_connection_index}: Final active session indices: {final_indices}")

            # Calculate the difference
            newly_opened_indices = sorted(list(set(final_indices) - set(initial_indices)))
            log.info(f"Identified newly opened session indices: {newly_opened_indices}")

            # --- Determine which indices to use ---
            # We need exactly 'effective_num_processes' indices. Use the newly opened ones.
            # If we opened fewer than requested due to limit/error, use only those opened.
            self._actual_session_indices_to_use = newly_opened_indices[:self.effective_num_processes]

            if len(self._actual_session_indices_to_use) < self.effective_num_processes:
                 log.warning(f"Could only identify {len(self._actual_session_indices_to_use)} new sessions to use, less than the target of {self.effective_num_processes}.")
                 # The effective number will be adjusted in run() before launch.

            log.info(f"Actual session indices targeted for workers: {self._actual_session_indices_to_use}")

        except AttachException as ae:
             log.error(f"Failed to attach to base session {base_session_index} on Connection {self.target_connection_index}: {ae}")
             raise
        except Exception as e:
             log.error(f"An error occurred during session opening process: {e}")
             raise

    def _launch_workers(self):
        """Launches the worker processes, targeting the correct sessions."""
        if not self._actual_session_indices_to_use:
             log.error("Cannot launch workers: No actual session indices determined to use.")
             return

        if len(self._actual_session_indices_to_use) != self.effective_num_processes:
            log.warning(f"Launching {len(self._actual_session_indices_to_use)} workers, as it differs from the initial effective count {self.effective_num_processes}.")
            self.effective_num_processes = len(self._actual_session_indices_to_use) # Final adjustment

        log.info(f"Launching {self.effective_num_processes} worker processes...")
        self._processes = []
        for i in range(self.effective_num_processes):
            # Determine the specific session index for this worker
            session_index_for_worker = self._actual_session_indices_to_use[i]

            # Get the temp file path for this worker's data chunk
            file_path = self._temp_files[i] if i < len(self._temp_files) else None

            log.info(f"Preparing worker process #{i} for Conn {self.target_connection_index}, Session {session_index_for_worker}...")
            # --- Pass target_connection_index AND specific session_index_for_worker ---
            args = (
                self.worker_function,
                file_path,
                self.target_connection_index, # Pass connection index
                session_index_for_worker     # Pass specific session index
            )
            process_name = f"SapWorker-{i}"
            try:
                p = multiprocessing.Process(
                    target=self._worker_process_target,
                    args=args,
                    name=process_name
                )
                p.start()
                self._processes.append(p)
                log.info(f"Launched process {process_name} (PID: {p.pid}) targeting Conn {self.target_connection_index}, Session {session_index_for_worker}.")
            except Exception as e:
                log.error(f"Failed to launch process {process_name}: {e}")
                raise # Critical error

    def _wait_for_workers(self):
        """Waits for all launched worker processes to complete."""
        # (No changes needed from previous version)
        log.info(f"Waiting for {len(self._processes)} worker processes to finish...")
        for p in self._processes:
            try:
                p.join()
                log.info(f"Process {p.name} (PID: {p.pid}) finished with exit code {p.exitcode}.")
                if p.exitcode != 0:
                     log.warning(f"Process {p.name} exited with non-zero code: {p.exitcode}. Check logs.")
            except Exception as e:
                 log.error(f"Error waiting for process {p.name}: {e}")
        log.info("All worker processes have finished.")

    def _cleanup_temp_files(self):
        """Removes any temporary data files created."""
        # (No changes needed from previous version)
        log.info("Cleaning up temporary data files...")
        cleaned_count = 0
        for f in self._temp_files:
            if f and os.path.exists(f):
                try:
                    os.remove(f)
                    cleaned_count += 1
                except Exception as e:
                    log.warning(f"Failed to remove temporary file {f}: {e}")
        log.info(f"Finished cleaning up {cleaned_count} temporary files.")

    @staticmethod
    def _worker_process_target(
        worker_function: WorkerFunctionType,
        file_path: Optional[str],
        # --- Added parameters ---
        connection_index: int,
        session_index: int):
        """
        The target function executed by each worker process.
        Connects to the specified SAP session and calls the user's worker function.
        """
        proc_name = multiprocessing.current_process().name
        log.info(f"--- [{proc_name}] Worker target started for Conn {connection_index}, Session {session_index} ---")
        sap: Optional[Sapscript] = None
        data: List[Any] = []
        attach_attempts = 3
        attach_delay = 5

        try:
            # --- Step 1: Read Data (if file path provided) ---
            # (Same logic as previous version)
            if file_path and os.path.exists(file_path):
                log.info(f"[{proc_name}] Reading data from {file_path}...")
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = [line.strip() for line in f if line.strip()]
                    log.info(f"[{proc_name}] Data read successfully ({len(data)} items).")
                except Exception as read_err:
                    log.error(f"[{proc_name}] Failed to read data file {file_path}: {read_err}")
                    return
            elif file_path:
                 log.warning(f"[{proc_name}] Provided data file path does not exist: {file_path}")
                 return
            else:
                log.info(f"[{proc_name}] No input data file provided for this worker.")


            # --- Step 2: Connect to SAP Session with Retries ---
            window: Optional[Window] = None
            for attempt in range(1, attach_attempts + 1):
                log.info(f"[{proc_name}] Attempt {attempt}/{attach_attempts} to connect to Conn {connection_index}, Session {session_index}...")
                try:
                    log.info(f"[{proc_name}] Creating Sapscript instance...")
                    sap = Sapscript()
                    log.info(f"[{proc_name}] Sapscript instance created.")

                    # --- COM Diagnostics (Optional) ---
                    # (Consider removing or simplifying if stable, keeping for now)
                    if win32com:
                        try:
                            log.info(f"[{proc_name}] Performing COM diagnostic check...")
                            sapgui_obj = win32com.client.GetObject("SAPGUI")
                            engine = sapgui_obj.GetScriptingEngine
                            log.info(f"[{proc_name}] ScriptingEngine obtained. Connections: {engine.Children.Count}")
                            if connection_index < engine.Children.Count:
                                conn = engine.Children(connection_index)
                                log.info(f"[{proc_name}] Target Connection {connection_index} Sessions: {conn.Children.Count}")
                                if session_index < conn.Children.Count:
                                    sess = conn.Children(session_index)
                                    log.info(f"[{proc_name}] Target Session {session_index} ID: {getattr(sess, 'ID', 'N/A')}")
                                    is_busy = sess.Busy if hasattr(sess, 'Busy') else 'N/A' # Просто читаем атрибут, не вызываем его
                                    log.info(f"[{proc_name}] Target Session {session_index} Busy: {is_busy}")
                                    if is_busy == True and attempt < attach_attempts:
                                        log.warning(f"[{proc_name}] Session {session_index} is busy. Retrying...")
                                        raise AttachException(f"Session {session_index} reported busy status.")
                                else:
                                     log.warning(f"[{proc_name}] Session index {session_index} out of bounds for Conn {connection_index}.")
                            else:
                                 log.warning(f"[{proc_name}] Connection index {connection_index} out of bounds.")
                        except Exception as diag_err:
                            log.error(f"[{proc_name}] COM diagnostic check failed: {diag_err}")
                    else:
                        log.warning(f"[{proc_name}] win32com not available, skipping COM diagnostics.")
                    # --- End COM Diagnostics ---

                    # --- Use correct connection and session index ---
                    log.info(f"[{proc_name}] Attaching to Conn {connection_index}, Session {session_index}...")
                    window = sap.attach_window(connection_index, session_index)
                    log.info(f"[{proc_name}] Attached successfully to Conn {connection_index}, Session {session_index}. Window object: {window}")
                    break # Success

                except (AttachException, SapGuiComException) as attach_err: # Catch specific attach/COM errors
                    log.warning(f"[{proc_name}] Attach attempt {attempt} failed: {attach_err}")
                    if attempt < attach_attempts:
                        log.info(f"[{proc_name}] Retrying in {attach_delay} seconds...")
                        time.sleep(attach_delay)
                    else:
                        log.error(f"[{proc_name}] All attach attempts failed for Conn {connection_index}, Session {session_index}. Worker cannot continue.")
                        # Don't raise, let it exit after loop
                except Exception as general_err:
                     log.error(f"[{proc_name}] Unexpected error during attach attempt {attempt}: {general_err}")
                     log.error(traceback.format_exc())
                     if attempt < attach_attempts:
                          log.info(f"[{proc_name}] Retrying in {attach_delay} seconds...")
                          time.sleep(attach_delay)
                     else:
                          log.error(f"[{proc_name}] Unexpected error persisted after {attach_attempts} attempts. Worker cannot continue.")


            # --- Step 3: Check if Attachment Succeeded ---
            if window is None:
                log.error(f"[{proc_name}] Could not attach to Conn {connection_index}, Session {session_index}. Worker exiting.")
                return

            # --- Step 4: Execute the User's Worker Function ---
            log.info(f"[{proc_name}] Calling worker function: {worker_function.__name__}...")
            worker_function(window, data)
            log.info(f"[{proc_name}] Worker function {worker_function.__name__} finished successfully.")

        except Exception as e:
            log.error(f"[{proc_name}] UNHANDLED EXCEPTION in worker target for Conn {connection_index}, Session {session_index}: {e}")
            log.error(f"[{proc_name}] Error Type: {type(e).__name__}")
            log.error(traceback.format_exc())
            if sap:
                try:
                    log.info(f"[{proc_name}] Attempting screenshot for error...")
                    sap.handle_exception_with_screenshot(e, filename_prefix=f"worker_error_{proc_name}_conn{connection_index}_sess{session_index}")
                except Exception as screen_err:
                    log.error(f"[{proc_name}] Failed to take screenshot: {screen_err}")
            else:
                log.warning(f"[{proc_name}] Cannot take screenshot: Sapscript instance not created.")

        finally:
            log.info(f"--- [{proc_name}] Worker target finished for Conn {connection_index}, Session {session_index} ---")
