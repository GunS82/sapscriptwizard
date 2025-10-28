# File: pysapscript/parallel/api.py

import logging
import sys
import time
from typing import Callable, List, Optional, Any, Dict

# Import the runner class using a relative path
from .runner import SapParallelRunner
# Import Sapscript and Window classes from the parent package
from ..sapscriptwizard import Sapscript
from ..window import Window
from ..types_.exceptions import AttachException # Import AttachException

# Set up logger for this module
log = logging.getLogger(__name__)

# Define the type for the worker function again for clarity within this file
WorkerFunctionType = Callable[[Window, List[Any]], Any] # Allow worker to return something

# --- Helper function to display session info ---
def _display_connections_and_sessions(connections_info: List[Dict[str, Any]]):
    """Formats and prints available connections and sessions."""
    print("-" * 60)
    print("Available SAP GUI Connections and Sessions:")
    if not connections_info:
        print("  No active connections found.")
        print("-" * 60)
        return

    for conn_data in connections_info:
        conn_idx = conn_data.get('index', 'N/A')
        conn_desc = conn_data.get('description', 'N/A')
        print(f"\nConnection {conn_idx}: {conn_desc}")
        sessions = conn_data.get('sessions', [])
        if sessions:
            for sess_info in sessions:
                sess_idx = sess_info.get('index', 'N/A')
                sess_id = sess_info.get('SystemSessionId', 'N/A')
                user = sess_info.get('User', 'N/A')
                sid = sess_info.get('SystemName', 'N/A')
                client = sess_info.get('Client', 'N/A')
                tcode = sess_info.get('Transaction', 'N/A')
                print(f"  -> Session {sess_idx}: User={user}, SID={sid}, Client={client}, TCode='{tcode}', ID={sess_id}")
        else:
            print("    (No active sessions found for this connection)")
    print("-" * 60)
# --- End Helper Function ---

# --- Helper function to parse session indices ---
def _parse_session_indices(input_str: str, available_indices: List[int]) -> Optional[List[int]]:
    """Parses comma-separated string into list of valid integers."""
    try:
        indices = [int(x.strip()) for x in input_str.split(',') if x.strip()]
        # Validate against available indices
        invalid = [idx for idx in indices if idx not in available_indices]
        if invalid:
            print(f"Error: Invalid or unavailable session indices entered: {invalid}")
            print(f"Available indices are: {available_indices}")
            return None
        if not indices: # Handle empty input after stripping
            print("Error: No session indices entered.")
            return None
        return sorted(list(set(indices))) # Return unique sorted list
    except ValueError:
        print("Error: Invalid input. Please enter comma-separated numbers.")
        return None
# --- End Helper Function ---


def run_parallel(
    enabled: bool,
    num_processes: int,
    worker_function: WorkerFunctionType,
    input_data_list: Optional[List[Any]] = None,
    input_data_file: Optional[str] = None,
    interactive: bool = False, # Default to non-interactive for scripting
    **runner_kwargs: Any
) -> Optional[Any]:
    """
    Runs a SAP GUI worker function either sequentially or in parallel,
    with optional interactive session selection.

    Args:
        enabled: If True, enables parallel execution. If False, runs sequentially.
        num_processes: Desired number of parallel processes (used if enabled=True and mode='new').
        worker_function: The user-defined function (accepts Window, List[Any]).
        input_data_list: List of data items for processing.
        input_data_file: Path to a file containing data items (one per line).
        interactive: If True, prompts the user to select connection, mode (new/existing),
                     and sessions (if mode='existing'). If False, uses defaults
                     (Connection 0, mode 'new').
        **runner_kwargs: Additional keyword arguments for SapParallelRunner
                         (e.g., popup_check_delay, wait_before_launch).

    Returns:
        Optional[Any]: In sequential mode, returns the worker_function's result.
                       In parallel mode, currently returns None.

    Raises:
        ValueError: For invalid input parameters.
        FileNotFoundError: If input_data_file not found (sequential mode).
        AttachException: If attaching to SAP session fails.
        Exception: Any exception from Sapscript or the worker_function.
        SystemExit: If initial scan fails or user interaction is aborted in interactive mode.
    """
    log.info(f"run_parallel called: enabled={enabled}, num_processes={num_processes}, interactive={interactive}")

    # --- Initial Scan and Setup ---
    target_connection_index: int = 0
    mode: str = 'new' # Default mode ('new' or 'existing')
    target_session_indices: Optional[List[int]] = None # Indices selected by user if mode='existing'
    effective_num_processes: int = 1 if not enabled else num_processes # Start with desired number

    log.info("Initializing Sapscript for scanning/execution...")
    try:
        # Create one Sapscript instance used for scanning and potentially sequential run
        sap = Sapscript()
        # --- Method assumed to be added to Sapscript ---
        # This method should return a list of dicts, e.g.:
        # [{'index': 0, 'description': '...', 'sessions': [{'index':0, 'User':'...', ...}, ...]}, ...]
        connections_info = sap.get_all_connections_info()
        # --- End of assumed method ---

        if not connections_info:
            msg = "No active SAP GUI connections found. Please start SAP Logon and log in."
            log.error(msg)
            print(f"Error: {msg}")
            # Exit cleanly instead of raising deep exception if no SAP GUI found at start
            sys.exit(1)

    except Exception as scan_err:
        log.exception(f"Critical error during initial SAP GUI scan: {scan_err}")
        print(f"Critical error during initial SAP GUI scan: {scan_err}")
        print("Ensure SAP GUI Scripting is enabled and SAP Logon is running.")
        sys.exit(1)

    # --- Interactive Mode ---
    if interactive:
        _display_connections_and_sessions(connections_info)
        available_connection_indices = [c['index'] for c in connections_info]

        # 1. Select Connection
        while True:
            try:
                conn_input = input(f"Enter the Connection index to use [{available_connection_indices[0]}]: ")
                if not conn_input.strip():
                    target_connection_index = available_connection_indices[0]
                    print(f"Using default connection index: {target_connection_index}")
                    break
                target_connection_index = int(conn_input)
                if target_connection_index not in available_connection_indices:
                    print(f"Error: Invalid connection index. Available: {available_connection_indices}")
                else:
                    break
            except ValueError:
                print("Error: Please enter a valid number.")
            except EOFError: # Handle Ctrl+D or unexpected end of input
                 print("\nInteraction aborted.")
                 sys.exit(0)

        # Get available sessions on the chosen connection
        selected_conn_info = next((c for c in connections_info if c['index'] == target_connection_index), None)
        available_sessions_on_target = [s['index'] for s in selected_conn_info.get('sessions', [])] if selected_conn_info else []

        # 2. Select Mode/Sessions (only if parallel enabled)
        if enabled:
            print(f"\nAvailable sessions on Connection {target_connection_index}: {available_sessions_on_target}")
            while True:
                try:
                    mode_input = input("Run in [N]ew sessions or use [E]xisting sessions? (N/E) [N]: ").strip().lower()
                    if not mode_input or mode_input == 'n':
                        mode = 'new'
                        print("Mode set to: new sessions")
                        # Check if лимит позволяет создать num_processes новых окон
                        current_session_count = len(available_sessions_on_target)
                        can_open = max(0, 6 - current_session_count)
                        if num_processes > can_open:
                            log.warning(f"Requested {num_processes} new sessions, but only {can_open} can be opened due to the 6-session limit (currently {current_session_count} open).")
                            effective_num_processes = can_open
                            if effective_num_processes == 0:
                                print("Error: Cannot open any new sessions (limit reached or exceeded). Try using existing sessions.")
                                # Loop back or exit? Let's loop back for now.
                                continue
                            else:
                                print(f"Will attempt to open {effective_num_processes} new sessions.")
                        else:
                             effective_num_processes = num_processes
                             print(f"Will attempt to open {effective_num_processes} new sessions.")
                        target_session_indices = None # Ensure this is None for 'new' mode
                        break # Exit mode selection loop
                    elif mode_input == 'e':
                        mode = 'existing'
                        print("Mode set to: existing sessions")
                        if not available_sessions_on_target:
                            print("Error: No existing sessions available on this connection to choose from.")
                            # Loop back to mode selection
                            continue

                        while True: # Loop for getting valid session indices
                            try:
                                indices_input = input(f"Enter comma-separated indices of EXISTING sessions to use (e.g., 0,1): ")
                                parsed_indices = _parse_session_indices(indices_input, available_sessions_on_target)
                                if parsed_indices:
                                    target_session_indices = parsed_indices
                                    effective_num_processes = len(target_session_indices) # Use the count of selected sessions
                                    print(f"Using existing sessions: {target_session_indices} ({effective_num_processes} processes)")
                                    break # Exit indices selection loop
                                # else: _parse_session_indices already printed error, loop again
                            except EOFError:
                                print("\nInteraction aborted.")
                                sys.exit(0)
                        break # Exit mode selection loop
                    else:
                        print("Invalid input. Please enter 'N' or 'E'.")
                except EOFError:
                    print("\nInteraction aborted.")
                    sys.exit(0)
        else: # Sequential mode
             mode = 'sequential' # Mark mode for clarity
             effective_num_processes = 1
             print(f"Sequential mode selected for Connection {target_connection_index}.")
             # In sequential, we typically use the first available session
             target_session_indices = None # Not directly used for session selection here


    # --- Non-Interactive Mode Defaults ---
    else: # not interactive
        target_connection_index = 0 # Default to first connection
        # Verify default connection exists
        if target_connection_index not in [c['index'] for c in connections_info]:
             alt_conn = [c['index'] for c in connections_info]
             if not alt_conn: # Should have been caught earlier, but double check
                  msg = "Non-interactive mode: Default connection 0 not found, and no other connections available."
                  log.error(msg)
                  raise AttachException(msg)
             target_connection_index = alt_conn[0]
             log.warning(f"Non-interactive mode: Default connection 0 not found. Using first available connection: {target_connection_index}")

        if enabled:
            mode = 'new' # Default to creating new sessions
            effective_num_processes = num_processes
            # Check limit for non-interactive 'new' mode
            selected_conn_info = next((c for c in connections_info if c['index'] == target_connection_index), None)
            available_sessions_on_target = [s['index'] for s in selected_conn_info.get('sessions', [])] if selected_conn_info else []
            current_session_count = len(available_sessions_on_target)
            can_open = max(0, 6 - current_session_count)
            if num_processes > can_open:
                 log.warning(f"Non-interactive mode: Requested {num_processes} new sessions, but only {can_open} can be opened due to limit. Reducing to {can_open}.")
                 effective_num_processes = can_open
                 if effective_num_processes == 0:
                      msg = "Non-interactive mode: Cannot open any new sessions (limit reached or exceeded)."
                      log.error(msg)
                      raise ValueError(msg) # Raise error in non-interactive if 0 processes
            target_session_indices = None
        else:
            mode = 'sequential'
            effective_num_processes = 1
            target_session_indices = None
        log.info(f"Non-interactive mode settings: Connection={target_connection_index}, Mode={mode}, Effective Processes={effective_num_processes}")


    # --- Execute Based on Mode ---

    # --- Sequential Execution Path ---
    if not enabled:
        log.info(f"Executing sequentially on Connection {target_connection_index}...")
        window = None
        try:
            # Determine session to use (e.g., first available on target connection)
            selected_conn_info = next((c for c in connections_info if c['index'] == target_connection_index), None)
            available_indices = [s['index'] for s in selected_conn_info.get('sessions', [])] if selected_conn_info else []
            if not available_indices:
                 raise AttachException(f"No active sessions found on Connection {target_connection_index} for sequential execution.")
            session_to_use = min(available_indices) # Use the lowest available index
            log.info(f"Attaching to session {session_to_use} on Connection {target_connection_index}...")

            # Attach using the Sapscript instance created earlier
            window = sap.attach_window(target_connection_index, session_to_use)
            log.info(f"Attached successfully: {window}")

            # Prepare data (same logic as before)
            data: List[Any] = []
            if input_data_file:
                log.info(f"Reading data from file: {input_data_file}")
                try:
                    with open(input_data_file, 'r', encoding='utf-8') as f:
                        data = [line.strip() for line in f]
                    log.info(f"Read {len(data)} lines from file.")
                except FileNotFoundError:
                    log.error(f"Input data file not found: {input_data_file}")
                    raise
                except Exception as e:
                    log.error(f"Error reading input file {input_data_file}: {e}")
                    raise
            elif input_data_list is not None:
                data = list(input_data_list)
                log.info(f"Using provided input_data_list with {len(data)} items.")
            else:
                data = []
                log.info("No input data provided. Worker will receive empty list.")

            # Execute worker function
            log.info(f"Executing worker function '{worker_function.__name__}'...")
            result = worker_function(window, data)
            log.info(f"Worker function '{worker_function.__name__}' finished.")
            return result

        except Exception as seq_err:
            log.exception(f"Error during sequential execution: {seq_err}")
            if sap: # sap instance should exist here
                try:
                    sap.handle_exception_with_screenshot(seq_err, filename_prefix=f"sequential_error_conn{target_connection_index}")
                except Exception as screen_err:
                    log.error(f"Failed to take screenshot: {screen_err}")
            raise

    # --- Parallel Execution Path ---
    else:
        if effective_num_processes == 0:
             # This case should ideally be prevented earlier (e.g., interactive loop or non-interactive error)
             log.error("Parallel execution requested, but effective number of processes is 0. Cannot proceed.")
             # Or raise ValueError("Cannot run parallel execution with zero effective processes.")
             return None # Exit gracefully

        log.info(f"Executing in parallel: Connection={target_connection_index}, Mode={mode}, Processes={effective_num_processes}")
        if target_session_indices:
             log.info(f"Using existing target sessions: {target_session_indices}")

        try:
            # Instantiate the runner with potentially modified parameters
            log.info("Initializing SapParallelRunner...")
            runner = SapParallelRunner(
                num_processes=effective_num_processes, # Use the calculated number
                worker_function=worker_function,
                input_data_file=input_data_file,
                input_data_list=input_data_list,
                # --- Pass new parameters ---
                target_connection_index=target_connection_index,
                mode=mode,
                target_session_indices=target_session_indices,
                # --- Pass other runner kwargs ---
                **runner_kwargs
            )

            log.info("Starting SapParallelRunner run()...")
            runner.run() # Blocks until completion
            log.info("SapParallelRunner run() completed.")
            return None # Parallel run currently doesn't return aggregated results

        except Exception as par_err:
            log.exception(f"Error during parallel execution setup or run: {par_err}")
            # Screenshots are handled within the worker process target
            raise
