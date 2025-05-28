import os
import time
import sys
import ctypes
from datetime import datetime
from pathlib import Path
import logging
from itertools import cycle

import pandas as pd

from sap_connection import get_last_session
from other_functions import append_status_to_excel, delete_file, vl10d_process_data, \
    run_excel_file_and_adjust_col_width, copy_df_column_to_clipboard, close_excel_file
from sap_transactions import vl10d_vl10c_load_variant_and_export_data, mb52_load_sap_numbers_and_export_data
from sap_functions import open_one_transaction, zsbe_load_and_export_data, simple_load_variant
from helper_program_functions import filter_out_items_booked_to_0004_spec_cust_requirement_location, fill_storage_location_quantities
from program_paths import ProgramPaths


paths_instance = ProgramPaths()
BASE_PATH = paths_instance.BASE_PATH
ERROR_LOG_PATH = paths_instance.ERROR_LOG_PATH

if __name__ == "__main__":
    username = os.getlogin()
    status_file = (
        f"C:/Users/{username}/OneDrive - Roto Frank DST/General/05_Automatyzacja_narzędzia/100_STATUS"
        f"/01_AUTOMATION_TOOLS_STATUS.xlsx"
    )

    today = datetime.today().strftime("%Y_%m_%d")
    start_time = datetime.now().strftime("%H:%M:%S")

    paths = paths_instance.paths

    program_status = dict()

    # Hide console window
    if sys.platform == "win32":
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        hWnd = kernel32.GetConsoleWindow()
        if hWnd:
            user32.ShowWindow(hWnd, 6)  # 6 = Minimize

    logging.basicConfig(
        filename=ERROR_LOG_PATH,
        level=logging.ERROR,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )

    try:
        sess1, tr1, nu1 = get_last_session(max_num_of_sessions=6)

        # TODO: get data from VL10D file
        vl10d_df = pd.read_excel(paths['vl10d_clean_data'])


        # Handle the information for status file
        # program_status["COHV_CONVERSION_SYSTEM_MESSAGE"] = result_sap_messages

    except Exception as e:
        print(e)
        logging.error("Error occurred", exc_info=True)

    finally:

        # Fill status file
        end_time = datetime.now().strftime("%H:%M:%S")
        program_status["start_time"] = start_time
        program_status["end_time"] = end_time
        # append_status_to_excel(
        #     status_file, program_status, ERROR_LOG_PATH, sheet_name="COHV_CONVERSION"
        # )
