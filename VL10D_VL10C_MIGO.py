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
from other_functions import append_status_to_excel, delete_file, vl10d_process_data, run_excel_file_and_adjust_col_width
from sap_transactions import vl10d_load_variant_and_export_data


BASE_PATH = Path(
    r"P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\05_VL10D_VL10C_MIGO"
)
ERROR_LOG_PATH = BASE_PATH / "error.log"


if __name__ == "__main__":
    username = os.getlogin()
    status_file = (
        f"C:/Users/{username}/OneDrive - Roto Frank DST/General/05_Automatyzacja_narzędzia/100_STATUS"
        f"/01_AUTOMATION_TOOLS_STATUS.xlsx"
    )

    today = datetime.today().strftime("%Y_%m_%d")
    start_time = datetime.now().strftime("%H:%M:%S")

    file_paths = {
        "vl10d_raw_data": f"historical_data/vl10d_all_items_{today}.xls",
        "vl10d_clean_data": f"vl10d_clean_data_{today}.xlsx",
        "historical_data": "historical_data",
    }

    paths = {key: BASE_PATH / filename for key, filename in file_paths.items()}

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

        # delete file
        delete_file(paths["vl10d_raw_data"])

        # TODO: export vl10d_all_items.xls from VL10D transaction
        vl10d_load_variant_and_export_data(
            session=sess1,
            file_path=str(paths["historical_data"]),
            file_name=paths["vl10d_raw_data"].name,
        )
        # TODO: process the data and save it to excel file vl10d.xlsx
        vl10d_process_data(file_name_raw_data=paths["vl10d_raw_data"], file_name_cleaned_data=paths["vl10d_clean_data"])

        run_excel_file_and_adjust_col_width(paths['vl10d_clean_data'])
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
