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
from sap_transactions import migo_lt06_lt04_booking_and_transfer
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

        mb52_doc_nums = []  # material documents
        to_numbers = []  # transfer orders numbers

        # TODO: get data from VL10D file
        vl10d_df = pd.read_excel(paths["vl10d_clean_data"])
        migo_lt06_lt04_booking_and_transfer(
            session=sess1,
            mat_nr="333914",
            source_storage_loc="0007",
            doc_header="Header",
            quantity="1",
            plant="2101",
            movement_type="313",
            is_multiple=True,
            is_last=False,
            is_first=True,
            quantities=[1, 2, 3],
            mb52_doc_nums=mb52_doc_nums,
            to_numbers=to_numbers
        )

        migo_lt06_lt04_booking_and_transfer(
            session=sess1,
            mat_nr="333914",
            source_storage_loc="0007",
            doc_header="Header 1",
            quantity="2",
            plant="2101",
            movement_type="313",
            is_multiple=True,
            is_last=False,
            is_first=False,
            quantities=[1, 2, 3],
            mb52_doc_nums=mb52_doc_nums,
            to_numbers=to_numbers
        )

        migo_lt06_lt04_booking_and_transfer(
            session=sess1,
            mat_nr="333914",
            source_storage_loc="0007",
            doc_header="Header 2",
            quantity="3",
            plant="2101",
            movement_type="313",
            is_multiple=True,
            is_last=True,
            is_first=False,
            quantities=[1, 2, 3],
            mb52_doc_nums=mb52_doc_nums,
            to_numbers=to_numbers
        )

        migo_lt06_lt04_booking_and_transfer(
            session=sess1,
            mat_nr="333914",
            source_storage_loc="0007",
            doc_header="Header 4",
            quantity="4",
            plant="2101",
            movement_type="313",
            is_multiple=False,
            is_last=True,
            is_first=True,
            quantities=[4],
            mb52_doc_nums=mb52_doc_nums,
            to_numbers=to_numbers
        )

        temp_df = pd.DataFrame({"mb52_mat_docs_nums": mb52_doc_nums})
        temp_df.to_excel(paths["mb52_mat_docs_nums"])

        temp_df = pd.DataFrame({"to_numbers": to_numbers})
        temp_df.to_excel(paths["to_numbers"])

        if True:
            print('debug')

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
