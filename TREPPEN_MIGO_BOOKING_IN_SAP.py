import os
import time
import sys
import ctypes
from datetime import datetime
import logging

import pandas as pd

from sap_connection import get_last_session
from VL10D_VL10C_MIGO_BOOKING_IN_SAP import migo_booking
from program_paths import ProgramPaths
from gui_manager import show_message


paths_instance = ProgramPaths()
BASE_PATH = paths_instance.BASE_PATH
ERROR_LOG_PATH = paths_instance.ERROR_LOG_PATH

ups_shipment_file = False


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

        mb02_doc_nums_313 = []  # material documents
        mb02_doc_nums_315 = []  # material documents
        to_numbers = []  # transfer orders numbers

        # TODO: get data from VL10D file
        vl10x_files_paths = ['vl10c_clean_data_treppen']
        # vl10x_files_paths = ["vl10d_clean_data"]
        for vl10x_file_path in vl10x_files_paths:
            migo_booking(paths[vl10x_file_path], sess1, mb02_doc_nums_313)
            migo_booking(paths[vl10x_file_path], sess1, mb02_doc_nums_315, movement_type="315", is_describtion=True)

        # save files
        temp_df = pd.DataFrame({"mb52_mat_docs_nums_313": mb02_doc_nums_313})
        temp_df.to_excel(paths["mb52_mat_docs_nums_313_treppen"])
        temp_df = pd.DataFrame({"mb52_mat_docs_nums_315": mb02_doc_nums_315})
        temp_df.to_excel(paths["mb52_mat_docs_nums_315_treppen"])

        temp_df = pd.DataFrame({"to_numbers": to_numbers})
        temp_df.to_excel(paths["to_numbers"])

        # message if there is belatronic item
        if ups_shipment_file:
            show_message('Uzupełnij plik UPS.\nSzczegóły w standardzie.')

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
