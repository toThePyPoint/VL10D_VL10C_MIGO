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
# BASE_PATH = Path(
#     r"P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\05_VL10D_VL10C_MIGO"
# )
# ERROR_LOG_PATH = BASE_PATH / "error.log"
VL10D_VARIANT_NAME = "SHIP_LU_PPS002"
VL10C_VARIANT_NAME = "SHIP_LU_PPS001"

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

    sales_offices_map = {
        "LV01": "Łotwa",
        "DE92": "Roto Treppen",
        "LT01": "Litwa",
        "FR01": "Francja",
        "IT03": "Włochy",
        "EE01": "Estonia",
        "CZ01": "Czechy",
        "PL21": "Polska-Export",
        "RU02": "Rosja(Kaliningrad)",
        "RO01": "Rumunia",
        "HU01": "Węgry",
        "PL01": "Polska",
        "ES01": "Hiszpania",
        "PT01": "Portugalia",
        "UA01": "Ukraina",
        "GB01": "Anglia",
        "SI01": "Słowenia",
        "BY01": "Białoruś",
        "SK01": "Słowacja",
        "HR01": "Chorwacja",
        "PL02": "Polska",
    }
    goods_recepients_map = {
        "100300": "0301",
        "103702": "3701",
        "101203": "1201"
    }

    storage_locations_list = ['0004', '0005', '0007', '0003', '0024', '0010', '0750', '0021']

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

        # delete files
        delete_file(paths["vl10d_raw_data"])
        delete_file(paths["zsbe_data_vl10d"])
        delete_file(paths["mb52_vl10d"])
        delete_file(paths["vl10c_raw_data"])
        delete_file(paths["zsbe_data_vl10c"])
        delete_file(paths["mb52_vl10c"])

        # TODO: export vl10d_all_items.xls from VL10D transaction
        vl10d_vl10c_load_variant_and_export_data(
            session=sess1,
            file_path=str(paths["temp_folder"]),
            file_name=paths["vl10d_raw_data"].name,
            transaction_name="vl10d",
            variant_name=VL10D_VARIANT_NAME
        )
        # TODO: process the data
        vl10d_df = vl10d_process_data(file_name_raw_data=paths["vl10d_raw_data"])

        # TODO: Match MRP controllers
        # copy SAP numbers to clipboard
        copy_df_column_to_clipboard(vl10d_df, "SAP_nr")
        # open ZSBE transaction
        open_one_transaction(session=sess1, transaction_name="ZSBE")
        # zsbe - load data and export it to excel file
        zsbe_load_and_export_data(session=sess1, file_path=str(paths['temp_folder']),
                                  file_name=paths['zsbe_data_vl10d'].name)
        # close Excel file which should be automatically opened
        time.sleep(3)
        close_excel_file(file_name=paths['zsbe_data_vl10d'].name)
        # load zsbe data into data frame
        zsbe_df = pd.read_excel(paths['zsbe_data_vl10d'])
        zsbe_df["Materiał"] = zsbe_df["Materiał"].astype(str)
        vl10d_merged_df = pd.merge(vl10d_df, zsbe_df, left_on="SAP_nr", right_on="Materiał", how="left")

        # TODO: Filter LS1 and LS2 MRP controllers out
        # drop unnecessary columns and rename new column
        columns_to_drop = [
            "Materiał",
        ]
        vl10d_merged_df.drop(columns=columns_to_drop, inplace=True)
        new_col_names = {
            "Kontroler MRP": "mrp_controller",
            "Rodzaj nabycia": "procurement_type"
        }
        vl10d_merged_df.rename(columns=new_col_names, inplace=True)
        # filter the data
        vl10d_merged_df = vl10d_merged_df[~vl10d_merged_df['mrp_controller'].isin(["LS1", "LS2"])]
        # remove rows where 'product_name' starts with 'EBR' or 'EDR' and 'procurement_type' == 'E'
        vl10d_merged_df = vl10d_merged_df[
            ~(
                    vl10d_merged_df['product_name'].str.startswith(('EBR', 'EDR')) &
                    (vl10d_merged_df['procurement_type'] == 'E')
            )
        ]

        # Add header column
        vl10d_merged_df['header'] = vl10d_merged_df['document_number'] + " " + vl10d_merged_df['goods_recepient_number'].apply(lambda x: goods_recepients_map[x])

        # ---------------------------------------------------
        # TODO: Add this to VL10C part
        # TODO: match quantities to storage locations
        # create columns
        vl10d_merged_df['header_suffix'] = ""
        for loc in storage_locations_list:
            vl10d_merged_df[f'loc_{loc}'] = 0
        # vl10d_merged_df['delete'] = False
        # vl10d_merged_df.to_pickle('excel_files/vl10d_merged_df.pkl')
        # copy SAP numbers to clipboard
        copy_df_column_to_clipboard(vl10d_merged_df, "SAP_nr")
        # open MB52 transaction
        open_one_transaction(session=sess1, transaction_name="MB52")
        simple_load_variant(sess1, "MISC_LU_PPS001", True)
        mb52_load_sap_numbers_and_export_data(session=sess1, file_path=str(paths['temp_folder']), file_name=paths['mb52_vl10d'].name)
        # close Excel file which should be automatically opened
        time.sleep(3)
        close_excel_file(file_name=paths['mb52_vl10d'].name)
        # load zsbe data into data frame
        mb52_df = pd.read_excel(paths['mb52_vl10d'], dtype={'Skład': str, 'Materiał': str})
        mb52_df.rename(columns={"Materiał": "SAP_nr", "Nieogranicz.wykorz.": "stock", "Skład": "storage_loc"},
                       inplace=True)
        # mb52_df.to_pickle('excel_files/mb52_df.pkl')
        filter_out_items_booked_to_0004_spec_cust_requirement_location(mb52_df, vl10d_merged_df)
        fill_storage_location_quantities(mb52_df, vl10d_merged_df)

        # TODO: Add this to VL10C part
        # ---------------------------------------------------

        # save vl10d_merged_df to Excel file
        vl10d_merged_df.to_excel(paths['vl10d_clean_data'], index=False)
        # open Excel file
        run_excel_file_and_adjust_col_width(paths['vl10d_clean_data'])

        # ---------------------------------------------
        # ----------------VL10C------------------------
        # ---------------------------------------------
        # TODO: export vl10c_all_items.xls from VL10C transaction
        vl10d_vl10c_load_variant_and_export_data(
            session=sess1,
            file_path=str(paths["temp_folder"]),
            file_name=paths["vl10c_raw_data"].name,
            transaction_name="vl10c",
            variant_name=VL10C_VARIANT_NAME
        )
        # TODO: process the data
        vl10c_df = vl10d_process_data(file_name_raw_data=paths["vl10c_raw_data"])

        # TODO: Match MRP controllers
        # copy SAP numbers to clipboard
        copy_df_column_to_clipboard(vl10c_df, "SAP_nr")
        # open ZSBE transaction
        open_one_transaction(session=sess1, transaction_name="ZSBE")
        # zsbe - load data and export it to excel file
        zsbe_load_and_export_data(session=sess1, file_path=str(paths['temp_folder']),
                                  file_name=paths['zsbe_data_vl10c'].name)
        # close Excel file which should be automatically opened
        time.sleep(3)
        close_excel_file(file_name=paths['zsbe_data_vl10c'].name)
        # load zsbe data into data frame
        zsbe_df = pd.read_excel(paths['zsbe_data_vl10c'])
        zsbe_df["Materiał"] = zsbe_df["Materiał"].astype(str)
        vl10c_merged_df = pd.merge(vl10c_df, zsbe_df, left_on="SAP_nr", right_on="Materiał", how="left")

        # TODO: Filter LS1 and LS2 MRP controllers out
        # drop unnecessary columns and rename new column
        columns_to_drop = [
            "Materiał",
        ]
        vl10c_merged_df.drop(columns=columns_to_drop, inplace=True)
        new_col_names = {
            "Kontroler MRP": "mrp_controller",
            "Rodzaj nabycia": "procurement_type"
        }
        vl10c_merged_df.rename(columns=new_col_names, inplace=True)
        # filter the data
        vl10c_merged_df = vl10c_merged_df[~vl10c_merged_df['mrp_controller'].isin(["LS1", "LS2"])]
        # remove rows where 'product_name' starts with 'EBR' or 'EDR' and 'procurement_type' == 'E'
        vl10c_merged_df = vl10c_merged_df[
            ~(
                    vl10c_merged_df['product_name'].str.startswith(('EBR', 'EDR')) &
                    (vl10c_merged_df['procurement_type'] == 'E')
            )
        ]

        # Add header column
        vl10c_merged_df['header'] = vl10c_merged_df['document_number'] + " " + vl10c_merged_df['sales_office'].apply(lambda x: sales_offices_map[x])

        # TODO: match quantities to storage locations
        # create columns
        vl10c_merged_df['header_suffix'] = ""
        for loc in storage_locations_list:
            vl10c_merged_df[f'loc_{loc}'] = 0
        # vl10c_merged_df.to_pickle('excel_files/vl10c_merged_df.pkl')
        # copy SAP numbers to clipboard
        copy_df_column_to_clipboard(vl10c_merged_df, "SAP_nr")
        # open MB52 transaction
        open_one_transaction(session=sess1, transaction_name="MB52")
        simple_load_variant(sess1, "MISC_LU_PPS001", True)
        mb52_load_sap_numbers_and_export_data(session=sess1, file_path=str(paths['temp_folder']),
                                              file_name=paths['mb52_vl10c'].name)
        # close Excel file which should be automatically opened
        time.sleep(3)
        close_excel_file(file_name=paths['mb52_vl10c'].name)
        # load zsbe data into data frame
        mb52_df = pd.read_excel(paths['mb52_vl10c'], dtype={'Skład': str, 'Materiał': str})
        mb52_df.rename(columns={"Materiał": "SAP_nr", "Nieogranicz.wykorz.": "stock", "Skład": "storage_loc"},
                       inplace=True)
        # mb52_df.to_pickle('excel_files/mb52_df.pkl')
        filter_out_items_booked_to_0004_spec_cust_requirement_location(mb52_df, vl10c_merged_df)
        fill_storage_location_quantities(mb52_df, vl10c_merged_df)

        # save vl10c_merged_df to Excel file
        vl10c_merged_df.to_excel(paths['vl10c_clean_data'], index=False)
        # open Excel file
        run_excel_file_and_adjust_col_width(paths['vl10c_clean_data'])

        # Handle the information for status file
        # program_status["COHV_CONVERSION_SYSTEM_MESSAGE"] = result_sap_messages

    except Exception as e:
        print(e)
        logging.error("Error occurred", exc_info=True)

    finally:
        # close unnecessary files
        close_excel_file(file_name=paths['zsbe_data_vl10c'].name)
        time.sleep(0.1)
        close_excel_file(file_name=paths['mb52_vl10c'].name)
        time.sleep(0.1)
        close_excel_file(file_name=paths['mb52_vl10d'].name)
        time.sleep(0.1)
        close_excel_file(file_name=paths['zsbe_data_vl10d'].name)

        # Fill status file
        end_time = datetime.now().strftime("%H:%M:%S")
        program_status["start_time"] = start_time
        program_status["end_time"] = end_time
        # append_status_to_excel(
        #     status_file, program_status, ERROR_LOG_PATH, sheet_name="COHV_CONVERSION"
        # )
