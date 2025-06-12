import os
from datetime import datetime
from pathlib import Path


class ProgramPaths:
    BASE_PATH = Path(
        r"P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\05_VL10D_VL10C_MIGO"
    )
    ERROR_LOG_PATH = BASE_PATH / "error.log"

    username = os.getlogin()
    status_file = (
        f"C:/Users/{username}/OneDrive - Roto Frank DST/General/05_Automatyzacja_narzędzia/100_STATUS"
        f"/01_AUTOMATION_TOOLS_STATUS.xlsx"
    )

    today = datetime.today().strftime("%Y_%m_%d")
    start_time = datetime.now().strftime("%H:%M:%S")

    file_paths = {
        "vl10d_raw_data": f"temp/vl10d_raw_data.xls",
        "vl10d_clean_data": f"historical_data/vl10d_clean_data_{today}.xlsx",
        "vl10c_raw_data": f"temp/vl10c_raw_data.xls",
        "vl10c_clean_data": f"historical_data/vl10c_clean_data_{today}.xlsx",
        "vl10c_raw_data_treppen": f"temp/vl10c_raw_data_treppen.xls",
        "vl10c_clean_data_treppen": f"historical_data/vl10c_clean_data_treppen_{today}.xlsx",
        "historical_data": "historical_data",
        "temp_folder": "temp",
        "zsbe_data_vl10d": "temp/zsbe_data_vl10d.xlsx",
        "zsbe_data_vl10c": "temp/zsbe_data_vl10c.xlsx",
        "zsbe_data_vl10c_treppen": "temp/zsbe_data_vl10c_treppen.xlsx",
        "mb52_vl10d": "temp/mb52_vl10d.xlsx",
        "mb52_vl10c": "temp/mb52_vl10c.xlsx",
        "mb52_vl10c_treppen": "temp/mb52_vl10c_treppen.xlsx",
        "to_numbers": f"historical_data/transfer_orders_numbers_{today}.xlsx",
        "mb52_mat_docs_nums_313": f"historical_data/mb52_mat_docs_nums_313_{today}.xlsx",
        "mb52_mat_docs_nums_315": f"historical_data/mb52_mat_docs_nums_315_{today}.xlsx",
        "mb52_mat_docs_nums_313_treppen": f"historical_data/mb52_mat_docs_nums_313_treppen_{today}.xlsx",
        "mb52_mat_docs_nums_315_treppen": f"historical_data/mb52_mat_docs_nums_315_treppen_{today}.xlsx",
        "mb51_vl10c": "temp/mb51_vl10c.xlsx"
    }

    def __init__(self):
        self.paths = {key: self.BASE_PATH / filename for key, filename in self.file_paths.items()}
