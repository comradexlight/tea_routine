from typing import List
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def _select_active_ws_names(input_wb: Workbook) -> List[str]:
    """
    The function asks the user which sheets frim incoming xlsx file are needed.
    """
    msg = "select the page index to be prepared for uploading," \
            "separated by space, like: 1 2 3: " 
    all_sheet_names = input_wb.sheetnames
    active_ws_names = []
    for index, name in enumerate(all_sheet_names):
        print(index, "-", name)
    choice = input(msg)
    if choice == "":
        active_ws_names.append(all_sheet_names[2])
    else:
        active_ws_index = [int(element) for element in choice.strip().split(" ")]
        for index, name in enumerate(all_sheet_names):
            if index in active_ws_index:
                active_ws_names.append(name)
    return active_ws_names


def get_active_ws(input_wb: Workbook) -> List[Worksheet]:
    """
    The function returns active worksheets by names of worksheets.
    """
    active_sheet_names = _select_active_ws_names(input_wb)
    return [input_wb[ws_name] for ws_name in active_sheet_names]
