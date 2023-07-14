import time
import pandas as pd

from sys import argv
from typing import List


def get_sheet_names(workbook: str) -> List[str]:
    wb = pd.ExcelFile(workbook)
    sheet_names = wb.sheet_names
    wb.close()
    return sheet_names


def get_active_sheets(sheet_names: List[str]) -> List[str]:
    MSG = "select the page index to be prepared for uploading," \
            "separated by space, like: 1 2 3: " 
    active_sheets = []
    for index, name in enumerate(sheet_names):
        print(index, "-", name)
    choice = input(MSG)
    if choice == "":
        active_sheets = sheet_names[2]
    else:
        active_sheets_index = [int(element) for element in choice.split(" ")]
        for index, name in enumerate(sheet_names):
            if index in active_sheets_index:
                active_sheets.append(name)
    return active_sheets 


def get_DataFrame_from_active_sheet(workbook_raw: str, active_sheets: List[str]) -> pd.DataFrame:
    df = pd.read_excel(
            io=workbook_raw,
            sheet_name=active_sheets,
            header=None,
            names=["title", "price", "qty", "amount", "discount"],
            usecols=[1, 2, 3, 4, 5],
            engine="openpyxl",
            skiprows=3,
            na_filter=False
            )
    return df


def clean_row_title(row: pd.Series):
    row["title"] = row["title"].removesuffix(", , кг").removesuffix(", , шт")


def filter_DataFrame(df: pd.DataFrame) -> pd.DataFrame:
    feltered_df = pd.DataFrame() 
    for index, row in df.iterrows():
        # if row["title"] != feltered_df.loc[f'{row["title"]}']:
            # feltered_df.add(row)
        clean_row_title(row)
        print(row["title"])
    return feltered_df


def main(path: str) -> None:
    workbook_raw = path
    sheet_names = get_sheet_names(workbook_raw)
    active_sheets = get_active_sheets(sheet_names)
    df = get_DataFrame_from_active_sheet(workbook_raw, active_sheets)
    print(filter_DataFrame(df))
    

if __name__ == '__main__':
    start_time = time.time()
    path = argv[1]
    main(path)
    print(time.time() - start_time)

