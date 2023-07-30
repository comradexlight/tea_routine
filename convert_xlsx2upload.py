"""
Prepares data from incoming xlsx files of sales.
Saves processed data in xlsx files for uploading to 1C.
"""
from time import time
from sys import argv
from dataclasses import dataclass
from typing import List

from icu import Collator, Locale
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import load_workbook


@dataclass(slots=True)
class SaleRecordItem:
    """
    A data model for storing a record of the sale of an item.
    """
    title: str
    price: int
    qty: float
    amount: float


def select_active_ws_names(input_wb: Workbook) -> List[str]:
    """
    The function asks the user which sheets frim incoming xlsx file are needed.
    """
    msg = "select the page index to be prepared for uploading," \
            "separated by space, like: 1 2 3: " 
    all_sheet_names = input_wb.sheetnames
    active_ws = []
    for index, name in enumerate(all_sheet_names):
        print(index, "-", name)
    choice = input(msg)
    if choice == "":
        active_ws.append(all_sheet_names[2])
    else:
        active_ws_index = [int(element) for element in choice.split(" ")]
        for index, name in enumerate(all_sheet_names):
            if index in active_ws_index:
                active_ws.append(name)
    return active_ws


def collect_data_from_ws(input_ws: Worksheet) -> List[SaleRecordItem]:
    """
    The function collects data from the sheet into the list of SaleRecordItems.
    """
    cache = []
    end_row = "Сумма денег за чаепития"
    sales_record_items = []
    for row in input_ws.iter_rows(min_col=2, min_row=4, max_col=5, values_only=True):
        if row[0] != end_row and is_row_need(row):
            if row[0] not in cache:
                sales_record_items.append(SaleRecordItem(
                    title=str(row[0]),
                    price=int(row[1]),
                    qty=float(row[2]),
                    amount=float(row[3]),
                ))
                cache.append(row[0])
            else:
                tmp_item = [item for item in sales_record_items if item.title == row[0]][0]
                sales_record_items[sales_record_items.index(tmp_item)].qty += float(row[2])
                sales_record_items[sales_record_items.index(tmp_item)].amount += float(row[3])
        elif row[0] == end_row:
            break

    return sales_record_items


def is_row_need(row: tuple) -> bool:
    """
    The fucntion checks if the data from the row is needed for uploading.
    """
    unnecessary_rows = [
        "итог работы чаепитие администратор утро",
        "Чаепитие 2 админа",
        "итог работы чаепитие 2 админа",
        "Чаепитие администратор вечер",
        "итог работы чаепитие администратор вечер",
        "Розница администратор утро",
        "итог работы розница администратор утро"
        "Розница 2 администратора",
        "итог работы розница 2 сотрудника",
        "Розница администратор вечер",
        "итог работы розница администратор вечер",
        "внутренние расходы",
    ]
    if [cell for cell in row if cell is None]:
        return False
    if row[0].strip() in unnecessary_rows:
        return False
    return True


def create_sheet_for_upload(wb2upload: Workbook,
                            list2upload: List[SaleRecordItem],
                            sheet_name: str) -> None:
    """
    The function inserts the data from the list to upload into the sheet of
    xlsx document to be uploaded.
    """
    new_ws = wb2upload.create_sheet(sheet_name)
    adjusted_width = 0
    for row, element in enumerate(list2upload, start=1):
        title = element.title.removesuffix(", , шт").removesuffix(", , кг")

        if adjusted_width < len(title):
            adjusted_width = len(title)

        new_ws.cell(row=row, column=1, value=title)
        new_ws.cell(row=row, column=2, value=element.price)
        new_ws.cell(row=row, column=3, value=element.qty)
        new_ws.cell(row=row, column=4, value=element.amount)
        new_ws.cell(row=row, column=5, value=f"=100*(1-D{row}/(B{row}*C{row}))")

    new_ws.column_dimensions['A'].width = adjusted_width
    new_ws.cell(row=len(list2upload) + 2, column=4, value=f"=SUM(D1:D{len(list2upload)})")


def sort_list2upload(list2upload: List[SaleRecordItem]) -> None:
    """
    The function sorts the items in the list of SaleRecordItems by title in ABC order.
    """
    collator = Collator.createInstance(Locale('ru_RU'))
    list2upload.sort(key=lambda element: collator.getSortKey(element.title))


def convert_xlsx(wb2upload: Workbook, input_ws: Worksheet) -> None:
    """
    The function collects and runs other functions to crete sheet in
    the xlsx document for uploading.
    """
    #TODO Need to cobine with main() function?
    list2upload = collect_data_from_ws(input_ws)
    sort_list2upload(list2upload)
    create_sheet_for_upload(wb2upload, list2upload, input_ws.title)


def main() -> None:
    """
    The Entry point.
    """
    path = argv[1]
    input_wb = load_workbook(filename=path, read_only=False)
    wb2upload = Workbook()
    active_ws_names = select_active_ws_names(input_wb)
    active_ws = [input_wb[ws_name] for ws_name in active_ws_names]
    for input_ws in active_ws:
        convert_xlsx(wb2upload, input_ws)
    wb2upload.save("2upload.xlsx")


if __name__ == '__main__':
    start_time = time()
    main()
    print(time() - start_time)
