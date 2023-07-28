from time import time
from sys import argv
from dataclasses import dataclass
from typing import List
# from multiprocessing import Process

from icu import Collator, Locale
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import load_workbook


@dataclass(slots=True)
class SaleRecordItem:
    title: str
    price: int
    qty: float
    amount: float
    

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

end_row = "Сумма денег за чаепития"


def read_xlsx2wb(path: str) -> Workbook:
    return load_workbook(filename=path, read_only=False)


def select_active_ws_names(wb: Workbook) -> List[str]:
    MSG = "select the page index to be prepared for uploading," \
            "separated by space, like: 1 2 3: " 
    all_sheet_names = wb.sheetnames
    active_ws = []
    for index, name in enumerate(all_sheet_names):
        print(index, "-", name)
    choice = input(MSG)
    if choice == "":
        active_ws.append(all_sheet_names[2])
    else:
        active_ws_index = [int(element) for element in choice.split(" ")]
        for index, name in enumerate(all_sheet_names):
            if index in active_ws_index:
                active_ws.append(name)
    return active_ws 


def collect_data_from_ws(ws: Worksheet) -> List[SaleRecordItem]:
    cache = []
    sales_record_items = []
    for row in ws.iter_rows(min_col=2, min_row=4, max_col=5, values_only=True):
        if row[0] != end_row and is_row_need(row):
            if row[0] not in cache:
                sales_record_items.append(SaleRecordItem(
                    title=row[0],
                    price=row[1],
                    qty=row[2],
                    amount=row[3],
                ))
                cache.append(row[0])
            else:
                tmp_item = [item for item in sales_record_items if item.title == row[0]][0]
                sales_record_items[sales_record_items.index(tmp_item)].qty += row[2]
                sales_record_items[sales_record_items.index(tmp_item)].amount += row[3]
        elif row[0] == end_row:
            break

    return sales_record_items


def is_row_need(row) -> bool:
    if [cell for cell in row if cell is None]:
        return False 
    elif row[0].strip() in unnecessary_rows:
        return False
    return True


def create_sheet_for_upload(wb2upload: Workbook,
                           list2upload: List[SaleRecordItem],
                           sheet_name: str) -> None:
    ws = wb2upload.create_sheet(sheet_name)
    print(ws.title, sheet_name)
    for row, element in enumerate(list2upload, start=1):
        ws.cell(row=row, column=1, value=element.title)
        ws.cell(row=row, column=2, value=element.price)
        ws.cell(row=row, column=3, value=element.qty)
        ws.cell(row=row, column=4, value=element.amount)


def sort_list2upload(list2upload: List[SaleRecordItem]) -> None:
    collator = Collator.createInstance(Locale('ru_RU'))
    list2upload.sort(key=lambda element: collator.getSortKey(element.title))
    return


def convert_xlsx(wb2upload: Workbook, ws: Worksheet) -> None:
    list2upload = collect_data_from_ws(ws)
    sort_list2upload(list2upload)
    create_sheet_for_upload(wb2upload, list2upload, ws.title)


def main() -> None:
    path = argv[1]
    wb = read_xlsx2wb(path)
    wb2upload = Workbook()  
    active_ws_names = select_active_ws_names(wb)
    active_ws = [wb[ws_name] for ws_name in active_ws_names]
    # processes = []
    for ws in active_ws:
        convert_xlsx(wb2upload, ws)
    wb2upload.save("2upload.xlsx")
    # for ws in active_ws:
        # processes.append(Process(target=convert_xlsx, args=(wb2upload, ws,),
                                 # daemon=True))
    # [process.start() for process in processes]
    # [process.join() for process in processes]


if __name__ == '__main__':
    start_time = time()
    main()
    print(time() - start_time)

