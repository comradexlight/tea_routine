from time import time
from sys import argv
from dataclasses import dataclass
from typing import List
from multiprocessing import Process

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import load_workbook


@dataclass(slots=True, frozen=True)
class SaleRecordItem:
    title: str
    price: int
    qty: float
    amount: float
    discount: int
    

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
    for row in ws.iter_rows(min_col=2, min_row=4, max_col=6, values_only=True):
        if row[0] not in cache:
            sales_record_items.append(SaleRecordItem(
                title=row[0],
                price=row[1],
                qty=row[2],
                amount=row[3],
                discount=row[4]
            ))
            cache.append(row[0])
        else:
            pass
            #TODO
    
    return sales_record_items


def convert_xlsx(ws: Worksheet) -> None:
    print(collect_data_from_ws(ws))


def main() -> None:
    path = argv[1]
    wb = read_xlsx2wb(path)
    active_ws_names = select_active_ws_names(wb)
    active_ws = [wb[ws_name] for ws_name in active_ws_names]
    processes = []
    for ws in active_ws:
        processes.append(Process(target=convert_xlsx, args=(ws,), daemon=True))
    [process.start() for process in processes]
    [process.join() for process in processes]


if __name__ == '__main__':
    start_time = time()
    main()
    print(time() - start_time)

