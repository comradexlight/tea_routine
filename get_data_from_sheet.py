from dataclasses import dataclass
from typing import List, Union
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(slots=True)
class SaleRecordItem:
    """
    A data model for storing a record of the sale of an item.
    """
    title: str
    price: int
    qty: float
    amount: float


def get_data_from_sheet(input_ws: Worksheet) -> List[SaleRecordItem]:
    """
    The function collects data from the sheet into the list of SaleRecordItems.
    """
    cache = []
    end_row = "Сумма денег за чаепития"
    sales_record_items = []
    for row in input_ws.iter_rows(min_col=2, min_row=4, max_col=5, values_only=True):
        if row[0] == end_row:
            break
        if _is_row_need(row):
            title = _get_data_from_row(row, "title")
            price = _get_data_from_row(row, "price")
            qty = _get_data_from_row(row, "qty")
            amount = _get_data_from_row(row, "amount")
            if row[0] not in cache:
                sales_record_items.append(SaleRecordItem(
                    title=title,
                    price=price,
                    qty=qty,
                    amount=amount
                ))
                cache.append(row[0])
            else:
                existing_item = [item for item in sales_record_items if item.title == row[0]][0]
                sales_record_items[sales_record_items.index(existing_item)].qty += qty
                sales_record_items[sales_record_items.index(existing_item)].amount += amount
    return sales_record_items


def _is_row_need(row: tuple) -> bool:
    """
    The fucntion checks if the data from the row is needed for uploading.
    """
    unnecessary_rows = [
        "итог работы чаепитие администратор утро",
        "итог работы розница администратор утро",
        "Чаепитие 2 админа",
        "итог работы чаепитие 2 админа",
        "Чаепитие администратор вечер",
        "итог работы чаепитие администратор вечер",
        "Розница администратор утро",
        "итог работы розница администратор утро",
        "Розница 2 администратора",
        "итог работы розница 2 сотрудника",
        "Розница администратор вечер",
        "итог работы розница администратор вечер",
        "внутренние расходы",
        "итог работы розница 1 администратор",
        "итог работы чаепитие 1 администратор",
        None
    ]
    if row[0] in unnecessary_rows:
        return False
    return True


def _get_data_from_row(row: tuple, field: str) -> Union[str, float, int]:
    match field:
        case "title":
            return row[0] if (row[0] is not None) else "Пропущено название"
        case "price":
            return row[1] if (row[1] is not None) else 0
        case "qty":
            return row[2] if (row[2] is not None) else 0
        case "amount":
            return row[3] if (row[3] is not None) else 0
