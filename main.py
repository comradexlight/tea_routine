import pandas as pd
from fix_1c_error import fix_1c_error

def main(path: str) -> None:
    workbook = pd.read_excel(
            io=path,
            engine="openpyxl",
            sheet_name=0,
            header=None,
            usecols="E, N, P, Q, R, S, T, U, W, Y, AA, AC",
            #dtype={"E": str, "N": object, "P": int, "Q": int, "R": int, "S": int, "T": int, "U": float, "W": float, "Y": float, "AA": float, "AC": float},
            skiprows=7
            )
    
    print(workbook)


if __name__ == "__main__":
    path = "baza.xlsx"
    fix_1c_error(path)
    main(path)
