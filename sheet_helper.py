import xlwings as xw
import pandas as pd
import datetime
from typing import Any, Optional, Dict, List

"""
TODO:
1: Polish code
2: Raise error in Excel msgbox
3. Anyway to call multiple columns? eg) range([A, C, F]).value
range(A).value + range(C).value + range(F).value is too slow
"""

class SheetHelper:
    def __init__(self, sheet: xw.main.Sheet):
        self.sheet: xw.main.Sheet = sheet
        self._col_headers_location: Dict[str, Dict[Any, List[str]]] = {}
        self._row_headers_location: Dict[str, Dict[Any, List[str]]] = {}
        self.__last_cell: xw.main.Range = sheet.cells.last_cell
        self.__last_row: int = self.__last_cell.row
        self.__last_col_num: int = self.__last_cell.column

    def range(self, cell: str) -> xw.main.Range:
        return self.sheet.range(cell)

    def cells(self, row: int, col: int) -> xw.main.Range:
        return self.sheet.cells(row, col)

    def iter_row(self, cell: xw.main.Range) -> Optional[xw.main.Range]:
        if cell.row == self.__last_row:
            return None
        else:
            # seems faster than cell.offset(1, 0)
            return self.cells(cell.row + 1, cell.column)

    def iter_col(self, cell: xw.main.Range) -> Optional[xw.main.Range]:
        if cell.column == self.__last_col_num:
            return None
        else:
            # seems faster than cell.offset(0, 1)
            return self.cells(cell.row, cell.column + 1)

    def get_range_in_col(self, col: str, start_row: Optional[int] = None, end_row: Optional[int] = None) -> xw.main.Range:
        if col.isalpha():
            upper = col.upper()
            col_num = self.column_number(col)
            if start_row is not None and end_row is not None:
                if start_row > 0:
                    if end_row >= start_row:
                        start_cell = upper + str(start_row)
                        end_cell = upper + str(end_row)
                        rng = start_cell + ":" + end_cell
                        return self.range(rng)
                    else:
                        raise Exception("end_row는 start_row 보다 같거나 커야합니다")
                else:
                    raise Exception("1 이상의 값을 받아야 합니다")
            elif start_row is None and end_row is not None:
                if end_row > 0:
                    start_cell = upper + str(1)
                    last_cell = upper + str(end_row)
                    rng = start_cell + ":" + last_cell
                    return self.range(rng)
                else:
                    raise Exception("1 이상의 값을 받아야 합니다")
            elif start_row is not None and end_row is None:
                if start_row > 0:
                    last_non_empty_row = self.cells(self.__last_row, col_num).end("up").row
                    if start_row > last_non_empty_row:
                        start_cell = upper + str(start_row)
                        rng = start_cell + ":" + start_cell
                        return self.range(rng)
                    else:
                        start_cell = upper + str(start_row)
                        last_cell = upper + str(last_non_empty_row)
                        rng = start_cell + ":" + last_cell
                        return self.range(rng)
                else:
                    raise Exception("1 이상의 값을 받아야 합니다")
            else: 
                last_non_empty_row = self.cells(self.__last_row, col_num).end("up").row
                start_cell = upper + str(1)
                last_cell = upper + str(last_non_empty_row)
                rng = start_cell + ":" + last_cell
                return self.range(rng)
        else:
            raise Exception("알파벳을 입력해 주십시오")

    def get_range_in_row(self, row: int, start_col: Optional[str] = None, end_col: Optional[str] = None) -> xw.main.Range:
        if row > 0 :
            row_str = str(row)
            if start_col is not None and end_col is not None:
                start_col_num = self.column_number(start_col)
                end_col_num = self.column_number(end_col)
                if end_col_num >= start_col_num:
                    start_cell = start_col + row_str
                    last_cell = end_col + row_str
                    rng = start_cell + ":" + last_cell
                    return self.range(rng)
                else:
                    raise Exception("end_col은 start_col 뒤에 있어야 합니다.")
            elif start_col is None and end_col is not None:
                start_cell = "A" + row_str
                last_cell = end_col + row_str
                rng = start_cell + ":" + last_cell
                return self.range(rng)
            elif start_col is not None and end_col is None:
                last_non_empty_col = self.cells(row, self.__last_col_num).end("left").column
                last_non_empty_col_letter = self.column_letter(last_non_empty_col)
                if self.column_number(start_col) > last_non_empty_col:
                    start_cell = start_col + row_str
                    rng = start_cell + ":" + start_cell
                    return self.range(rng)
                else:
                    start_cell = start_col + row_str
                    last_cell = last_non_empty_col_letter + row_str
                    rng = start_cell + ":" + last_cell
                    return self.range(rng)
            else: # None, None
                last_non_empty_col = self.cells(row, self.__last_col_num).end("left").column
                last_non_empty_col_letter = self.column_letter(last_non_empty_col)
                start_cell = "A" + row_str
                last_cell = last_non_empty_col_letter + row_str
                rng = start_cell + ":" + last_cell
                return self.range(rng)
        else:
            raise Exception("1 이상의 값을 받아야 합니다")

    def get_values_in_col(self, col: str, start_row: Optional[int] = None, end_row: Optional[int] = None) -> pd.Series:
        return self.get_range_in_col(col, start_row, end_row).options(pd.DataFrame, index=False, header=False).value.squeeze()

    def get_values_in_row(self, row: int, start_col: Optional[str] = None, end_col: Optional[str] = None) -> pd.Series:
        return self.get_range_in_row(row, start_col, end_col).options(pd.DataFrame, index=False, header=False).value.squeeze()

    def get_all_values_in_col(self, col: str) -> pd.Series:
        return self.get_values_in_col(col)

    def get_all_values_in_row(self, row: int) -> pd.Series:
        return self.get_values_in_row(row)

    def get_value_idx_in_col(self, col: str, start_row: Optional[int] = None, end_row: Optional[int] = None) -> Dict[Any, List[str]]:
        column_values = self.get_values_in_col(col, start_row, end_row)
        dict: Dict[Any, List[str]] = {}
        for idx, value in enumerate(column_values.values):
            if value is not None:
                cell = col + str(idx + 1)
                if type(value) == datetime.datetime:
                    date_string = value.strftime("%Y-%m-%d")
                    if date_string in dict:
                        dict[date_string].append(cell)
                    else:
                        dict[date_string] = [cell]
                else:
                    if value in dict:
                        dict[value].append(cell)
                    else:
                        dict[value] = [cell]
            else:
                continue
        return dict

    def get_value_idx_in_row(self, row: int, start_col: Optional[str] = None, end_col: Optional[str] = None) -> Dict[Any, List[str]]:
        row_values = self.get_values_in_row(row, start_col, end_col)
        dict: Dict[Any, List[str]] = {}
        for idx, value in enumerate(row_values.values):
            if value is not None:
                cell = self.column_letter(idx + 1) + str(row)
                if type(value) == datetime.datetime:
                    date_string = value.strftime("%Y-%m-%d")
                    if date_string in dict:
                        dict[date_string].append(cell)
                    else:
                        dict[date_string] = [cell]
                else:
                    if value in dict:
                        dict[value].append(cell)
                    else:
                        dict[value] = [cell]
            else:
                continue
        return dict

    def get_col_from_cell(self, cell: str) -> str:
        return self.column_letter(self.range(cell).column)

    def get_col_num_from_cell(self, cell: str) -> int:
        return self.range(cell).column

    def get_row_from_cell(self, cell: str) -> int:
        return self.range(cell).row

    def filter_cells_from_col(self, cell_list: List[str], col: str) -> List[str]:
        if col.isalpha():
            col_num = self.column_number(col)
            res = []
            for cell in cell_list:
                if self.get_col_num_from_cell(cell) >= col_num:
                    res.append(cell)
                else:
                    continue
            return res
        else:
            raise Exception("알파벳을 입력해 주십시오")

    def filter_cells_after_col(self, cell_list: List[str], col: str) -> List[str]:
        if col.isalpha():
            next_col_num = self.column_number(col) + 1
            next_col_letter = self.column_letter(next_col_num)
            return self.filter_cells_from_col(cell_list, next_col_letter)
        else:
            raise Exception("알파벳을 입력해 주십시오")

    def filter_cells_from_row(self, cell_list: List[str], row: int) -> List[str]:
        if row > 0:
            res = []
            for cell in cell_list:
                cell_row = self.get_row_from_cell(cell)
                if cell_row >= row:
                    res.append(cell)
                else:
                    continue
            return res
        else:
            raise Exception("1 이상의 값을 받아야 합니다")

    def filter_cells_after_row(self, cell_list: List[str], row: int) -> List[str]:
        if row >= 0:
            return self.filter_cells_from_row(cell_list, row + 1) 
        else:
            raise Exception("0 이상의 값을 받아야 합니다")

    def find_first_location_in_col(self, col: str, value: Any, start_row: Optional[int] = None, value_idx_col: Optional[Dict[Any, List[str]]] = None) -> str:
        row = start_row if start_row is not None else 1
        all_locations = value_idx_col[value] if value_idx_col is not None else self.get_value_idx_in_col(col)[value]
        upper = col.upper()
        filtered_locations = self.filter_cells_from_row(all_locations, row)
        if len(filtered_locations) == 0:
            raise Exception(f"{upper}열에서 {value} 값을 찾을 수 없습니다.")
        else:
            return filtered_locations[0]

    def find_first_location_in_row(self, row: int, value: Any, start_col: Optional[str] = None, value_idx_row: Optional[Dict[Any, List[str]]] = None) -> str:
        col = start_col if start_col is not None else "A"
        all_locations = value_idx_row[value] if value_idx_row is not None else self.get_value_idx_in_row(row)[value]
        filtered_locations = self.filter_cells_from_col(all_locations, col)
        if len(filtered_locations) == 0:
            raise Exception(f"{row}행에서 {value} 값을 찾을 수 없습니다.")
        else:
            return filtered_locations[0]

    @property
    def col_headers_location(self):
        return self._col_headers_location

    @property
    def row_headers_location(self):
        return self._row_headers_location

    # @staticmethod
    # def update_location(location_dict: Dict, value: Any, location: str) -> None:
    #     if value is not None:
    #         if type(value) == datetime.datetime:
    #             date_string = value.strftime("%Y-%m-%d")
    #             if date_string in location_dict:
    #                 location_dict[date_string].append(location)
    #             else:
    #                 location_dict[date_string] = [location]
    #         else:
    #             if value in location_dict:
    #                 if value not in location_dict[value]:
    #                     location_dict[value].append(value)
    #             else:
    #                 location_dict[value] = [location]

    # def update_col_headers_location(self, value: Any, location: str):
    #     col_letter = self.get_col_from_cell(location)
    #     if col_letter not in self.col_headers_location:

    #     pass

    # def update_row_headers_location(self, value: Any, location: str):
    #     row = self.get_row_from_cell(location)
    #     pass

    @staticmethod
    def column_letter(n: int) -> str:
        if n < 1:
            raise Exception("1 이상의 값을 받아야합니다.")
        else:
            string = ""
            ascii_A = ord("A")
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(ascii_A + remainder) + string
            return string

    @staticmethod
    def column_number(col: str) -> int:
        if col.isalpha():
            ascii_A = ord("A")
            n = 0
            for alpha in col.upper():
                if n is None:
                    n = 1 + ord(alpha) - ascii_A
                else:
                    n = n * 26 + 1 + ord(alpha) - ascii_A
            return n
        else:
            raise Exception("알파벳을 입력해 주십시오")
