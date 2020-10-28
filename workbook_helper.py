import xlwings as xw
#import win32api
from typing import List, Optional

class WorkBookHelper:
    def __init__(self, book: xw.main.Book):
        self.book: xw.main.Book = book
        self.all_sheet_names: List[str] = [sheet.name for sheet in self.book.sheets]
        self.all_cell_names: List[str] = [cell_name.name for cell_name in self.book.names]
        self.__invalid_counter: int = 0
        self.__number_of_sheet_checks: int = 0
        self.__number_of_cell_name_checks: int = 0

    @property
    def invalid_counter(self) -> int:
        return self.__invalid_counter

    @property
    def number_of_sheet_checks(self) -> int:
        return self.__number_of_sheet_checks

    @property
    def number_of_cell_name_checks(self) -> int:
        return self.__number_of_sheet_checks

    def check_sheet(self, sheet_name: str) -> Optional[xw.main.Sheet]:
        self.__number_of_sheet_checks += 1
        if sheet_name in self.all_sheet_names:
            return self.book.sheets[sheet_name]
        else:
            self.__invalid_counter += 1
            #win32api.MessageBox(self.book.app.hwnd, f"{sheet_name} 시트를 찾을 수 없습니다.")
            return None

    def check_cell_name(self, cell_name: str) -> Optional[str]:
        self.__number_of_cell_name_checks += 1
        if cell_name in self.all_cell_names:
            return cell_name
        else:
            self.__invalid_counter += 1
            #win32api.MessageBox(self.book.app.hwnd, f"{cell_name} 이름를 찾을 수 없습니다.")
            return None

    def number_of_checks(self) -> int:
        return self.__number_of_cell_name_checks + self.__number_of_sheet_checks

    def all_valid(self) -> bool:
        return self.number_of_checks() > 0 and self.__invalid_counter == 0

# active_book = xw.books.active
# checker = WorkBookHelper(xw.books.active)
