# МогучийДД (MoguchiyDD)
# 2024.07.28, 12:06:35 PM
# convert_from_excel_to_json.py


from enum import Enum

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, coordinate_to_tuple

from json import load
from time import sleep


class ConvertFromExcelToJSON:
    """
    Convert from EXCEL to JSON

    PARAMETERS:
    - **filepath** : an EXCEL file
    - **json_schema** : a JSON SCHEMA file

    FUNCTIONS:
    - **cmd**(type: Commands, sheet: int = 0) -> Workbook | Worksheet |
    list[dict] | None : Running a class via available commands
    """

    class Commands(Enum):
        open = "open"
        sheet = "sheet"
        schema = "schema"
        getData = "getData"
        close = "close"

    def __init__(self, filepath: str, json_schema: str) -> None:
        self.filepath: str = filepath
        self.schema: str = json_schema
        self.data_schema: dict = {}
        self.data_for_save: list = []
        self.workbook: None | Workbook = None
        self.worksheet: None | Worksheet = None
        self.sheet: int = 0

    def cmd(
        self,
        type: Commands,
        sheet: int = 0
    ) -> Workbook | Worksheet | list[dict] | None:
        """
        Running a class via available commands

        PARAMETERS:
        - **type** : Variable from the **Command** class
        - **sheet** : Opening sheet number from an EXCEL file

        COMMANDS:
        - «**open**» : Opening an EXCEL file
        - «**sheet**» : Taking a sheet from an EXCEL file for processing
        (default 1st)
        - «**schema**» : Opening a SCHEMA (JSON file)
        - «**getData**» : Get data from an EXCEL file
        - «**close**» : Closing an EXCEL file

        RETURN: the commands «**open**», «**sheet**» and «**getData**» have
        data return, and the rest nothing
        """

        if type is self.Commands.open:
            self.__enter__()
            return self.workbook
        elif type is self.Commands.sheet:
            self.sheet = sheet
            self.__sheet__()
            return self.worksheet
        elif type is self.Commands.schema:
            self.__schema__()
        elif type is self.Commands.getData:
            self.__get_data__()
            return self.data_for_save
        elif type is self.Commands.close:
            self.__exit__()
        
        return None

    def __enter__(self) -> None:
        """
        Opening an EXCEL file \\
        **Launch via the «open» command**
        """

        self.workbook: Workbook = load_workbook(self.filepath, data_only=True)

    def __sheet__(self) -> None:
        """
        Taking a sheet from an EXCEL file for processing (default 1st) \\
        **Launch via the «sheet» command**
        """

        self.worksheet: Worksheet = self.workbook.worksheets[self.sheet]

    def __schema__(self) -> None:
        """
        Opening a SCHEMA (JSON file) \\
        **Launch via the «schema» command**
        """

        with open(self.schema, "r") as schema:
            self.data_schema: dict = load(schema)

    def __get_data__(self) -> None:
        """
        Extracting data from an EXCEL file of a specific 1 sheet into
        a dictionary \\
        **Launch via the «getData» command**
        """

        def find_merge_cells(field: str) -> tuple[int] | None:
            """
            Determine by field whether it is included in the merge of fields

            PARAMETERS:
            - **field** : For example, A9

            RETURN: (minimum row, minimum column, maximum row, maximum column)
            """

            coordinates: tuple[int, int] = coordinate_to_tuple(field)
            row: int = coordinates[0]
            column: int = coordinates[1]

            for merged_cell_range in self.worksheet.merged_cells.ranges:
                min_col, min_row, max_col, max_row = merged_cell_range.bounds
                if min_row <= row <= max_row and min_col <= column <= max_col:
                    merge_cells: tuple[int] = (
                        min_row, min_col,
                        max_row, max_col
                    )
                    return merge_cells
                sleep(0.01)

            return None

        if len(self.data_schema) == 0:
            print("Schema is empty")
            exit(-1)

        # Read Schema
        rows_to: int = self.data_schema["rows_to"]
        rows_from: int = self.data_schema["rows_from"]
        rows_except: list[int] | None = self.data_schema["rows_except"]
        fields_to: str = self.data_schema["fields_to"]
        fields_from: str = self.data_schema["fields_from"]
        fields_except: list[str] | None = [
            self.data_schema["fields_except"]
            if self.data_schema["fields_except"] else []
        ][0]

        # Important Variables
        is_loop: bool = True
        stop_count: int = 1
        count: int = 1
        cur_row: int = rows_to
        fields_data: dict = {}

        while is_loop:  # for 'field_to'
            col_letter: str = get_column_letter(count)
            if col_letter == fields_to:
                is_loop = False
                stop_count, count = count, stop_count
            else:
                count += 1
            sleep(0.01)

        is_loop = True
        while is_loop:  # main
            col_letter = get_column_letter(count)
            if col_letter not in fields_except:  # Save Data from Excel
                field: str = f"{ col_letter }{ cur_row }"
                is_merge_field: tuple[int] | None = find_merge_cells(field)
                if is_merge_field:
                    col_letter = get_column_letter(is_merge_field[1])
                    field = f"{ col_letter }{ is_merge_field[0] }"
                fields_data[field] = self.worksheet[field].value

            if col_letter == fields_from:  # Finish work for Row
                self.data_for_save.append({cur_row: fields_data})
                fields_data = {}
                cur_row += 1

                if rows_except:  # Excpect
                    while cur_row in rows_except:
                        cur_row += 1

                count = stop_count

                if cur_row == (rows_from + 1):
                    is_loop = False
            else:
                count += 1
            sleep(0.01)

    def __exit__(self) -> None:
        """
        Closing an EXCEL file \\
        **Launch via the «close» command**
        """

        if self.workbook:
            self.workbook.close()
