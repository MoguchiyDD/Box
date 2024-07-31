# МогучийДД (MoguchiyDD)
# 2024.07.28, 12:06:35 PM
# convert_from_excel_to_json.py


from enum import Enum

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from json import load


class ConvertFromExcelToJSON:
    """
    Convert from EXCEL to JSON

    PARAMETERS:
    - **filepath** : an EXCEL file
    - **json_schema** : a JSON SCHEMA file

    FUNCTIONS:
    - **cmd**(type: Commands) -> Workbook : Running a class via available
    commands
    """

    class Commands(Enum):
        open = "open"
        sheet = "sheet"
        schema = "schema"
        close = "close"

    def __init__(self, filepath: str, json_schema: str) -> None:
        self.filepath: str = filepath
        self.schema: str = json_schema
        self.workbook: None | Workbook = None
        self.worksheet: None | Worksheet = None
        self.sheet: int = 0

    def cmd(
        self,
        type: Commands,
        sheet: int = 0
    ) -> Workbook | Worksheet | None:
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
        - «**close**» : Closing an EXCEL file

        RETURN: the commands «**open**» and «**sheet**»have data return,
        and the rest nothing
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

    def __exit__(self) -> None:
        """
        Closing an EXCEL file \\
        **Launch via the «close» command**
        """

        if self.workbook:
            self.workbook.close()
