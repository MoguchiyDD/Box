# МогучийДД (MoguchiyDD)
# 2024.07.28, 12:06:35 PM
# convert_from_excel_to_json.py


from enum import Enum

from openpyxl import Workbook, load_workbook


class ConvertFromExcelToJSON:
    """
    Convert from EXCEL to JSON

    PARAMETERS:
    - **filepath** : an EXCEL file

    FUNCTIONS:
    - **cmd**(type: Commands) -> Workbook : Running a class via available
    commands
    """

    class Commands(Enum):
        open = "open"
        close = "close"

    def __init__(self, filepath: str) -> None:
        self.filepath: str = filepath
        self.workbook: None | Workbook = None

    def cmd(self, type: Commands) -> Workbook | None:
        """
        Running a class via available commands

        PARAMETERS:
        - **type** : Variable from the **Command** class

        COMMANDS:
        - «**open**» : Opening an EXCEL file
        - «**close**» : Closing an EXCEL file

        RETURN: the commands «**open**» have data return, and the rest nothing
        """

        if type is self.Commands.open:
            self.__enter__()
            return self.workbook
        elif type is self.Commands.close:
            self.__exit__()
        
        return None

    def __enter__(self) -> None:
        """
        Opening an EXCEL file \\
        **Launch via the «open» command**
        """

        self.workbook: Workbook = load_workbook(self.filepath, data_only=True)

    def __exit__(self) -> None:
        """
        Closing an EXCEL file \\
        **Launch via the «close» command**
        """

        if self.workbook:
            self.workbook.close()
