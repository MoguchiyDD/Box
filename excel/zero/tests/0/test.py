# МогучийДД (MoguchiyDD)
# 2024.07.28, 08:42:44 PM
# test.py


from sys import path as syspath
from os import path as ospath

syspath.append(
    ospath.abspath(
        ospath.join(
            ospath.dirname(__file__),
            "../../../"
        )
    )
)

from zero.convert_from_excel_to_json import ConvertFromExcelToJSON


# ------------ VARIABLES ------------

EXCEL = "../excelfile.test.xlsx"
SCHEMA = "schema.test.json"

# -----------------------------------


if __name__ == "__main__":
    convert = ConvertFromExcelToJSON(EXCEL, SCHEMA, "result.test")
    convert.cmd(convert.Commands.open)
    convert.cmd(convert.Commands.schema)
    convert.cmd(convert.Commands.sheet)
    convert.cmd(convert.Commands.getDataAndLoad)
    convert.cmd(convert.Commands.close)
