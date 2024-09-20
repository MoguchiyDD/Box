# Convert from EXCEL to JSON through schema
The «**ConvertFromExcelToJSON**» class

## Schema
```bash
# Schema for reading data from EXCEL and saving in JSON format
{
  "rows_from": integer,  # Beginning to read rows
  "rows_to": integer,  # End of reading rows (inclusive)
  "rows_except": integers array or NULL,  # Which rows not to take
  "fields_from": character,  # Beginning to read columns
  "fields_to": character,  # End of reading columns (inclusive)
  "fields_except": characters array or NULL,  # Which columns not to take
  "fields_percentage": characters array or NULL,  # Columns to calculate percentage
  "save": {  # Diagram of how to save data in JSON format
    "field1": null,  # 1 column is saved in 1 field
    "fi1ld2": ["f1", "f2", "f3"],  # 3 columns are saved in 3 fields
    ...
  }
}
```

## Documents
| FUNCTIONS |      COMMAND       |                          DESCRIPTION                           |                                  INPUT                                  |         OUTPUT          |
| --------- | ------------------ | -------------------------------------------------------------- | ----------------------------------------------------------------------- | ----------------------- |
| **cmd**   | **open**           | Opening an EXCEL file                                          | **type**: ConvertFromExcelToJSON.Commands.open                          | **Workbook**            |
| **cmd**   | **sheet**          | Taking a sheet from an EXCEL file for processing (default 1st) | **type**: ConvertFromExcelToJSON.Commands.sheet<br />**sheet**: integer | **Worksheet**           |
| **cmd**   | **schema**         | Opening a SCHEMA (JSON file)                                   | **type**: ConvertFromExcelToJSON.Commands.schema                        | **None**                |
| **cmd**   | **getData**        | Get data from an EXCEL file                                    | **type**: ConvertFromExcelToJSON.Commands.getData                       | **Array[Dictionaries]** |
| **cmd**   | **load**           | Write data via schema to a new JSON file                       | **type**: ConvertFromExcelToJSON.Commands.load                          | **None**                |
| **cmd**   | **getDataAndLoad** | Running the «getData» and «load» commands                      | **type**: ConvertFromExcelToJSON.Commands.getDataAndLoad                | **None**                |
| **cmd**   | **close**          | Closing an EXCEL file                                          | **type**: ConvertFromExcelToJSON.Commands.close                         | **None**                |

## Tests
- **0** : No exceptions
- **1** : There is an exception (_there are more columns than fields to save data in JSON_)
- **2** : There is an exception (_there are less columns than fields to save data in JSON_)
