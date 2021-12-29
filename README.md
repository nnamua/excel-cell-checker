# excel-cell-checker

## Description

This tool checks a given `.xlsx` file has the structure specified in a `.json` file.

## Requirements

Python 3 is required, at least 3.7.
The required modules can be installed with:

```
$ pip install -r requirements.txt
```

## Usage

First, you must create a `.json` file containing the structure of your excel file. The root of the structure is an array with the key `"cols"`:

```json
{
  "cols" : [
    {
      "name" : "id",
      "type" : "string",
      "regex" : "[0-9]{5}",
      "non-null" : true
    },
    {
      "name" : "first_name",
      "type" : "string"
    },
    {
      "name" : "age",
      "type" : "number"
    }
  ]
}
```

The elements of the `cols` array are the columns of your excel file, aswell as their respective data type. The currently supported data types are `string`, `number` and `date`.

The tool can also optionally check the content of cells, but right now this feature is limited to regular expressions for `string` columns.

Run `checker.py` and supply a `.xlsx` file aswell as a `.json` structure file:

```
$ py checker.py <excel file> <structure file>
```

If you want to check a specific sheet in your excel file, supply the sheet name using `-s <sheetname>`.

First, the tool will check if the excel file contains the same rows as specified in the `.json` structure (it is assumed, that the first row contains column names and all remaining rows contain data). If this is succesful, each cells type (and content) will be examined.
If you don't want a column to be checked, you can specifiy `skip` in your structure file:
```json
{
  "name" : "useless"
  "type" : "string"
  "skip" : true
}
```

After examining the excel sheet, a summary of all found violations is printed.
This summary can be modified by the following parameters:
  * `--hide-skipped` Hides skipped columns
  * `--hide-ok` Hides columns with no violations

## Examples

Example source files can be found in the `examples` directory.

Running the tools on these files should yield:

```
$ py .\checker.py .\examples\example.xlsx .\examples\structure.json
Loading structure file structure.json ..
Loading excel file example.xlsx ..
Loaded file with 5 data rows.
Checking basic column structure ..   Done!
Checking row 5 of 5 ..
Done!

> id
[ERROR] : 2 violations found

  The following cells did not match the regular expression:

      Row  Value
    -----  -------
        5  '42'

  The following cells are empty, even though non-null is set to true:

      Row
    -----
        4

> first_name
[OK] : No violations found

> age
[ERROR] : 1 violations found
  The following cells did not match the expected type (number) :

      Row  Value    Type
    -----  -------  ------
        6  '17'     str



> useless
[SKIPPED]
```