# excel-table-to-word

Import data from an Excel table into multiple tables in Word.

## Description

Quick and dirty module that loops through all the paragraphs in a Word document, filtering for paragraphs with a certain heading level value that has a table following it. 

The text of that heading is used as the name of the group. 

Any existing rows in the table are removed except for the heading. 

Each row in the Excel table that has the same group name is then appended to the table in Word.

## Notes

* Cells are imported as text strings - number formats from Excel are not honoured.
* If any cells in Excel have errors this will probably cause an exception.
* The groupname column in Excel must not be between the data columns that will be copied across.

## TODO

* Add Application.ScreenRefresh and DoEvents for a minor speed improvement
* Add proper handling of error cells in Excel
* Consider checking VarType() and NumberFormatting of Excel data array to apply the number formatting