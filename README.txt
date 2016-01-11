This program 'excel_wrap_final.py' is designed
to search for excel files within the 'spath' directory
matching the 'glob' parameters.

Then it loops through those files and looks for
a specific sheet name and copies data from specified
cells in the sheet name.  (Currently 'Sheet1' and 'A1:B1')

It then copies that data to a list 'name_scores'.

It creates a workbook in memory, pastes that data into
the first worksheet within the workbook, and then saves
it to the 'final_wb' location.  