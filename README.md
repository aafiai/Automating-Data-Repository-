# Automating-Data-Repository-
**Routine SelectAndConvertCSV()**

*The provided VBA code is a macro designed to perform several tasks on a CSV file selected by the user. Below is a detailed explanation of what each part of the code does:*

1. File Selection:

The user is prompted to select a CSV file. This is done using the Application.GetOpenFilename method, which opens a file dialog allowing the user to choose a file with a .csv extension.

If the user cancels the file selection, the macro exits immediately with If csvFile = False Then Exit Sub.

2.  Opening the CSV File:

Once a file is selected, it is opened as a workbook using Workbooks.Open.

The first worksheet of the workbook is accessed and set as ws.

3. Freeze Panes:

The code freezes the top row (row 2) of the worksheet using ActiveWindow.FreezePanes = True. This keeps the row visible while scrolling through the rest of the data.

4. AutoFilter:

AutoFilter is applied to the first row of the worksheet using ws.Rows(1).AutoFilter, which allows filtering of data based on criteria.

5. Cell Selection:

Cell A1 is selected to remove any previous selections or highlights using ws.Range("A1").Select.

6. File Name and Date Formatting:

The file name is extracted without its extension using string manipulation functions (Mid, InStrRev).

The current date is formatted in "8th October 2024" format using Day(Date) and Format(Date, "mmmm yyyy").

7. File Saving:

A new file path is constructed with the original name, current date, and .xlsb extension (a binary Excel format).

The workbook is saved in binary format using wb.SaveAs with FileFormat:=xlExcel12.

8. Closing the Workbook:

The workbook is closed without saving changes again (wb.Close SaveChanges:=False), as the saving operation has already been performed.

9. Notification:

A message box informs the user that the file has been successfully converted and saved with the new path using MsgBox.




This macro automates the process of converting a CSV file into a binary Excel file, applying useful features like freezing panes and autofilter, and appropriately naming the file with the current date included.




**Routine EventDetails()**
*The VBA code you provided is a subroutine named EventDetails(), which performs several operations on a CSV file that the user selects. Here's a breakdown of what the code does:*

1. File Selection: The code prompts the user to select a CSV file. If no file is selected, the subroutine exits.

2. Open CSV File: The selected CSV file is opened as a workbook, and operations are performed on the first worksheet.

3. Freeze Panes: The panes are frozen at cell B2, meaning the top row and first column remain visible when scrolling.

4. Enable Filters: Filters are enabled on the first row to allow easy sorting and filtering of the data.

5. Conditional Formatting: Conditional formatting is applied to highlight duplicate values in column A with a light red fill.

6. Date Adjustments: For columns Y, Z, AA, and AB, the code checks each cell to see if it contains a date. If the year of the date is between 1930 and 1985, it adds 100 years to the year component of the date and updates the cell with the new date.

7. File Name Construction: The original file name (without extension) is extracted, and the current date is formatted as "8th October 2024".

8. Save As XLSB: The workbook is saved as a binary file (.xlsb) with the same name, appended by the current date.

9. Close Workbook: After saving, the workbook is closed without saving any further changes.

10. Notification: A message box informs the user that the file is ready, with all the specified changes applied.



This code is useful for processing event details stored in CSV format, applying specific formatting and data transformations, and saving the results in a more efficient binary format.



**Routine EventDetailsLITE()**

*The VBA code is designed to process and transform data from a CSV file using Excel. Below is a step-by-step explanation of what the code accomplishes:*

1. File Selection:

The user is prompted to select a CSV file to process. If no file is selected, the procedure exits.

2. CSV File Opening:

The selected CSV file is opened, and its first worksheet is accessed for further processing.

3. Freeze Panes:

The top row and the first column are frozen at cell B2 for easier navigation and viewing within the worksheet.

4. Enable Filters:

Filters are applied to the first row of the worksheet to facilitate data analysis.

5. Conditional Formatting:

Conditional formatting is applied to highlight duplicate values in Column A with a light red fill color.

6. Remove Rows:

Rows are removed if Column U contains the value "Cancelled."

Rows are also removed if Column AV does not contain the value "Core."

7. Date Adjustment:

Dates in Columns Y, Z, AA, and AB are adjusted. If a date's year is between 1930 and 1985, it is increased by 100 years (e.g., 1950 becomes 2050).

8. File Naming and Saving:

The original CSV filename is extracted (without its extension).

The current date is formatted and appended to the filename.

The processed data is saved as an Excel Binary Workbook (.xlsb) with a new name that includes the date and the tag "Light."

The workbook is then closed without saving changes again.

9. Notification:

A message box informs the user that the processing is complete and describes the rows that were removed.




This script automates data transformation tasks and ensures that the output file is cleaned and saved in a specified format for further use or analysis.




**Routine SortCSVDescending()**

*The VBA macro is designed to automate the process of selecting a CSV file, sorting its data, and saving it as an Excel Binary Workbook (.xlsb) with a date-stamped file name. Below is a breakdown of its functionality:*

1. File Selection:

The macro prompts the user to select a CSV file using Application.GetOpenFilename.

2. CSV File Handling:

If a file is selected, it opens the CSV in Excel and sets the first sheet as the active worksheet.

3. Determine Last Row:

It calculates the last row in column A using ws.Cells(ws.Rows.Count, "A").End(xlUp).Row.

4. Freeze Panes and Enable Filters:

The macro freezes panes at the second row and column, and applies filters to the first row.

5. Sort Data:

It sorts the data in column A in descending order.

6. File Naming:

Extracts the file name without extension and constructs a name that includes the current date in the format "8th October 2024".

7. Save As Excel Binary Workbook:

Saves the sorted data as an .xlsb file with the newly constructed name.

8. User Notification:

Displays a message box confirming the operation and showing the path of the saved file.





**Routine NEWRECORDSCompareColumnAAndExtractAdditions()**
*This VBA macro is designed to compare two Excel files, specifically focusing on the values in Column A. The goal is to identify and extract entries from today's file (`todaysFile`) that are not present in yesterday's file (`yesterdaysFile`). The extracted entries are then saved in a new Excel workbook.*

 

Here's a brief overview of how the macro operates:

1. File Selection: Prompts the user to select today's and yesterday's files using a file dialog.

 

2. Open Files: Opens both Excel files and identifies the last row and column in today's file and the last row in yesterday's file.

 

3. Store Yesterday's Data: Collects unique values from Column A in yesterday's file using a VBA `Collection`.

 

4. Create New Workbook: Initializes a new workbook to store the results with a sheet named "Additions".

 

5. Header Copy: Copies the header row from today's file to the new workbook.

 

6. Comparison: Iterates through today's Column A, checking for entries not found in yesterday's data, and copies these rows to the new workbook.

 

7. Close Source Files: Closes both source files without saving changes.

 

8. Save New Workbook: Saves the new workbook with a filename format including "NEW" and the current date.

 

9. User Notification: Displays a message box to inform the user that the comparison is complete.











 
