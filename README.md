# Automating-Data-Repository-
1. CSV File Processing and Conversion
Task: Write a VBA macro to open a CSV file, apply some operations (e.g., filling blank cells with formulas), and then save it as an Excel binary file.
Details: We handled tasks like:
Freezing panes and enabling filters on the CSV file.
Applying a formula to empty cells in the last column of the CSV.
Saving the file as an .xlsb format.
Issues: We encountered some problems where the original code didn't work as expected, leading to some modifications in the approach.
Outcome: We simplified the task to focus on the formula filling part, but this was later deferred.
2. Compare Today's and Yesterday's Data (Column A)
Task: Write a VBA macro to compare Column A from two different files (today's and yesterday's), and extract the new records from today’s file that don’t exist in yesterday’s file.
Details: The task involved:
Comparing the unique values in Column A of both files.
Creating a new workbook and copying rows from today's file that are not present in yesterday's file into this new workbook.
Storing unique values from yesterday’s file in a collection.
Outcome: The macro successfully extracted and saved the new rows in a new workbook, and the user was prompted to name the new file.
3. Save New Workbook with "NEW" Prefix
Task: Modify the previously written macro to save the new workbook with "NEW_" added to the name of today’s file, followed by "_Additions" and the current date.
Outcome: We successfully updated the code to save the new file with the desired naming convention (e.g., NEW_TodaysFile_Additions_dd-mmm-yyyy.xlsx).

Key Highlights:
Error Handling: We used error handling (On Error Resume Next) to ensure no interruptions occurred during data processing, especially when checking the uniqueness of values in Column A.
Workbook Operations: Operations like copying rows, freezing panes, and writing formulas were included.
File Saving: Focused on saving the new file with a clear and structured naming convention, ensuring that the user could easily identify and differentiate files.
