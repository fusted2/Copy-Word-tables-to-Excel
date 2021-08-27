# Copy-Word-tables-to-Excel

Hi everyone!

This project is created on 27-Aug-2021.

At this time, I am an auditor who begins to learn Python to solve my real-life problems.

# Problem Description

I have Word report files that usually contain around 80+ tables. Sometimes, I have to copy & paste those tables into Excel to re-calculate if the Sum/Total rows are correctly calculated (due to group's mistakes or decimal round up/down,...).

# What do I cover in this project?
## 1. Copy & Paste the content/text of the Word tables

I use --python-docx-- and --openpyxl-- packages to copy all Word tables' content (plain text) into one/multiple Excel sheet(s) in consecutive order and separate them by 3 blank rows.

> pip install python-docx
> 
> pip install openpyxl

## 2. Format the Excel cells (optional)

The Excel cells are formatted as Bold, Italic, Underline ('single') as theirs in Word.

Apply border ('thin') to all Excel cells.

This part is just optional as it may slow down the progress (~4 minutes for 70 formatted tables in my case compared with ~2.5 minutes for only plain text).

## 3. Numeric issue

I have 2 kinds of report files: *(123.456.780,99)* and *(123,456,780.99)*

I convert the content type into (negative) integer/float or string before paste into Excel using --locale-- package.

## 4. Check merged cells

When I use --docx--, the merged cells' content maybe read & write more than once.

I check if the content in the merged cell has been read & written, if yes, skip it (using *._tc* attribute/method (I don't know what it is?)).

# Example

Look at my --example-- folder.

I have a _blankExcelReport.xlsx file_. Please note that sometimes this file may contain some Excel function such as sum, countif,... and I use this file as a template to copy the number from Word tables and update the data for those functions.

The _wordToExcel.py_ will copy the tables from _wordReport.docx_ to the _blankExcelReport.xlsx_ and save as _newExcelReport.xlsx_ (so that the original Excel template won't be messed up).

# What I have tried and fail?

- Format the text's color and the cell's color (Excel).
- "Copy" the font and size of Word tables' content.
- Make the optional options easier to use. For now, I only know to comment/uncomment the code.

# Acknowledgement

Thank you Al Sweigart for teaching me in this course https://www.udemy.com/course/automate/
