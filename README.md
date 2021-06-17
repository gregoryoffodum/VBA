# VBA SHEET AUTOMATION
Automation using VBA programming

## PROBLEM STATEMENT
- Six (6) workbooks (Addis Ababa, Havana, Kano, Lagos, Oslo, Tokyo) represent different subsidiaries.
- Each workbook consisting of five (5) columns each depict name, staff ID, role, and salary information
- Automation using VBA is required to summarize and export vital information from each sheet


## OBJECTIVES
- Populate the sum salaries of individual roles for every subsidiary
- Generate a pdf summary for each workbook/subsidiary
- Generate a general sum total report of all spreadsheets

## Folder contains: 
- "General Report" contains the VBA codes 
- 6 Original workbooks: Addis Ababa, Havana, Kano, Lagos, Oslo, Tokyo (excluding summary report)

A total of 14 files were created, excel and pdf reports (12) for each subsidiary and general summary (2):

- 6 Automated workbooks: Addis Ababa, Havana, Kano, Lagos, Oslo, Tokyo (including summary report)
- 6 Automated individual pdf reports: Addis Ababa, Havana, Kano, Lagos, Oslo, Tokyo 
- 2 general summary reports

Private Sub CommandButton1_Click()

### Declare variables accordingy

Dim vwbsubsidiary As Workbook
Dim vwssubsheet As Worksheet
Dim vstrpath As String
Dim file As String
Dim i As Integer

### Declare file path

vstrpath = "C:\Users\GREGORY\Downloads\ACADEMY FINAL PROJECT\"

### Utilize a FOR loop throughout the code for dynamic automation
*‘file’ variable is the names of workbooks to be iterated located in the cells - 'i' of the current worksheet*


For i = 2 To 7
file = ThisWorkbook.Sheets(1).Cells(i, 1).Value
vstrpath = "C:\Users\GREGORY\Downloads\ACADEMY FINAL PROJECT\"

### Set variables to open file in file path

    Set vwbsubsidiary = Workbooks.Open(vstrpath & file & ".xlsx")
    Set vwssubsheet = vwbsubsidiary.Worksheets(file)
    
### Locate/create cells for labels and computed sum values

vwssubsheet.Range("G1").Value = file & " SUBSIDIARY " & " SUMMARY "
vwssubsheet.Range("H2").Value = "SALARY" & " PER " & " ROLE "
vwssubsheet.Range("H2").Font.Bold = True


vwssubsheet.Range("G2").Value = "ROLE"
vwssubsheet.Range("G2").Font.Bold = True

vwssubsheet.Range("G3").Value = "ASSOCIATE"
vwssubsheet.Range("G3").Font.Bold = True

vwssubsheet.Range("G5").Value = "MANAGER"
vwssubsheet.Range("G5").Font.Bold = True

vwssubsheet.Range("G7").Value = "SENIOR"
vwssubsheet.Range("G7").Font.Bold = True

vwssubsheet.Range("G9").Value = "TOTAL" & " SUM "
vwssubsheet.Range("G9").Font.Bold = True

### Append values to the created labels and cells

vwssubsheet.Range("H3").Value = Application.WorksheetFunction.SumIf(Range("D2").EntireColumn, "ASSOCIATE", Range("E2").EntireColumn)
vwssubsheet.Range("H5").Value = Application.WorksheetFunction.SumIf(Range("D2").EntireColumn, "MANAGER", Range("E2").EntireColumn)
vwssubsheet.Range("H7").Value = Application.WorksheetFunction.SumIf(Range("D2").EntireColumn, "SENIOR", Range("E2").EntireColumn)
vwssubsheet.Range("H9").Value = Application.WorksheetFunction.Sum(Range("E2").EntireColumn)

Application.ScreenUpdating = True

### Autofit and center align columns

vwssubsheet.Cells.EntireColumn.AutoFit
Columns("H").HorizontalAlignment = xlCenter

### Export PDF for each ‘file’ accordingly


vwssubsheet.Range("G1:H9").ExportAsFixedFormat xlTypePDF, vstrpath & file & "_REPORT" & ".pdf"

### Copy and paste sum information for each subsidiary and sum cummulatively after each iteration

vwbsubsidiary.Sheets(file).Range("H3:H9").Copy
ThisWorkbook.Sheets(1).Range("H3:H9").PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks:=True, Transpose:=False

### Locate/create cells for labels and computed sum values


ThisWorkbook.Sheets(1).Range("G1").Value = "GENERAL REPORT PAGE"
ThisWorkbook.Sheets(1).Range("G1").Font.Bold = True
ThisWorkbook.Sheets(1).Range("H1").Value = "TOTAL SALARY"
ThisWorkbook.Sheets(1).Range("H1").Font.Bold = True
    
ThisWorkbook.Sheets(1).Range("G3").Value = "TOTAL ASSOCIATE SALARY"
ThisWorkbook.Sheets(1).Range("G3").Font.Bold = True
ThisWorkbook.Sheets(1).Range("G5").Value = "TOTAL MANAGER SALARY"
ThisWorkbook.Sheets(1).Range("G5").Font.Bold = True
ThisWorkbook.Sheets(1).Range("G7").Value = "TOTAL SENIOR SALARY"
ThisWorkbook.Sheets(1).Range("G7").Font.Bold = True
ThisWorkbook.Sheets(1).Range("G9").Value = "GRAND TOTAL SALARY"
ThisWorkbook.Sheets(1).Range("G9").Font.Bold = True

### Autofit and center align columns

ThisWorkbook.Sheets(1).Cells.EntireColumn.AutoFit
Columns("H").HorizontalAlignment = xlCenter

### Export grand total summary to PDF

ThisWorkbook.Sheets(1).Range("G1:H10").ExportAsFixedFormat xlTypePDF, "C:\Users\GREGORY\Downloads\ACADEMY FINAL PROJECT\GENERAL REPORT SHEET.pdf"


Next i

End Sub
