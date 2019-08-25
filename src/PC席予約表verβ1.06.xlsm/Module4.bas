Attribute VB_Name = "Module4"
Option Explicit

Sub import_Shift()

Application.Calculation = xlCalculationManual

Dim wb As Workbook
Set wb = ThisWorkbook

Dim Open_Filepath As String
Open_Filepath = Application.GetOpenFilename
Workbooks.Open Open_Filepath

Dim Shift_Filename As String
Shift_Filename = Dir(Open_Filepath)

Dim Shift_BookName As Workbook

Set Shift_BookName = Workbooks(Shift_Filename)
Application.Calculation = xlCalculationAutomatic
Shift_BookName.Sheet1.Cells(5, 3) = 3
wb.Sheet8.Cells(2, 1) = 1
wb.Sheet8.Cells(3, 1) = Shift_BookName.Sheet1.Cells(4, 3)

Application.Calculation = xlCalculationAutomatic


End Sub
