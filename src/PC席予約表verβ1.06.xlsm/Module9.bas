Attribute VB_Name = "Module9"
Option Explicit

Sub refresh_diplicate_sheet()
Worksheets("メイン").EnableCalculation = False
Dim main As Worksheet
Dim Duplicate As Worksheet
Set main = Worksheets("メイン")
Set Duplicate = Worksheets("重複チェック")

Duplicate.Cells(1, 1).Value = 19900101

Call check_res_day
End Sub
