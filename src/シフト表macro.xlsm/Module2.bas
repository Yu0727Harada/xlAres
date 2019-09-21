Attribute VB_Name = "Module2"
Option Explicit

Sub Shift_reset()

Dim OkCancel As VbMsgBoxResult
OkCancel = MsgBox("シフト表をリセットしてよろしいですか？", vbOKCancel)
If OkCancel = vbOK Then
    
    Dim i As Integer
    Dim j As Integer
    i = 0
    Do While Cells(i + shift_table_number_start_row, shift_table_number_start_colomn) <> ""
        i = i + 1
    Loop
    j = 0
    Do While Cells(shift_table_date_start_row, shift_table_time_start_colomn + j) <> ""
        j = j + 1
    Loop
    
    With Range("A1").Offset(shift_table_time_start_row - 1, shift_table_time_start_colomn - 1).Resize(i, j)
        .Clear
        .NumberFormatLocal = "@"
    End With
End If

End Sub
