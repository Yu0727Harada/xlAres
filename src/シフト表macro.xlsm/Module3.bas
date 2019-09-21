Attribute VB_Name = "Module3"
Option Explicit

Sub Set_color()

Dim i As Integer
Dim j As Integer
Dim match_row As Integer

j = shift_table_date_start_row + 1
i = shift_table_time_start_colomn

Do While Cells(shift_table_date_start_row, i) <> ""
    If Cells(j, i).Value = "“y" Then
        Cells(j, i).Font.ColorIndex = 5
    ElseIf Cells(j, i).Value = "“ú" Then
        Cells(j, i).Font.ColorIndex = 3
    Else
        Cells(j, i).Font.ColorIndex = 1
    End If
    
    i = i + 1
Loop

j = shift_table_number_start_row
Do While Cells(j, 1) <> "" '
i = shift_table_time_start_colomn
    Do While Cells(shift_table_date_start_row, i) <> ""
            If Cells(j, i) = "" Then
                Cells(j, i).Interior.ColorIndex = 0
            Else
                On Error GoTo error
                    match_row = WorksheetFunction.Match(Cells(j, i), Range("B:B"), 0)
                On Error GoTo 0
                Cells(j, i).Interior.Color = Cells(match_row, 3).Interior.Color
            End If
        i = i + 1
    Loop
j = j + 1
Loop

Exit Sub

error:

match_row = 1
Resume Next
End Sub



