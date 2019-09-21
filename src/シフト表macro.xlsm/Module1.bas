Attribute VB_Name = "Module1"
Option Explicit
Public shift_table_number_start_row As Integer
Public shift_table_number_start_colomn As Integer
Public shift_table_time_start_row As Integer
Public shift_table_time_start_colomn As Integer
Public shift_table_date_start_row As Integer

Sub Obtain_ShiftTime()
Dim i As Integer
Dim j As Integer
Dim search_row As Integer

j = shift_table_number_start_row


Do While Cells(j, shift_table_number_start_colomn) <> ""
'For j = 4 To 16
    i = shift_table_time_start_colomn
    Do While Cells(shift_table_date_start_row, i) <> ""
        
        If Cells(j, i) <> "" Then
            On Error GoTo error
                search_row = WorksheetFunction.Match(Cells(j, i), Range("B:B"), 0)
            On Error GoTo 0
        End If
        i = i + 1
    Loop
'Next j
j = j + 1
Loop
    
    
Exit Sub

error:
Dim LastRow As Integer
LastRow = Cells(Rows.Count, 2).End(xlUp).Row + 1
Cells(LastRow, 2) = Cells(j, i)
Resume Next



End Sub
                    
