Attribute VB_Name = "Module1"
Option Explicit

Sub Obtain_ShiftTime()
Dim i As Integer
Dim j As Integer
Dim search_row As Integer


For j = 4 To 16
    For i = 3 To 17
        
        If Cells(j, i) <> "" Then
            On Error GoTo error
                search_row = WorksheetFunction.Match(Cells(j, i), Range("B21:B40"), 0)
            On Error GoTo 0
        End If
    Next i
Next j
    
    
Exit Sub

error:
Dim LastRow As Integer
LastRow = Cells(Rows.Count, 2).End(xlUp).Row + 1
Cells(LastRow, 2) = Cells(j, i)
Resume Next



End Sub
                    
