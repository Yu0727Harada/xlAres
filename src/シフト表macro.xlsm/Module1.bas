Attribute VB_Name = "Module1"
Option Explicit

Sub printbottom()
Dim i As Integer
Dim j As Integer
Dim search_row As Integer

Range("B21:B50").Clear


For j = 3 To 16
    For i = 3 To 17
        
        If j = 3 Then
            If Cells(j, i).Value = "“y" Then
                Cells(j, i).Font.ColorIndex = 5
            ElseIf Cells(j, i).Value = "“ú" Then
                Cells(j, i).Font.ColorIndex = 3
            Else
                Cells(j, i).Font.ColorIndex = 1
            End If
        Else
            If Cells(j, i) <> "" Then
                On Error GoTo error
                    search_row = WorksheetFunction.Match(Cells(j, i), Range("B21:B40"), 0)
                On Error GoTo 0
            End If
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
                    
