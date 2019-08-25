Attribute VB_Name = "Module2"
Option Explicit

Sub colour()

Dim i As Integer
Dim j As Integer
Dim match_row As Integer

i = 3
j = 4

    For j = 4 To 16
        For i = 3 To 17
            If Cells(j, i) = "" Then
                Cells(j, i).Interior.ColorIndex = 0
            Else
                On Error GoTo error
                match_row = WorksheetFunction.Match(Cells(j, i), Range("B21:B40"), 0)
                On Error GoTo 0
                Cells(j, i).Interior.Color = Cells(match_row + 20, 3).Interior.Color
            End If
        Next i
    Next j

Exit Sub

error:

match_row = 0
Resume Next
End Sub

