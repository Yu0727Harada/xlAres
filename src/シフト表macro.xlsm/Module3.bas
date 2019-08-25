Attribute VB_Name = "Module3"
Option Explicit

Sub Set_color()

Dim i As Integer
Dim j As Integer
Dim match_row As Integer

i = 3
j = 4

Do While Cells(j, 1) <> ""
'    For j = 3 To 16
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
                    If Cells(j, i) = "" Then
                        Cells(j, i).Interior.ColorIndex = 0
                    Else
                        On Error GoTo error
                        match_row = WorksheetFunction.Match(Cells(j, i), Range("B:B"), 0)
                        On Error GoTo 0
                        Cells(j, i).Interior.Color = Cells(match_row, 3).Interior.Color
                    End If
                End If
        Next i
'    Next j
j = j + 1
Loop

Exit Sub

error:

match_row = 0
Resume Next
End Sub



