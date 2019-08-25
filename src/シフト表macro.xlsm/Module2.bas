Attribute VB_Name = "Module2"
Option Explicit

Sub Shift_reset()

Dim i As Integer

i = 1

Do While Cells(i + 3, 1) <> ""
i = i + 1
Loop

With Range("C4:Q4").Resize(i - 1, 15)
    .Clear
    .NumberFormatLocal = "@"
End With

End Sub
