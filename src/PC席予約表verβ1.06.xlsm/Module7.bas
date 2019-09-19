Attribute VB_Name = "Module7"
Option Explicit

Sub worksheet_recalculate()
If Worksheets("メイン").EnableCalculation = False Then
    Worksheets("メイン").EnableCalculation = True
End If
End Sub
Sub past_enable_switch()
    Dim inputpass As String
    inputpass = InputBox("パスコードを入力してください", "パスコードの入力")
    If inputpass = passcord Then
        If Cells(4, 20) = "off" Then
            Cells(4, 20) = "on"
        ElseIf Cells(4, 20) = "on" Then
            Cells(4, 20) = "off"
        Else
            Cells(4, 20) = "off"
        End If
    ElseIf inputpass = "" Then
    
    Else
        MsgBox ("パスコードが一致しません")
    End If
End Sub
Sub main_sheet_sort()

    Call Worksheets("生データ").Range("A:F").Sort(key1:=Worksheets("生データ").Range("D:D"), order1:=xlAscending, Header:=xlYes)

End Sub
