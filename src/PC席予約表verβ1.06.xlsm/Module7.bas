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

Sub selction_move()

    Dim inputpass As String
    inputpass = InputBox("パスコードを入力してください", "パスコードの入力")
    If inputpass = passcord Then
        If Cells(5, 20) = "off" Then
            Cells(5, 20) = "on"
        ElseIf Cells(5, 20) = "on" Then
            Cells(5, 20) = "off"
        Else
            Cells(5, 20) = "off"
        End If
    ElseIf inputpass = "" Then
    
    Else
        MsgBox ("パスコードが一致しません")
    End If


End Sub

Sub refresh_diplicate_sheet()
Worksheets("メイン").EnableCalculation = False
Dim main As Worksheet
Dim Duplicate As Worksheet
Set main = Worksheets("メイン")
Set Duplicate = Worksheets("重複チェック")

Duplicate.Cells(1, 1).Value = 19900101

Call check_res_day
End Sub

Sub show_profile()
Profile.Show

End Sub

