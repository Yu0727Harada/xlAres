Attribute VB_Name = "Module7"
Option Explicit

Sub worksheet_recalculate()
'ワークシート再計算を行うプロシージャ

If Worksheets("メイン").EnableCalculation = False Then
    Worksheets("メイン").EnableCalculation = True
End If
Application.Calculate
End Sub
Sub past_enable_switch()
'マスター入力モードのオンオフプロシージャ

    Dim inputpass As String
    inputpass = InputBox("パスコードを入力してください", "パスコードの入力")
    If inputpass = passcord Then
        If Range(master_on_off).Value = "off" Then
            Range(master_on_off).Value = "on"
        ElseIf Range(master_on_off).Value = "on" Then
            Range(master_on_off).Value = "off"
        Else
            Range(master_on_off).Value = "off"
        End If
    ElseIf inputpass = "" Then
    
    Else
        MsgBox ("パスコードが一致しません")
    End If
End Sub
Sub main_sheet_sort()
'生データをソートするプロシージャ
Call Worksheets("生データ").Range("A:AA").Sort(key1:=Worksheets("生データ").Cells(1, data_sheet.reserve_code).EntireColumn, order1:=xlAscending, Header:=xlYes)

End Sub

Sub selction_move()
'カーソル強制カーソル移動のオンオフを切り替えるプロシージャ

    Dim inputpass As String
    inputpass = InputBox("パスコードを入力してください", "パスコードの入力")
    If inputpass = passcord Then
        If Range(cell_corsor_move).Value = "off" Then
            Range(cell_corsor_move).Value = "on"
        ElseIf Range(cell_corsor_move).Value = "on" Then
            Range(cell_corsor_move).Value = "off"
        Else
            Range(cell_corsor_move).Value = "off"
        End If
    ElseIf inputpass = "" Then
    
    Else
        MsgBox ("パスコードが一致しません")
    End If


End Sub

Sub refresh_diplicate_sheet()
'重複チェックシートを一度削除してもう一度入れなおすプロシージャ

'Worksheets("メイン").EnableCalculation = False
Dim main As Worksheet
Dim duplicate As Worksheet
Set main = Worksheets("メイン")
Set duplicate = Worksheets("重複チェック")

duplicate.Cells(1, 1).Value = 19900101

Call check_res_day
End Sub

Sub show_profile()
Profile.Show

End Sub

Sub limit_res_on_off_pass()

    Dim inputpass As String
    inputpass = InputBox("パスコードを入力してください", "パスコードの入力")
    If inputpass = passcord Then
        If Range(limit_res_on_off).Value = "off" Then
            Range(limit_res_on_off).Value = "on"
        ElseIf Range(limit_res_on_off).Value = "on" Then
            Range(limit_res_on_off).Value = "off"
        Else
            Range(limit_res_on_off).Value = "off"
        End If
    ElseIf inputpass = "" Then
    
    Else
        MsgBox ("パスコードが一致しません")
    End If
    
End Sub

