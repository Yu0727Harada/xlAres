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

Sub formulabar_display()

If Application.DisplayFormulaBar = True Then
    Application.DisplayFormulaBar = False
Else
    Application.DisplayFormulaBar = True
End If

End Sub

Sub headings_display()

If ActiveWindow.DisplayHeadings = True Then
    ActiveWindow.DisplayHeadings = False
Else
    ActiveWindow.DisplayHeadings = True
End If

End Sub

Sub statusbar_display()
If Application.DisplayStatusBar = True Then
    Application.DisplayStatusBar = False
Else
    Application.DisplayStatusBar = True
End If

End Sub

Sub scrollbar_display()
Dim inputpass As String
    inputpass = InputBox("パスコードを入力してください", "パスコードの入力")
    If inputpass = passcord Then
        If Application.DisplayScrollBars = True Then
            Application.DisplayScrollBars = False
        Else
            Application.DisplayScrollBars = True
        End If
   ElseIf inputpass = "" Then
    
    Else
        MsgBox ("パスコードが一致しません")
    End If
    
End Sub

Sub tabs_display()

If ActiveWindow.DisplayWorkbookTabs = True Then
    ActiveWindow.DisplayWorkbookTabs = False
Else
    ActiveWindow.DisplayWorkbookTabs = True
End If

End Sub

Sub un_protect()
Call Sheets("メイン").Unprotect
End Sub

Sub protect()
Call Sheets("メイン").protect(UserInterfaceOnly:=True)
End Sub

Sub ribbon_display()

Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"

End Sub

Sub ribbon_undisplay()

Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"

End Sub

Sub vbe_open()
Application.VBE.Windows(1).SetFocus
End Sub
Public Sub input_new_passcord()

Dim pass_yesno As Integer

pass_yesno = passcord_inputform
If pass_yesno = 2 Then
    Exit Sub
ElseIf pass_yesno = 1 Then
    Exit Sub
Else

    Dim input_new_pass As String
    input_new_pass = InputBox("パスコードとして追加する学籍番号を入力するか、学生証のバーコードをスキャンしてください。すでに登録されている番号を入力することで削除することもできます。")
    Dim trans_input_new_pass As Variant
    trans_input_new_pass = translate_number(input_new_pass)
    If trans_input_new_pass = "" Then
        Exit Sub
    ElseIf Int(trans_input_new_pass) = -1 Then
        MsgBox ("有効な学籍番号ではありません")
        Exit Sub
    Else
        Dim search As Integer
        On Error GoTo match_error
        search = WorksheetFunction.Match(Int(trans_input_new_pass), Sheets("passcord").Range("A:A"), 1)
        On Error GoTo 0
        If Int(trans_input_new_pass) = WorksheetFunction.Index(Sheets("passcord").Range("A:A"), search) Then
            Dim delete_yesno As String
            delete_yesno = MsgBox("この番号はすでに登録されています。この番号を削除しますか？", vbYesNo + vbQuestion, "番号の削除の確認")
            If delete_yesno = vbNo Then
                Exit Sub
            Else
                Call Sheets("passcord").Cells(search, 1).EntireRow.Delete(xlShiftUp)
                Exit Sub
            End If
        Else
            Sheets("passcord").Rows(search + 1).Insert
            Sheets("passcord").Cells(search + 1, 1).Value = trans_input_new_pass
        End If
    End If

End If

Exit Sub
match_error:
search = 0
Sheets("passcord").Rows(search + 1).Insert
Sheets("passcord").Cells(search + 1, 1).Value = trans_input_new_pass
End Sub
