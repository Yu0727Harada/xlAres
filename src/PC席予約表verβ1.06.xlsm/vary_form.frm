VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vary_form 
   Caption         =   "予約の変更など"
   ClientHeight    =   8352.001
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   5160
   OleObjectBlob   =   "vary_form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "vary_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub delete_use_button_Click()

Unload vary_form
delete_user.Show

End Sub


Private Sub UserForm_Initialize()

Dim targt_time_text As String

If (時間帯 = 1) Then
    target_time_text = "1時限目"
ElseIf (時間帯 = 2) Then
    target_time_text = "2時限目"
ElseIf (時間帯 = 3) Then
    target_time_text = "昼休み"
ElseIf (時間帯 = 4) Then
    target_time_text = "3時限目"
ElseIf (時間帯 = 5) Then
    target_time_text = "4時限目"
ElseIf (時間帯 = 6) Then
    target_time_text = "5時限目"
ElseIf (時間帯 = 7) Then
    target_time_text = "閉室まで"
End If

Dim target_pc_text As String

If (席番号 = 1) Then
    target_pc_text = "１番のPC席"
ElseIf (席番号 = 2) Then
    target_pc_text = "2番のPC席"
ElseIf (席番号 = 3) Then
    target_pc_text = "3番のPC席"
ElseIf (席番号 = 4) Then
    target_pc_text = "4番のPC席"
ElseIf (席番号 = 5) Then
    target_pc_text = "5番のPC席"
End If


Label3 = target_time_text + "の" + target_pc_text + "の" + vbCrLf + "予約変更"



追加ボタン.SetFocus


End Sub

Private Sub 延長ボタン_Click()

If 連続可能か = False Then

'    MsgBox ("延長できません")
    post_confirm.Show
    Unload vary_form
    Exit Sub
ElseIf Sheets("メイン").Range(limit_res_on_off).Value = "on" Then
    limit_res_inform.Show
    Unload vary_form
    Exit Sub
Else
    Dim 予約コード As Long
    Dim 現在の位置 As Long
    予約コード = resreve_day * 100 + 時間帯 * 10 + 席番号
    現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
    Dim 予約確認 As Integer
    Dim CNT(10) As Integer
    Dim student_number_list(10) As Variant
    Dim cable_check As Boolean
    Dim j As Integer
    Dim k As Integer
    Dim L As Integer
    k = 0
    L = 0
    
    j = 0
    While Sheets("生データ").Cells(現在の位置, j + 6) <> ""
            student_number_list(j) = Sheets("生データ").Cells(現在の位置, j + 6)
            j = j + 1
    Wend

If Sheets("生データ").Cells(現在の位置, 5) = 0 Then
    cable_check = False
Else
    cable_check = True
End If

    Call check_res_day
    Call check_res_num(student_number_list(), data_num, CNT())
    
    Dim bl_res_dup_check As Boolean
    bl_res_dup_check = res_duplicate_check(j - 1, 0, CNT())
    If bl_res_dup_check = False Then
        Unload vary_form
        Exit Sub
    End If

    Dim bl_res_input As Boolean
    bl_res_input = res_input_rawsheet(resreve_day, 時間帯 + 1, 席番号, cable_check, student_number_list(), j - 1)
    
    Unload vary_form

End If

End Sub

Private Sub 退席ボタン_Click()

Unload vary_form
Call leave
            

End Sub

Private Sub 貸出ボタン_Click()

Call cable
Unload vary_form

End Sub

Private Sub 追加ボタン_Click()

Unload vary_form
利用者追加フォーム.Show

End Sub

