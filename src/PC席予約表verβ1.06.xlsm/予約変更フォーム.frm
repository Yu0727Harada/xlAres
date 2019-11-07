VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 予約変更フォーム 
   Caption         =   "予約の変更など"
   ClientHeight    =   6139
   ClientLeft      =   105
   ClientTop       =   448
   ClientWidth     =   5292
   OleObjectBlob   =   "予約変更フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "予約変更フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()

End Sub

Private Sub UserForm_Initialize()
追加ボタン.SetFocus
End Sub

Private Sub 延長ボタン_Click()

If 連続可能か = False Then

'    MsgBox ("延長できません")
    post_confirm.Show
    Unload 予約変更フォーム
    Exit Sub
ElseIf Sheets("メイン").Range(limit_res_on_off).Value = "on" Then
    limit_res_inform.Show
    Unload 予約変更フォーム
    Exit Sub
Else
    Dim 予約コード As Long
    Dim 現在の位置 As Long
    予約コード = resreve_day * 100 + 時間帯 * 10 + 席番号
    現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
    Dim 予約確認 As Integer
    Dim CNT(10) As Integer
    Dim 学籍番号リスト(10) As Variant
    Dim cable_check As Boolean
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    k = 0
    l = 0
    
    j = 0
    While Sheets("生データ").Cells(現在の位置, j + 6) <> ""
            学籍番号リスト(j) = Sheets("生データ").Cells(現在の位置, j + 6)
            j = j + 1
    Wend

If Sheets("生データ").Cells(現在の位置, 5) = 0 Then
    cable_check = False
Else
    cable_check = True
End If

    Call check_res_day
    Call check_res_num(学籍番号リスト(), data_num, CNT())
    
    Dim bl_res_dup_check As Boolean
    bl_res_dup_check = res_duplicate_check(j - 1, 0, CNT())
    If bl_res_dup_check = False Then
        Unload 予約変更フォーム
        Exit Sub
    End If

    Dim bl_res_input As Boolean
    bl_res_input = res_input_rawsheet(resreve_day, 時間帯 + 1, 席番号, cable_check, 学籍番号リスト(), j - 1)
    
    Unload 予約変更フォーム

End If

End Sub

Private Sub 取消ボタン_Click()

Unload 予約変更フォーム
取消フォーム.Show
            

End Sub

Private Sub 貸出ボタン_Click()

Call cable
Unload 予約変更フォーム

End Sub

Private Sub 追加ボタン_Click()

Unload 予約変更フォーム
利用者追加フォーム.Show

End Sub
