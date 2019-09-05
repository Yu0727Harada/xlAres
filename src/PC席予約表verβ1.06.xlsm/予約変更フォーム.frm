VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 予約変更フォーム 
   Caption         =   "予約の変更など"
   ClientHeight    =   4200
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   4501
   OleObjectBlob   =   "予約変更フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "予約変更フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
追加ボタン.SetFocus
End Sub

Private Sub 延長ボタン_Click()

If 連続可能か = False Then

'    MsgBox ("延長できません")
    post_confirm.Show
    Unload 予約変更フォーム
    Exit Sub

Else
    Worksheets("メイン").EnableCalculation = False
    Dim 予約コード As Long
    Dim 現在の位置 As Long
    予約コード = 予約日 * 100 + 時間帯 * 10 + 席番号
    現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
    Dim 予約確認 As Integer
    Dim CNT(10) As Integer
    Dim 学籍番号リスト(10) As Variant
    Dim cable_check As Integer
    Dim j As Integer
    Dim k As Integer
    Dim L As Integer
    j = 0
    k = 0
    L = 0

'    Dim 複数人表示参照 As Worksheet
'    Set 複数人表示参照 = Worksheets("複数人表示参照")

    
    While Sheets("生データ").Cells(現在の位置, j + 6) <> ""
            学籍番号リスト(j) = Sheets("生データ").Cells(現在の位置, j + 6)
            j = j + 1
    Wend

    cable_check = Sheets("生データ").Cells(現在の位置, 5)

    Call check_res_day
    Call check_res_num(学籍番号リスト(), data_num, CNT())

        Do While Sheets("生データ").Cells(現在の位置, k + 6) <> ""
            If CNT(k) >= 2 And 予約確認 = 0 Then
                予約確認 = MsgBox("既に２コマ以上予約していますが、予約してよろしいですか？", vbYesNo + vbwuestion, "予約の確認")
                    If 予約確認 = vbNo Then
                        Worksheets("メイン").EnableCalculation = True
                        Unload 予約変更フォーム
                        Exit Sub
                    End If
            End If
            k = k + 1
        Loop

    予約コード = 予約日 * 100 + (時間帯 + 1) * 10 + 席番号
    現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
    Sheets("生データ").Rows(現在の位置 + 1).Insert

'    LastRow = Sheets("生データ").Cells(Rows.Count, 1).End(xlUp).Row + 1
    Sheets("生データ").Cells(現在の位置 + 1, 1).Value = 予約日
    Sheets("生データ").Cells(現在の位置 + 1, 2).Value = 時間帯 + 1
    Sheets("生データ").Cells(現在の位置 + 1, 3).Value = 席番号
    Sheets("生データ").Cells(現在の位置 + 1, 4).Value = 予約コード
    Sheets("生データ").Cells(現在の位置 + 1, 5).Value = cable_check

    
        For L = 0 To k - 1
            Sheets("生データ").Cells(現在の位置 + 1, 6 + L).Value = 学籍番号リスト(L)
        Next L
    Call input_res_num(学籍番号リスト(), k - 1)
    Worksheets("メイン").EnableCalculation = True
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
