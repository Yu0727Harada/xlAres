VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 取消フォーム 
   Caption         =   "予約の取消の確認"
   ClientHeight    =   5999
   ClientLeft      =   91
   ClientTop       =   420
   ClientWidth     =   7217
   OleObjectBlob   =   "取消フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "取消フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim 予約コード As Long
Dim 現在の位置 As Long
Dim confirm_number As Variant
Set duplicate = Worksheets("重複チェック")
Dim search_stu_row
Dim i As Integer
Dim target_stu_list(10) As Variant

予約コード = resreve_day * 100 + 時間帯 * 10 + 席番号
現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
If 予約コード = WorksheetFunction.Index(Sheets("生データ").Range("D:D"), 現在の位置) Then
    
    For i = 0 To 10
        If Sheets("生データ").Cells(現在の位置, i + 6).Value = "" Then
            Exit For
        End If
        target_stu_list(i) = Sheets("生データ").Cells(現在の位置, i + 6).Value
    Next i
    If TextBox1 = passcord Then
        Call Sheets("生データ").Cells(現在の位置, 1).EntireRow.Delete(xlShiftUp)
        Call delete_res_num(target_stu_list, i - 1)
'        MsgBox ("予約を取り消しました")
        delete_confirm.Show
        Unload 取消フォーム
        Exit Sub
    End If
    
    'テキストボックスを変換する前にパスコードが入力されているかを場合分けする
    
'    Call textbox_restrict(TextBox1, 確認番号)
    confirm_number = translate_number(TextBox1, 1)
    '入力された番号を変換する
    
    If confirm_number = "" Then
        MsgBox ("学籍番号を入力してください")
        Exit Sub
    ElseIf confirm_number = -1 Then
        TextBox1 = ""
        Exit Sub
    End If
    If number_valid <> 0 Then
        number_valid = 0
        Exit Sub
    End If
    
    
'    If 確認番号 = Sheets("生データ").Cells(現在の位置, i).Value Then

    Dim result_list As Variant
    result_list = Filter(target_stu_list, Int(confirm_number))
    If UBound(result_list) <> -1 Then
        Call Sheets("生データ").Cells(現在の位置, 1).EntireRow.Delete(xlShiftUp)
        '        予約したデータを削除
        Call delete_res_num(target_stu_list, i - 1)
        Worksheets("メイン").EnableCalculation = True
        Application.Calculate
'        MsgBox ("予約を取り消しました")
        delete_confirm.Show
        Unload 取消フォーム
        Exit Sub
    End If
    
    Dim search As Integer
    On Error GoTo error_nothing
    search = WorksheetFunction.Match(Int(confirm_number), Sheets("passcord").Cells(1, 1).EntireColumn, 1)
    On Error GoTo 0
    If Int(confirm_number) = WorksheetFunction.Index(Sheets("passcord").Cells(1, 1).EntireColumn, search) Then
        Call Sheets("生データ").Cells(現在の位置, 1).EntireRow.Delete(xlShiftUp)
        '        予約したデータを削除
        Call delete_res_num(target_stu_list, i - 1)
        Worksheets("メイン").EnableCalculation = True
        Application.Calculate
'        MsgBox ("予約を取り消しました")
        delete_confirm.Show
        Unload 取消フォーム
        Exit Sub
    End If
    
    '何事もなくループをでたら入力した内容を削除
    MsgBox ("学籍番号が一致しません。もう一度入力してください")
    TextBox1 = ""

Else
    MsgBox ("予約がありません。LAに確認を依頼してください。エラー番号１０３")
End If

Exit Sub

error_nothing:
search = 1
Resume Next

End Sub


Private Sub Label2_Click()

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Call keypressrestrict(KeyAscii)

End Sub

Private Sub UserForm_Click()

End Sub

