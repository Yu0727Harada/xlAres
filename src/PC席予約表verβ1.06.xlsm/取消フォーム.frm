VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 取消フォーム 
   Caption         =   "予約の取消の確認"
   ClientHeight    =   6006
   ClientLeft      =   91
   ClientTop       =   413
   ClientWidth     =   7217
   OleObjectBlob   =   "取消フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "取消フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Private Sub CommandButton1_Click()
'
'Worksheets("メイン").EnableCalculation = False
'
'Dim 予約コード As Long
'Dim 現在の位置 As Long
'Dim 確認番号 As String
'
'予約コード = 予約日 * 100 + 時間帯 * 10 + 席番号
'現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
'If 予約コード = WorksheetFunction.Index(Sheets("生データ").Range("D:D"), 現在の位置) Then
'
'    If TextBox1 = passcord Then
'        Call Sheets("生データ").Cells(現在の位置, 1).EntireRow.Delete(xlShiftUp)
'        MsgBox ("予約を取り消しました")
'        Worksheets("メイン").EnableCalculation = True
'        Unload 取消フォーム
'        Exit Sub
'    End If
'
'    'テキストボックスを変換する前にパスコードが入力されているかを場合分けする
'
'    Call textbox_restrict(TextBox1, 確認番号)
'    '入力された番号を変換する
'
'    If 確認番号 = "" Then
'        MsgBox ("学籍番号を入力してください")
'        Worksheets("メイン").EnableCalculation = True
'        Exit Sub
'    End If
'    If number_valid <> 0 Then
'        number_valid = 0
'        Worksheets("メイン").EnableCalculation = True
'        Exit Sub
'    End If
'
'    Dim i As Integer
'    For i = 6 To 10
'    'とりあえず適当に1０まで回してる
'    '    LastRow = Sheets("生データ").Cells(Rows.Count, 1).End(xlUp).Row + 1
'    '    現在の最終行を取得
'
'    '    Sheets("生データ").Cells(LastRow, 6).Formula = "=mid(indirect(address(" + Str(現在の位置) + "," + Str(i) + ")),1,8)"
'    '    Sheets("生データ").Cells(LastRow, 7).Formula = "=mid(indirect(address(" + Str(現在の位置) + "," + Str(i) + ")),9,16)"
'    ''    最終行に該当の予約の学籍番号を２つに分けるための式を入れる
'    '    '桁数が多すぎて普通に取得するとE＋形式になってしまい、stringに変換したりしてもうまくいかなかった。ワークシート上でmid関数を使えば普通に取得できたので苦肉の策としてこうしてます。
'    '    k = Sheets("生データ").Cells(LastRow, 6).Value
'    '    l = Sheets("生データ").Cells(LastRow, 7).Value
'    ''    最終行の計算結果をk、ｌに格納
'    '
'    '    If k = Mid(確認番号, 1, 8) And l = Mid(確認番号, 9) Then
'    '            Sheets("生データ").Cells(LastRow, 6).Clear
'    '            Sheets("生データ").Cells(LastRow, 7).Clear
'    '            入力した式を削除する
'        If Sheets("生データ").Cells(現在の位置, i).Value = "" Then
'            Exit For
'        End If
'        If 確認番号 = Sheets("生データ").Cells(現在の位置, i).Value Then
'            Call Sheets("生データ").Cells(現在の位置, 1).EntireRow.Delete(xlShiftUp)
'    '        予約したデータを削除
'            Worksheets("メイン").EnableCalculation = True
'            MsgBox ("予約を取り消しました")
'            Unload 取消フォーム
'            Exit Sub
'        End If
'    Next i
'
'    '何事もなくループをでたら入力した内容を削除
'    MsgBox ("学籍番号が一致しません。もう一度入力してください")
'    Worksheets("メイン").EnableCalculation = True
'    TextBox1 = ""
'
'Else
'    MsgBox ("予約がありません。LAに確認を依頼してください。")
'End If
'
'End Sub
'
'Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'
'Call keypressrestrict(KeyAscii)
'
'End Sub
'
'Private Sub UserForm_Click()
'
'End Sub



Private Sub CommandButton1_Click()

Worksheets("メイン").EnableCalculation = False

Dim 予約コード As Long
Dim 現在の位置 As Long
Dim 確認番号 As String
Set Duplicate = Worksheets("重複チェック")
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
        Worksheets("メイン").EnableCalculation = True
        delete_confirm.Show
        Unload 取消フォーム
        Exit Sub
    End If
    
    'テキストボックスを変換する前にパスコードが入力されているかを場合分けする
    
    Call textbox_restrict(TextBox1, 確認番号)
    '入力された番号を変換する
    
    If 確認番号 = "" Then
        MsgBox ("学籍番号を入力してください")
        Worksheets("メイン").EnableCalculation = True
        Exit Sub
    End If
    If number_valid <> 0 Then
        number_valid = 0
        Worksheets("メイン").EnableCalculation = True
        Exit Sub
    End If
    
    
'    If 確認番号 = Sheets("生データ").Cells(現在の位置, i).Value Then

    Dim result_list As Variant
    result_list = Filter(target_stu_list, Int(確認番号))
    If UBound(result_list) <> -1 Then
        Call Sheets("生データ").Cells(現在の位置, 1).EntireRow.Delete(xlShiftUp)
        '        予約したデータを削除
        Call delete_res_num(target_stu_list, i - 1)
        Worksheets("メイン").EnableCalculation = True
'        MsgBox ("予約を取り消しました")
        delete_confirm.Show
        Unload 取消フォーム
        Exit Sub
    End If

    
    '何事もなくループをでたら入力した内容を削除
    MsgBox ("学籍番号が一致しません。もう一度入力してください")
    Worksheets("メイン").EnableCalculation = True
    TextBox1 = ""

Else
    MsgBox ("予約がありません。LAに確認を依頼してください。")
End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Call keypressrestrict(KeyAscii)

End Sub

Private Sub UserForm_Click()

End Sub

