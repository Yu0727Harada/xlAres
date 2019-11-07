VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 利用者追加フォーム 
   Caption         =   "利用者の追加"
   ClientHeight    =   6585
   ClientLeft      =   105
   ClientTop       =   448
   ClientWidth     =   6482
   OleObjectBlob   =   "利用者追加フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "利用者追加フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
number_valid = 0
学籍番号テキストボックス6.SetFocus
End Sub
Private Sub 学籍番号テキストボックス6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)

End Sub

Private Sub 学籍番号テキストボックス7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)

End Sub


Private Sub 学籍番号テキストボックス8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)

End Sub

Private Sub 学籍番号テキストボックス9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)

End Sub


Private Sub 学籍番号テキストボックス10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)

End Sub

Private Sub 追加登録ボタン_Click()

'Worksheets("メイン").EnableCalculation = False
'生データにデータを入れるたびにメインシート再計算が起こると処理が重くなるので再計算を停止する

Dim 学籍番号6 As Variant
Dim 学籍番号7 As Variant
Dim 学籍番号8 As Variant
Dim 学籍番号9 As Variant
Dim 学籍番号10 As Variant

Call textbox_restrict(学籍番号テキストボックス6, 学籍番号6)
Call textbox_restrict(学籍番号テキストボックス7, 学籍番号7)
Call textbox_restrict(学籍番号テキストボックス8, 学籍番号8)
Call textbox_restrict(学籍番号テキストボックス9, 学籍番号9)
Call textbox_restrict(学籍番号テキストボックス10, 学籍番号10)

If number_valid <> 0 Then
    number_valid = 0
    Exit Sub
End If

Dim 複数人表示参照 As Worksheet
Dim CNT(5) As Integer
Dim 予約確認 As Integer
Dim 学籍番号リスト(5) As Variant
Dim data_num As Integer
data_num = -1

If 学籍番号6 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号6
End If
If 学籍番号7 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号7
End If
If 学籍番号8 <> "" Then
    data_num = data_num + 1
     学籍番号リスト(data_num) = 学籍番号8
End If
If 学籍番号9 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号9
End If
If 学籍番号10 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号10
End If

If data_num = -1 Then
    MsgBox ("学籍番号を入力してください")
    Exit Sub
End If

Call check_res_day
Call check_res_num(学籍番号リスト(), data_num, CNT())

Dim bl_res_dup_check As Boolean
bl_res_dup_check = res_duplicate_check(data_num, 0, CNT())
If bl_res_dup_check = False Then
    Worksheets("メイン").EnableCalculation = True
    Unload 予約フォーム
    Exit Sub
End If

    
Dim 予約コード As Long
Dim add_search As Long
予約コード = resreve_day * 100 + 時間帯 * 10 + 席番号

add_search = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
'予約コードを生成して、それが何番目にあるのか取得

Call stu_num_list_input_rawsheet(add_search, 学籍番号リスト(), data_num)

Worksheets("メイン").EnableCalculation = True


adduser.Show

Unload 利用者追加フォーム
Application.Calculate

End Sub

Private Sub キャンセル追加ボタン_Click()

Unload 利用者追加フォーム


End Sub
