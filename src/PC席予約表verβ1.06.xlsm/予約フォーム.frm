VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 予約フォーム 
   Caption         =   "予約フォーム"
   ClientHeight    =   8442
   ClientLeft      =   -462
   ClientTop       =   -1799
   ClientWidth     =   8022
   OleObjectBlob   =   "予約フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "予約フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim extend_check As Boolean
Dim cable_check As Boolean

Private Sub Label4_Click()

If checBox = "" Then
    checBox = "●"
    extend_check = True
Else
    checBox = ""
    extend_check = False
End If
End Sub

Private Sub Label5_Click()
If checkbox2 = "" Then
    checkbox2 = "●"
    cable_check = True
Else
    checkbox2 = ""
    cable_check = False
End If
End Sub

Private Sub UserForm_Initialize()

number_valid = 0
学籍番号テキストボックス1.SetFocus
cable_check = False
extend_check = False
'一番最初にフォームが開いたときにテキストボックス位置に入力出来る状態にしておく
End Sub

Private Sub キャンセル_Click()

Unload 予約フォーム

End Sub

Private Sub 学籍番号テキストボックス1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)
End Sub
Private Sub 学籍番号テキストボックス2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)
End Sub
Private Sub 学籍番号テキストボックス3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)
End Sub
Private Sub 学籍番号テキストボックス4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)
End Sub
Private Sub 学籍番号テキストボックス5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Call keypressrestrict(KeyAscii)
End Sub
Private Sub 登録_Click()

'Worksheets("メイン").EnableCalculation = False
'生データにデータを入れるたびにメインシート再計算が起こると処理が重くなるので再計算を停止する

Dim student_number_1 As Variant
Dim student_number_2 As Variant
Dim student_number_3 As Variant
Dim student_number_4 As Variant
Dim student_number_5 As Variant

student_number_1 = translate_number(学籍番号テキストボックス1, 1)
If student_number_1 = -1 Then
    学籍番号テキストボックス1 = ""
End If
student_number_2 = translate_number(学籍番号テキストボックス2, 1)
If student_number_2 = -1 Then
    学籍番号テキストボックス2 = ""
End If
student_number_3 = translate_number(学籍番号テキストボックス3, 1)
If student_number_3 = -1 Then
    学籍番号テキストボックス3 = ""
End If
student_number_4 = translate_number(学籍番号テキストボックス4, 1)
If student_number_4 = -1 Then
    学籍番号テキストボックス4 = ""
End If
student_number_5 = translate_number(学籍番号テキストボックス5, 1)
If student_number_5 = -1 Then
    学籍番号テキストボックス5 = ""
End If

If number_valid <> 0 Then
    number_valid = 0
    Exit Sub
End If

'テキストボックスに入力されたのを変換しstudent_number_1 -5にいれる
'ループでやらないのは無効な番号が入力されたときに入力された内容を削除するときにめんどいから

'Dim 複数人表示参照 As Worksheet
Dim CNT(5) As Integer
'Dim 予約確認 As Integer
Dim student_number_list(5) As Variant
Dim data_num As Integer
data_num = -1

Dim 現在の位置 As Integer
 

If student_number_1 <> "" Then
    data_num = data_num + 1
    student_number_list(data_num) = student_number_1
End If
If student_number_2 <> "" Then
    data_num = data_num + 1
    student_number_list(data_num) = student_number_2
End If
If student_number_3 <> "" Then
    data_num = data_num + 1
     student_number_list(data_num) = student_number_3
End If
If student_number_4 <> "" Then
    data_num = data_num + 1
    student_number_list(data_num) = student_number_4
End If
If student_number_5 <> "" Then
    data_num = data_num + 1
    student_number_list(data_num) = student_number_5
End If

If data_num = -1 Then
    MsgBox ("学籍番号を入力してください")
    Exit Sub
End If

'student_number_listに変換した番号を０から順に格納。何も入力されてないdata_num=-1の時にはプロシージャを抜ける
Dim extend_bl As String

If extend_check = True Then
    If 連続可能か = False Then
        extend_bl = MsgBox("次の時間帯は予約できません。一コマだけ予約しますか？", vbYesNo + vbQuestion, "予約の確認")
        If extend_bl = vbNo Then
            MsgBox ("予約画面に移動します。")
            Unload 予約フォーム
            Exit Sub
        Else
            extend_check = False
        End If
    ElseIf Worksheets("メイン").Range(limit_res_on_off).Value = "on" Then
            extend_bl = MsgBox("現在、混雑のため予約の制限をしています。１コマだけ予約しますか？", vbYesNo + vbQuestion, "予約の確認")
        If extend_bl = vbNo Then
            MsgBox ("予約画面に移動します。")
            Unload 予約フォーム
            Exit Sub
        Else
            extend_check = False
        End If
    End If
End If
    
Call check_res_day
Call check_res_num(student_number_list(), data_num, CNT())
    
Dim bl_dup_check As Boolean
bl_dup_check = res_duplicate_check(data_num, 0, CNT())
    
If bl_dup_check = False Then
    Worksheets("メイン").EnableCalculation = True
    Unload 予約フォーム
    Exit Sub
End If
       
Dim bl_res_input_raw As Boolean
bl_res_input_raw = res_input_rawsheet(resreve_day, 時間帯, 席番号, cable_check, student_number_list(), data_num)
If bl_res_input_raw = False Then
    Unload 予約フォーム
    Exit Sub
End If
      
If extend_check = True Then
    bl_res_input_raw = res_input_rawsheet(resreve_day, 時間帯 + 1, 席番号, cable_check, student_number_list(), data_num)
    If bl_res_input_raw = False Then
        Unload 予約フォーム
        Exit Sub
    End If
    Unload 予約フォーム
    Exit Sub
End If

Unload 予約フォーム


End Sub
