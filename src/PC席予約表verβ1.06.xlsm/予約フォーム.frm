VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 予約フォーム 
   Caption         =   "予約フォーム"
   ClientHeight    =   7497
   ClientLeft      =   -462
   ClientTop       =   -1799
   ClientWidth     =   6671
   OleObjectBlob   =   "予約フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "予約フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub checBox_Change()

End Sub

Private Sub CommandButton1_Click()

'If checBox = "" Then
'    checBox = "●"
'Else
'    checBox = ""
'End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

If checBox = "" Then
    checBox = "●"
    チェックボックス2コマ = True
Else
    checBox = ""
    チェックボックス2コマ = False
End If
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub Label5_Click()
If checkbox2 = "" Then
    checkbox2 = "●"
    ケーブルチェック = True
Else
    checkbox2 = ""
    ケーブルチェック = False
End If
End Sub

Private Sub UserForm_Initialize()

number_valid = 0
学籍番号テキストボックス1.SetFocus
'一番最初にフォームが開いたときにテキストボックス位置に入力出来る状態にしておく
End Sub

Private Sub キャンセル_Click()


Unload 予約フォーム

End Sub




Private Sub チェックボックス2コマ_Click()

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

Worksheets("メイン").EnableCalculation = False
'生データにデータを入れるたびにメインシート再計算が起こると処理が重くなるので再計算を停止する

Dim 学籍番号1 As Variant
Dim 学籍番号2 As Variant
Dim 学籍番号3 As Variant
Dim 学籍番号4 As Variant
Dim 学籍番号5 As Variant

Call textbox_restrict(学籍番号テキストボックス1, 学籍番号1)
Call textbox_restrict(学籍番号テキストボックス2, 学籍番号2)
Call textbox_restrict(学籍番号テキストボックス3, 学籍番号3)
Call textbox_restrict(学籍番号テキストボックス4, 学籍番号4)
Call textbox_restrict(学籍番号テキストボックス5, 学籍番号5)
If number_valid <> 0 Then
    number_valid = 0
    Worksheets("メイン").EnableCalculation = True
    Exit Sub
End If

'テキストボックスに入力されたのを変換し学籍番号1 -5にいれる
'ループでやらないのは無効な番号が入力されたときに入力された内容を削除するときにめんどいから


Dim 複数人表示参照 As Worksheet
Dim CNT(5) As Integer
Dim 予約確認 As Integer
Dim 学籍番号リスト(5) As Variant
Dim data_num As Integer
data_num = -1
Dim 予約コード As Long
Dim 現在の位置 As Integer
 

If 学籍番号1 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号1
End If
If 学籍番号2 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号2
End If
If 学籍番号3 <> "" Then
    data_num = data_num + 1
     学籍番号リスト(data_num) = 学籍番号3
End If
If 学籍番号4 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号4
End If
If 学籍番号5 <> "" Then
    data_num = data_num + 1
    学籍番号リスト(data_num) = 学籍番号5
End If

If data_num = -1 Then
    MsgBox ("学籍番号を入力してください")
    Worksheets("メイン").EnableCalculation = True
    Exit Sub
End If

'学籍番号リストに変換した番号を０から順に格納。何も入力されてないdata_num=-1の時にはプロシージャを抜ける

'Set 複数人表示参照 = Worksheets("複数人表示参照")

If チェックボックス2コマ = False Then

    Call check_res_day
    Call check_res_num(学籍番号リスト(), data_num, CNT())
    
    Dim k As Integer

    For k = 0 To data_num
    
        If CNT(k) >= 2 Then
            予約確認 = MsgBox("既に２コマ以上予約していますが、予約してよろしいですか？", vbYesNo + vbwuestion, "予約の確認")
                If 予約確認 = vbNo Then
                    Worksheets("メイン").EnableCalculation = True
                    Unload 予約フォーム
                    Exit Sub
                Else
                    Exit For
                End If
        End If
    Next k


'
'For k = 0 To data_num
'    CNT(k) = WorksheetFunction.CountIf(複数人表示参照.Range("C16:L62"), 学籍番号リスト(k))
'        If CNT(k) >= 2 Then
'
'            予約確認 = MsgBox("既に２コマ以上予約していますが、予約してよろしいですか？", vbYesNo + vbwuestion, "予約の確認")
'                If 予約確認 = vbNo Then
'                    Worksheets("メイン").EnableCalculation = True
'                    Unload 予約フォーム
'                    Exit Sub
'                Else
'                    Exit For
'                End If
'        End If
'Next k

'data_numの数､データの数ぶんループをまわす
'CNT配列に順に重複数を数える
'2 以上だったら確認する

   予約コード = 予約日 * 100 + 時間帯 * 10 + 席番号
    

    On Error GoTo error_process
    現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
    On Error GoTo 0
        If 予約コード = WorksheetFunction.Index(Sheets("生データ").Range("D:D"), 現在の位置) Then
            MsgBox ("すでにこの枠の予約があるため予約ができません。LAに確認を依頼してください")
            Worksheets("メイン").EnableCalculation = True
            Unload 予約フォーム
            Exit Sub
        End If
    
    Sheets("生データ").Rows(現在の位置 + 1).Insert
'    LastRow = Sheets("生データ").Cells(Rows.Count, 1).End(xlUp).Row + 1
    Dim Lastcolumn As Long
    
    Sheets("生データ").Cells(現在の位置 + 1, 1).Value = 予約日
    Sheets("生データ").Cells(現在の位置 + 1, 2).Value = 時間帯
    Sheets("生データ").Cells(現在の位置 + 1, 3).Value = 席番号
    Sheets("生データ").Cells(現在の位置 + 1, 4).Value = 予約コード
    
    
    Call cable_new(ケーブルチェック, 現在の位置 + 1)
    Lastcolumn = Sheets("生データ").Cells(現在の位置 + 1, Columns.Count).End(xlToLeft).Column + 1

    For m = 0 To data_num
        Sheets("生データ").Cells(現在の位置 + 1, Lastcolumn + m).Value = 学籍番号リスト(m)
    Next m
    
    Call input_res_num(学籍番号リスト(), data_num)

    
End If
        
If チェックボックス2コマ = True Then
    If 連続可能か = True Then
'   シート3で代入した値を確認して次の予約が空いてるか確認

        Call check_res_day
        Call check_res_num(学籍番号リスト(), data_num, CNT())
'    For o = 0 To data_num
'        CNT(o) = WorksheetFunction.CountIf(複数人表示参照.Range("C16:L62"), 学籍番号リスト(k))
'            If CNT(o) >= 1 Then
'                予約確認 = MsgBox("既に1コマ以上予約していますが、予約してよろしいですか？", vbYesNo + vbwuestion, "予約の確認")
'                    If 予約確認 = vbNo Then
'                        Worksheets("メイン").EnableCalculation = True
'                        Unload 予約フォーム
'                        Exit Sub
'                    Else
'                        Exit For
'                    End If
'            End If
'    Next o
        Dim L As Integer

        For L = 0 To data_num
        
            If CNT(L) >= 1 Then
                予約確認 = MsgBox("既に1コマ以上予約していますが、予約してよろしいですか？", vbYesNo + vbwuestion, "予約の確認")
                    If 予約確認 = vbNo Then
                        Worksheets("メイン").EnableCalculation = True
                        Unload 予約フォーム
                        Exit Sub
                    Else
                        Exit For
                    End If
            End If
        Next L

        予約コード = 予約日 * 100 + 時間帯 * 10 + 席番号
        On Error GoTo error_process
         現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
        On Error GoTo 0

             If 予約コード = WorksheetFunction.Index(Sheets("生データ").Range("D:D"), 現在の位置) Then
                 MsgBox ("すでにこの枠の予約があるため予約ができません。LAに確認を依頼してください")
                 Worksheets("メイン").EnableCalculation = True
                 Unload 予約フォーム
                 Exit Sub
             End If

            Sheets("生データ").Rows(現在の位置 + 1).Insert
            Sheets("生データ").Cells(現在の位置 + 1, 1).Value = 予約日
            Sheets("生データ").Cells(現在の位置 + 1, 2).Value = 時間帯
            Sheets("生データ").Cells(現在の位置 + 1, 3).Value = 席番号
            Sheets("生データ").Cells(現在の位置 + 1, 4).Value = 予約コード
            Call cable_new(ケーブルチェック, 現在の位置 + 1)
            Lastcolumn = Sheets("生データ").Cells(現在の位置 + 1, Columns.Count).End(xlToLeft).Column + 1

            For m = 0 To data_num
                Sheets("生データ").Cells(現在の位置 + 1, Lastcolumn + m).Value = 学籍番号リスト(m)
            Next m
            Call input_res_num(学籍番号リスト(), data_num)
          
            予約コード = 予約日 * 100 + (時間帯 + 1) * 10 + 席番号
            On Error GoTo error_process
            現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
            On Error GoTo 0
            Sheets("生データ").Rows(現在の位置 + 1).Insert
            Sheets("生データ").Cells(現在の位置 + 1, 1).Value = 予約日
            Sheets("生データ").Cells(現在の位置 + 1, 2).Value = 時間帯 + 1
            Sheets("生データ").Cells(現在の位置 + 1, 3).Value = 席番号
            Sheets("生データ").Cells(現在の位置 + 1, 4).Value = 予約コード
            Call cable_new(ケーブルチェック, 現在の位置 + 1)
            Lastcolumn = Sheets("生データ").Cells(現在の位置 + 1, Columns.Count).End(xlToLeft).Column + 1

            For m = 0 To data_num
                Sheets("生データ").Cells(現在の位置 + 1, Lastcolumn + m).Value = 学籍番号リスト(m)
            Next m
            Call input_res_num(学籍番号リスト(), data_num)


    Else
        MsgBox ("２コマ予約できません。")
    End If
End If

    Worksheets("メイン").EnableCalculation = True
    Call sheet_color_check
    Unload 予約フォーム
    Exit Sub

error_process:

現在の位置 = 1
Resume Next

End Sub
