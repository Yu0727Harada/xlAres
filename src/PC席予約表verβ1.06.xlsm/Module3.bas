Attribute VB_Name = "Module3"
Option Explicit
Public Sub textbox_restrict(ByVal textbox_name As Object, student_number As Variant)

Dim 入学年度 As Integer
Dim 学科 As Integer
'number_valid変数はこのプロシージャでexitsubをしても元のプロシージャを抜けることはできないので、１以上の値ならエラーを出したというフラグとして使っています。
'ここで二重のプロシージャを抜ける方法もあるのかもしれないが、そうすると、複数のテキストボックスに入力に誤りがあった際に、最初にエラーを見つけた時点で抜けてしまうと、そのテキストボックスしか空白にできないので現状このようにしています。

If textbox_name <= 0 Then
'    MsgBox ("有効な学籍番号を入力してください")
    stunum_error.Show
    number_valid = number_valid + 1
    textbox_name = ""
    ''0以下が入力された場合はテキストボックスを空にした入力フォームに戻る
    Exit Sub
End If

If Len(textbox_name) = 7 Then
'学籍番号は７桁なので、７桁入力された場合は下の処理を行って変換する
'    入学年度の処理
    If IsNumeric(Mid(textbox_name, 3, 2)) = True Then
        入学年度 = Mid(textbox_name, 3, 2)
    Else
'        MsgBox ("有効な学籍番号を入力してください")
        stunum_error.Show
        number_valid = number_valid + 1
        textbox_name = ""
        Exit Sub
    End If
    
        '5文字目がMだった場合はマスターの処理
    If Mid(textbox_name, 5, 1) = "M" Or Mid(textbox_name, 5, 1) = "m" Then
        If Mid(textbox_name, 1, 2) <= 10 Then
            学科 = Mid(textbox_name, 1, 2) + 2000
        ElseIf Mid(textbox_name, 1, 2) = "61" Then
            学科 = 2201
        ElseIf Mid(textbox_name, 1, 2) = "62" Then
            学科 = 2202
        ElseIf Mid(textbox_name, 1, 2) = "51" Then
            学科 = 2101
        ElseIf IsNumeric(Mid(textbox_name, 1, 2)) = False Then
            stunum_error.Show
'            MsgBox ("有効な学籍番号を入力してください")
            number_valid = number_valid + 1
            textbox_name = ""
            Exit Sub
        Else
            学科 = 2099        '2099は予想しない学籍番号来た場合のワイルドカード
        End If
        If IsNumeric(Mid(textbox_name, 6, 2)) = False Then
'            MsgBox ("有効な学籍番号を入力してください")
            stunum_error.Show
            number_valid = number_valid + 1
            textbox_name = ""
            Exit Sub
        End If
        student_number = 入学年度 & 学科 & "0" & Mid(textbox_name, 6, 2)
'ドクターの処理
    ElseIf Mid(textbox_name, 5, 1) = "D" Or Mid(textbox_name, 5, 1) = "d" Then
        If Mid(textbox_name, 1, 2) = 1 Then
            学科 = Mid(textbox_name, 1, 2) + 2010
        ElseIf Mid(textbox_name, 1, 2) >= 2 Then
            学科 = Mid(textbox_name, 1, 2) + 2011
        ElseIf Mid(textbox_name, 1, 2) = "61" Then
            学科 = 2211
        ElseIf Mid(textbox_name, 1, 2) = "62" Then
            学科 = 2212
        ElseIf Mid(textbox_name, 1, 2) = "51" Then
            学科 = 2111
        ElseIf IsNumeric(Mid(textbox_name, 1, 2)) = False Then
'            MsgBox ("有効な学籍番号を入力してください")
            stunum_error.Show
            number_valid = number_valid + 1
            textbox_name = ""
            Exit Sub
        Else
            学科 = 2199  'ワイルドカード
        End If
        If IsNumeric(Mid(textbox_name, 6, 2)) = False Then
'                MsgBox ("有効な学籍番号を入力してください")
            stunum_error.Show
            number_valid = number_valid + 1
            textbox_name = ""
            Exit Sub
        End If
            student_number = 入学年度 & 学科 & "0" & Mid(textbox_name, 6, 2)
    ElseIf Mid(textbox_name, 5, 1) = "s" Or Mid(textbox_name, 5, 1) = "S" Then
        If Mid(textbox_name, 1, 2) <= 10 Then
            学科 = Mid(textbox_name, 1, 2) + 2500
        ElseIf Mid(textbox_name, 1, 2) >= 51 And Mid(textbox_name, 1, 2) <= 57 Then
            学科 = Mid(textbox_name, 1, 2) - 40 + 2500
        ElseIf Mid(textbox_name, 1, 2) = 11 Then
            学科 = 2521
        ElseIf IsNumeric(Mid(textbox_name, 1, 2)) = False Then
'                MsgBox ("有効な学籍番号を入力してください")
            stunum_error.Show
            number_valid = number_valid + 1
            textbox_name = ""
            Exit Sub
        Else
            学科 = 2599 'ワイルドカード
        End If
            
        If IsNumeric(Mid(textbox_name, 6, 2)) = False Then
'               MsgBox ("有効な学籍番号を入力してください")
            stunum_error.Show
            number_valid = number_valid + 1
            textbox_name = ""
            Exit Sub
        End If
            '５文字目がSの交換留学生に対応。めんどくさいので学科番号のあとすぐを９にすることで解決。多分ほとんど来ないでしょう
            student_number = 入学年度 & 学科 & "9" & Mid(textbox_name, 6, 2)
        ' MでもDでもSでもない場合の処理
    Else
        If Mid(textbox_name, 1, 2) <= 10 Then
            学科 = Mid(textbox_name, 1, 2) + 2500
        ElseIf Mid(textbox_name, 1, 2) >= 51 And Mid(textbox_name, 1, 2) <= 57 Then
            学科 = Mid(textbox_name, 1, 2) - 40 + 2500
        ElseIf Mid(textbox_name, 1, 2) = 11 Then
            学科 = 2521
        ElseIf IsNumeric(Mid(textbox_name, 1, 2)) = False Then
'                MsgBox ("有効な学籍番号を入力してください")
            stunum_error.Show
            number_valid = number_valid + 1
            textbox_name = ""
            Exit Sub
        Else
            学科 = 2599 'ワイルドカード
        End If
            If IsNumeric(Mid(textbox_name, 6, 2)) = False Then
'                    MsgBox ("有効な学籍番号を入力してください")
                stunum_error.Show
                number_valid = number_valid + 1
                textbox_name = ""
                Exit Sub
            End If
        student_number = 入学年度 & 学科 & Mid(textbox_name, 5, 3)
        
    '   社会福祉学科だけ台帳番号が謎なので最後に特例的に処理を書いた
    End If

ElseIf Len(textbox_name) = 16 Then
    
        入学年度 = Mid(textbox_name, 3, 2)
        学科 = Mid(textbox_name, 8, 4)
        student_number = 入学年度 & 学科 & Mid(textbox_name, 13, 3)
 Else
 
    If textbox_name <> "" Then
    student_number = textbox_name
    
    'テキストボックスに何か入力されていて７桁の文字列でない場合はそのまま代入する
    End If

End If

If textbox_name <> "" Then
'テキストボックスに何も入力されていない場合は処理を行わない
    
    If student_number <= 0 Then
'        MsgBox ("有効な学籍番号を入力してください")
        stunum_error.Show
        number_valid = number_valid + 1
        textbox_name = ""
        ''0以下が入力された場合はテキストボックスを空にした入力フォームに戻る
    Exit Sub
    ElseIf IsNumeric(student_number) = False And student_number <> "" Then
'        MsgBox ("有効な学籍番号を入力してください")
        stunum_error.Show
        number_valid = number_valid + 1
        textbox_name = ""
        Exit Sub
    ElseIf Len(student_number) <> 9 Then
'        MsgBox ("有効な学籍番号を入力してください")
        stunum_error.Show
        number_valid = number_valid + 1
        textbox_name = ""
        Exit Sub
    End If
End If

End Sub
Public Sub keypressrestrict(ByVal KeyAscii As MSForms.ReturnInteger)
'キー入力でM、D、Sと数字以外が入力できないようにしてる
If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" And Chr(KeyAscii) <> "M" And Chr(KeyAscii) <> "D" And Chr(KeyAscii) <> "m" And Chr(KeyAscii) <> "d" And Chr(KeyAscii) <> "s" And Chr(KeyAscii) <> "S" Then
'    If Chr(KeyAscii) = "M" Or Chr(KeyAscii) = "D" Or Chr(KeyAscii) = "m" Or Chr(KeyAscii) = "d" Or Chr(KeyAscii) = "s" Or Chr(KeyAscii) = "S" Then
'        Exit Sub
        
'    End If
    KeyAscii = 0
End If
End Sub
Public Sub cable()

    Dim 予約コード As Long
    Dim 現在の位置 As Long
    予約コード = resreve_day * 100 + 時間帯 * 10 + 席番号
    現在の位置 = WorksheetFunction.Match(予約コード, Sheets("生データ").Range("D:D"), 1)
    If Sheets("生データ").Cells(現在の位置, 5).Value = 0 Then
'        Dim 貸出確認 As Integer
'        貸出確認 = MsgBox("HDMIケーブル等を貸し出します", vbYesNo + vbwuestion, "貸出の確認")
'            If 貸出確認 = vbNo Then
'                Unload 予約変更フォーム
'                Exit Sub
'            Else
                Sheets("生データ").Cells(現在の位置, 5).Value = 1
'            End If
    Else
'        Dim 返却確認 As Integer
'        返却確認 = MsgBox("HDMIケーブル等の返却を受け付けました", vbYesNo + vbwuestion, "返却の確認")
'            If 返却確認 = vbNo Then
'                Unload 予約変更フォーム
'                Exit Sub
'            Else
                Sheets("生データ").Cells(現在の位置, 5).Value = 0
'            End If
    End If

End Sub
Public Sub cable_new(ByVal check As Boolean, ByVal Row As Variant)

'新規予約したときにケーブル貸し出しするときのプロシージャ
If check = True Then
    Sheets("生データ").Cells(Row, 5).Value = 1
Else
    Sheets("生データ").Cells(Row, 5).Value = 0
End If

End Sub

Function res_duplicate_check(ByVal data_number As Integer, situation As Integer, count() As Integer)

Dim i As Integer
Dim bl_res As String
For i = 0 To data_number
    If count(i) >= 2 - situation Then
            bl_res = MsgBox("１日に予約できるコマ上限数をオーバーしてしまいます。予約を続けますか？", vbYesNo + vbQuestion, "予約の確認")
                If bl_res = vbNo Then
                    res_duplicate_check = False
                    Exit Function
                Else
                    Dim inputpass As String
                    inputpass = InputBox("予約を続ける場合はLAを呼び、パスコードの入力を依頼してください", "パスコードの入力")
                    If inputpass = passcord Then
                        Exit For
                    ElseIf inputpass = "" Then
                        MsgBox ("予約画面に移動します。")
                        res_duplicate_check = False
                        Exit Function
                    Else
                        MsgBox ("パスコードが一致しません。予約画面に移動します。")
                        res_duplicate_check = False
                        Exit Function
                    End If
                End If
    End If
Next i

res_duplicate_check = True

End Function

Function res_input_rawsheet(ByVal resrve_day_for_input As Long, ByVal time_cord As Integer, ByVal seat_number As Integer, cable As Boolean, stu_list() As Variant, data_number As Integer)

Dim resrve_code_number As Long
resrve_code_number = resrve_day_for_input * 100 + time_cord * 10 + seat_number
        
Dim search As Integer
On Error GoTo error_process
    search = WorksheetFunction.Match(resrve_code_number, Sheets("生データ").Range("D:D"), 1)
On Error GoTo 0
If resrve_code_number = WorksheetFunction.Index(Sheets("生データ").Range("D:D"), search) Then
    MsgBox ("すでにこの枠の予約があるため予約ができません。LAに確認を依頼してください(error code:001)")
    res_input_rawsheet = False
    Exit Function
End If
    
Sheets("生データ").Rows(search + 1).Insert

Sheets("生データ").Cells(search + 1, 1).Value = resreve_day
Sheets("生データ").Cells(search + 1, 2).Value = time_cord
Sheets("生データ").Cells(search + 1, 3).Value = seat_number
Sheets("生データ").Cells(search + 1, 4).Value = resrve_code_number

Call cable_new(cable, search + 1)
Call stu_num_list_input_rawsheet(search + 1, stu_list(), data_number)

res_input_rawsheet = True
Exit Function

error_process:
search = 1
Resume Next
End Function

Function stu_num_list_input_rawsheet(ByVal Row As Integer, stu_list() As Variant, data_number)

Dim Lastcolumn As Long
Dim i As Integer
Lastcolumn = Sheets("生データ").Cells(Row, Columns.count).End(xlToLeft).Column + 1
For i = 0 To data_number
    Sheets("生データ").Cells(Row, Lastcolumn + i).Value = stu_list(i)
Next i
    
Call input_res_num(stu_list(), data_number)

End Function

