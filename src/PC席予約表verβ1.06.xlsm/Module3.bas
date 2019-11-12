Attribute VB_Name = "Module3"
Option Explicit

Enum data_sheet
day_code = 1
time_zone
seat_num
reserve_code
cable_frag
student_num_start

End Enum

Function translate_number(ByVal raw_number As String, option_number As Integer)
'テキストボックスに入力された番号を変換するプロシージャ
'student_numberのデータ型はvaliantじゃないとダメ。桁数がおおいので。
'when raw_number is a invalid number,this procedure return -1
'option_number 0の時は無効な学籍番号が入力されてもエラーメッセージを表示しない。１の時はエラーメッセージを表示する

Dim enter_year As Integer
Dim subject As Integer
'number_valid変数はこのプロシージャでexitsubをしても元のプロシージャを抜けることはできないので、１以上の値ならエラーを出したというフラグとして使っています。
'ここで二重のプロシージャを抜ける方法もあるのかもしれないが、そうすると、複数のテキストボックスに入力に誤りがあった際に、最初にエラーを見つけた時点で抜けてしまうと、そのテキストボックスしか空白にできないので現状このようにしています。

'If raw_number <= 0 Then
''    MsgBox ("有効な学籍番号を入力してください")
'    stunum_error.Show
'    number_valid = number_valid + 1
'    translate_number = -1
'    ''0以下が入力された場合は-1
'    Exit Function
'End If

If Len(raw_number) = 7 Then
'学籍番号は７桁なので、７桁入力された場合は下の処理を行って変換する
'    enter_yearの処理
    If IsNumeric(Mid(raw_number, 3, 2)) = True Then
        enter_year = Mid(raw_number, 3, 2)
    Else
        If option_number = 1 Then
            stunum_error.Show
        End If
        number_valid = number_valid + 1
        translate_number = -1
        Exit Function
    End If
    
        '5文字目がMだった場合はマスターの処理
    If Mid(raw_number, 5, 1) = "M" Or Mid(raw_number, 5, 1) = "m" Then
        If Mid(raw_number, 1, 2) <= 10 Then
            subject = Mid(raw_number, 1, 2) + 2000
        ElseIf Mid(raw_number, 1, 2) = "61" Then
            subject = 2201
        ElseIf Mid(raw_number, 1, 2) = "62" Then
            subject = 2202
        ElseIf Mid(raw_number, 1, 2) = "51" Then
            subject = 2101
        ElseIf IsNumeric(Mid(raw_number, 1, 2)) = False Then
            If option_number = 1 Then
                stunum_error.Show
            End If
            number_valid = number_valid + 1
            translate_number = -1
            Exit Function
        Else
            subject = 2099        '2099は予想しない学籍番号来た場合のワイルドカード
        End If
        If IsNumeric(Mid(raw_number, 6, 2)) = False Then
            If option_number = 1 Then
                stunum_error.Show
            End If
            number_valid = number_valid + 1
            translate_number = -1
            Exit Function
        End If
        translate_number = enter_year & subject & "0" & Mid(raw_number, 6, 2)
'ドクターの処理
    ElseIf Mid(raw_number, 5, 1) = "D" Or Mid(raw_number, 5, 1) = "d" Then
        If Mid(raw_number, 1, 2) = 1 Then
            subject = Mid(raw_number, 1, 2) + 2010
        ElseIf Mid(raw_number, 1, 2) >= 2 Then
            subject = Mid(raw_number, 1, 2) + 2011
        ElseIf Mid(raw_number, 1, 2) = "61" Then
            subject = 2211
        ElseIf Mid(raw_number, 1, 2) = "62" Then
            subject = 2212
        ElseIf Mid(raw_number, 1, 2) = "51" Then
            subject = 2111
        ElseIf IsNumeric(Mid(raw_number, 1, 2)) = False Then
            If option_number = 1 Then
                stunum_error.Show
            End If
            number_valid = number_valid + 1
            translate_number = -1
            Exit Function
        Else
            subject = 2199  'ワイルドカード
        End If
        If IsNumeric(Mid(raw_number, 6, 2)) = False Then
            If option_number = 1 Then
                stunum_error.Show
            End If
            number_valid = number_valid + 1
            translate_number = -1
            Exit Function
        End If
            translate_number = enter_year & subject & "0" & Mid(raw_number, 6, 2)
    ElseIf Mid(raw_number, 5, 1) = "s" Or Mid(raw_number, 5, 1) = "S" Then
        If Mid(raw_number, 1, 2) <= 10 Then
            subject = Mid(raw_number, 1, 2) + 2500
        ElseIf Mid(raw_number, 1, 2) >= 51 And Mid(raw_number, 1, 2) <= 57 Then
            subject = Mid(raw_number, 1, 2) - 40 + 2500
        ElseIf Mid(raw_number, 1, 2) = 11 Then
            subject = 2521
        ElseIf IsNumeric(Mid(raw_number, 1, 2)) = False Then
            If option_number = 1 Then
                stunum_error.Show
            End If
            number_valid = number_valid + 1
            translate_number = -1
            Exit Function
        Else
            subject = 2599 'ワイルドカード
        End If
            
        If IsNumeric(Mid(raw_number, 6, 2)) = False Then
            If option_number = 1 Then
                stunum_error.Show
            End If
            number_valid = number_valid + 1
            translate_number = -1
            Exit Function
        End If
            '５文字目がSの交換留学生に対応。めんどくさいので学科番号のあとすぐを９にすることで解決。多分ほとんど来ないでしょう
            translate_number = enter_year & subject & "9" & Mid(raw_number, 6, 2)
        ' MでもDでもSでもない場合の処理
    Else
        If Mid(raw_number, 1, 2) <= 10 Then
            subject = Mid(raw_number, 1, 2) + 2500
        ElseIf Mid(raw_number, 1, 2) >= 51 And Mid(raw_number, 1, 2) <= 57 Then
            subject = Mid(raw_number, 1, 2) - 40 + 2500
        ElseIf Mid(raw_number, 1, 2) = 11 Then
            subject = 2521
        ElseIf IsNumeric(Mid(raw_number, 1, 2)) = False Then
            If option_number = 1 Then
                stunum_error.Show
            End If
            number_valid = number_valid + 1
            translate_number = -1
            Exit Function
        Else
            subject = 2599 'ワイルドカード
        End If
            If IsNumeric(Mid(raw_number, 6, 2)) = False Then
                If option_number = 1 Then
                    stunum_error.Show
                End If
                number_valid = number_valid + 1
                translate_number = -1
                Exit Function
            End If
        translate_number = enter_year & subject & Mid(raw_number, 5, 3)
        
    '   社会福祉学科だけ台帳番号が謎なので最後に特例的に処理を書いた
    End If

ElseIf Len(raw_number) = 16 Then
    
        enter_year = Mid(raw_number, 3, 2)
        subject = Mid(raw_number, 8, 4)
        translate_number = enter_year & subject & Mid(raw_number, 13, 3)
 Else
 
    If raw_number <> "" Then
    translate_number = raw_number
    
    'テキストボックスに何か入力されていて７桁の文字列でない場合はそのまま代入する
    End If

End If

If raw_number <> "" Then
'テキストボックスに何も入力されていない場合は処理を行わない
    
    If translate_number <= 0 Then
        If option_number = 1 Then
            stunum_error.Show
        End If
        number_valid = number_valid + 1
        translate_number = -1
        ''0以下が入力された場合はテキストボックスを空にした入力フォームに戻る
    Exit Function
    ElseIf IsNumeric(translate_number) = False And translate_number <> "" Then
        If option_number = 1 Then
            stunum_error.Show
        End If
        number_valid = number_valid + 1
        translate_number = -1
        Exit Function
    ElseIf Len(translate_number) <> 9 Then
        If option_number = 1 Then
            stunum_error.Show
        End If
        number_valid = number_valid + 1
        translate_number = -1
        Exit Function
    End If
End If

End Function
Public Sub keypressrestrict(ByVal KeyAscii As MSForms.ReturnInteger)
'キー入力でM、D、Sと数字以外が入力できないようにしてる
If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" And Chr(KeyAscii) <> "M" And Chr(KeyAscii) <> "D" And Chr(KeyAscii) <> "m" And Chr(KeyAscii) <> "d" And Chr(KeyAscii) <> "s" And Chr(KeyAscii) <> "S" Then
    KeyAscii = 0
End If
End Sub
Public Sub cable()
'ケーブル貸し出しをするプロシージャ

    Dim search_target_code As Long
    Dim result_row  As Long
    search_target_code = resreve_day * 100 + 時間帯 * 10 + 席番号
    result_row = WorksheetFunction.Match(search_target_code, Sheets("生データ").Cells(1, data_sheet.reserve_code).EntireColumn, 1)
    If Sheets("生データ").Cells(result_row, data_sheet.cable_frag).Value = 0 Then
        Sheets("生データ").Cells(result_row, data_sheet.cable_frag).Value = 1
    Else
        Sheets("生データ").Cells(result_row, data_sheet.cable_frag).Value = 0
    End If

End Sub
Public Sub cable_new(ByVal check As Boolean, ByVal Row As Variant)

'新規予約したときにケーブル貸し出しするときのプロシージャ
If check = True Then
    Sheets("生データ").Cells(Row, data_sheet.cable_frag).Value = 1
Else
    Sheets("生データ").Cells(Row, data_sheet.cable_frag).Value = 0
End If

End Sub

Function res_duplicate_check(ByVal data_number As Integer, situation As Integer, count() As Integer)
'既定の予約数を超えていないかチェックするプロシージャ
Dim i As Integer
Dim bl_res As String
Dim limit_res_day As Integer

If IsNumeric(Range(limit_reserve_count).Value) = False Then
    limit_res_day = 36
Else
    limit_res_day = Range(limit_reserve_count).Value
End If

For i = 0 To data_number
    If count(i) >= limit_res_day - situation Then
            bl_res = MsgBox("１日に予約できるコマ上限数をオーバーしてしまいます。予約を続けますか？", vbYesNo + vbQuestion, "予約の確認")
                If bl_res = vbNo Then
                    res_duplicate_check = False
                    Exit Function
                Else
'                    Dim inputpass As String
'                    inputpass = InputBox("予約を続ける場合はLAを呼び、パスコードの入力を依頼してください", "パスコードの入力")
'                    If inputpass = passcord Then
                    Dim get_pass As Integer
                    get_pass = passcord_inputform
                    If get_pass = 0 Then
                        Exit For
                    ElseIf get_pass = 1 Then
                        'MsgBox ("パスコードが一致しません。予約画面に移動します。")
                        res_duplicate_check = False
                        Exit Function
                    ElseIf get_pass = 2 Or get_pass = 3 Then
                        res_duplicate_check = False
                        Exit Function
                    End If
                End If
    End If
Next i

res_duplicate_check = True

End Function

Function res_input_rawsheet(ByVal resrve_day_for_input As Long, ByVal time_cord As Integer, ByVal seat_number As Integer, cable As Boolean, stu_list() As Variant, data_number As Integer)
'生データシートにデータを入力するプロシージャ
Dim resrve_code_number As Long
resrve_code_number = resrve_day_for_input * 100 + time_cord * 10 + seat_number
        
Dim search As Integer
On Error GoTo error_process
    search = WorksheetFunction.Match(resrve_code_number, Sheets("生データ").Cells(1, data_sheet.reserve_code).EntireColumn, 1)
On Error GoTo 0
If resrve_code_number = WorksheetFunction.Index(Sheets("生データ").Cells(1, data_sheet.reserve_code).EntireColumn, search) Then
    MsgBox ("すでにこの枠の予約があるため予約ができません。LAに確認を依頼してください(エラー番号:１01)")
    res_input_rawsheet = False
    Exit Function
End If
    
Sheets("生データ").Rows(search + 1).Insert

Sheets("生データ").Cells(search + 1, data_sheet.day_code).Value = resreve_day
Sheets("生データ").Cells(search + 1, data_sheet.time_zone).Value = time_cord
Sheets("生データ").Cells(search + 1, data_sheet.seat_num).Value = seat_number
Sheets("生データ").Cells(search + 1, data_sheet.reserve_code).Value = resrve_code_number

Call cable_new(cable, search + 1)
Call stu_num_list_input_rawsheet(search + 1, stu_list(), data_number)

res_input_rawsheet = True
Exit Function

error_process:
search = 1
Resume Next
End Function

Function stu_num_list_input_rawsheet(ByVal Row As Integer, stu_list() As Variant, data_number)
'生データシートに学籍番号を入力するプロシージャ

Dim Lastcolumn As Long
Dim i As Integer
Lastcolumn = Sheets("生データ").Cells(Row, Columns.count).End(xlToLeft).Column + 1
For i = 0 To data_number
    Sheets("生データ").Cells(Row, Lastcolumn + i).Value = stu_list(i)
Next i

Call input_res_num(stu_list(), data_number)

End Function

