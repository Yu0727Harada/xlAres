Attribute VB_Name = "Module1"
Public 予約日 As Long
Public 時間帯 As Integer
Public 席番号 As Integer
Public 連続可能か As String
Public frag As Integer
Public number_valid As Integer
Public tm As Double
Public on_time As Integer
Public Const passcord As String = 1907

Public Sub setting_time()

Dim now_time As Date

On Error GoTo sheet_cal_error
now_time = Sheets("メイン").Cells(2, 12).Value
On Error GoTo 0

If now_time > 0.4375 And now_time <= 0.50694444 Then
    on_time = 3
ElseIf now_time > 0.5069444 And now_time <= 0.5416 Then
    on_time = 4
ElseIf now_time > 0.5416 And now_time <= 0.60416 Then
    on_time = 5
ElseIf now_time > 0.60416 And now_time <= 0.6736 Then
    on_time = 6
ElseIf now_time > 0.6736 And now_time <= 0.74305 Then
    on_time = 7
ElseIf now_time > 0.74305 And now_time <= 0.79166 Then
    on_time = 8
ElseIf now_time > 0.79166 Then
    on_time = 9
Else
    on_time = 2
End If



'現在の時刻を取得して数字を代入。数時は時間帯に対応するもの

'If Time > 0.4375 And Time <= 0.50694444 Then
'    on_time = 3
'ElseIf Time > 0.5069444 And Time <= 0.5416 Then
'    on_time = 4
'ElseIf Time > 0.5416 And Time <= 0.60416 Then
'    on_time = 5
'ElseIf Time > 0.60416 And Time <= 0.6736 Then
'    on_time = 6
'ElseIf Time > 0.6736 And Time <= 0.74305 Then
'    on_time = 7
'ElseIf Time > 0.74305 And Time <= 0.79166 Then
'    on_time = 8
'ElseIf Time > 0.79166 Then
'    on_time = 9
''ElseIf Time < 0.44 Then
'    on_time = 5
'Else
'    on_time = 2
'
'Timeを用いればコンピューター時計準拠に､
'変数now_timeを使えばセルでいじれます

'End If

Exit Sub


sheet_cal_error:
Exit Sub
End Sub
Public Sub shift_check()

Worksheets("メイン").EnableCalculation = False
Dim now_time As Date

On Error GoTo sheet_cal_error
'now_time = Sheets("メイン").Cells(2, 12).Value
now_time = Time
On Error GoTo 0

Dim i As Integer


For i = 0 To 48
If now_time > i * 1 / 48 And now_time < i * 1 / 48 + 1 / 24 / 60 Then
Call shift_output_mainsheet(now_time)
Exit For
'    now_date = Sheets("メイン").Cells(2, 11).Value
'    search = WorksheetFunction.Match(CDbl(now_date), Sheets("シフト表").Range("B:B"), 1) + 1
'    If Int(now_date) <> Int(WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search)) Then
'    Exit For
'    Else
'        Do While now_date = Int(WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search))
'
'            end_time = WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search) - now_date
'            start_time = WorksheetFunction.Index(Sheets("シフト表").Range("A:A"), search) - now_date
'            If now_time < end_time And now_time > start_time Then
'                shift(j) = WorksheetFunction.Index(Sheets("シフト表").Range("C:C"), search)
'                j = j + 1
'            End If
'            search = search + 1
'        Loop
'            Dim k As Integer
'            k = 0
'            For k = 0 To j
'                Cells(1, 15 + k).Value = shift(k)
'            Next k
'
'    End If
End If
Next i

'i = shift(0)
'Sheets("メイン").Shapes.Range(Array("Picture 4")).Formula = "=出力!inditrct(address(," + Str(i) + ",2)"
'=mid(indirect(address(" + Str(現在の位置) + "," + Str(i) + ")),1,8)
Worksheets("メイン").EnableCalculation = True
sheet_cal_error:
Exit Sub
End Sub

Public Sub shift_output_mainsheet(ByVal now_time As Date)

Dim j As Integer
Dim now_date As Date
Dim search As Integer
Dim end_time As Date
Dim start_time As Date
Dim shift(4) As Integer
j = 0

   now_date = Sheets("メイン").Cells(2, 11).Value
    search = WorksheetFunction.Match(CDbl(now_date), Sheets("シフト表").Range("B:B"), 1) + 1
    If Int(now_date) <> Int(WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search)) Then
    Exit Sub
    Else
        Do While now_date = Int(WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search))
                   
            end_time = WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search) - now_date
            start_time = WorksheetFunction.Index(Sheets("シフト表").Range("A:A"), search) - now_date
            If now_time < end_time And now_time > start_time Then
                shift(j) = WorksheetFunction.Index(Sheets("シフト表").Range("C:C"), search)
                j = j + 1
            End If
            search = search + 1
        Loop
            Dim k As Integer
            k = 0
            For k = 0 To j - 1
                Cells(1, 15 + k).Value = shift(k)
            Next k

    End If

End Sub

Public Sub recal()

Application.Calculate
'シートの再計算を行う
Call shift_check
tm = Now() + TimeValue("00:01:00")
Application.OnTime EarliestTime:=tm, Procedure:="recal", Schedule:=True
'tm変数に一分後をセット
'ontime関数で一分後にまたrecalプロシージャを実行
'なぜかシートモジュールにrecalプロシージャを書くとうまくいかない

End Sub


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
    予約コード = 予約日 * 100 + 時間帯 * 10 + 席番号
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

Public Sub cable_new(ByVal check As Object, ByVal Row As Variant)

'新規予約したときにケーブル貸し出しするときのプロシージャ
If check.Value = True Then
    Sheets("生データ").Cells(Row, 5).Value = 1
Else
    Sheets("生データ").Cells(Row, 5).Value = 0
End If

End Sub

Sub input_res_num(ByRef student_num_list() As Variant, ByVal stu_data_num As Integer)
'重複チェックシートにstudent_num_listの学籍番号があるかどうかチェックし、あったらB列に１を足し、なかったら昇順で位置する場所に番号を挿入する
Set Duplicate = Worksheets("重複チェック")
Dim search_stu_row
Dim i As Integer

    For i = 0 To stu_data_num
        search_stu_row = WorksheetFunction.Match(Int(student_num_list(i)), Duplicate.Range("A:A"), 1)
        If Int(student_num_list(i)) <> WorksheetFunction.Index(Duplicate.Range("A:A"), search_stu_row) Then
            Duplicate.Rows(search_stu_row + 1).Insert
            Duplicate.Cells(search_stu_row + 1, 1) = student_num_list(i)
            Duplicate.Cells(search_stu_row + 1, 2) = Duplicate.Cells(search_stu_row + 1, 2) + 1
        Else
            Duplicate.Cells(search_stu_row, 2) = Duplicate.Cells(search_stu_row, 2) + 1
        End If
    Next i

End Sub
Sub delete_res_num(ByRef student_num_list() As Variant, ByVal stu_data_num As Integer)
'予約が削除されたときに重複チェックシートの学籍番号のB列から１を引き、０になった場合は行を削除する
Set Duplicate = Worksheets("重複チェック")
Dim search_stu_row
Dim i As Integer

    For i = 0 To stu_data_num
        search_stu_row = WorksheetFunction.Match(Int(student_num_list(i)), Duplicate.Range("A:A"), 1)
        If Int(student_num_list(i)) <> WorksheetFunction.Index(Duplicate.Range("A:A"), search_stu_row) Then
        MsgBox ("該当の学籍番号が重複チェックシートで見つかりませんでした")
        Else
            Duplicate.Cells(search_stu_row, 2) = Duplicate.Cells(search_stu_row, 2) - 1
            If Duplicate.Cells(search_stu_row, 2) <= 0 Then
                Call Duplicate.Cells(search_stu_row, 1).EntireRow.Delete(xlShiftUp)
            End If
        End If
    Next i

End Sub

Sub check_res_num(ByRef student_num_list() As Variant, ByVal stu_data_num As Integer, ByRef CNT() As Integer)
'重複チェックシートに学籍番号が登録されているかチェックする
Set Duplicate = Worksheets("重複チェック")
Dim search_stu_row
Dim i As Integer

    For i = 0 To stu_data_num
        search_stu_row = WorksheetFunction.Match(Int(student_num_list(i)), Duplicate.Range("A:A"), 1)
        If Int(student_num_list(i)) <> WorksheetFunction.Index(Duplicate.Range("A:A"), search_stu_row) Then
            CNT(i) = 0
        Else
            CNT(i) = Duplicate.Cells(search_stu_row, 2)
        End If
    Next i

End Sub

Sub check_res_day()
'日付が変わったときに変更された日の重複チェックシートに更新する

Worksheets("メイン").EnableCalculation = False
Set main = Worksheets("メイン")
Set Duplicate = Worksheets("重複チェック")
Set data = Worksheets("生データ")

Duplicate.Cells(1, 2).Value = Format(main.Cells(2, 11), "yyyymmdd")
If Duplicate.Cells(1, 1) = Duplicate.Cells(1, 2) Then
    Exit Sub
End If
Duplicate.Cells.Clear
Duplicate.Cells(1, 1) = Format(main.Cells(2, 11), "yyyymmdd")

Call data.Range("A:F").Sort(key1:=data.Range("D:D"), order1:=xlAscending, Header:=xlYes)

Dim search_up As Integer

On Error GoTo error_process
search_up = WorksheetFunction.Match(Duplicate.Cells(1, 1).Value, data.Range("A:A"), 1)
On Error GoTo 0

If Duplicate.Cells(1, 1).Value <> WorksheetFunction.Index(data.Range("A:A"), search_up) Then
    Exit Sub
End If

Dim search_target_Row As Integer
Dim i As Integer
Dim j As Integer

i = 0
    While data.Cells(search_up - i, 1) = Duplicate.Cells(1, 1).Value
        j = 0
        While data.Cells(search_up - i, 6 + j).Value <> ""
            On Error GoTo error_process_2
            search_target_Row = WorksheetFunction.Match(data.Cells(search_up - i, 6 + j), Duplicate.Range("A:A"), 1)
            On Error GoTo 0
                If data.Cells(search_up - i, 6 + j) = Duplicate.Cells(search_target_Row, 1) Then
                    Duplicate.Cells(search_target_Row, 2) = Duplicate.Cells(search_target_Row, 2) + 1
                Else
                    Duplicate.Rows(search_target_Row + 1).Insert
                    Duplicate.Cells(search_target_Row + 1, 1) = data.Cells(search_up - i, 6 + j)
                    Duplicate.Cells(search_target_Row + 1, 2) = Duplicate.Cells(search_target_Row + 1, 2) + 1
                End If
            j = j + 1
        Wend
        
        i = i + 1
        
    Wend

Exit Sub

error_process:
search_up = 1
Resume Next

error_process_2:
search_target_Row = 1
Resume Next
End Sub
