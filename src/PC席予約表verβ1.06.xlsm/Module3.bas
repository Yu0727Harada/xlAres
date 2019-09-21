Attribute VB_Name = "Module3"
Option Explicit

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

