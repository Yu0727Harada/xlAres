Attribute VB_Name = "Module9"
Option Explicit
Enum duplicate_sheet
student_num = 1
reserve_count
End Enum



Sub input_res_num(ByRef student_num_list() As Variant, ByVal stu_data_num As Integer)
'重複チェックシートにstudent_num_listの学籍番号があるかどうかチェックし、あったらB列に１を足し、なかったら昇順で位置する場所に番号を挿入する
Dim duplicate As Worksheet
Set duplicate = Worksheets("重複チェック")
Dim search_stu_row
Dim i As Integer

    For i = 0 To stu_data_num
        search_stu_row = WorksheetFunction.Match(Int(student_num_list(i)), duplicate.Cells(1, duplicate_sheet.student_num).EntireColumn, 1)
        If Int(student_num_list(i)) <> WorksheetFunction.Index(duplicate.Cells(1, duplicate_sheet.student_num).EntireColumn, search_stu_row) Then
            duplicate.Rows(search_stu_row + 1).Insert
            duplicate.Cells(search_stu_row + 1, duplicate_sheet.student_num) = student_num_list(i)
            duplicate.Cells(search_stu_row + 1, duplicate_sheet.reserve_count) = duplicate.Cells(search_stu_row + 1, duplicate_sheet.reserve_count) + 1
        Else
            duplicate.Cells(search_stu_row, duplicate_sheet.reserve_count) = duplicate.Cells(search_stu_row, duplicate_sheet.reserve_count) + 1
        End If
    Next i

End Sub

Sub delete_res_num(ByRef student_num_list() As Variant, ByVal stu_data_num As Integer)
'予約が削除されたときに重複チェックシートの学籍番号のB列から１を引き、０になった場合は行を削除する
Dim duplicate As Worksheet
Set duplicate = Worksheets("重複チェック")
Dim search_stu_row
Dim i As Integer

    For i = 0 To stu_data_num
       On Error GoTo invalid_number_change

        search_stu_row = WorksheetFunction.Match(Int(student_num_list(i)), duplicate.Cells(1, duplicate_sheet.student_num).EntireColumn, 1)
        On Error GoTo 0
        If Int(student_num_list(i)) <> WorksheetFunction.Index(duplicate.Cells(1, duplicate_sheet.student_num).EntireColumn, search_stu_row) Then
        MsgBox ("該当の学籍番号が重複チェックシートで見つかりませんでした。エラー番号１０２")
        Else
            duplicate.Cells(search_stu_row, duplicate_sheet.reserve_count) = duplicate.Cells(search_stu_row, duplicate_sheet.reserve_count) - 1
            If duplicate.Cells(search_stu_row, duplicate_sheet.reserve_count) <= 0 Then
                Call duplicate.Cells(search_stu_row, 1).EntireRow.Delete(xlShiftUp)
            End If
        End If
    Next i
Exit Sub
invalid_number_change:
MsgBox ("該当の学籍番号が重複チェックシートで見つかりませんでした。エラー番号１０２")

End Sub

Sub check_res_num(ByRef student_num_list() As Variant, ByVal stu_data_num As Integer, ByRef CNT() As Integer)
'重複チェックシートに学籍番号が登録されているかチェックする
Dim duplicate As Worksheet
Set duplicate = Worksheets("重複チェック")
Dim search_stu_row
Dim i As Integer

    For i = 0 To stu_data_num
        search_stu_row = WorksheetFunction.Match(Int(student_num_list(i)), duplicate.Cells(1, duplicate_sheet.student_num).EntireColumn, 1)
        If Int(student_num_list(i)) <> WorksheetFunction.Index(duplicate.Cells(1, duplicate_sheet.student_num).EntireColumn, search_stu_row) Then
            CNT(i) = 0
        Else
            CNT(i) = duplicate.Cells(search_stu_row, duplicate_sheet.reserve_count)
        End If
    Next i

End Sub

Sub check_res_day()
'日付が変わったときに変更された日の重複チェックシートに更新する

'Worksheets("メイン").EnableCalculation = False
Dim main As Worksheet
Dim duplicate As Worksheet
Dim data As Worksheet
Set main = Worksheets("メイン")
Set duplicate = Worksheets("重複チェック")
Set data = Worksheets("生データ")

duplicate.Cells(1, 2).Value = Format(main.Cells(2, 11), "yyyymmdd")
If duplicate.Cells(1, 1) = duplicate.Cells(1, 2) Then
    Exit Sub
End If
duplicate.Cells.Clear
duplicate.Cells(1, 1) = Format(main.Cells(2, 11), "yyyymmdd")

Call main_sheet_sort

'Call data.Range("A:F").Sort(key1:=data.Range("D:D"), order1:=xlAscending, Header:=xlYes)

Dim search_up As Integer

On Error GoTo error_process
search_up = WorksheetFunction.Match(duplicate.Cells(1, 1).Value, data.Cells(1, data_sheet.day_code).EntireColumn, 1)
On Error GoTo 0

If duplicate.Cells(1, 1).Value <> WorksheetFunction.Index(data.Cells(1, data_sheet.day_code).EntireColumn, search_up) Then
    Exit Sub
End If

Dim search_target_Row As Integer
Dim i As Integer
Dim j As Integer

i = 0
While data.Cells(search_up - i, 1) = duplicate.Cells(1, 1).Value
    j = 0
    While data.Cells(search_up - i, data_sheet.student_num_start + j).Value <> ""
        On Error GoTo error_process_2
        search_target_Row = WorksheetFunction.Match(data.Cells(search_up - i, data_sheet.student_num_start + j), duplicate.Cells(1, duplicate_sheet.student_num).EntireColumn, 1)
        On Error GoTo 0
            If data.Cells(search_up - i, data_sheet.student_num_start + j) = duplicate.Cells(search_target_Row, duplicate_sheet.student_num) Then
                duplicate.Cells(search_target_Row, duplicate_sheet.reserve_count) = duplicate.Cells(search_target_Row, duplicate_sheet.reserve_count) + 1
            Else
                duplicate.Rows(search_target_Row + 1).Insert
                duplicate.Cells(search_target_Row + 1, duplicate_sheet.student_num) = data.Cells(search_up - i, data_sheet.student_num_start + j)
                duplicate.Cells(search_target_Row + 1, duplicate_sheet.reserve_count) = duplicate.Cells(search_target_Row + 1, duplicate_sheet.reserve_count) + 1
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

