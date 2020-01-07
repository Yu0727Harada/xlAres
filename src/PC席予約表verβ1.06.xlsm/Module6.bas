Attribute VB_Name = "Module6"
Option Explicit

Public Sub export_date()

Dim law_data As Workbook
Dim export_data As Workbook

Set law_data = Workbooks(Application.ThisWorkbook.name)

Workbooks.Add
Set export_data = Workbooks(ActiveWorkbook.name)

Dim search_start_date As String
search_start_date = InputBox("データをエクスポートの日付の範囲の始点をyyyymmdd形式で入力してください。例）２０２０年１月７日→20200107")
Dim search_end_date As String
search_end_date = InputBox("データをエクスポートの日付の範囲の終点をyyyymmdd形式で入力してください。例）２０２０年１月７日→20200107")

Dim start_row As Integer
Dim end_row As Integer

start_row = WorksheetFunction.Match(CLng(search_start_date) - 1, law_data.Sheets("生データ").Cells(1, data_sheet.day_code).EntireColumn, 1)
If CLng(search_start_date) <> WorksheetFunction.Index(law_data.Sheets("生データ").Cells(1, data_sheet.day_code).EntireColumn, start_row) Then
    start_row = start_row + 1
End If
end_row = WorksheetFunction.Match(CLng(search_end_date), law_data.Sheets("生データ").Cells(1, data_sheet.day_code).EntireColumn, 1)

law_data.Sheets("生データ").Activate
law_data.Sheets("生データ").Range(Cells(start_row, data_sheet.day_code), Cells(end_row, data_sheet.day_code)).EntireRow.Copy
export_data.Sheets(1).Activate
export_data.Sheets(1).Range(Cells(1, 1), Cells(end_row - start_row, 1)).EntireRow.PasteSpecial Paste:=xlPasteValues
law_data.Sheets("生データ").Activate
law_data.Sheets("生データ").Range(Cells(start_row, data_sheet.day_code), Cells(end_row, data_sheet.day_code)).EntireRow.Delete (xlShiftUp)

export_data.Worksheets.Add
export_data.Sheets("Sheet2").Cells(1, data_sheet.day_code).Value = "予約日"
export_data.Sheets("Sheet2").Cells(1, data_sheet.cable_frag).Value = "ケーブル貸し出し"
export_data.Sheets("Sheet2").Cells(1, data_sheet.reserve_code).Value = "予約コード"
export_data.Sheets("Sheet2").Cells(1, data_sheet.seat_num).Value = "席番号"
export_data.Sheets("Sheet2").Cells(1, data_sheet.student_num_start).Value = "学年"
export_data.Sheets("Sheet2").Cells(1, data_sheet.student_num_start + 1).Value = "学科"
export_data.Sheets("Sheet2").Cells(1, data_sheet.time_zone).Value = "時間帯"
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim grade As Long
Dim subjuct_row As String

i = 1
j = 2

Do While export_data.Sheets(1).Cells(i, data_sheet.student_num_start).Value <> ""
    k = 0
    Do While export_data.Sheets("Sheet1").Cells(i, data_sheet.student_num_start + k).Value <> ""
        export_data.Sheets("Sheet2").Cells(j, data_sheet.day_code).Value = export_data.Sheets("Sheet1").Cells(i, data_sheet.day_code)
        export_data.Sheets("Sheet2").Cells(j, data_sheet.cable_frag).Value = export_data.Sheets("Sheet1").Cells(i, data_sheet.cable_frag)
        export_data.Sheets("Sheet2").Cells(j, data_sheet.reserve_code).Value = export_data.Sheets("Sheet1").Cells(i, data_sheet.reserve_code)
        export_data.Sheets("Sheet2").Cells(j, data_sheet.seat_num).Value = export_data.Sheets("Sheet1").Cells(i, data_sheet.seat_num)
        export_data.Sheets("Sheet2").Cells(j, data_sheet.time_zone).Value = export_data.Sheets("Sheet1").Cells(i, data_sheet.time_zone)
        
        'enter_year = (Mid(CLng(export_data.Sheets("Sheet1").Cells(i, data_sheet.student_num_start + k).Value), 1, 2) + 2000) - (Mid(CLng(export_data.Sheets("Sheet1").Cells(i, data_sheet.day_code).Value), 1, 4)) + 1
        grade = enter_year(CLng(export_data.Sheets("Sheet1").Cells(i, data_sheet.student_num_start + k).Value), CLng(export_data.Sheets("Sheet1").Cells(i, data_sheet.day_code).Value))
        export_data.Sheets("Sheet2").Cells(j, data_sheet.student_num_start).Value = grade
        Dim temp As Long
        temp = Mid(CLng(export_data.Sheets("Sheet1").Cells(i, data_sheet.student_num_start + k).Value), 3, 4)
        subjuct_row = WorksheetFunction.VLookup(temp, law_data.Sheets("学科コード表").Range("A2:B46"), 2, False)
        export_data.Sheets("Sheet2").Cells(j, data_sheet.student_num_start + 1).Value = subjuct_row
    
        k = k + 1
        j = j + 1
    Loop
    i = i + 1
Loop


End Sub

Function enter_year(ByVal student_number As Long, ByVal date_code As Long)

Dim year As Long

If Mid(date_code, 5, 2) < 4 Then
    year = Mid(date_code, 1, 4) - 1
Else
    year = Mid(date_code, 1, 4)
End If

enter_year = year - (Mid(student_number, 1, 2) + 2000) + 1

End Function
