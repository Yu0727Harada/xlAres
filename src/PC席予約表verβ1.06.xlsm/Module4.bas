Attribute VB_Name = "Module4"
Option Explicit

Sub import_Shift()
'シフトを読み込むプロシージャ

Dim shift_year As range
Dim shift_month As range

Application.Calculation = xlCalculationManual

Dim wb As Workbook
Set wb = Workbooks(Application.ThisWorkbook.name)

Dim open_filepath As String
open_filepath = Application.GetOpenFilename(filefilter:="microsoft excelbook,*.xlsm", Title:="シフト表のエクセルファイルを選んでください")
If open_filepath = "False" Then
    Exit Sub
End If
Workbooks.Open open_filepath

Dim Shift_Filename As String
Shift_Filename = Dir(open_filepath)

Dim Shift_BookName As Workbook

Set Shift_BookName = Workbooks(Shift_Filename)

If Shift_Filename = Application.ThisWorkbook.name Then
    MsgBox ("シフト表のエクセルファイルと、元のエクセルファイルのファイル名が同じです。異なるファイル名に変更してください")
    Exit Sub
End If

'With Shift_BookName.Worksheets(1)
Set shift_year = Shift_BookName.Worksheets(1).range("C1") '読み込むシフト表の年度が入力されているセルを設定
Set shift_month = Shift_BookName.Worksheets(1).range("F1") '月が入力されているセルを設定
'End With

Dim i As Integer
Dim j As Integer
i = shift_table_time_start_colomn

Dim Shift_data_aray() As Variant
Dim Shift_data_num As Integer
Shift_data_num = 0

With Shift_BookName.Worksheets(1)
    Do While .Cells(shift_table_date_start_row, i) <> ""
    j = shift_table_number_start_row
        Do While .Cells(j, shift_table_number_start_colomn) <> ""
    
            If .Cells(j, i) <> "" Then
                ReDim Preserve Shift_data_aray(2, Shift_data_num)
                Dim year As Integer
                Dim month As Integer
                Dim dates As Integer
                year = shift_year.Value
                month = shift_month.Value
                dates = .Cells(shift_table_date_start_row, i).Value
                Dim yyyymmdd As Date
                yyyymmdd = Replace(Str(year) & "/" & Str(month) & "/" & Str(dates), " ", "")
'                Shift_data_aray(0, Shift_data_num) = yyyymmdd
'                なぜかreplace関数で空白スペースを置換しないと空白が入るので置換
                Dim tmp As Variant
                tmp = Split(.Cells(j, i), "-")
                Shift_data_aray(0, Shift_data_num) = yyyymmdd & " " & tmp(0) & ":00"
                Shift_data_aray(1, Shift_data_num) = yyyymmdd & " " & tmp(1) & ":00"
                Shift_data_aray(2, Shift_data_num) = .Cells(j, shift_table_number_start_colomn)
                
                Shift_data_num = Shift_data_num + 1
            End If
        j = j + 1
        Loop
    i = i + 1
    Loop
End With

Dim Shift_data_aray_trans() As Variant
ReDim Shift_data_aray_trans(UBound(Shift_data_aray, 2), 2)

Dim k As Integer
Dim date_temp As Date

For k = 0 To UBound(Shift_data_aray, 2)
    date_temp = CDate(Shift_data_aray(0, k))
    Shift_data_aray_trans(k, 0) = date_temp
    date_temp = CDate(Shift_data_aray(1, k))
    Shift_data_aray_trans(k, 1) = date_temp
    Shift_data_aray_trans(k, 2) = Shift_data_aray(2, k)
Next k

Call Quick_sort(Shift_data_aray_trans(), 1, 0, UBound(Shift_data_aray_trans, 1))

Dim shift_time_end As range
Set shift_time_end = wb.Worksheets("シフト表").Columns(勤務時間帯終了)
Dim search_up As Integer
Dim search_down As Integer
On Error GoTo data_nothing_up
    search_up = WorksheetFunction.Match(CDbl(Shift_data_aray_trans(0, 1)) - 1, shift_time_end, 1) + 1 '   CDblで型を変換しないとうまくmatch検索できない｡
On Error GoTo 0
On Error GoTo data_notihng_down
    search_down = WorksheetFunction.Match(CDbl(Shift_data_aray_trans(UBound(Shift_data_aray_trans, 1) - 1, 1)), shift_time_end, 1) + 1
On Error GoTo 0

Do While Int(Shift_data_aray_trans(UBound(Shift_data_aray_trans, 1) - 1, 1)) = Int(WorksheetFunction.Index(shift_time_end, search_down))
    search_down = search_down + 1
Loop

Dim Okcancel As Integer
'If Int(WorksheetFunction.Index(shift_time_end, search_up)) <> Int(WorksheetFunction.Index(shift_time_end, search_down)) Then
 If search_up <> search_down Then
    Okcancel = MsgBox("すでに読み込まれている期間のシフトです。以前のデータを上書きしますが、よろしいですか？", vbOKCancel)
    If Okcancel = 2 Then
        MsgBox ("処理を終了します")
        Exit Sub
    Else
        wb.Worksheets("シフト表").Activate
        wb.Worksheets("シフト表").range(Cells(search_up, 1), Cells(search_down - 1, 1)).EntireRow.Delete (xlShiftUp)
    End If
End If
With wb.Worksheets("シフト表")
    .Activate
    .range(Cells(search_up, 1), Cells(search_up + UBound(Shift_data_aray_trans, 1) - 1, 1)).EntireRow.Insert
    '.Resize(UBound(Shift_data_aray_trans, 1), 3).NumberFormatLocal = "G/標準"
    .range("A1").Offset(search_up - 1, 0).Resize(UBound(Shift_data_aray_trans, 1), 3) = Shift_data_aray_trans
End With

Application.Calculation = xlCalculationAutomatic

Exit Sub
data_nothing_up:
search_up = 2
Resume Next
data_notihng_down:
search_down = 2
Resume Next
End Sub
