VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_new_shift_form 
   Caption         =   "シフトの追加"
   ClientHeight    =   4635
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   7095
   OleObjectBlob   =   "add_new_shift_form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "add_new_shift_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub ComboBox2_Change()

End Sub

Private Sub CommandButton1_Click()

If ComboBox2.ListIndex = -1 Then
    MsgBox ("名前を選択してください")
    Exit Sub
End If

Dim end_time As Date

If ComboBox1.ListIndex = -1 Then
    end_time = DateAdd("n", 120, Date + Time)
ElseIf ComboBox1.ListIndex = 0 Then
    end_time = DateAdd("n", 15, Date + Time)
Else
    end_time = DateAdd("n", 30 * ComboBox1.ListIndex, Date + Time)
End If

Dim search As Integer
Dim shift_time_end As Range
Set shift_time_end = Worksheets("シフト表").Columns(shift_table.勤務時間帯終了)
On Error GoTo data_nothing
search = WorksheetFunction.Match(CDbl(end_time), shift_time_end, 1)
With Worksheets("シフト表")
    .Cells(search + 1, 2).EntireRow.Insert
    .Cells(search + 1, 1) = Date + Time
    .Cells(search + 1, 2) = end_time
    .Cells(search + 1, 3) = Sheets("入力").Cells(2 + ComboBox2.ListIndex, 1)
End With

MsgBox ("シフトに追加しました")
Call shift_output_mainsheet(Time)
Unload add_new_shift_form

Exit Sub

data_nothing:
search = 2

End Sub


Private Sub UserForm_Initialize()

CommandButton1.SetFocus

With ComboBox1
    .AddItem "15分勤務"
    .AddItem "30分勤務"
    .AddItem "60分勤務"
    .AddItem "90分勤務"
    .AddItem "120分勤務"
End With

Dim i As Integer
i = 2

With ComboBox2
    While Sheets("入力").Cells(i, 1) <> ""
        .AddItem Sheets("入力").Cells(i, 4).Value
        i = i + 1
    Wend
End With

ComboBox2.ListRows = 5

End Sub
