VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} edit_shift_form 
   Caption         =   "現在のシフトの編集"
   ClientHeight    =   4221
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5068
   OleObjectBlob   =   "edit_shift_form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "edit_shift_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim shift() As Integer
Dim shift_row() As Integer

Private Sub ComboBox2_Change()

End Sub

Private Sub CommandButton1_Click()

If ComboBox2.ListIndex = -1 Or ComboBox1.ListIndex = -1 Then
    MsgBox ("項目を選択してください")
    Exit Sub
End If

If ComboBox2.ListIndex = 0 Then

    Call Sheets("シフト表").Cells(shift_row(ComboBox1.ListIndex), 1).EntireRow.Delete(xlShiftUp)

ElseIf ComboBox2.ListIndex = 1 Then

    Sheets("シフト表").Cells(shift_row(ComboBox1.ListIndex), 2) = Date + Time
    
End If
Call shift_output_mainsheet(Time)
Unload edit_shift_form

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()

CommandButton1.SetFocus

With ComboBox2
    .AddItem "削除"
    .AddItem "退勤"
End With

Call get_shift(Time, Date, shift(), shift_row())

Dim i As Integer
Dim name_row As Integer
Dim name As String

For i = 0 To UBound(shift_row) - 1
    On Error GoTo data_nothing
    name_row = WorksheetFunction.Match(shift(i), Sheets("入力").Range("A:A"), 1)
    On Error GoTo 0
    If WorksheetFunction.Index(Sheets("入力").Range("A:A"), name_row) = shift(i) Then
        name = WorksheetFunction.Index(Sheets("入力").Range("D:D"), name_row)
    Else
        name = "No." + CStr(shift(i))
    End If
    With ComboBox1
        .AddItem (Format(CDate(WorksheetFunction.Index(Sheets("シフト表").Cells(1, 1).EntireColumn, shift_row(i))), "hh:mm") + "~" + Format(CDate(WorksheetFunction.Index(Sheets("シフト表").Cells(1, 2).EntireColumn, shift_row(i))), "hh:mm") + " " + name)
    End With
Next i

ComboBox1.ListRows = 5

Exit Sub
data_nothing:
name_row = 2
Resume Next

End Sub

