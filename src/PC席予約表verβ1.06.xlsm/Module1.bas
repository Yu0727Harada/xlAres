Attribute VB_Name = "Module1"
Public 予約日 As Long
Public resreve_day As Long
Public 時間帯 As Integer
Public 席番号 As Integer
Public 連続可能か As String
Public frag As Integer
Public number_valid As Integer
Public tm As Double
Public on_time As Integer
Public Const passcord As String = 1907
Public time_sheet As Range
Public res_table_start_row As Integer
Public res_table_start_colomn As Integer
Public res_table_width_row As Integer
Public res_table_width_colomn As Integer
Public shift_table_number_start_row As Integer
Public shift_table_number_start_colomn As Integer
Public shift_table_time_start_row As Integer
Public shift_table_time_start_colomn As Integer
Public shift_table_date_start_row As Integer

Enum shift_table
'読み込んだシフト表の列の位置を上から昇順で振り分け
勤務時間帯開始 = 1
勤務時間帯終了
勤務No
End Enum

Public Sub setting_time()

Dim now_time As Date

'違うブックをひらいて作業している場合はメインシートが見つからないためエラーになるので、エラー回避
On Error GoTo sheet_cal_error
now_time = time_sheet.Value
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
Public Sub sheet_color_check()
'On Error GoTo error
'now_time = time_sheet.Value
'On Error GoTo 0

Dim 色セルのRow As Integer
Dim 色セルのcolumn As Integer

色セルのRow = res_table_start_row
色セルのcolumn = res_table_start_colomn

Call setting_time

With Sheets("メイン")
Do While 色セルのcolumn < res_table_start_colomn + res_table_width_colomn
    Do While 色セルのRow < res_table_start_row + res_table_width_row
        If on_time >= 色セルのcolumn And .Range("K2") = Date Then
            If .Cells(色セルのRow, 色セルのcolumn).Text = "予約済" Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(104, 109, 37) '黄色（影）
            ElseIf InStr(.Cells(色セルのRow, 色セルのcolumn).Text, "貸出中") > 0 Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(104, 73, 37) 'オレンジ（影）
            ElseIf .Cells(色セルのRow, 色セルのcolumn).Text = "" Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(104, 115, 123) '影
            Else
        '        If Cells(色セルのRow, 色セルのcolumn).Text <> "" And Cells(色セルのRow, 色セルのcolumn).Text <> "予約済" And Cells(色セルのRow, 色セルのcolumn).Text <> "予約済(貸出中)" Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(73, 106, 121) '水色（影）
    '           どれにも当てはまらない場合の色設定
            End If
        Else
            If .Cells(色セルのRow, 色セルのcolumn).Text = "予約済" Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(255, 240, 76) '黄色
            ElseIf InStr(.Cells(色セルのRow, 色セルのcolumn).Text, "貸出中") > 0 Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(255, 160, 76) 'オレンジ
            ElseIf .Cells(色セルのRow, 色セルのcolumn).Text = "" Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = xlNone '透明
            Else
        '        If Cells(色セルのRow, 色セルのcolumn).Text <> "" And Cells(色セルのRow, 色セルのcolumn).Text <> "予約済" And Cells(色セルのRow, 色セルのcolumn).Text <> "予約済(貸出中)" Then
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(180, 235, 250) '水色
    '           どれにも当てはまらない場合の色設定
            End If
        End If
        
        色セルのRow = 色セルのRow + 1
    
'    席番号１～５までの中身をみて計算
    
    Loop

色セルのRow = res_table_start_row
色セルのcolumn = 色セルのcolumn + 1
'席番号は４列目に戻して次の時間帯の計算に移る

Loop

End With

Exit Sub

'error:
'
'Exit Sub

End Sub

Public Sub shift_check()
On Error GoTo sheet_cal_error
Worksheets("メイン").EnableCalculation = False
Dim now_time As Date
now_time = Time 'TimeはＰＣ上の時刻
On Error GoTo 0

Dim i As Integer

'今の時間が３０分区切りの近辺か判断するif文。now time変数にはシリアル数が入っている。シリアル数は一日＝１なので、３０分は１／４８。
'iに0から48まで代入してすべての時刻において００分かあるいは３０分からその後の1/24/60（この場合１分）の間に今の時刻が入っていないか調べる。
'毎分シフトをチェックすると重すぎる気がしたので30分区切りにした｡実行しても大して処理の重さは変わらない気もする。
'もっと賢い書き方がある気がするがdoble型で約数かどうかの判断するのが怖かったので愚直に実装した。
For i = 0 To 48
If now_time > i * 1 / 48 And now_time < i * 1 / 48 + 1 / 24 / 60 Then
    Call shift_output_mainsheet(now_time)
    Exit For
End If
Next i

Worksheets("メイン").EnableCalculation = True
sheet_cal_error:
Exit Sub
End Sub

Public Sub shift_output_mainsheet(ByVal now_time As Date)
Worksheets("メイン").EnableCalculation = False
Dim j As Integer
Dim now_date As Date
Dim search As Integer
Dim end_time As Date
Dim start_time As Date
Dim Shift(4) As Integer
Dim shp As Shape
Dim rng As Range
Dim k As Integer
Dim L As Integer
j = 0
Dim shift_time_end As Range
Set shift_time_end = Sheets("シフト表").Columns(勤務時間帯終了)

   now_date = Date 'Dateはコンピューター上の日付
   On Error GoTo sheet_cal_error
    search = WorksheetFunction.Match(CDbl(now_date), shift_time_end, 1) + 1 '   CDblで型を変換しないとうまくmatch検索できない｡
    On Error GoTo 0
    If Int(now_date) <> Int(WorksheetFunction.Index(shift_time_end, search)) Then 'doble型だと時刻まで含、Int型なら日付のみになる
        
            k = 0
            For k = 0 To 1
                If Cells(7, 20 + k).Value <> Shift(k) Then
                    Cells(7, 20 + k).Value = Shift(k)
                    Call shapes_delete(Sheets("メイン").Range(Cells(5 + L * 3, 11), Cells(5 + L * 3, 11)))
'                    For Each shp In Sheets("メイン").shapes
'                        Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
'                        If Not (Intersect(rng, Sheets("メイン").Range(Cells(5 + k * 3, 11), Cells(5 + k * 3, 11))) Is Nothing) Then
'                            shp.Delete
'                        End If
'                    Next
                End If
            Next k
    Exit Sub
    Else
        Do While now_date = Int(WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search))
            end_time = WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search) - now_date
            start_time = WorksheetFunction.Index(Sheets("シフト表").Range("A:A"), search) - now_date
            If now_time < end_time And now_time > start_time Then
                Shift(j) = WorksheetFunction.Index(Sheets("シフト表").Range("C:C"), search)
                j = j + 1
            End If
            search = search + 1
        Loop

            L = 0
            For L = 0 To 1
                If Cells(7, 20 + L).Value <> Shift(L) Then
                    Cells(7, 20 + L).Value = Shift(L)
                    
                    Call shapes_delete(Sheets("メイン").Range(Cells(5 + L * 3, 11), Cells(5 + L * 3, 11)))
'                    For Each shp In Sheets("メイン").shapes
'                        Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
'                        If Not (Intersect(rng, Sheets("メイン").Range(Cells(5 + L * 3, 11), Cells(5 + L * 3, 11))) Is Nothing) Then
'                            shp.Delete
'                        End If
'                    Next
                    Sheets("出力").Cells(Shift(L) + 1, 2).CopyPicture
                    Sheets("メイン").Paste Cells(5 + L * 3, 11)
                End If
            Next L

    End If

Sheets("メイン").Cells(12, 2).Select
Exit Sub
sheet_cal_error:
search = 2
Resume Next
                
End Sub

Function shapes_delete(ByVal delete_area As Range)
For Each shp In Sheets("メイン").shapes
    Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
    If shp.Name <> "state" Then
        If Not (Intersect(rng, delete_area) Is Nothing) Then
            shp.Delete
        End If
    End If
Next


End Function


Public Sub recal()

Application.Calculate
'シートの再計算を行う
Call shift_check
tm = now() + TimeValue("00:01:00")
Application.OnTime EarliestTime:=tm, Procedure:="recal", Schedule:=True
'tm変数に一分後をセット
'ontime関数で一分後にまたrecalプロシージャを実行
'なぜかシートモジュールにrecalプロシージャを書くとうまくいかない

End Sub

