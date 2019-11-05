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
Public passcord_input As Variant

Public Const passcord As String = 1907

Public Const time_sheet As String = "L2" '時刻セルの位置
Public Const date_sheet As String = "K2" '日付セルの位置
Public Const master_on_off As String = "T4" 'マスター入力モードのオンオフを記述してるセルの位置
Public Const cell_corsor_move As String = "T5" '強制カーソル移動オンオフを記述してるセルの位置
Public Const corsor_move_target As String = "B12" '強制カーソル移動の移動先
Public Const limit_reserve_count As String = "T14" '一日の予約上限数を入力しているセルの位置
Public Const limit_res_on_off As String = "T16" '予約制限モードのオンオフ
Public Const res_table_start_row As Integer = 4 '予約表の開始位置（左上セル）
Public Const res_table_start_colomn As Integer = 3 '予約表の開始位置
Public Const res_table_width_row As Integer = 5 '予約表の長さ＝席番号の数
Public Const res_table_width_colomn As Integer = 7 '予約表のながさ＝利用時間の区間数
Public Const now_shift_number_row As Integer = 7 'LAコントロール部分の現在のシフトNoを表示するセルの行の位置
Public Const now_shift_number_column As Integer = 20 '上の列の位置。現状はこの左に順に表示されます
Public Const on_time_output As String = "AA3"

Public Const now_shift_menber_profile_output_row As Integer = 5 'プロフィールを表示するセルの行
Public Const now_shift_menber_profile_output_column As Integer = 11 '上の列
Public Const now_shift_menber_profile_output_row_move As Integer = 3 '二人目を表示するときにいくつ移動した行に表示するか
Public Const now_shift_menber_profile_output_column_move As Integer = 0 '上の列バージョン
Public Const shift_profile_count As Integer = 2 '表示するプロフィールをいくつにするか

Public Const shift_table_number_start_row As Integer = 4 '勤務ナンバーの開始位置。長さは空白のセルが出るまで処理するので設定しなくてもよい。※Noに直下セルにに何か置くとそこまで処理します
Public Const shift_table_number_start_colomn As Integer = 1
Public Const shift_table_time_start_row As Integer = 4 '１３－１４などのシフトを入力するセルの開始位置。長さはNo列の長さまで処理する
Public Const shift_table_time_start_colomn As Integer = 3
Public Const shift_table_date_start_row As Integer = 2 '日付を入力している位置。これが空白になるまでシフトの読み込みを続ける


Enum shift_table
'読み込んだシフト表の列の位置を上から昇順で振り分け
勤務時間帯開始 = 1
勤務時間帯終了
勤務No
End Enum

Public Sub setting_sheet()
'シートをひらいたときに自動で実行されるプロシージャ

'シートの保護
Call Sheets("メイン").Protect(UserInterfaceOnly:=True)
Call sheet_color_check

End Sub

Public Sub setting_time()
'現在の時刻がコマごとにした場合いくつになるか入力するプロシージャ

Dim now_time As Date

'違うブックをひらいて作業している場合はメインシートが見つからないためエラーになるので、エラー回避
On Error GoTo sheet_cal_error
now_time = Sheets("メイン").Range(time_sheet).Value
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

'Timeを用いればコンピューター時計準拠に､
'変数now_timeを使えばセルでいじれます

Sheets("メイン").Range(on_time_output).Value = on_time

Exit Sub


sheet_cal_error:
Exit Sub
End Sub
Public Sub sheet_color_check()
'表に入力されているテキストに従ってセルの背景色を設定するプロシージャ

'On Error GoTo error
'now_time = time_sheet.Value
'On Error GoTo 0

Dim 色セルのRow As Integer
Dim 色セルのcolumn As Integer

色セルのRow = res_table_start_row
色セルのcolumn = res_table_start_colomn

Call setting_time

On Error GoTo Sheet_protect_error

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
                .Cells(色セルのRow, 色セルのcolumn).Interior.Color = RGB(180, 235, 250) '水色
    '           どれにも当てはまらない場合の色設定
            End If
        End If
        On Error GoTo 0
        
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
Sheet_protect_error:
MsgBox ("シートが保護されているため、セルの背景色を変更できません。マニュアルのエラー番号００２をみて対処してください")
Exit Sub

End Sub

Public Sub shift_check()
'現在のシフトを更新するべきか判断するプロシージャ

On Error GoTo sheet_cal_error
'Worksheets("メイン").EnableCalculation = False
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
'現在のシフトを取得して、シフトの変更があったらシフトを表示するセルのオブジェクトを削除してあらたにプロフィールを出力する

'Worksheets("メイン").EnableCalculation = False
Dim j As Integer
Dim now_date As Date
Dim search As Integer
Dim end_time As Date
Dim start_time As Date
Dim Shift() As Integer
Dim shp As Shape
Dim rng As Range
Dim k As Integer
Dim l As Integer
j = 0
ReDim Preserve Shift(0)
Dim shift_time_end As Range
Set shift_time_end = Sheets("シフト表").Columns(勤務時間帯終了)
Dim shift_row As Integer

   now_date = Date 'Dateはコンピューター上の日付
   On Error GoTo sheet_cal_error
    search = WorksheetFunction.Match(CDbl(now_date), shift_time_end, 1) + 1 '   CDblで型を変換しないとうまくmatch検索できない｡
    On Error GoTo 0
    If Int(now_date) <> Int(WorksheetFunction.Index(shift_time_end, search)) Then 'doble型だと時刻まで含、Int型なら日付のみになる
'            '日付が一致しない場合、すなわちその日のシフトがなかった場合の処理
'            k = 0
'            For k = 0 To 5 '同時に勤務する人数の最大値５までループを回してメインシートの番号を出力するところが空白でないなら０を入力
'                'If Cells(now_shift_number_row, now_shift_number_column + k).Value <> Shift(k) Then
'                '    Cells(now_shift_number_row, now_shift_number_column + k).Value = Shift(k)
'                If Cells(now_shift_number_row, now_shift_number_column + k).Value <> 0 Then
'                    Cells(now_shift_number_row, now_shift_number_column + k).Value = 0
'                End If
'            Next k
'
'            For k = 0 To shift_profile_count  '表示されているプロフィールを削除
'                Call shapes_delete(Sheets("メイン").Range(Cells(now_shift_menber_profile_output_row + k * now_shift_menber_profile_output_row_move, now_shift_menber_profile_output_column + k * now_shift_menber_profile_output_column_move), Cells(now_shift_menber_profile_output_row + k * now_shift_menber_profile_output_row_move, now_shift_menber_profile_output_column + k * now_shift_menber_profile_output_column_move)))
'            Next k
'
''            shift_row = WorksheetFunction.Match(0, Sheets("出力").Cells(1, 1).EntireColumn, 1)
''            If 0 <> WorksheetFunction.Index(Sheets("出力").Cells(1, 1).EntireColumn, shift_row) Then
''                MsgBox ("エラー番号２０２　番号が出力シートに存在しません。このまま処理を実行します")
''            Else
''                Sheets("出力").Cells(1, 2).CopyPicture
''                Sheets("メイン").Paste Cells(now_shift_menber_profile_output_row, now_shift_menber_profile_output_column)
''
'                Exit Sub
'            End If
    Else '日付が一致した場合、すなわち当日のシフトがにゅうりょくされていたばあい
        Do While now_date = Int(WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search)) '
            end_time = WorksheetFunction.Index(Sheets("シフト表").Range("B:B"), search) - now_date
            start_time = WorksheetFunction.Index(Sheets("シフト表").Range("A:A"), search) - now_date
            If now_time < end_time And now_time > start_time Then
                ReDim Preserve Shift(j + 1)
                Shift(j) = WorksheetFunction.Index(Sheets("シフト表").Range("C:C"), search)
                j = j + 1
                If j > 5 Then 'シフト人数が５人より多い場合はループを抜ける
                    Exit Do
                End If
            End If
            search = search + 1
        Loop

    End If
            
    Dim profile_count As Integer '表示したプロフィールの数を記録
    profile_count = 0
    
    For k = 0 To shift_profile_count  '表示されているプロフィールを削除
            On Error GoTo object_error
            Call shapes_delete(Sheets("メイン").Range(Cells(now_shift_menber_profile_output_row + k * now_shift_menber_profile_output_row_move, now_shift_menber_profile_output_column + k * now_shift_menber_profile_output_column_move), Cells(now_shift_menber_profile_output_row + k * now_shift_menber_profile_output_row_move, now_shift_menber_profile_output_column + k * now_shift_menber_profile_output_column_move)))
            On Error GoTo 0
    Next k
    
    If UBound(Shift) = 0 Then 'シフト配列の要素数が０かどうか
        shift_row = WorksheetFunction.Match(0, Sheets("出力").Cells(1, 1).EntireColumn, 1)
            If 0 = WorksheetFunction.Index(Sheets("出力").Cells(1, 1).EntireColumn, shift_row) Then 'シフトが０だったら番号０のプロフィールを表示。０のプロフィールがないなら表示しない
                Sheets("出力").Cells(shift_row, 2).CopyPicture
                Sheets("メイン").Paste Cells(now_shift_menber_profile_output_row, now_shift_menber_profile_output_column)
            End If
    Else
        'Call Quick_sort_single(Shift(), 0, UBound(Shift))
        For l = 0 To UBound(Shift) 'シフト配列の要素数だけ回す
            'If Cells(now_shift_number_row, now_shift_number_column + L).Value <> Shift(L) Then 'シフト番号の変化がないなら以下の操作はしない
                Cells(now_shift_number_row, now_shift_number_column + l).Value = Shift(l)
                
                If profile_count < shift_profile_count And Shift(l) <> 0 Then 'まだプロフィール表示数が設定以下かつシフト番号が０以外なら以下の処理を行う
                    shift_row = WorksheetFunction.Match(Shift(l), Sheets("出力").Cells(1, 1).EntireColumn, 1)
                    If Shift(l) <> WorksheetFunction.Index(Sheets("出力").Cells(1, 1).EntireColumn, shift_row) Then
                        MsgBox ("エラー番号２０２　番号が出力シートに存在しません。このまま処理を実行します")
                    Else
                        Sheets("出力").Cells(shift_row, 2).CopyPicture
                        Sheets("メイン").Paste Cells(now_shift_menber_profile_output_row + profile_count * now_shift_menber_profile_output_row_move, now_shift_menber_profile_output_column + profile_count * now_shift_menber_profile_output_column_move)
                        profile_count = profile_count + 1
                    End If
                End If
            'End If
        Next l
    End If
    
'            表示したプロフィールの数が足りていないようならすでに表示してあるところを削除
'    For k = profile_count To shift_profile_count   '表示されているプロフィールを削除
'        Call shapes_delete(Sheets("メイン").Range(Cells(now_shift_menber_profile_output_row + k * now_shift_menber_profile_output_row_move, now_shift_menber_profile_output_column + k * now_shift_menber_profile_output_column_move), Cells(now_shift_menber_profile_output_row + k * now_shift_menber_profile_output_row_move, now_shift_menber_profile_output_column + k * now_shift_menber_profile_output_column_move)))
'    Next k
'
'

Dim m As Integer
m = UBound(Shift) '前に表示していたシフトの人数が今のシフトの人数が多かったら前のシフトのコントロールパネルの要素数を削除
Do While Cells(now_shift_number_row, now_shift_number_column + m) <> ""
    Cells(now_shift_number_row, now_shift_number_column + m).Value = ""
    m = m + 1
Loop


Sheets("メイン").Range(corsor_move_target).Select
Exit Sub
sheet_cal_error:
search = 2
Resume Next
object_error:
Exit Sub
                
End Sub

Function shapes_delete(ByVal delete_area As Range)
'対象の範囲にある図形を削除。ただし図形の名前がstateの場合は削除しない

Call Sheets("メイン").Unprotect
For Each shp In Sheets("メイン").shapes
    Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
    If shp.Name <> "state" Then
        If Not (Intersect(rng, delete_area) Is Nothing) Then
            shp.Delete
        End If
    End If
Next

Call Sheets("メイン").Protect(UserInterfaceOnly:=True)

End Function


Public Sub recal()
'定期的にシートの再計算を行うためのプロシージャ
If Worksheets("メイン").EnableCalculation = False Then
    Worksheets("メイン").EnableCalculation = True
End If
Application.Calculate
'シートの再計算を行う
Call shift_check
tm = now() + TimeValue("00:01:00")
Application.OnTime EarliestTime:=tm, Procedure:="recal", Schedule:=True
'tm変数に一分後をセット
'ontime関数で一分後にまたrecalプロシージャを実行
'なぜかシートモジュールにrecalプロシージャを書くとうまくいかない

End Sub

Function passcord_inputform()

passcordform.Show

End Function
