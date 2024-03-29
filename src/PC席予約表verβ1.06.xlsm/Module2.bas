Attribute VB_Name = "Module2"
Option Explicit

Sub profile_update()
'入力シートのプロフィール更新ボタン

Dim inputsheet As Object
Dim outputsheet As Object
Set inputsheet = ThisWorkbook.Worksheets("入力")
Set outputsheet = ThisWorkbook.Worksheets("出力")

'Worksheets("メイン").EnableCalculation = False

'If outputsheet.Cells(1, 1) = 0 Then
'    outputsheet.Activate
'    Call shapes_delete(outputsheet.range(Cells(1, 1), Cells(1, 2)), outputsheet, False)
'    Call delete_sheet_data(2, 2, outputsheet)
'Else
'    outputsheet.Cells.Clear
'    Dim shp As Shape
'    For Each shp In outputsheet.shapes
'        shp.Delete
'    Next shp
'    outputsheet.Cells(1, 1) = 0
'End If

Call Worksheets("入力").Range("A:F").Sort(key1:=Worksheets("入力").Cells(1, 1).EntireColumn, order1:=xlAscending, Header:=xlYes)

Dim i As Integer
Dim j As Integer
i = 1
j = 2
Do While inputsheet.Cells(j, 1) <> ""
    If i = 1 Then
        If outputsheet.Cells(i, 1).Value = 0 And inputsheet.Cells(j, 1) <> 0 Then
            outputsheet.Activate
            Call shapes_delete(outputsheet.Range(Cells(1, 1), Cells(1, 2)), outputsheet, False)
            Call delete_sheet_data(2, 2, outputsheet)
            GoTo skip_insert_0_profile
        Else
            outputsheet.Cells.Clear
            Dim shp As Shape
            For Each shp In outputsheet.shapes
                shp.Delete
            Next shp
        End If
    End If
    
    outputsheet.Range(Cells(i, 1), Cells(i, 2)).EntireRow.RowHeight = profile_height
    
    outputsheet.Cells(i, 1).Value = inputsheet.Cells(j, 1).Value
    Dim cellT As Variant
    Dim cellL As Variant
    Dim cellW As Variant
    Dim cellH As Variant
    
    With outputsheet.Cells(i, 2)
        cellT = .Top
        cellL = .Left
        cellW = .Width
        cellH = .Height
    End With
    
    Dim T As Variant
    Dim L As Variant
    Dim W As Variant
    Dim H As Variant
    
    T = cellT + cellH * 0.02 '名前と所属を表示するテキストボックスの左上の位置のｙ軸方向の位置。cellTは表示するセルの左上の位置。cellHはセルの高さ。かける数字を調整することで位置を調整できる
    L = cellL + cellW * 0.45 '上記のｘ軸方向の位置。かける数字を調整することで位置を調整できる
    W = cellW / 2 'テキストボックスの幅の大きさ。
    H = cellH - cellH / 4 'テキストボックスの高さの大きさ
    
    With outputsheet.shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=L, Top:=T, Width:=W, Height:=H)
        .TextFrame.Characters.Text = inputsheet.Cells(j, 2) & vbLf & inputsheet.Cells(j, 3) & vbLf & inputsheet.Cells(j, 4)
        .Fill.Visible = False
        .Line.Visible = False
            With .TextFrame.Characters.Font  'テキストボックスのフォントの設定
                        .Size = 14
                        .name = "BIZ UDゴシック"
            End With
    End With
    
    Dim T2 As Variant
    Dim L2 As Variant
    Dim W2 As Variant
    Dim H2 As Variant
    
    T2 = cellT + cellH * 0.55 'コメントのテキストボックスの左上の位置のｙ軸方向の位置。cellTは表示するセルの左上の位置。cellHはセルの高さ。かける数字を調整することで位置を調整できる
    L2 = cellL '上記のｘ軸方向の位置。かける数字を調整することで位置を調整できる
    W2 = cellW 'テキストボックスの幅の大きさ。
    H2 = cellH * 0.45 'テキストボックスの高さの大きさ
    
    With outputsheet.shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=L2, Top:=T2, Width:=W2, Height:=H2)
        .TextFrame.Characters.Text = inputsheet.Cells(j, 6)
        .Fill.Visible = False
        .Line.Visible = False
            With .TextFrame.Characters.Font 'フォントの設定
                        .Size = 10
                        .name = "BIZ UDゴシック"
            End With
    End With
    
    Dim picfile_path As String
    
'        Dim pic As Object
    If inputsheet.Cells(j, 5).Value = "" Then
        inputsheet.Cells(j, 5).Value = "\Noimage.png"
        'picfile_path = "�C:\Users\haradayuii\Documents\GitHub\PCresrveSystem\bin\Noimage.png"
    Else

    End If
        picfile_path = ThisWorkbook.Path + inputsheet.Cells(j, 5).Value
'            Set pic = LoadPicture(picfile_path)
        Dim pic_W As Variant
        Dim pic_H As Variant
        Dim trim_T As Variant
        Dim trim_L As Variant
        Dim trim_R As Variant
        Dim trim_B As Variant
'            pic_H = pic.Height * 0.0378
'            pic_W = pic.Width * 0.0378
        
        outputsheet.Activate
'            なんかよくわからないけど､出力シートをアクティブにしないとうまくいかない｡たぶんselectionで処理しているからだと思うけど､shapeオブジェクト難解でよくわからない
        On Error GoTo insert_error
        ActiveSheet.Pictures.Insert(picfile_path).Select
        On Error GoTo 0
        With Application.CommandBars
            If .GetEnabledMso("PictureResetAndSize") = True Then .ExecuteMso "PictureResetAndSize"
        End With
'            画像によってエクセルに挿入した時点で縦横が逆になっている場合がある｡（おそらくiPhoneなどの写真アプリでの編集のせい）その場合、以下の正方形に加工する処理で問題が発生するので、画像のリセットをかけて、元に戻す処理。
        
        pic_H = Selection.Height
        pic_W = Selection.Width
        If pic_H > pic_W Then
            trim_L = 0
            trim_R = 0
            trim_T = (pic_H - pic_W) / 2
            trim_B = (pic_H - pic_W) / 2
            With Selection.ShapeRange.PictureFormat
                .CropTop = trim_T
                .CropLeft = trim_L
                .CropRight = trim_R
                .CropBottom = trim_B
            End With
        ElseIf pic_H < pic_W Then
            trim_T = 0
            trim_B = 0
            trim_L = (pic_W - pic_H) / 2
            trim_R = (pic_W - pic_H) / 2
            With Selection.ShapeRange.PictureFormat
                .CropTop = trim_T
                .CropLeft = trim_L
                .CropRight = trim_R
                .CropBottom = trim_B
            End With

        End If
'            MsgBox CLng(Selection.Width) & "*" & CLng(Selection.Height)
        Selection.ShapeRange.AutoShapeType = msoShapeOval '画像の形
        Selection.Top = cellT + cellH * 0.05 '画像の左上のｙ軸の位置
        Selection.Left = cellL + cellL * 0.1 '画像の左上のｘ軸の位置
        Selection.Width = cellW * 0.37 '画像の幅
    
skip_insert_picture:
    j = j + 1
skip_insert_0_profile:
    i = i + 1
Loop

If i = 1 Then '何もデータがなかった場合の処理
    outputsheet.Activate
    Call shapes_delete(outputsheet.Range(Cells(1, 1), Cells(1, 2)), outputsheet, False)
    Call delete_sheet_data(2, 2, outputsheet)
End If


MsgBox ("出力完了しました")

Exit Sub

error:
MsgBox ("Noimageのファイルパスが間違ってるみたいなので直してください")
Resume Next

insert_error:
MsgBox ("入力シートに入力されている画像ファイルパスが間違っているか、Noimageファイルがエクセルファイルの入っているフォルダにありません。")
GoTo skip_insert_picture

End Sub
