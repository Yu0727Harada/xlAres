Attribute VB_Name = "Module2"
Option Explicit

Sub profile_update()
'入力シートのプロフィール更新ボタン

Dim inputsheet As Object
Dim outputsheet As Object
Set inputsheet = Worksheets("入力")
Set outputsheet = Worksheets("出力")

Worksheets("メイン").EnableCalculation = False

Dim shp As Shape

For Each shp In outputsheet.shapes
    shp.Delete
Next shp

Dim i As Integer
i = 2
    Do While inputsheet.Cells(i, 1) <> ""
        outputsheet.Cells(i, 1).Value = inputsheet.Cells(i, 1).Value
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
            .TextFrame.Characters.Text = inputsheet.Cells(i, 2) & vbLf & inputsheet.Cells(i, 3) & vbLf & inputsheet.Cells(i, 4)
            .Fill.Visible = False
            .Line.Visible = False
                With .TextFrame.Characters.Font  'テキストボックスのフォントの設定
                            .Size = 14
                            .Name = "源ノ角ゴシック JP"
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
            .TextFrame.Characters.Text = inputsheet.Cells(i, 6)
            .Fill.Visible = False
            .Line.Visible = False
                With .TextFrame.Characters.Font 'フォントの設定
                            .Size = 10
                            .Name = "源ノ角ゴシック JP"
                End With
        End With
        
        Dim picfile_path As String
'        Dim pic As Object
        picfile_path = inputsheet.Cells(i, 5).Value
        If picfile_path = "" Then
            MsgBox ("画像のパスがありません")
        Else
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
            ActiveSheet.Pictures.Insert(picfile_path).Select
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
        End If
                
        i = i + 1
    Loop

Worksheets("メイン").EnableCalculation = True
End Sub
