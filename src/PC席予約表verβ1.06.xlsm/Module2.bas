Attribute VB_Name = "Module2"
Option Explicit

Sub profile_update()
Dim inputsheet As Object
Dim outputsheet As Object
Set inputsheet = Worksheets("入力")
Set outputsheet = Worksheets("出力")

Worksheets("メイン").EnableCalculation = False

Dim i As Integer
i = 2
    Do While inputsheet.Cells(i, 1) <> ""
        outputsheet.Cells(i - 1, 1).Value = inputsheet.Cells(i, 1).Value
        Dim cellT As Variant
        Dim cellL As Variant
        Dim cellW As Variant
        Dim cellH As Variant
        
        With outputsheet.Cells(i - 1, 2)
            cellT = .Top
            cellL = .Left
            cellW = .Width
            cellH = .Height
        End With
        
        Dim T As Variant
        Dim L As Variant
        Dim W As Variant
        Dim H As Variant
        
        T = cellT + cellH / 10
        L = cellL + cellW / 2
        W = cellW / 2
        H = cellH - cellH / 4
        
        With outputsheet.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=L, Top:=T, Width:=W, Height:=H)
            .TextFrame.Characters.Text = inputsheet.Cells(i, 2) & vbLf & inputsheet.Cells(i, 3) & vbLf & inputsheet.Cells(i, 4)
            .Fill.Visible = False
            .Line.Visible = False
                With .TextFrame.Characters.Font
                            .Size = 18
                            .Name = "源ノ角ゴシック JP"
                End With
        End With
        
        Dim T2 As Variant
        Dim L2 As Variant
        Dim W2 As Variant
        Dim H2 As Variant
        
        T2 = cellT + cellH * 0.75
        L2 = cellL
        W2 = cellW
        H2 = cellH * 0.25
        With outputsheet.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=L2, Top:=T2, Width:=W2, Height:=H2)
            .TextFrame.Characters.Text = inputsheet.Cells(i, 6)
            .Fill.Visible = False
            .Line.Visible = False
                With .TextFrame.Characters.Font
                            .Size = 12
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
            Selection.ShapeRange.AutoShapeType = msoShapeOval
            Selection.Top = cellT + cellH * 0.1
            Selection.Left = cellL + cellL * 0.2
            Selection.Width = cellW * 0.4
        End If
                
        i = i + 1
    Loop

Worksheets("メイン").EnableCalculation = True
End Sub
