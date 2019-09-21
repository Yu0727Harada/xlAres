Attribute VB_Name = "Module5"
Option Explicit

Sub Quick_sort(ByRef Data() As Variant, key As Integer, min As Integer, max As Integer)

Dim i As Integer
Dim j As Integer
Dim R As Variant
Dim temp As Variant

i = min
j = max

R = (Data(min, key) + Data(max, key)) / 2

Do

    Do While Data(i, key) < R
        i = i + 1
    Loop
    Do While Data(j, key) > R
        j = j - 1
    Loop
    
    If i >= j Then Exit Do
    
    Dim k As Integer
    
    For k = LBound(Data, 2) To UBound(Data, 2)
    '入れ替えるのは並び変えるキーだけでなくすべての次元なのでループを回す
        temp = Data(i, k)
        Data(i, k) = Data(j, k)
        Data(j, k) = temp
    Next k
    
    i = i + 1
    j = j - 1

Loop

If min < i - 1 Then
    Call Quick_sort(Data, 1, min, i - 1)
End If
If max > j + 1 Then
    Call Quick_sort(Data, 1, j + 1, max)
End If

End Sub


