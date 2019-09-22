Attribute VB_Name = "Module5"
Option Explicit

Sub Quick_sort(ByRef data() As Variant, key As Integer, min As Integer, max As Integer)
'クイックソート行うプロシージャ

Dim i As Integer
Dim j As Integer
Dim R As Variant
Dim temp As Variant

i = min
j = max

R = (data(min, key) + data(max, key)) / 2

Do

    Do While data(i, key) < R
        i = i + 1
    Loop
    Do While data(j, key) > R
        j = j - 1
    Loop
    
    If i >= j Then Exit Do
    
    Dim k As Integer
    
    For k = LBound(data, 2) To UBound(data, 2)
    '入れ替えるのは並び変えるキーだけでなくすべての次元なのでループを回す
        temp = data(i, k)
        data(i, k) = data(j, k)
        data(j, k) = temp
    Next k
    
    i = i + 1
    j = j - 1

Loop

If min < i - 1 Then
    Call Quick_sort(data, 1, min, i - 1)
End If
If max > j + 1 Then
    Call Quick_sort(data, 1, j + 1, max)
End If

End Sub


