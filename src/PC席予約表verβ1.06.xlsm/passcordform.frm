VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} passcordform 
   Caption         =   "パスコードの入力"
   ClientHeight    =   3563
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5558
   OleObjectBlob   =   "passcordform.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "passcordform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

passcord_input = TextBox1.Text
Unload passcordform

End Sub

Private Sub UserForm_Initialize()
TextBox1.PasswordChar = "*"
End Sub
