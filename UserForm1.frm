VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ロギング中"
   ClientHeight    =   885
   ClientLeft      =   9960.001
   ClientTop       =   3870
   ClientWidth     =   2630
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'フォームの閉じるボタンの無効化
    If CloseMode = 0 Then Cancel = True
End Sub
