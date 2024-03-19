VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Обновить"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2850
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public i As Byte

Function choises() As Byte
    i = 0
    Menu.Show
    choises = i
End Function

Private Sub both_Click()
    i = 2
    Menu.Hide
End Sub

Private Sub Only_LiArt_Click()
    i = 1
    Menu.Hide
End Sub
