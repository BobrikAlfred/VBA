VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ND 
   Caption         =   "Проверка вычислений"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3615
   OleObjectBlob   =   "ND.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "ND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub userform_Initialize()
    Label1.Caption = "Проверьте правильность вычислений и нажмите"
End Sub

Private Sub CommandButton1_Click()
    ND.Hide
    Call Leftovers.FPTA
End Sub
