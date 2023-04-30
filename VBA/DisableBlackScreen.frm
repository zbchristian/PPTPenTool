VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DisableBlackScreen 
   Caption         =   "Back to pesentation"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "DisableBlackScreen.frx":0000
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "DisableBlackScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.Height = 102
    Me.Width = 102
    PenTool.InitForm (Me.caption)
End Sub

Private Sub DisableBlackScreen_Click()
    PenTool.ToggleBlackScreen
End Sub


