VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PenToolbar 
   Caption         =   "Pen Properties"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   OleObjectBlob   =   "PenToolbar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "PenToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Quck Access Toolbar for Powerpoint slideshow
'
' zbchristian 2023
'

Private Const ToolBar_Height As Integer = 40

Private Sub MoveToolbar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PenTool.MoveToolbarSelect Button, X, Y
End Sub

Private Sub MoveToolbar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PenTool.MoveToolbar Button, X, Y
End Sub

Private Sub UserForm_Initialize()
    Me.Height = ToolBar_Height
    PenTool.FormCaption = Me.caption
    PenTool.InitHandle
    PenTool.InitForm (Me.caption)
End Sub

Private Sub CmdBlack_Click()
    PenTool.SetPenBlack
End Sub

Private Sub CmdBlue_Click()
    PenTool.SetPenBlue
End Sub

Private Sub CmdGreen_Click()
    PenTool.SetPenGreen
End Sub

Private Sub CmdRed_Click()
    PenTool.SetPenRed
End Sub

Private Sub LaserPointer_Click()
    PenTool.SetLaserPointer
End Sub

Private Sub Eraser_Click()
    PenTool.SetEraser
End Sub

Private Sub Marker_Click()
    PenTool.SetMarker
End Sub

Private Sub NewSlide_Click()
    PenTool.CreateNewSlide
End Sub

Private Sub PrevSlide_Click()
    PenTool.GotoPrevSlide
End Sub

Private Sub NextSlide_Click()
    PenTool.GotoNextSlide
End Sub

Private Sub AllSlides_Click()
    PenTool.ShowAllSlides
End Sub

Private Sub ExitSlideShow_Click()
    PenTool.ExitSlideShow
End Sub

Private Sub Turn_Click()
    PenTool.FormCaption = PenToolbarVert.caption
    Me.Hide
    PenToolbarVert.Show
    PenTool.SetActiveToolbar "vertical"
End Sub

Private Sub BlackScreen_Click()
    PenTool.ToggleBlackScreen
End Sub


