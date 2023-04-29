Attribute VB_Name = "PenTool_Init"
Option Explicit
'
' Quck Access Toolbar for Powerpoint slideshow
'
' zbchristian 2023
'

Public PenTool As New PenToolClass
Sub InitializeApp()
    Set PenTool.App = Application
    PenTool.SlideHWnd = 0
End Sub

