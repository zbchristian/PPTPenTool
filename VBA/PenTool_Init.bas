Attribute VB_Name = "PenTool_Init"
Option Explicit
'
' Quck Access Toolbar for Powerpoint slideshow
'
' zbchristian 2023
'
Public PenToolEnabled As Boolean

Public PenTool As New PenToolClass
Sub InitializeApp()
    PenToolEnabled = True
    Set PenTool.App = Application
    PenTool.Init
End Sub

Sub DisableToolbar()
    PenToolEnabled = False
    PenTool.Init
End Sub

