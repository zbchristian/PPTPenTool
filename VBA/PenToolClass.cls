VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PenToolClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Quck Access Toolbar for Powerpoint slideshow
'
' zbchristian 2023
'
Const GWL_STYLE As Integer = -16
Const GWL_EXSTYLE As Integer = -20
Const WS_CAPTION As Long = &HC00000
Const WS_BORDER As Long = &H800000

Private Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

Private Declare PtrSafe Function SetActiveWindow Lib "user32.dll" _
    (ByVal hWnd As LongPtr) As LongPtr

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Private Declare PtrSafe Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIndex As Long) As Long

Private Declare PtrSafe Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As LongPtr, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
   ByVal hWnd As LongPtr, _
   hWndInsertAfter As LongPtr, X As Long, Y As Long, _
   cx As Long, cy As Long, uFlags As Long) As Boolean

Const SWP_NOZORDER As Long = &H4
Const SWP_SHOWWINDOW As Long = &H40

Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal hWnd As LongPtr, lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum States
    isPen = 0
    isLaserPointer = 1
    isHighlighter = 2
    isEraser = 3
    isArrow = 4
End Enum

Public SlideHWnd As LongPtr

Private StateOfPen As States

Public FormCaption As String

Private ActiveToolbar As String

Private XPos As Single
Private YPos As Single

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private Sub App_SlideShowBegin(ByVal Wn As SlideShowWindow)
    If PenToolEnabled Then
        PenToolbar.Show (vbModeless)
        With ActivePresentation.SlideShowSettings.Run.View
            .LaserPointerEnabled = False
            .PointerColor.RGB = RGB(0, 0, 0)
            .PointerType = ppSlideShowPointerPen
        End With
    End If
End Sub

Public Sub Init()
    PenTool.SlideHWnd = 0
    ActiveToolbar = "horizontal"
End Sub

Public Sub InitHandle()
    Dim hWnd As LongPtr
     SlideHWnd = 0
     hWnd = FindWindow("ThunderDFrame", FormCaption)
     hWnd = GetParent(hWnd)
     If hWnd > 0 Then
        SlideHWnd = hWnd
     End If
End Sub

Public Sub InitForm(ByVal caption As String)
    Dim hWnd As LongPtr
    Dim lStyle As Long
    Dim X As Long, Y As Long, W As Long, H As Long
    Dim R As RECT
    
    hWnd = FindWindow("ThunderDFrame", caption)
    
    ' remove title bar and border
    lStyle = GetWindowLong(hWnd, GWL_STYLE)
    lStyle = lStyle And Not WS_CAPTION And Not WS_BORDER
    SetWindowLong hWnd, GWL_STYLE, lStyle
    SetWindowLong hWnd, GWL_EXSTYLE, 0
    
    ' get x,y, width and height of form
    GetWindowRect hWnd, R
    X = R.Left
    Y = R.Top
'    W = R.Right - R.Left
    W = 250
    H = R.Bottom - R.Top
    SetWindowPos hWnd, 0, X, Y, W, H, SWP_SHOWWINDOW
    
    DrawMenuBar hWnd
End Sub

Public Sub SetActiveToolbar(toolbar As String)
    ActiveToolbar = toolbar
End Sub

Public Sub SendESC()
    SendCmd ("{ESC}")
End Sub

Public Sub ExitSlideShow()
    ActivePresentation.SlideShowWindow.View.Exit
End Sub

Private Sub SetState(ByVal st As States)
    StateOfPen = st
End Sub

Private Sub SendCmd(ByVal c As String)
    If SlideHWnd > 0 Then
       SetActiveWindow (SlideHWnd)
       SendKeys c, False
    End If
End Sub

Private Sub EnablePen()
    If StateOfPen = isHighlighter Then
' Getting back from the highlighter to pen is not working correctly in Office 2016
' Workaround : switch to Laserpointer first and then to Pen
        SendCmd ("^l")
        Sleep 100
        SendCmd "^p"
        SetState (isPen)
    End If
End Sub

Public Sub SetDefaultCursor()
' There is a problem to switch from highlighter to Arrow cursor
' Switch first to laser pointer
    SendCmd ("^l")
    Sleep 100
    SendCmd ("^a")
    SetState (isArrow)
End Sub

Public Sub SetPenBlack()
    EnablePen
    With ActivePresentation.SlideShowSettings.Run.View
        .LaserPointerEnabled = False
        .PointerColor.RGB = RGB(0, 0, 0)
        .PointerType = ppSlideShowPointerPen
    End With
End Sub

Public Sub SetPenBlue()
    EnablePen
    With ActivePresentation.SlideShowSettings.Run.View
        .LaserPointerEnabled = False
        .PointerColor.RGB = RGB(0, 0, 255)
        .PointerType = ppSlideShowPointerPen
    End With
End Sub

Public Sub SetPenGreen()
    EnablePen
    With ActivePresentation.SlideShowSettings.Run.View
        .LaserPointerEnabled = False
        .PointerColor.RGB = RGB(0, 255, 0)
        .PointerType = ppSlideShowPointerPen
    End With
End Sub

Public Sub SetPenRed()
    EnablePen
    With ActivePresentation.SlideShowSettings.Run.View
        .PointerColor.RGB = RGB(255, 0, 0)
        .PointerType = ppSlideShowPointerPen
    End With
End Sub

Public Sub SetLaserPointer()
    SendCmd ("^l")
' VBA code works, but there is interference wih the Highlighter
'    EnablePen
'    With ActivePresentation.SlideShowSettings.Run.View
'        .PointerType = ppSlideShowPointerArrow
'        .LaserPointerEnabled = True
'    End With
    SetState (isLaserPointer)
End Sub

Public Sub SetEraser()
    SendCmd ("^e")
    SetState (isEraser)
' VBA code works, but does not show the eraser icon
'    With ActivePresentation.SlideShowSettings.Run.View
'        .PointerColor.RGB = RGB(255, 0, 0)
'        .PointerType = ppSlideShowPointerEraser
'        .LaserPointerEnabled = True
'    End With
End Sub

Public Sub SetMarker()
    SendCmd ("^i")
    SetState (isHighlighter)
End Sub

Public Sub CreateNewSlide()
    Dim lCurrentSlide As Long
    lCurrentSlide = SlideShowWindows(1).View.Slide.SlideIndex
    Call ActivePresentation.Slides.AddSlide(lCurrentSlide + 1, _
        ActivePresentation.SlideMaster.CustomLayouts(7))
    SlideShowWindows(1).View.GotoSlide (lCurrentSlide + 1)
    SetPenBlack
End Sub

Public Sub GotoPrevSlide()
    SetPenBlack
    SendCmd ("{UP}")
End Sub

Public Sub GotoNextSlide()
    SetPenBlack
    SendCmd ("{DOWN}")
End Sub

Public Sub ShowAllSlides()
    SetPenBlack
    SendCmd ("^-")
End Sub
    
Private Sub ShowHideToolbar(ByVal doShow As Boolean)
    If ActiveToolbar = "horizontal" Then
        If doShow Then
            PenToolbar.Show
        Else
            PenToolbar.Hide
        End If
    Else
        If doShow Then
            PenToolbarVert.Show
        Else
            PenToolbarVert.Hide
        End If
    End If
End Sub

Public Sub ToggleBlackScreen()
    With SlideShowWindows(1).View
        If .state = ppSlideShowBlackScreen Then
            .state = ppSlideShowRunning
            DisableBlackScreen.Hide
            ShowHideToolbar (True)
            SetPenBlack
        Else
            SetPenBlack
            .state = ppSlideShowBlackScreen
            ShowHideToolbar (False)
            DisableBlackScreen.Show
            ActivePresentation.SlideShowSettings.Run.View.PointerType = ppSlideShowPointerArrow
        End If
    End With
End Sub

Public Sub MoveToolbarSelect(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        XPos = X
        YPos = Y
    End If
End Sub

Public Sub MoveToolbar(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    If Button And 1 Then
        PenToolbar.Left = PenToolbar.Left + (X - XPos)
        PenToolbar.Top = PenToolbar.Top + (Y - YPos)
        PenToolbarVert.Left = PenToolbarVert.Left + (X - XPos)
        PenToolbarVert.Top = PenToolbarVert.Top + (Y - YPos)
    End If
End Sub
