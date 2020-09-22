VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
Public WithEvents lblCaption As Label
Attribute lblCaption.VB_VarHelpID = -1
Public WithEvents img As Image
Attribute img.VB_VarHelpID = -1
Public fMax As FormControl.max
Public fMin As FormControl.Min
Public fClose As FormControl.Close
Public fTray As FormControl.Tray
Public WithEvents Pic As FormControl.Pic
Attribute Pic.VB_VarHelpID = -1
Public StatusBar As Shape
Public lblStatusBar As Label
Dim rs As Boolean
Dim aX As Integer, aY As Integer
Dim hr As Long
Private Sub frm_Load()
    frm.AutoRedraw = True
    frm.ScaleMode = vbPixels
    frm.Cls
    Dim lStyle As Long
    'Delete Caption Bar And Border Of The Form
    lStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    lStyle = lStyle And Not (WS_BORDER Or WS_CAPTION Or WS_DLGFRAME Or WS_THICKFRAME)
    SetWindowLong frm.hwnd, GWL_STYLE, lStyle
    
    'Shape Up The Form
    hr = CreateRoundRectRgn(0, 0, frm.ScaleWidth, frm.ScaleHeight, 10, 10)
    SetWindowRgn frm.hwnd, hr, True
    SetWindowPos frm.hwnd, HWND_TOP, frm.ScaleLeft, frm.ScaleTop, frm.ScaleWidth, frm.ScaleHeight, SWP_NOMOVE Or SWP_SHOWWINDOW
    
    Call Paint_Form
    
    fClose.Move frm.ScaleWidth - fClose.Width - 4, 6
    fMax.Move fClose.Left - fMax.Width - 1, 6
    fMin.Move fMax.Left - fMin.Width - 1, 6
    fTray.Move fMin.Left - fTray.Width - 1, 6
    With lblCaption
        .Move 0, 4, frm.ScaleWidth, 24
        .BackStyle = 0
        .Alignment = 2
        .Font = "Verdana"
        .FontSize = 12
        .FontBold = True
        .ForeColor = vbWhite
        .Caption = frm.Caption
        .Visible = True
    End With
    With img
        .Left = 5
        .Top = 5
        .Stretch = True
        .Width = 16
        .Height = 16
        .Picture = frm.Icon
        .Visible = True
    End With
    With StatusBar
        .BorderColor = RGB(0, 0, 160)
        .BorderWidth = 1
        .FillStyle = 1
        .Shape = 4
        .Move 12, frm.ScaleHeight - 34, frm.ScaleWidth - 24, 24
        .Visible = True
    End With
    With lblStatusBar
        .Move 18, frm.ScaleHeight - 30, frm.ScaleWidth - 50
        .BackStyle = 0
        .AutoSize = True
        .Visible = True
        .Font = "Tahoma"
        .FontBold = True
        .FontSize = 8
        .ForeColor = vbWhite
    End With
    Pic.Left = frm.ScaleWidth - Pic.ScaleWidth
    Pic.Top = frm.ScaleHeight - Pic.ScaleHeight
End Sub

Private Sub frm_Resize()
    frm.AutoRedraw = True
    frm.ScaleMode = vbPixels
    frm.Cls
    'Delete Caption Bar And Border Of The Form
    Dim lStyle As Long
    lStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    lStyle = lStyle And Not (WS_BORDER Or WS_CAPTION Or WS_DLGFRAME Or WS_THICKFRAME)
    SetWindowLong frm.hwnd, GWL_STYLE, lStyle
    'Shape Up The Form
    hr = CreateRoundRectRgn(0, 0, frm.ScaleWidth, frm.ScaleHeight, 10, 10)
    SetWindowRgn frm.hwnd, hr, True
    SetWindowPos frm.hwnd, HWND_TOP, frm.ScaleLeft, frm.ScaleTop, frm.ScaleWidth, frm.ScaleHeight, SWP_NOMOVE Or SWP_SHOWWINDOW
    
    Call Paint_Form
    
    fClose.Move frm.ScaleWidth - fClose.Width - 4, 6
    fMax.Move fClose.Left - fMax.Width - 1, 6
    fMin.Move fMax.Left - fMin.Width - 1, 6
    fTray.Move fMin.Left - fTray.Width - 1, 6
    'Positioning Label,Image,Shape
    With lblCaption
        .Move 0, 4, frm.ScaleWidth, 24
        .BackStyle = 0
        .Alignment = 2
        .Font = "Verdana"
        .FontSize = 12
        .FontBold = True
        .ForeColor = vbWhite
        .Caption = frm.Caption
        .Visible = True
    End With
    With img
        .Left = 5
        .Top = 5
        .Stretch = True
        .Width = 16
        .Height = 16
        .Picture = frm.Icon
        .Visible = True
    End With
    With StatusBar
        .BorderColor = RGB(0, 0, 160)
        .BorderWidth = 1
        .FillStyle = 1
        .Shape = 4
        .Move 12, frm.ScaleHeight - 34, frm.ScaleWidth - 24, 24
        .Visible = True
    End With
    With lblStatusBar
        .Move 18, frm.ScaleHeight - 30, frm.ScaleWidth - 50
        .BackStyle = 0
        .AutoSize = True
        .Visible = True
        .Font = "Tahoma"
        .FontBold = True
        .FontSize = 8
        .ForeColor = vbWhite
    End With
    Pic.Left = frm.ScaleWidth - Pic.ScaleWidth
    Pic.Top = frm.ScaleHeight - Pic.ScaleHeight
    fClose.ToolTipText = "Close"
    If frm.WindowState = vbNormal Then
        fMax.ToolTipText = "Maximize"
        Pic.Visible = True
    Else
        fMax.ToolTipText = "Restore"
        Pic.Visible = False
    End If
    fMin.ToolTipText = "Minimize"
    fTray.ToolTipText = "Send To Tray"
End Sub

Private Sub frm_Unload(Cancel As Integer)
    DeleteObject hr
End Sub

Private Sub lblCaption_DblClick()
    If frm.WindowState = vbMaximized Then
        frm.WindowState = vbNormal
    Else
        frm.WindowState = vbMaximized
    End If
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim P As POINTAPI
    Dim r As RECT
    Dim scW, scH As Integer
    frmSysmenu.ScaleMode = vbPixels
    scW = Screen.Width / 15 - frmSysmenu.ScaleWidth
    Dim LX, LY As Long
    If Button = vbRightButton Then
        'Sys Menu
        GetCursorPos P
        If P.X < scW Then
            MoveWindow frmSysmenu.hwnd, P.X, P.Y, frmSysmenu.ScaleWidth, frmSysmenu.ScaleHeight, 1
            SetWindowPos frmSysmenu.hwnd, HWND_TOP, P.X, P.Y, frmSysmenu.Width, frmSysmenu.Height, SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        Else
            MoveWindow frmSysmenu.hwnd, scW, P.Y, frmSysmenu.ScaleWidth, frmSysmenu.ScaleHeight, 1
            SetWindowPos frmSysmenu.hwnd, HWND_TOP, scW, P.Y, frmSysmenu.Width, frmSysmenu.Height, SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        End If
    Else
        'Move
        If Button = vbLeftButton Then
            ReleaseCapture
            SendMessage frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
        End If
    End If
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Resize Form
    On Error GoTo picError
    Dim Result As Long
    Dim Pos As POINTAPI
    Dim c As Long
        rs = True
        Do
        Result = GetCursorPos(Pos)
        aX% = Pos.X
        aY% = Pos.Y
        DoEvents
        Result = GetCursorPos(Pos)
        frm.Width = frm.Width + (Pos.X - aX%) * 20
        frm.Height = frm.Height + (Pos.Y - aY%) * 20
        Pic.Left = frm.ScaleWidth - Pic.ScaleWidth
        Pic.Top = frm.ScaleHeight - Pic.ScaleHeight
        Loop Until rs = False
        Exit Sub
picError:
        rs = False
        Exit Sub
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    rs = False
End Sub
Private Sub Paint_Form()
'Paint Form
    frm.Cls
    Dim frmWidth As Integer, frmHeight As Integer
    Dim X, Y
    Dim i
    Dim rndRG As Long
    Dim max As Long
    Dim l As Long
    Dim Gx, Bx As Integer
    Dim lp, fg As Long
    
    Gx = 220
        For i = 0 To 28
            frm.Line (0, i)-(frm.ScaleWidth, i), RGB(0, Gx, 0)
            Gx = Gx - 4
            If Gx < 4 Then
                Gx = 4
            End If
        Next i
    Bx = 160
        For i = 29 To 36
            frm.Line (0, i)-(frm.ScaleWidth, i), RGB(0, 0, Bx)
            Bx = Bx + 10
        Next i
        For i = 37 To frm.ScaleHeight
            frm.Line (0, i)-(frm.ScaleWidth, i), RGB(0, 0, Bx)
        Next i
   
    Bx = 200
    For Y = 0 To 5
        frm.Line (Y, 37)-(Y, frm.ScaleHeight), RGB(0, 0, Bx)
        Bx = Bx + 10
    Next Y
    Bx = 150
    For Y = frm.ScaleHeight To frm.ScaleHeight - 10 Step -1
        frm.Line (0, Y)-(frm.ScaleWidth, Y), RGB(0, 0, Bx)
        Bx = Bx + 10
    Next Y
    Bx = 220
    For Y = frm.ScaleWidth - 7 To frm.ScaleWidth
        frm.Line (Y, 36)-(Y, frm.ScaleHeight), RGB(0, 0, Bx)
        Bx = Bx - 10
        If Bx < 10 Then
            Bx = 10
        End If
    Next Y
End Sub
