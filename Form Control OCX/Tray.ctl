VERSION 5.00
Begin VB.UserControl Tray 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   MouseIcon       =   "Tray.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   435
   ScaleWidth      =   480
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   540
      Top             =   495
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   60
      MouseIcon       =   "Tray.ctx":0152
      MousePointer    =   99  'Custom
      Picture         =   "Tray.ctx":02A4
      Top             =   75
      Width           =   300
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   330
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
Private t_Icon As NOTIFYICONDATA
Private w_State As Long
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click

Private Function MouseOver() As Boolean
    Dim p As POINTAPI
    Dim d As Long
    On Error Resume Next
    d = GetCursorPos(p)
    If WindowFromPoint(p.x, p.y) = UserControl.hwnd Then
         MouseOver = True
    End If
End Function
Private Sub frm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        Dim Result As Long
        Dim msg As Long
            If frm.ScaleMode = vbPixels Then
                msg = x
            Else
                msg = x / Screen.TwipsPerPixelX
            End If
    Select Case msg
        Case WM_LBUTTONUP
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon
        Case WM_LBUTTONDBLCLK    '515 restore form window
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon
        Case WM_RBUTTONUP
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon
        Case WM_RBUTTONDBLCLK
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon
       End Select
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub Timer1_Timer()
    Dim hRgn As Long
    If Not MouseOver Then
        Shape1.FillColor = vbWhite
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.Height = 255
    UserControl.Width = 270
    UserControl.ScaleMode = vbPixels
    UserControl.AutoRedraw = True
    UserControl.Cls
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Image1.Move (UserControl.ScaleWidth - Image1.Width) / 2, (UserControl.ScaleHeight - Image1.Height) / 2
    Image1.ToolTipText = "Send To Tray"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Shape1.FillColor = &HE0E0E0
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer1.Enabled = True
    If x < 0 Or x > UserControl.ScaleWidth Or y < 0 Or y > UserControl.ScaleHeight Then
        Shape1.FillColor = &H80C0FF
    Else
        Shape1.FillColor = vbWhite
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        With t_Icon
        .cbSize = Len(t_Icon)
        .hwnd = frm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm.Icon
        .szTip = "Developed By Debasis Ghosh" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "Developed By Debasis Ghosh" & vbCrLf & "Click Here" & Chr(0)
        .szInfoTitle = "" & frm.Caption & ""
        .dwInfoFlags = NIIF_INFO
        .uTimeout = 3000
   End With
        Shell_NotifyIcon NIM_ADD, t_Icon
        w_State = frm.WindowState
        frm.Hide
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        Set frm = Parent
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 255
    UserControl.Width = 270
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

