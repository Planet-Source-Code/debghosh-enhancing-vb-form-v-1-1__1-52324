VERSION 5.00
Begin VB.UserControl Max 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   MouseIcon       =   "Max.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   420
   ScaleWidth      =   495
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   555
      Top             =   480
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   225
      Shape           =   4  'Rounded Rectangle
      Top             =   225
      Width           =   180
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   135
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   150
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   390
   End
End
Attribute VB_Name = "Max"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Dim frmState As Boolean
Private Function MouseOver() As Boolean
    Dim p As POINTAPI
    Dim d As Long
    On Error Resume Next
    d = GetCursorPos(p)
    If WindowFromPoint(p.X, p.Y) = UserControl.hwnd Then
         MouseOver = True
    End If
End Function

Private Sub frm_Load()
    Call UcState
End Sub

Private Sub frm_Resize()
    Call UcState
End Sub

Private Sub Timer1_Timer()
    If Not MouseOver Then
    If UserControl.Enabled = True Then
        If frmState = True Then
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbWhite
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbGreen
            End With
            With Shape3
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbGreen
            End With
        Else
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbWhite
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbGreen
            End With
        End If
        Else
            With Shape1
            .ZOrder 1
            .FillStyle = 0
            .FillColor = &HE0E0E0
        End With
        Shape3.Visible = False
        With Shape2
            .ZOrder 0
            .FillStyle = 0
            .FillColor = &HC0C0C0
        End With
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
    End If
        Timer1.Enabled = False
    End If
End Sub
Private Sub UcState()
    If frm.WindowState = vbMaximized Then
        frmState = True
        Shape2.Visible = True
        Shape3.Visible = True
        Shape2.Move 2, 2, 10, 9
        Shape3.Move 6, 6, Shape2.Width, Shape2.Height
    ElseIf frm.WindowState = vbNormal Then
        frmState = False
        Shape3.Visible = False
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    UserControl.Height = 255
    UserControl.Width = 270
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Call UcPaint
End Sub
Private Sub UcPaint()
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    UserControl.Cls
    UserControl.Height = 255
    UserControl.Width = 270
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If UserControl.Enabled = True Then
        Shape3.Visible = False
        With Shape1
            .ZOrder 1
            .FillStyle = 0
            .FillColor = vbWhite
        End With
        With Shape2
            .ZOrder 0
            .FillStyle = 0
            .FillColor = vbGreen
        End With
        With Shape3
            .ZOrder 0
            .FillStyle = 0
            .FillColor = vbGreen
        End With
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
    Else
        If UserControl.Enabled = False Then
        With Shape1
            .ZOrder 1
            .FillStyle = 0
            .FillColor = &HE0E0E0
        End With
        Shape3.Visible = False
        With Shape2
            .ZOrder 0
            .FillStyle = 0
            .FillColor = &HC0C0C0
        End With
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If frmState = True Then
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = &HE0E0E0
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = &H808080
            End With
            With Shape3
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = &H808080
            End With
        Else
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = &HE0E0E0
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = &H808080
            End With
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X > 0 Or X < UserControl.ScaleWidth Or Y > 0 Or Y < UserControl.ScaleHeight Then
        If frmState = True Then
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbGreen
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbWhite
            End With
            With Shape3
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbWhite
            End With
        Else
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbGreen
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbWhite
            End With
        End If
    Else
        Timer1.Enabled = False
        Call UcPaint
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If frmState = True Then
            frmState = False
            Shape2.Visible = True
            Shape3.Visible = True
            Shape2.Move 2, 2, 10, 9
            Shape3.Move 6, 6, Shape2.Width, Shape2.Height
            ShowWindow frm.hwnd, SW_RESTORE Or SW_SHOWNORMAL
        Else
            frmState = True
            Shape3.Visible = False
            Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
            ShowWindow frm.hwnd, SW_SHOWMAXIMIZED
        End If
        
        RaiseEvent Click
    End If
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        Set frm = Parent
        Call UcState
    End If
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 255
    UserControl.Width = 270
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call UcPaint
    PropertyChanged "Enabled"
End Property

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

