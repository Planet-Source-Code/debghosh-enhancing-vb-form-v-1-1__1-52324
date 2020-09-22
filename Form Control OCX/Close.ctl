VERSION 5.00
Begin VB.UserControl Close 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   MouseIcon       =   "Close.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   450
   ScaleWidth      =   480
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   540
      Top             =   510
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   315
      X2              =   135
      Y1              =   135
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   105
      X2              =   330
      Y1              =   150
      Y2              =   315
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   330
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   360
   End
End
Attribute VB_Name = "Close"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Private Sub Timer1_Timer()
    If Not MouseOver Then
        Timer1.Enabled = False
        Call UCEnabled
    End If
End Sub
Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
    UserControl.Height = 255
    UserControl.Width = 270
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Call UCEnabled
 End Sub
Private Sub UserControl_InitProperties()
    UserControl.ScaleMode = vbPixels
    UserControl.Height = 255
    UserControl.Width = 270
    Shape1.ZOrder 1
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Shape1.FillColor = &HE0E0E0
        Line1.BorderColor = &H808080
        Line2.BorderColor = &H808080
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X < 0 Or X > UserControl.ScaleWidth Or Y < 0 Or Y > UserControl.ScaleHeight Then
        Call UCEnabled
    Else
        Call UcMouseOver
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If UserControl.Enabled = True Then
            Unload UserControl.Parent
            RaiseEvent Click
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 255
    UserControl.Width = 270
    Exit Sub
End Sub
Private Function MouseOver() As Boolean
    Dim p As POINTAPI
    Dim d As Long
    On Error Resume Next
    d = GetCursorPos(p)
    If WindowFromPoint(p.X, p.Y) = UserControl.hwnd Then
         MouseOver = True
    End If
End Function
Private Sub UcMouseOver()
    Shape1.FillColor = vbRed
    Line1.BorderColor = vbWhite
    Line2.BorderColor = vbWhite
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call UCEnabled
    PropertyChanged "Enabled"
End Property
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        Set frm = Parent
    End If
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Call UCEnabled
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
Private Sub UCEnabled()
If UserControl.Enabled = True Then
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    UserControl.Cls
    Shape1.ZOrder 1
    Shape1.FillStyle = 0
    Shape1.FillColor = vbWhite
    With Line1
        .X1 = 4
        .X2 = UserControl.ScaleWidth - 5
        .Y1 = 4
        .Y2 = UserControl.ScaleHeight - 5
        .BorderColor = vbRed
        .ZOrder 0
    End With
    With Line2
        .X1 = UserControl.ScaleWidth - 5
        .X2 = 4
        .Y1 = 4
        .Y2 = UserControl.ScaleHeight - 5
        .BorderColor = vbRed
        .ZOrder 0
    End With
Else
    Shape1.ZOrder 1
    Shape1.FillStyle = 0
    Shape1.FillColor = &HE0E0E0
    With Line1
        .X1 = 4
        .X2 = UserControl.ScaleWidth - 5
        .Y1 = 4
        .Y2 = UserControl.ScaleHeight - 5
        .BorderColor = &HC0C0C0
        .ZOrder 0
    End With
    With Line2
        .X1 = UserControl.ScaleWidth - 5
        .X2 = 4
        .Y1 = 4
        .Y2 = UserControl.ScaleHeight - 5
        .BorderColor = &HC0C0C0
        .ZOrder 0
    End With
End If
End Sub
