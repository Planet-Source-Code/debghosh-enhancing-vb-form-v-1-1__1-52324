VERSION 5.00
Begin VB.UserControl DBtn 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   990
   ScaleHeight     =   375
   ScaleWidth      =   990
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1020
      Top             =   375
   End
   Begin VB.Image img 
      Height          =   255
      Left            =   30
      Top             =   105
      Width           =   195
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debasis"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   345
      TabIndex        =   0
      Top             =   135
      Width           =   675
   End
End
Attribute VB_Name = "DBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim m_Fc As OLE_COLOR
Dim m_d_Fc As OLE_COLOR
'Declare the API-Functions
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Event Declarations:
Event Click()
Private Sub PicLabelPos()
If UserControl.Enabled = True Then
    img.Visible = True
    lblCaption.Enabled = True
    lblCaption.Top = (UserControl.ScaleHeight - lblCaption.Height) / 2
    If img.Picture <> 0 Then
        img.Top = (UserControl.ScaleHeight - img.Height) / 2
        If lblCaption.Caption <> "" Then
            lblCaption.Enabled = True
            img.Left = (UserControl.ScaleWidth - img.Width - lblCaption.Width - 5) / 2
        Else
            img.Left = (UserControl.ScaleWidth - img.Width) / 2
        End If
        lblCaption.Left = img.Left + img.Width + 5
    Else
        lblCaption.Left = (UserControl.ScaleWidth - lblCaption.Width) / 2
    End If
Else
    With img
        .Visible = False
    End With
    lblCaption.Enabled = False
    lblCaption.Top = (UserControl.ScaleHeight - lblCaption.Height) / 2
    lblCaption.Left = (UserControl.ScaleWidth - lblCaption.Width) / 2
End If
End Sub
 Private Sub UcMouseMove()
    UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), &HFFFFFF
    UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), &HFFFFFF
    UserControl.Line (UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, 0), &H808080
    UserControl.Line (UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)-(0, UserControl.ScaleHeight - 1), &H808080
 End Sub
 Private Sub UcMouseDown()
    UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), &H808080
    UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), &H808080
    UserControl.Line (UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, 0), &HFFFFFF
    UserControl.Line (UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)-(0, UserControl.ScaleHeight - 1), &HFFFFFF
 End Sub
Private Function IsMouseOver() As Boolean
    Dim P As POINTAPI
    Dim db As Long
    On Error Resume Next
    db = GetCursorPos(P)
    If WindowFromPoint(P.X, P.Y) = UserControl.hWnd Then
         IsMouseOver = True
    End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    Call PicLabelPos
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call PicLabelPos
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    lblCaption.FontUnderline = False
    Call PicLabelPos
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    m_d_Fc = lblCaption.ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get OnMouseMoveForeColor() As OLE_COLOR
    OnMouseMoveForeColor = m_Fc
End Property
Public Property Let OnMouseMoveForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_Fc = New_ForeColor
    PropertyChanged "OnMouseMoveForeColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    If Not IsMouseOver Then
        UserControl.Cls
        Call PicLabelPos
        lblCaption.ForeColor = m_d_Fc
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If UserControl.Enabled = True Then
        If (KeyAscii = 13 Or KeyAscii = 27) Then
            RaiseEvent Click
            Exit Sub
        End If
    End If
End Sub

Private Sub UserControl_EnterFocus()
    Call UserControl_GotFocus
End Sub

Private Sub UserControl_ExitFocus()
    Call UserControl_LostFocus
End Sub

Private Sub UserControl_GotFocus()
    lblCaption.FontUnderline = True
End Sub

Private Sub UserControl_Initialize()
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    Call PicLabelPos
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            RaiseEvent Click
        Case vbKeySpace
            Call UcMouseDown
        Case 39, 40
            SendKeys "{Tab}"
        Case 37, 38
            SendKeys "+{Tab}"
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeySpace) Then
        UserControl.Cls
        Call PicLabelPos
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_LostFocus()
    lblCaption.FontUnderline = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call UcMouseDown
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X > 0 Or X < UserControl.ScaleWidth Or Y > 0 Or Y < UserControl.ScaleHeight Then
        Call UcMouseMove
        lblCaption.ForeColor = m_Fc
    Else
        UserControl.Cls
        Call PicLabelPos
        lblCaption.ForeColor = m_d_Fc
        Timer1.Enabled = False
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        RaiseEvent Click
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=img,img,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = img.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set img.Picture = New_Picture
    Call PicLabelPos
    PropertyChanged "Picture"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Debasis")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_d_Fc = lblCaption.ForeColor
    m_Fc = PropBag.ReadProperty("OnMouseMoveForeColor", m_d_Fc)
End Sub

Private Sub UserControl_Resize()
    Call PicLabelPos
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Debasis")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("OnMouseMoveForeColor", m_Fc, m_d_Fc)
End Sub


