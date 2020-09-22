VERSION 5.00
Begin VB.Form frmSysmenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Sysmenu"
   ClientHeight    =   1725
   ClientLeft      =   3840
   ClientTop       =   2385
   ClientWidth     =   2190
   LinkTopic       =   "Form2"
   ScaleHeight     =   1725
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2250
      Top             =   1845
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   690
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      Height          =   330
      Left            =   165
      Top             =   2145
      Width           =   315
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   30
      TabIndex        =   3
      Top             =   1380
      Width           =   2130
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Maximize"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   30
      TabIndex        =   2
      Top             =   870
      Width           =   2130
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   2130
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2130
   End
End
Attribute VB_Name = "frmSysmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function IsMouseOver() As Boolean
    Dim P As POINTAPI
    Dim db As Long
    On Error Resume Next
    db = GetCursorPos(P)
    If WindowFromPoint(P.X, P.Y) = frmSysmenu.hwnd Then
         IsMouseOver = True
    End If
End Function
Private Sub Form_Load()
    Dim i, t As Integer
    frmSysmenu.AutoRedraw = True
    frmSysmenu.ScaleMode = vbPixels
    frmSysmenu.Cls
    Shape1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    t = 2
    For i = 0 To 3
        With lblMenu.Item(i)
            .Left = 2
            .Top = t
            .Width = 142
            .Height = 22
            .ForeColor = vbBlack
            .Visible = True
        End With
        t = t + lblMenu.Item(i).Height + 7
    Next i
    Me.Line (2, lblMenu.Item(3).Top - 4)-(lblMenu.Item(3).Width, lblMenu.Item(3).Top - 4), vbBlack
    If Screen.ActiveForm.WindowState = vbNormal Then
        lblMenu.Item(0).Enabled = False
        lblMenu.Item(2).Enabled = True
    ElseIf Screen.ActiveForm.WindowState = vbMaximized Then
        lblMenu.Item(0).Enabled = True
        lblMenu.Item(2).Enabled = False
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    Dim i As Integer
    Shape2.Visible = False
    For i = 0 To 3
        lblMenu.Item(i).FontBold = False
    Next i
End Sub
Private Sub lblMenu_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
            Screen.ActiveForm.WindowState = vbNormal
        Case 1
            Unload Me
            ShowWindow Screen.ActiveForm.hwnd, SW_MINIMIZE
        Case 2
            Unload Me
            ShowWindow Screen.ActiveForm.hwnd, SW_MAXIMIZE
        Case 3
            Unload Me
            Unload Screen.ActiveForm
    End Select
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblMenu.Item(Index).Enabled = True Then
    Shape2.Visible = True
    lblMenu.Item(Index).ZOrder 0
    lblMenu.Item(Index).FontBold = True
    With lblMenu.Item(Index)
        Shape2.Move .Left, .Top, .Width, .Height
    End With
End If
End Sub
Private Sub Timer1_Timer()
    If Not IsMouseOver Then
        Timer1.Enabled = False
        Unload Me
    End If
End Sub

