VERSION 5.00
Object = "*\A..\Form Control OCX\Form Control.vbp"
Object = "*\A..\Button Control\Grad Button\GradBtn.vbp"
Object = "*\A..\Button Control\Deba Button\Deba Button.vbp"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Deba Enhanced Vb Form"
   ClientHeight    =   7605
   ClientLeft      =   1485
   ClientTop       =   825
   ClientWidth     =   9570
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   Begin GradBtn.FancyBtn FancyBtn6 
      Height          =   495
      Left            =   4935
      TabIndex        =   18
      Top             =   6705
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   873
      DUCColor        =   1
      Caption         =   "Disabled Button"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":3072
      ScaleHeight     =   33
      ScaleMode       =   0
      ScaleWidth      =   229
   End
   Begin GradBtn.FancyBtn FancyBtn7 
      Height          =   1170
      Left            =   5700
      TabIndex        =   17
      Top             =   5475
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2064
      DUCColor        =   3
      OnMouseMoveUColor=   1
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "frmMain.frx":3146
      ScaleHeight     =   78
      ScaleMode       =   0
      ScaleWidth      =   127
   End
   Begin GradBtn.FancyBtn FancyBtn5 
      Height          =   1185
      Left            =   5535
      TabIndex        =   16
      Top             =   4215
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   2090
      DUCColor        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":47A0
      ScaleHeight     =   79
      ScaleMode       =   0
      ScaleWidth      =   152
   End
   Begin GradBtn.FancyBtn FancyBtn4 
      Height          =   405
      Left            =   5025
      TabIndex        =   15
      Top             =   3750
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   714
      OnMouseMoveUColor=   3
      Caption         =   "Button With Picture"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      Picture         =   "frmMain.frx":5DFA
      ScaleHeight     =   27
      ScaleMode       =   0
      ScaleWidth      =   210
      OnMouseMoveForeColor=   16777215
   End
   Begin GradBtn.FancyBtn FancyBtn3 
      Height          =   465
      Left            =   6720
      TabIndex        =   14
      Top             =   3210
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   820
      OnMouseMoveUColor=   1
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      ScaleHeight     =   31
      ScaleMode       =   0
      ScaleWidth      =   88
      OnMouseMoveForeColor=   16777215
   End
   Begin GradBtn.FancyBtn FancyBtn2 
      Height          =   465
      Left            =   5280
      TabIndex        =   13
      Top             =   3225
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   820
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   31
      ScaleMode       =   0
      ScaleWidth      =   88
   End
   Begin GradBtn.FancyBtn FancyBtn1 
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   2790
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   661
      DUCColor        =   3
      Caption         =   "Normal Button"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   25
      ScaleMode       =   0
      ScaleWidth      =   218
   End
   Begin DebaButton.DBtn DBtn4 
      Height          =   420
      Left            =   1455
      TabIndex        =   10
      Top             =   6705
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   741
      BackColor       =   16711680
      Caption         =   "Disabled Button"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   255
   End
   Begin DebaButton.DBtn DBtn3 
      Height          =   1020
      Left            =   1860
      TabIndex        =   9
      Top             =   5550
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   1799
      BackColor       =   16711680
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Picture         =   "frmMain.frx":5ECE
      OnMouseMoveForeColor=   255
   End
   Begin DebaButton.DBtn DBtn2 
      Height          =   1155
      Left            =   1065
      TabIndex        =   8
      Top             =   4335
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   2037
      BackColor       =   16711680
      Caption         =   "Button With Picture"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Picture         =   "frmMain.frx":7528
      OnMouseMoveForeColor=   33023
   End
   Begin DebaButton.DBtn DBtn1 
      Height          =   555
      Left            =   1080
      TabIndex        =   7
      Top             =   3690
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   979
      BackColor       =   16711680
      Caption         =   "Normal Button"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   255
   End
   Begin FormControl.Pic Pic1 
      Height          =   375
      Left            =   9450
      TabIndex        =   5
      Top             =   7560
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      MousePointer    =   8
      ScaleHeight     =   25
      ScaleWidth      =   27
   End
   Begin FormControl.Tray Tray1 
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   -240
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin FormControl.Close Close1 
      Height          =   255
      Left            =   8355
      TabIndex        =   3
      Top             =   -225
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin FormControl.Min Min1 
      Height          =   255
      Left            =   7935
      TabIndex        =   2
      Top             =   -225
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin FormControl.Max Max1 
      Height          =   255
      Left            =   8160
      TabIndex        =   1
      Top             =   -225
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Goto Right Side Corner and Drag it to resize the Form. Check Out the Gradient Buttons."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   570
      TabIndex        =   20
      Top             =   1830
      Width           =   8820
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click On The Buttons"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5550
      TabIndex        =   19
      Top             =   2490
      Width           =   2235
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   4890
      Left            =   4680
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click On The Buttons"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   1515
      TabIndex        =   11
      Top             =   3345
      Width           =   2235
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   4035
      Left            =   825
      Top             =   3255
      Width           =   3720
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   345
      Left            =   2865
      TabIndex        =   6
      Top             =   -390
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   9585
      Top             =   15
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   975
      Top             =   -465
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   390
      Left            =   1755
      TabIndex        =   0
      Top             =   -360
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim d As New clsForm

Private Sub DBtn1_Click()
    MsgBox "Normal Button", vbInformation
End Sub

Private Sub DBtn2_Click()
    MsgBox "Button With Picture", vbInformation
End Sub

Private Sub Form_Load()
    Set d = New clsForm
    Set d.frm = Me
    Set d.fMax = Max1
    Set d.fMin = Min1
    Set d.fClose = Close1
    Set d.fTray = Tray1
    Set d.img = Image1
    Set d.lblCaption = Label1
    Set d.Pic = Pic1
    Set d.lblStatusBar = Label2
    Set d.StatusBar = Shape1
    Label2.Caption = Format(Now, "dddd,  dd MMM YYYY")
    DBtn1.BackColor = RGB(0, 0, 240)
    DBtn2.BackColor = RGB(0, 0, 240)
    DBtn3.BackColor = RGB(0, 0, 240)
    DBtn4.BackColor = RGB(0, 0, 240)
End Sub
