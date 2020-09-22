VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Help"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1680
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":030A
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hddcontact@hirdhav.com"
      Height          =   225
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address:"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":03A2
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5300
      X2              =   5300
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Help"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2310
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5295
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblOkSupport_Click()
frmLogin.Show
Unload Me
End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = vbBlack
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeOk.BackColor = RGB(145, 155, 100)
lblOk.ForeColor = vbBlack
End Sub
