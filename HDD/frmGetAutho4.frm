VERSION 5.00
Begin VB.Form frmGetAutho4 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Get Autho (Step 4)"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
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
   Icon            =   "frmGetAutho4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblFinishSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label lblFinish 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Finish"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Shape shapeFinish 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1680
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODE:"
      Height          =   225
      Left            =   720
      TabIndex        =   7
      Top             =   2640
      Width           =   525
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      Height          =   225
      Left            =   1320
      TabIndex        =   6
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   225
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label lblCODE 
      BackStyle       =   0  'Transparent
      Caption         =   "Authentication CODE."
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Authentication Information:"
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   2730
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5900
      X2              =   5900
      Y1              =   240
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetAutho4.frx":030A
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Get Autho (Step 4)"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3465
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   4800
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5895
   End
End
Attribute VB_Name = "frmGetAutho4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblName.Caption = frmGetAutho2.txtFName.Text + " " + frmGetAutho2.txtLName.Text
lblCODE.Caption = frmGetAutho3.txtAutho.Text
shapeFinish.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblFinishSupport_Click()
'Unload frmGetAutho2
'Unload frmGetAutho3
frmAutho.Show
frmAutho.txtName.Text = lblName.Caption
frmAutho.txtAutho.Text = lblCODE.Caption
Unload Me
End Sub

Private Sub lblFinishSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblFinish.ForeColor = RGB(145, 155, 100)
shapeFinish.BackColor = vbBlack
End Sub

Private Sub lblFinishSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeFinish.BackColor = RGB(145, 155, 100)
lblFinish.ForeColor = vbBlack
End Sub
