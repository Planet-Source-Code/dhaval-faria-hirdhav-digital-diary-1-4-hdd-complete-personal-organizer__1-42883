VERSION 5.00
Begin VB.Form frmAuthoHelp 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Autho Help"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
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
   Icon            =   "frmAuthoHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line5 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   120
      X2              =   4920
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblForAuthoCODESupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblForAuthoCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Authentication CODE"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Shape shapeForAuthoCODE 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1320
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "If you forgot your Authentication CODE than click on the below given Button."
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1665
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To get your Authentication CODE click on the Next Button else click on Cancel Button."
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next >"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3240
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAuthoHelp.frx":030A
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5055
      X2              =   5040
      Y1              =   240
      Y2              =   6000
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   6000
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Autho Help"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2850
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5055
   End
End
Attribute VB_Name = "frmAuthoHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeForAuthoCODE.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
frmAutho.Show
Unload Me
End Sub

Private Sub lblCancelSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = vbBlack
End Sub

Private Sub lblCancelSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel.BackColor = RGB(145, 155, 100)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblExistUserSupport_Click()
frmExistUser1.Show
Unload Me
End Sub

Private Sub lblForAuthoCODESupport_Click()
frmForAutho1.Show
Unload Me
End Sub

Private Sub lblForAuthoCODESupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblForAuthoCode.ForeColor = RGB(145, 155, 100)
shapeForAuthoCODE.BackColor = vbBlack
End Sub

Private Sub lblForAuthoCODESupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeForAuthoCODE.BackColor = RGB(145, 155, 100)
lblForAuthoCode.ForeColor = vbBlack
End Sub

Private Sub lblOkSupport_Click()
frmGetAutho1.Show
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
