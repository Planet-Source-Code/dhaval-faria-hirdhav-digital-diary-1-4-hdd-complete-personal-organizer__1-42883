VERSION 5.00
Begin VB.Form frmYesNoBox 
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
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
   Icon            =   "frmYesNoBox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblNoSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblYesSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      Height          =   225
      Left            =   4725
      TabIndex        =   3
      Top             =   975
      Width           =   225
   End
   Begin VB.Label lblYes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      Height          =   225
      Left            =   1905
      TabIndex        =   2
      Top             =   975
      Width           =   315
   End
   Begin VB.Shape shapeNo 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   4080
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape shapeYes 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1320
      Top             =   840
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6500
      X2              =   6500
      Y1              =   240
      Y2              =   1680
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   1680
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   6495
   End
End
Attribute VB_Name = "frmYesNoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblQuestion.Caption = MyMessage1
shapeYes.BackColor = RGB(145, 155, 100)
shapeNo.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblNoSupport_Click()
Yes = False
Unload Me
End Sub

Private Sub lblNoSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNo.ForeColor = RGB(145, 155, 100)
shapeNo.BackColor = vbBlack
End Sub

Private Sub lblNoSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNo.ForeColor = vbBlack
shapeNo.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblYesSupport_Click()
Yes = True
Unload Me
End Sub

Private Sub lblYesSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblYes.ForeColor = RGB(145, 155, 100)
shapeYes.BackColor = vbBlack
End Sub

Private Sub lblYesSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblYes.ForeColor = vbBlack
shapeYes.BackColor = RGB(145, 155, 100)
End Sub
