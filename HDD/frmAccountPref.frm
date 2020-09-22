VERSION 5.00
Begin VB.Form frmAccountPref 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Account Editor"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
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
   Icon            =   "frmAccountPref.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblQuestionName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Q / A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1605
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblMainMenuSupport 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   3795
      Width           =   4755
   End
   Begin VB.Shape shapeMainMenu 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   240
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label lblInfoName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   2265
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Å"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   60
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   240
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblPasswordName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2115
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "—"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   60
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Account Editor"
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   15
      Width           =   3180
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   4320
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5160
   End
End
Attribute VB_Name = "frmAccountPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strUsername

Private Sub Form_Load()
'Take Username from the frmMain
strUsername = frmMain.lblUsername.Caption

'Change the color and Design of the controls.
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeMainMenu.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
lblPassword.BackColor = RGB(145, 155, 100)
lblInfo.BackColor = RGB(145, 155, 100)
lblQuestion.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblInfo_Click()
frmChangeInfo.Show
Unload Me
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblInfo.ForeColor = RGB(145, 155, 100)
lblInfo.BackColor = vbBlack
End Sub

Private Sub lblInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblInfo.BackColor = RGB(145, 155, 100)
lblInfo.ForeColor = vbBlack
End Sub

Private Sub lblMainMenuSupport_Click()
frmMain.Show
Unload Me
End Sub

Private Sub lblMainMenuSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMainMenu.ForeColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = vbBlack
End Sub

Private Sub lblMainMenuSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeMainMenu.BackColor = RGB(145, 155, 100)
lblMainMenu.ForeColor = vbBlack
End Sub

Private Sub lblPassword_Click()
frmChangePassword.Show
Unload Me
End Sub

Private Sub lblPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPassword.ForeColor = RGB(145, 155, 100)
lblPassword.BackColor = vbBlack
End Sub

Private Sub lblPassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPassword.ForeColor = vbBlack
lblPassword.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblQuestion_Click()
frmChangeQA.Show
Unload Me
End Sub

Private Sub lblQuestion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblQuestion.ForeColor = RGB(145, 155, 100)
lblQuestion.BackColor = vbBlack
End Sub

Private Sub lblQuestion_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblQuestion.BackColor = RGB(145, 155, 100)
lblQuestion.ForeColor = vbBlack
End Sub
