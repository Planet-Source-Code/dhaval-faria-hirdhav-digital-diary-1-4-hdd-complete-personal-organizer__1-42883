VERSION 5.00
Begin VB.Form frmReminders 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Reminders"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
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
   Icon            =   "frmReminders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMainMenuSupport 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   3320
      Width           =   4035
   End
   Begin VB.Shape shapeMainMenu 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   240
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   240
      Y2              =   3960
   End
   Begin VB.Label lblARName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anivarsary Reminder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   1920
   End
   Begin VB.Label lblBRName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday Reminder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label lblAR 
      AutoSize        =   -1  'True
      Caption         =   "Z"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblBR 
      AutoSize        =   -1  'True
      Caption         =   "e"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Reminders"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2865
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4560
   End
End
Attribute VB_Name = "frmReminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblAR.BackColor = RGB(145, 155, 100)
lblBR.BackColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblAR_Click()
frmAniRemind.Show
Unload Me
End Sub

Private Sub lblAR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAR.ForeColor = RGB(145, 155, 100)
lblAR.BackColor = vbBlack
End Sub

Private Sub lblAR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAR.BackColor = RGB(145, 155, 100)
lblAR.ForeColor = vbBlack
End Sub

Private Sub lblBR_Click()
frmBirthRemind.Show
Unload Me
End Sub

Private Sub lblBR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBR.ForeColor = RGB(145, 155, 100)
lblBR.BackColor = vbBlack
End Sub

Private Sub lblBR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBR.BackColor = RGB(145, 155, 100)
lblBR.ForeColor = vbBlack
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
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
