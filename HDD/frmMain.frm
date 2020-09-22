VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3375
      Begin VB.Label lblAccountName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Account Editor"
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
         Left            =   2040
         TabIndex        =   23
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblAccount 
         AutoSize        =   -1  'True
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   60
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   2040
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblCalenderName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scheduler"
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
         Left            =   1920
         TabIndex        =   19
         Top             =   1200
         Width           =   1410
      End
      Begin VB.Label lblCalender 
         AutoSize        =   -1  'True
         Caption         =   "¶"
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
         Left            =   2040
         TabIndex        =   18
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblContacts 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "»"
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
         Left            =   285
         TabIndex        =   7
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lblContactsName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contacts"
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
         Left            =   0
         TabIndex        =   6
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label lblMemo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ù"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label lblMemoName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Memo"
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
         Left            =   -120
         TabIndex        =   4
         Top             =   2760
         Width           =   2010
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblAUName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Update"
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
         Left            =   1680
         TabIndex        =   16
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label lblAU 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "˝"
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
         Left            =   1965
         TabIndex        =   15
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lblAbtUsName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About Us"
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
         Left            =   -120
         TabIndex        =   14
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblAbtUs 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "i"
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
         Left            =   285
         TabIndex        =   13
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Label lblLogOutSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1320
      TabIndex        =   21
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblLogOut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log Out"
      Height          =   225
      Left            =   1800
      TabIndex        =   20
      Top             =   4440
      Width           =   660
   End
   Begin VB.Shape shapeLogOut 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1320
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   3840
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblDownSupport 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   3640
      TabIndex        =   10
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label lblDown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3640
      TabIndex        =   11
      Top             =   2730
      Width           =   375
   End
   Begin VB.Label lblUpSupport 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   3640
      TabIndex        =   9
      Top             =   1080
      Width           =   350
   End
   Begin VB.Label lblUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3640
      TabIndex        =   8
      Top             =   1280
      Width           =   375
   End
   Begin VB.Shape shapeUp 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   3640
      Top             =   1080
      Width           =   350
   End
   Begin VB.Line Line4 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   3840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4080
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name Will Be Here"
      Height          =   225
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   930
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  v 1.3"
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   15
      Width           =   2325
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4080
   End
   Begin VB.Shape shapeDown 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   3640
      Top             =   2520
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'Take a username from frmLogin and then unload frmLogin
lblUsername.Caption = frmLogin.txtUsername.Text
Unload frmLogin

'Change Color and set Design property to controls
lblAccount.BackColor = RGB(145, 155, 100)
lblContacts.BackColor = RGB(145, 155, 100)
Frame2.BackColor = RGB(145, 155, 100)
shapeUp.BackColor = RGB(145, 155, 100)
Me.BackColor = RGB(145, 155, 100)
shapeDown.BackColor = RGB(145, 155, 100)
shapeLogOut.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
Frame1.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
lblAU.BackColor = RGB(145, 155, 100)
lblMemo.BackColor = RGB(145, 155, 100)
lblAbtUs.BackColor = RGB(145, 155, 100)
lblCalender.BackColor = RGB(145, 155, 100)

'Disable Up Arrow so no one can go up
lblUp.Enabled = False
lblUpSupport.Enabled = False
End Sub

Private Sub lblAbtUs_Click()
frmAbout.Show
Me.Hide
End Sub

Private Sub lblAbtUs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbtUs.ForeColor = RGB(145, 155, 100)
lblAbtUs.BackColor = vbBlack
End Sub

Private Sub lblAbtUs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbtUs.ForeColor = vbBlack
lblAbtUs.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblAccount_Click()
frmAccountPref.Show
Me.Hide
End Sub

Private Sub lblAccount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAccount.ForeColor = RGB(145, 155, 100)
lblAccount.BackColor = vbBlack
End Sub

Private Sub lblAccount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAccount.BackColor = RGB(145, 155, 100)
lblAccount.ForeColor = vbBlack
End Sub

Private Sub lblAU_Click()
Shell App.Path + "\HDAU.exe", vbNormalFocus
End
End Sub

Private Sub lblAU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAU.BackColor = vbBlack
lblAU.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblAU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAU.ForeColor = vbBlack
lblAU.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCalender_Click()
frmCalender.Show
Me.Hide
End Sub

Private Sub lblCalender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCalender.ForeColor = RGB(145, 155, 100)
lblCalender.BackColor = vbBlack
End Sub

Private Sub lblCalender_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCalender.BackColor = RGB(145, 155, 100)
lblCalender.ForeColor = vbBlack
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblContacts_Click()
frmContacts.Show
Me.Hide
End Sub

Private Sub lblContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblContacts.BackColor = vbBlack
lblContacts.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblContacts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblContacts.BackColor = RGB(145, 155, 100)
lblContacts.ForeColor = vbBlack
End Sub

Private Sub lblDownSupport_Click()
If Frame1.Visible = True Then
    Frame2.Visible = True
    Frame1.Visible = False
    lblUp.Enabled = True
    lblUpSupport.Enabled = True
    lblDown.Enabled = False
    lblDownSupport.Enabled = False
End If
End Sub

Private Sub lblDownSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDown.ForeColor = RGB(145, 155, 100)
shapeDown.BackColor = vbBlack
End Sub

Private Sub lblDownSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeDown.BackColor = RGB(145, 155, 100)
lblDown.ForeColor = vbBlack
End Sub

Private Sub lblLogOutSupport_Click()
HDDYesNoBox "Are you sure do you want to Log Out?"
    If Yes Then
        ChDir "c:\"
        End
    Else
        Exit Sub
    End If
End Sub

Private Sub lblLogOutSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogOut.ForeColor = RGB(145, 155, 100)
shapeLogOut.BackColor = vbBlack
End Sub

Private Sub lblLogOutSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeLogOut.BackColor = RGB(145, 155, 100)
lblLogOut.ForeColor = vbBlack
End Sub

Private Sub lblMemo_Click()
frmMemo.Show
Me.Hide
End Sub

Private Sub lblMemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMemo.ForeColor = RGB(145, 155, 100)
lblMemo.BackColor = vbBlack
End Sub

Private Sub lblMemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMemo.BackColor = RGB(145, 155, 100)
lblMemo.ForeColor = vbBlack
End Sub

Private Sub lblUpSupport_Click()
If Frame2.Visible = True Then
    Frame1.Visible = True
    Frame2.Visible = False
    lblDown.Enabled = True
    lblDownSupport.Enabled = True
    lblUp.Enabled = False
    lblUpSupport.Enabled = False
End If
End Sub

Private Sub lblUpSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUp.ForeColor = RGB(145, 155, 100)
shapeUp.BackColor = vbBlack
End Sub

Private Sub lblUpSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeUp.BackColor = RGB(145, 155, 100)
lblUp.ForeColor = vbBlack
End Sub
