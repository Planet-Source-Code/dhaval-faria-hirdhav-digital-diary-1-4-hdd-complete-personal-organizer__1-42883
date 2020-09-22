VERSION 5.00
Begin VB.Form frmBRDetail 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Detail Birthday Reminder Details"
   ClientHeight    =   4095
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
   Icon            =   "frmBRDetail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSS 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtMM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtHH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtRDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtBDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtAP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "AM/PM"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2855
      Width           =   2415
   End
   Begin VB.TextBox txtPHNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2375
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   455
      Width           =   2415
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3450
      Width           =   1455
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD/MM/YYYY"
      Height          =   225
      Left            =   3240
      TabIndex        =   16
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD/MM/YYYY"
      Height          =   225
      Left            =   3240
      TabIndex        =   15
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   3580
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   215
      Y2              =   4055
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   1200
      TabIndex        =   13
      Top             =   2855
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      Height          =   225
      Left            =   480
      TabIndex        =   12
      Top             =   2375
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   225
      Left            =   3000
      TabIndex        =   11
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   225
      Left            =   2400
      TabIndex        =   10
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminding Time:"
      Height          =   225
      Left            =   360
      TabIndex        =   9
      Top             =   1895
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminding Date:"
      Height          =   225
      Left            =   360
      TabIndex        =   8
      Top             =   1415
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      Height          =   225
      Left            =   840
      TabIndex        =   7
      Top             =   935
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   225
      Left            =   1200
      TabIndex        =   6
      Top             =   455
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   215
      Y2              =   4055
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Detail Birthday Reminder"
      Height          =   225
      Left            =   195
      TabIndex        =   5
      Top             =   15
      Width           =   4035
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5160
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2040
      Top             =   3450
      Width           =   1455
   End
End
Attribute VB_Name = "frmBRDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strUsername

Private Sub Form_Load()
strUsername = frmMain.lblUsername.Caption
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
Me.BackColor = RGB(145, 155, 100)
txtName.BackColor = RGB(145, 155, 100)
txtPHNo.BackColor = RGB(145, 155, 100)
txtHH.BackColor = RGB(145, 155, 100)
txtMM.BackColor = RGB(145, 155, 100)
txtSS.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtAP.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtBDate.BackColor = RGB(145, 155, 100)
txtRDate.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\BR.dat")
Set ReS = db.OpenRecordset("BR")

ReS.Move (frmBirthRemind.lstInfo.ListIndex)
txtName.Text = ReS("Name")
txtPHNo.Text = ReS("PHNo")
txtEMail.Text = ReS("EMail")
txtBDate.Text = ReS("BDate")
txtRDate.Text = ReS("RDate")
txtHH.Text = ReS("TimeH")
txtMM.Text = ReS("TimeM")
txtSS.Text = ReS("TimeS")
txtAP.Text = ReS("TimeAP")

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

End Sub

Private Sub lblCancelSupport_Click()
frmBirthRemind.Show
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
