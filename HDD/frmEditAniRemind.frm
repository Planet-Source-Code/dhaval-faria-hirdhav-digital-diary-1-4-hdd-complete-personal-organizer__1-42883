VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEditAniRemind 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Edit Anivarsary Reminder"
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
   Icon            =   "frmEditAniRemind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   470
      Width           =   2415
   End
   Begin VB.TextBox txtPHNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   2390
      Width           =   2415
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   2870
      Width           =   2415
   End
   Begin VB.TextBox txtAP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "AM/PM"
      Top             =   1910
      Width           =   735
   End
   Begin MSMask.MaskEdBox txtRDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1430
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtBDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   950
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtSS 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   1910
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   2
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtMM 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   1910
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   2
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtHH 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1910
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   2
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   3470
      Width           =   1455
   End
   Begin VB.Label lblEditSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   3470
      Width           =   1455
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Edit Anivarsary Reminder"
      Height          =   225
      Left            =   195
      TabIndex        =   23
      Top             =   15
      Width           =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   230
      Y2              =   4070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   225
      Left            =   1200
      TabIndex        =   22
      Top             =   470
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      Height          =   225
      Left            =   840
      TabIndex        =   21
      Top             =   950
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminding Date:"
      Height          =   225
      Left            =   360
      TabIndex        =   20
      Top             =   1430
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminding Time:"
      Height          =   225
      Left            =   360
      TabIndex        =   19
      Top             =   1910
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   225
      Left            =   2400
      TabIndex        =   18
      Top             =   1910
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   225
      Left            =   3000
      TabIndex        =   17
      Top             =   1910
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      Height          =   225
      Left            =   480
      TabIndex        =   16
      Top             =   2390
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   1200
      TabIndex        =   15
      Top             =   2870
      Width           =   555
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   230
      Y2              =   4070
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   225
      Left            =   1155
      TabIndex        =   14
      Top             =   3615
      Width           =   330
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3610
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD/MM/YYYY"
      Height          =   225
      Left            =   3240
      TabIndex        =   10
      Top             =   950
      Width           =   1050
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD/MM/YYYY"
      Height          =   225
      Left            =   3240
      TabIndex        =   8
      Top             =   1430
      Width           =   1050
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5160
   End
   Begin VB.Shape shapeEdit 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   600
      Top             =   3470
      Width           =   1455
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3120
      Top             =   3470
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditAniRemind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strUsername

Private Sub Form_Load()
strUsername = frmMain.lblUsername.Caption
Me.BackColor = RGB(145, 155, 100)
txtName.BackColor = RGB(145, 155, 100)
shapeEdit.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
txtAP.Width = "375"
txtPHNo.BackColor = RGB(145, 155, 100)
txtBDate.BackColor = RGB(145, 155, 100)
txtRDate.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtHH.BackColor = RGB(145, 155, 100)
txtMM.BackColor = RGB(145, 155, 100)
txtSS.BackColor = RGB(145, 155, 100)
txtAP.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\AR.dat")
Set ReS = db.OpenRecordset("AR")

ReS.Move (frmAniRemind.lstInfo.ListIndex)
txtName.Text = ReS("Name")
txtPHNo.Text = ReS("PHNo")
txtEMail.Text = ReS("EMail")
txtBDate.Text = ReS("ADate")
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
frmAniRemind.Show
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

Private Sub lblEditSupport_Click()
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\AR.dat")
Set ReS = db.OpenRecordset("AR")

ReS.Move (frmAniRemind.lstInfo.ListIndex)
ReS.Edit
ReS("Name") = txtName.Text
ReS("EMail") = txtEMail.Text
ReS("ADate") = txtBDate.Text
ReS("RDate") = txtRDate.Text
ReS("TimeH") = txtHH.Text
ReS("TimeS") = txtSS.Text
ReS("TimeM") = txtMM.Text
ReS("TimeAP") = txtAP.Text
ReS("PHNo") = txtPHNo.Text
ReS.Update

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing
HDDMsgBox "Record Edited Successfully."
Unload frmAniRemind
frmAniRemind.Show
Unload Me
End Sub

Private Sub lblEditSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = RGB(145, 155, 100)
shapeEdit.BackColor = vbBlack
End Sub

Private Sub lblEditSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = vbBlack
shapeEdit.BackColor = RGB(145, 155, 100)
End Sub
