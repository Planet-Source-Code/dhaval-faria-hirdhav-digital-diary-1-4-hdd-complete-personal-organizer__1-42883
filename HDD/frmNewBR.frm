VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewBR 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - New Birthday Reminder"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
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
   Icon            =   "frmNewBR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAP3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "AM"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtAP1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "AM"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin MSMask.MaskEdBox txtAlarmTime 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFrom 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4920
      X2              =   4920
      Y1              =   240
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   4920
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduler Type:"
      Height          =   225
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday Reminder"
      Height          =   225
      Left            =   1800
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   225
      Left            =   1155
      TabIndex        =   14
      Top             =   840
      Width           =   435
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14/Mar/1985"
      Height          =   225
      Left            =   1800
      TabIndex        =   13
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   1275
      TabIndex        =   12
      Top             =   4080
      Width           =   240
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3240
      TabIndex        =   11
      Top             =   4080
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Time:"
      Height          =   225
      Left            =   480
      TabIndex        =   10
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time From:"
      Height          =   225
      Left            =   645
      TabIndex        =   9
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   225
      Left            =   555
      TabIndex        =   8
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Dairy  -  New Birthday Reminder"
      Height          =   225
      Left            =   100
      TabIndex        =   0
      Top             =   10
      Width           =   3930
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   4560
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4920
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   600
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2760
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
txtFrom.BackColor = RGB(145, 155, 100)
txtAlarmTime.BackColor = RGB(145, 155, 100)
txtAP1.BackColor = RGB(145, 155, 100)
txtAP3.BackColor = RGB(145, 155, 100)
txtDesc.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption
lblDate.Caption = SchDate & "/" & SchMonth & "/" & SchYear
End Sub

Private Sub lblCancelSupport_Click()
frmCalender.Show
Unload Me
frmSchType.Show vbModal
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

Private Sub lblOkSupport_Click()

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(SchMonth)

ReS.AddNew
ReS("SchType") = lblType.Caption
ReS("Date") = lblDate.Caption
ReS("TF") = txtFrom.Text
ReS("AT") = txtAlarmTime.Text
ReS("AP1") = txtAP1.Text
ReS("AP3") = txtAP3.Text
ReS("Description") = txtDesc.Text
ReS.Update

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

HDDMsgBox "Birthday Reminder inserted successfully."

Unload frmSchType
frmCalender.Show
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
