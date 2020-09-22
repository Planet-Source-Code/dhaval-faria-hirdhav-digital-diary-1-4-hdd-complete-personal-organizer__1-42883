VERSION 5.00
Begin VB.Form frmMemoEdit 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Memo Edit"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
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
   Icon            =   "frmMemoEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   590
      Width           =   4335
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1190
      Width           =   4335
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3360
      TabIndex        =   7
      Top             =   4940
      Width           =   585
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2880
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   225
      Left            =   1245
      TabIndex        =   3
      Top             =   4935
      Width           =   315
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Memo Details"
      Height          =   225
      Left            =   200
      TabIndex        =   6
      Top             =   10
      Width           =   3090
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   20
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   230
      Y2              =   5390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   350
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   950
      Width           =   1020
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4815
      X2              =   4815
      Y1              =   230
      Y2              =   5390
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   600
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4800
      Y1              =   5390
      Y2              =   5390
   End
End
Attribute VB_Name = "frmMemoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strUsername As String

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = vbBlack
txtTitle.BackColor = RGB(145, 155, 100)
txtDescription.BackColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Memo.dat")
Set ReS = db.OpenRecordset("Memo")

ReS.Move (frmMemo.lstMemo.ListIndex)
txtTitle.Text = ReS("Title")
txtDescription.Text = ReS("Details")

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

End Sub

Private Sub lblCancelSupport_Click()
frmMemo.Show
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

Private Sub lblOkSupport_Click()

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Memo.dat")
Set ReS = db.OpenRecordset("Memo")

ReS.Move (frmMemo.lstMemo.ListIndex)
ReS.Edit
ReS("Details") = txtDescription.Text
ReS.Update

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

HDDMsgBox "Memo edited successfully."

Unload Me
Unload frmMemo
frmMemo.Show

End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = vbBlack
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeOk.BackColor = RGB(145, 155, 100)
lblOk.ForeColor = vbBlack
End Sub
