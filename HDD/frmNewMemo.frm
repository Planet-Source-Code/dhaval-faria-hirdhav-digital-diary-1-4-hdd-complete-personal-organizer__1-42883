VERSION 5.00
Begin VB.Form frmNewMemo 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - New Memo"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
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
   Icon            =   "frmNewMemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   3255
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblSaveSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3420
      TabIndex        =   6
      Top             =   4810
      Width           =   585
   End
   Begin VB.Label lblSave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Top             =   4810
      Width           =   420
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4920
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3000
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Shape shapeSave 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   600
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1020
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4935
      X2              =   4920
      Y1              =   240
      Y2              =   5400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   405
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   5400
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  New Memo"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2880
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   4935
   End
End
Attribute VB_Name = "frmNewMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeCancel.BackColor = RGB(145, 155, 100)
shapeSave.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
txtTitle.BackColor = RGB(145, 155, 100)
txtDescription.BackColor = RGB(145, 155, 100)
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

Private Sub lblSaveSupport_Click()
If txtTitle.Text = "" Then
    HDDMsgBox "Please enter Title of New Memo."
    Exit Sub
ElseIf txtTitle.Text = " " Then
    HDDMsgBox "Please enter Title of New Memo."
    Exit Sub
End If
Dim strUsername As String
Dim db As Database
Dim ReS As Recordset
strUsername = frmMain.lblUsername.Caption
Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Memo.dat")
Set ReS = db.OpenRecordset("Memo")
ReS.AddNew
ReS("Title") = txtTitle.Text
ReS("Details") = txtDescription.Text
ReS.Update
ReS.Close
db.Close
Set ReS = Nothing
Set db = Nothing
HDDMsgBox "New Memo inserted successfully."
frmMemo.Show
Unload Me
End Sub

Private Sub lblSaveSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSave.ForeColor = RGB(145, 155, 100)
shapeSave.BackColor = vbBlack
End Sub

Private Sub lblSaveSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeSave.BackColor = RGB(145, 155, 100)
lblSave.ForeColor = vbBlack
End Sub
