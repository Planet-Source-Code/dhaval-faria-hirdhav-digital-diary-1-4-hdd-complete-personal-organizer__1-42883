VERSION 5.00
Begin VB.Form frmMemoDetails 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - memo Details"
   ClientHeight    =   5415
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
   Icon            =   "frmMemoDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2100
      TabIndex        =   5
      Top             =   4920
      Width           =   240
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4800
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1440
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4815
      X2              =   4815
      Y1              =   240
      Y2              =   5400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1020
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Memo Details"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3090
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   5400
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   15
      Width           =   4815
   End
End
Attribute VB_Name = "frmMemoDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
txtDescription.BackColor = RGB(145, 155, 100)
txtTitle.BackColor = RGB(145, 155, 100)
txtTitle.Text = frmMemo.lstMemo.Text
shapeOk.BackColor = RGB(145, 155, 100)

Dim strUsername As String
strUsername = frmMain.lblUsername.Caption

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Memo.dat")
Set ReS = db.OpenRecordset("Memo")

ReS.Move (frmMemo.lstMemo.ListIndex)
txtDescription.Text = ReS("Details")
Exit Sub
db.Close
ReS.Close
Set db = Nothing
Set ReS = Nothing

End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblOkSupport_Click()
frmMemo.Show
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
