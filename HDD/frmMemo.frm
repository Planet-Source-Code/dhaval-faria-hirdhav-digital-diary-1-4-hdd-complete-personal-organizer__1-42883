VERSION 5.00
Begin VB.Form frmMemo 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Memo"
   ClientHeight    =   5295
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
   Icon            =   "frmMemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMemo 
      Appearance      =   0  'Flat
      Height          =   3180
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblDeleteSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblEditSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblDelete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      Height          =   225
      Left            =   2800
      TabIndex        =   7
      Top             =   4340
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   225
      Left            =   840
      TabIndex        =   6
      Top             =   4340
      Width           =   315
   End
   Begin VB.Shape shapeDelete 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2400
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Shape shapeEdit 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   4080
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblMainMenuSupport 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label lblMainMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      Height          =   225
      Left            =   1500
      TabIndex        =   4
      Top             =   4875
      Width           =   915
   End
   Begin VB.Shape shapeMainMenu 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   240
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label lblNewSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblNew 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Memo"
      Height          =   225
      Left            =   1485
      TabIndex        =   1
      Top             =   495
      Width           =   945
   End
   Begin VB.Shape shapeNew 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1200
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Memo"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2460
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4080
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strUsername As String

Private Sub Form_Load()
shapeEdit.BackColor = RGB(145, 155, 100)
shapeDelete.BackColor = RGB(145, 155, 100)
Me.BackColor = RGB(145, 155, 100)
lstMemo.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeNew.BackColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = RGB(145, 155, 100)

'Store the username
strUsername = frmMain.lblUsername.Caption

'Open database
Dim db As Database
Dim ReS As Recordset

On Error GoTo HanErr

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Memo.dat")
Set ReS = db.OpenRecordset("Memo")

Do
    lstMemo.AddItem ReS("Title")
    ReS.MoveNext
Loop
ReS.Close
db.Close
Set db = Nothing
Set ReS = Nothing
HanErr:
    If Err.Number = 3021 Then
        Exit Sub
    End If
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblDeleteSupport_Click()
If lstMemo.ListIndex = "-1" Then
    HDDMsgBox "Please select the memo from the list."
Else
    HDDYesNoBox "Do you want to delete it?"
    If Yes Then
        Dim db As Database
        Dim ReS As Recordset
        
        Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Memo.dat")
        Set ReS = db.OpenRecordset("Memo")
        
        ReS.Move (lstMemo.ListIndex)
        ReS.Delete
        ReS.Close
        db.Close
        Set ReS = Nothing
        Set db = Nothing
        HDDMsgBox "Successfully Deleted."
        Unload Me
        frmMemo.Show
    End If
End If
End Sub

Private Sub lblDeleteSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDelete.ForeColor = RGB(145, 155, 100)
shapeDelete.BackColor = vbBlack
End Sub

Private Sub lblDeleteSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeDelete.BackColor = RGB(145, 155, 100)
lblDelete.ForeColor = vbBlack
End Sub

Private Sub lblEditSupport_Click()
If lstMemo.ListIndex = "-1" Then
    HDDMsgBox "Please select the memo from the list."
Else
    frmMemoEdit.Show
    Me.Hide
End If
End Sub

Private Sub lblEditSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = RGB(145, 155, 100)
shapeEdit.BackColor = vbBlack
End Sub

Private Sub lblEditSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeEdit.BackColor = RGB(145, 155, 100)
lblEdit.ForeColor = vbBlack
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

Private Sub lblNewSupport_Click()
frmNewMemo.Show
Unload Me
End Sub

Private Sub lblNewSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNew.ForeColor = RGB(145, 155, 100)
shapeNew.BackColor = vbBlack
End Sub

Private Sub lblNewSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeNew.BackColor = RGB(145, 155, 100)
lblNew.ForeColor = vbBlack
End Sub

Private Sub lstMemo_DblClick()
frmMemoDetails.Show
Me.Hide
End Sub
