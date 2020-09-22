VERSION 5.00
Begin VB.Form frmBirthRemind 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary  -  Birthday Reminder"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
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
   Icon            =   "frmBirthRemind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstInfo 
      Appearance      =   0  'Flat
      Height          =   2280
      ItemData        =   "frmBirthRemind.frx":030A
      Left            =   360
      List            =   "frmBirthRemind.frx":030C
      TabIndex        =   1
      Top             =   480
      Width           =   4335
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
      Left            =   3480
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   3020
      Width           =   1215
   End
   Begin VB.Label lblEditSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   225
      Left            =   1920
      TabIndex        =   6
      Top             =   3015
      Width           =   1245
   End
   Begin VB.Label lblNewSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3020
      Width           =   1215
   End
   Begin VB.Label lblMainMenuSupport 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reminder Menu"
      Height          =   225
      Left            =   360
      TabIndex        =   2
      Top             =   3680
      Width           =   4275
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Shape shapeMainMenu 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Shape shapeDelete 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3480
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Shape shapeEdit 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1920
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Shape shapeNew 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Birthday Reminder"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3510
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5040
   End
End
Attribute VB_Name = "frmBirthRemind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername

Private Sub Form_Load()
lblCaption.ForeColor = RGB(145, 155, 100)
lstInfo.BackColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeNew.BackColor = RGB(145, 155, 100)
shapeEdit.BackColor = RGB(145, 155, 100)
shapeDelete.BackColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption
Me.BackColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

On Error GoTo ErrHan:

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\BR.dat")
Set ReS = db.OpenRecordset("BR")

Do
    lstInfo.AddItem ReS("RDate") & " - " & ReS("Name") & " - " & ReS("PHNo")
    ReS.MoveNext
Loop

ReS.Close
db.Close

Set db = Nothing
Set ReS = Nothing

ErrHan:
    If Err.Number = "3021" Then
        Exit Sub
    End If
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblDeleteSupport_Click()
If lstInfo.ListIndex = "-1" Then
    HDDMsgBox "Please select the item from the list box."
Else
    HDDYesNoBox "Are you sure? Do you want to delete it?"
    
    If Yes Then
        Dim db As Database
        Dim ReS As Recordset
        
        Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\BR.dat")
        Set ReS = db.OpenRecordset("BR")
        
        ReS.Move (lstInfo.ListIndex)
        ReS.Delete
        
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        Unload Me
        Me.Show
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
If lstInfo.ListIndex = "-1" Then
    HDDMsgBox "Please select the item from the list box."
Else
    frmEditBirthRemind.Show
    Me.Hide
End If
End Sub

Private Sub lblEditSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = RGB(145, 155, 100)
shapeEdit.BackColor = vbBlack
End Sub

Private Sub lblEditSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = vbBlack
shapeEdit.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblMainMenuSupport_Click()
frmReminders.Show
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
frmNewBR.Show
Me.Hide
End Sub

Private Sub lblNewSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNew.ForeColor = RGB(145, 155, 100)
shapeNew.BackColor = vbBlack
End Sub

Private Sub lblNewSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeNew.BackColor = RGB(145, 155, 100)
lblNew.ForeColor = vbBlack
End Sub

Private Sub lstInfo_DblClick()
frmBRDetail.Show
Me.Hide
End Sub
