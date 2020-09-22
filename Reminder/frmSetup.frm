VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Dairy - Reminder Setup"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
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
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAgain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Don't Ask Me Username and Password Again."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1560
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3240
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   480
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Here, You can change the configuration of the Hirdhav Digital Diary - Reminder."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Reminder Setup"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3300
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5280
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
chkAgain.BackColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption
shapeOk.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Reminder.dat")
Set ReS = db.OpenRecordset("Reminder")

If ReS("Auto") = "Yes" Then
    chkAgain.Value = 1
Else
    chkAgain.Value = 0
End If

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing
End Sub

Private Sub lblCancelSupport_Click()
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

Private Sub lblOkSupport_Click()
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Reminder.dat")
Set ReS = db.OpenRecordset("Reminder")

If chkAgain.Value = 1 Then
    ReS.Edit
    ReS("Auto") = "Yes"
    ReS("Username") = strUsername
    ReS.Update
Else
    ReS.Edit
    ReS("Auto") = "No"
    ReS.Update
End If

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

MsgBox "Settings has been changed successfully."
Unload Me
End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeOk.BackColor = vbBlack
lblOk.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = vbBlack
shapeOk.BackColor = RGB(145, 155, 100)
End Sub
