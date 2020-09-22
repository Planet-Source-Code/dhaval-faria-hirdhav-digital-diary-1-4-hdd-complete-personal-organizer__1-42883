VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Reminder Login"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
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
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAgain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Don't ask me Username and Password again."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3600
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblLogInSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblLogIn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Log In"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6600
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Shape shapeLogIn 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1320
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   225
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   930
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   240
      Y2              =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":030A
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6375
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Reminder Login"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3270
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   2880
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   6600
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
App.TaskVisible = False
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeCancel.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
txtUsername.BackColor = RGB(145, 155, 100)
txtPassword.BackColor = RGB(145, 155, 100)
chkAgain.BackColor = RGB(145, 155, 100)
shapeLogIn.BackColor = RGB(145, 155, 100)

On Error GoTo AA

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Reminder.dat")
Set ReS = db.OpenRecordset("Reminder")

If ReS("Auto") = "Yes" Then
    txtUsername.Text = ReS("Username")
    ReS.Close
    db.Close
    Set ReS = Nothing
    Set db = Nothing
    frmMain.Show
Else
    ReS.Close
    db.Close
    Set ReS = Nothing
    Set db = Nothing
    Exit Sub
End If

AA:
If Err.Number = 3021 Then
    Exit Sub
End If
End Sub

Private Sub lblCancelSupport_Click()
End
End Sub

Private Sub lblCancelSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = vbBlack
End Sub

Private Sub lblCancelSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel.BackColor = RGB(145, 155, 100)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblLogInSupport_Click()

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

On Error GoTo ErrHan
ReS.MoveFirst
Do
    If txtUsername.Text & txtPassword.Text = ReS("Username") & ReS("Password") Then
        ReS.Close
        db.Close
        
        Set db = OpenDatabase(App.Path + "\Reminder.dat")
        Set ReS = db.OpenRecordset("Reminder")
        
        ReS.Edit
        ReS("Username") = txtUsername.Text
        ReS("Password") = txtPassword.Text
        If chkAgain.Value = 1 Then
            ReS("Auto") = "Yes"
        Else
            ReS("Auto") = "No"
        End If
        ReS.Update
        
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        
        frmMain.Show
        Exit Sub
    Else
        ReS.MoveNext
    End If
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

ErrHan:
    If Err.Number = 3021 Then
        MsgBox "Invalid Username or Password. Please try again."
        ReS.Close
        db.Close
        Set ReS = Nothing
        Set db = Nothing
    End If
End Sub

Private Sub lblLogInSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogIn.ForeColor = RGB(145, 155, 100)
shapeLogIn.BackColor = vbBlack
End Sub

Private Sub lblLogInSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeLogIn.BackColor = RGB(145, 155, 100)
lblLogIn.ForeColor = vbBlack
End Sub
