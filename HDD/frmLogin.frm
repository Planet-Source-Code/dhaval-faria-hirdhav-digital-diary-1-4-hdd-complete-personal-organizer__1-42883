VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Login"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
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
   ScaleHeight     =   2535
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblForgotPassSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label lblForgotPass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password?"
      Height          =   465
      Left            =   4800
      TabIndex        =   14
      Top             =   1275
      Width           =   1410
   End
   Begin VB.Shape shapeForgotPass 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   4800
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label lblHelpSupport 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblNewUserSupport 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblNewUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New User"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   10
      Top             =   2060
      Width           =   825
   End
   Begin VB.Shape shapeNewUser 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblLogInSupport 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5265
      TabIndex        =   8
      Top             =   2055
      Width           =   405
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3680
      TabIndex        =   7
      Top             =   2055
      Width           =   585
   End
   Begin VB.Label lblLogIn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log In"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   2055
      Width           =   510
   End
   Begin VB.Shape shapeHelp 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   4800
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3240
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Shape shapeLogIn 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1680
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":030A
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6360
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6380
      X2              =   6380
      Y1              =   240
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   2520
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Login"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2400
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   6375
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
ChDir App.Path
MkDir "Data"
lblCaption.ForeColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
Me.BackColor = RGB(145, 155, 100)
shapeNewUser.BackColor = RGB(145, 155, 100)
shapeLogIn.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeHelp.BackColor = RGB(145, 155, 100)
txtUserName.BackColor = RGB(145, 155, 100)
txtPassword.BackColor = RGB(145, 155, 100)
shapeForgotPass.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
End
End Sub

Private Sub lblCancelSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = vbBlack
End Sub

Private Sub lblCancelSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = vbBlack
shapeCancel.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblForgotPassSupport_Click()
frmForgotPass.Show
Me.Hide
End Sub

Private Sub lblForgotPassSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblForgotPass.ForeColor = RGB(145, 155, 100)
shapeForgotPass.BackColor = vbBlack
End Sub

Private Sub lblForgotPassSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblForgotPass.ForeColor = vbBlack
shapeForgotPass.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblHelpSupport_Click()
frmHelp.Show
Unload Me
End Sub

Private Sub lblHelpSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.ForeColor = RGB(145, 155, 100)
shapeHelp.BackColor = vbBlack
End Sub

Private Sub lblHelpSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.ForeColor = vbBlack
shapeHelp.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblLogInSupport_Click()

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

On Error GoTo ErrHan

Do
        If txtUserName.Text & txtPassword.Text = ReS("Username") & ReS("Password") Then
            frmMain.Show
            Exit Sub
        End If
ReS.MoveNext
Loop

ErrHan:
If Err.Number = 3021 Then
    HDDMsgBox "Invalid Username or Password. Please try again."
End If


ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing
End Sub

Private Sub lblLogInSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogIn.ForeColor = RGB(145, 155, 100)
shapeLogIn.BackColor = vbBlack
End Sub

Private Sub lblLogInSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogIn.ForeColor = vbBlack
shapeLogIn.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblNewUserSupport_Click()
frmNewUser.Show
Unload Me
End Sub

Private Sub lblNewUserSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNewUser.ForeColor = RGB(145, 155, 100)
shapeNewUser.BackColor = vbBlack
End Sub

Private Sub lblNewUserSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNewUser.ForeColor = vbBlack
shapeNewUser.BackColor = RGB(145, 155, 100)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

If txtUserName.Text = "" Then
    txtUserName.SetFocus
    Exit Sub
End If

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

On Error GoTo ErrHan

Do
        If txtUserName.Text & txtPassword.Text = ReS("Username") & ReS("Password") Then
            frmMain.Show
            Exit Sub
        End If
ReS.MoveNext
Loop

ErrHan:
If Err.Number = 3021 Then
    HDDMsgBox "Invalid Username or Password. Please try again."
End If

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtPassword.SetFocus
If txtPassword.Text <> "" Then
    Dim db As Database
    Dim ReS As Recordset
    
    Set db = OpenDatabase(App.Path + "\HDD.dat")
    Set ReS = db.OpenRecordset("Users")
    
    On Error GoTo ErrHan
    
    Do
            If txtUserName.Text & txtPassword.Text = ReS("Username") & ReS("Password") Then
                frmMain.Show
                Exit Sub
            End If
    ReS.MoveNext
    Loop
    
ErrHan:
    If Err.Number = 3021 Then
        HDDMsgBox "Invalid Username or Password. Please try again."
    End If
    
    ReS.Close
    db.Close
    
    Set ReS = Nothing
    Set db = Nothing
End If
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub
