VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Change Password"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
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
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtNewPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtOldPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblChangeSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblChange 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   2530
      Width           =   1335
   End
   Begin VB.Shape shapeChange 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3260
      TabIndex        =   9
      Top             =   2660
      Width           =   585
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2880
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      Height          =   225
      Left            =   675
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      Height          =   225
      Left            =   765
      TabIndex        =   3
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   225
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Change Password"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3495
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4680
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lblUsername.Caption = frmMain.lblUsername.Caption
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeChange.BackColor = RGB(145, 155, 100)
txtOldPassword.BackColor = RGB(145, 155, 100)
txtCPassword.BackColor = RGB(145, 155, 100)
txtNewPassword.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
frmAccountPref.Show
Unload Me
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

Private Sub lblChangeSupport_Click()
If txtOldPassword.Text = "" Then
    HDDMsgBox "Please enter the old password."
    Exit Sub
End If

HDDYesNoBox "Are you sure, do you want to change your password?"

If Yes Then
    GoTo 1
Else
    Exit Sub
End If

1 Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

5 If ReS("Username") = lblUsername.Caption Then
    If txtOldPassword.Text <> ReS("Password") Then
        HDDMsgBox "Sorry, your Old Password is not matching."
        
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        Exit Sub
    ElseIf txtNewPassword.Text <> txtCPassword.Text Then
        HDDMsgBox "Sorry, your new password is not matching."
        
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        Exit Sub
    Else
        ReS.Edit
        ReS("Password") = txtNewPassword.Text
        ReS.Update
        HDDMsgBox "Your Password is successfully changed."
        
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
    
        frmAccountPref.Show
        Unload Me
    End If
Else
    ReS.MoveNext
    GoTo 5
End If

End Sub

Private Sub lblChangeSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblChange.ForeColor = RGB(145, 155, 100)
shapeChange.BackColor = vbBlack
End Sub

Private Sub lblChangeSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeChange.BackColor = RGB(145, 155, 100)
lblChange.ForeColor = vbBlack
End Sub

Private Sub txtCPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtOldPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub
