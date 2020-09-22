VERSION 5.00
Begin VB.Form frmGetAutho2 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Get Autho (Step 2)"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
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
   Icon            =   "frmGetAutho2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGender 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "M/F"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   23
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   21
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtUName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtCountry 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox txtState 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox txtCity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Male / Female"
      Height          =   225
      Left            =   2640
      TabIndex        =   31
      Top             =   2880
      Width           =   1155
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender :"
      Height          =   225
      Left            =   1200
      TabIndex        =   30
      Top             =   2880
      Width           =   705
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   29
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblCheckINETSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2040
      TabIndex        =   28
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblNextSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4320
      TabIndex        =   27
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lblCheckINET 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check Internet Status"
      Height          =   255
      Left            =   2040
      TabIndex        =   25
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblNext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next >"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Shape shapeCheckINET 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2040
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Shape shapeNext 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   4320
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   6120
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hint Answer :"
      Height          =   225
      Left            =   840
      TabIndex        =   22
      Top             =   5040
      Width           =   1125
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hint Question :"
      Height          =   225
      Left            =   720
      TabIndex        =   20
      Top             =   4680
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   225
      Left            =   960
      TabIndex        =   19
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   225
      Left            =   960
      TabIndex        =   18
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country :"
      Height          =   225
      Left            =   1200
      TabIndex        =   16
      Top             =   4320
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   225
      Left            =   1440
      TabIndex        =   14
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   225
      Left            =   1560
      TabIndex        =   12
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
      Height          =   225
      Left            =   1320
      TabIndex        =   10
      Top             =   3240
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   225
      Left            =   960
      TabIndex        =   6
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   225
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   960
   End
   Begin VB.Line Line3 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   5880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6135
      X2              =   6120
      Y1              =   240
      Y2              =   6120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetAutho2.frx":030A
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   5895
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   6120
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Get Autho (Step 2)"
      Height          =   225
      Left            =   200
      TabIndex        =   2
      Top             =   10
      Width           =   3465
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   10
      Top             =   10
      Width           =   6135
   End
End
Attribute VB_Name = "frmGetAutho2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
txtUName.BackColor = RGB(145, 155, 100)
txtPassword.BackColor = RGB(145, 155, 100)
txtFName.BackColor = RGB(145, 155, 100)
txtLName.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtCity.BackColor = RGB(145, 155, 100)
txtState.BackColor = RGB(145, 155, 100)
txtCountry.BackColor = RGB(145, 155, 100)
txtQuestion.BackColor = RGB(145, 155, 100)
txtAnswer.BackColor = RGB(145, 155, 100)
txtGender.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeNext.BackColor = RGB(145, 155, 100)
shapeCheckINET.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
frmGetAutho1.Show
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

Private Sub lblCheckINETSupport_Click()
Dim flags As Long
Dim result As Boolean

result = InternetGetConnectedState(flags, 0)

If result Then
    HDDMsgBox "Internet Connection is Working..."
    Exit Sub
Else
    HDDMsgBox "Internet Connection is not working..."
    Exit Sub
End If
End Sub

Private Sub lblCheckINETSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCheckINET.ForeColor = RGB(145, 155, 100)
shapeCheckINET.BackColor = vbBlack
End Sub

Private Sub lblCheckINETSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCheckINET.BackColor = RGB(145, 155, 100)
lblCheckINET.ForeColor = vbBlack
End Sub

Private Sub lblNextSupport_Click()
If txtUName.Text = "" Then
    HDDMsgBox "Please insert Username."
    Exit Sub
ElseIf InStr(1, txtUName.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
ElseIf txtPassword.Text = "" Then
    HDDMsgBox "Please insert Password."
    Exit Sub
ElseIf InStr(1, txtPassword.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
ElseIf txtFName.Text = "" Then
    HDDMsgBox "Please enter your First Name."
    Exit Sub
ElseIf InStr(1, txtFName.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
ElseIf txtLName.Text = "" Then
    HDDMsgBox "Please enter your Last Name."
    Exit Sub
ElseIf InStr(1, txtLName.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
ElseIf txtGender.Text <> "M" Then
    If txtGender.Text <> "F" Then
        HDDMsgBox "Invalid Gender."
        Exit Sub
    End If
ElseIf InStr(1, txtEMail.Text, " ") Then
    HDDMsgBox "Invalid E-Mail Address."
    Exit Sub
ElseIf InStr(1, txtCity.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
ElseIf txtCity.Text = "" Then
    HDDMsgBox "Please enter your city name."
    Exit Sub
ElseIf txtState.Text = "" Then
    HDDMsgBox "Please enter your state name."
    Exit Sub
ElseIf InStr(1, txtState.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
ElseIf txtCountry.Text = "" Then
    HDDMsgBox "Please enter your Country name."
    Exit Sub
ElseIf InStr(1, txtCountry.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
ElseIf txtQuestion.Text = "" Then
    HDDMsgBox "Please enter your Hint Question."
    Exit Sub
ElseIf txtAnswer.Text = "" Then
    HDDMsgBox "Please enter your Hint Answer."
    Exit Sub
End If
DoEvents
frmGetAutho3.Show
frmGetAutho3.Timer1.Enabled = True
DoEvents
Me.Hide
DoEvents
End Sub

Private Sub lblNextSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = RGB(145, 155, 100)
shapeNext.BackColor = vbBlack
End Sub

Private Sub lblNextSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeNext.BackColor = RGB(145, 155, 100)
lblNext.ForeColor = vbBlack
End Sub

Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEMail_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtGender_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub
