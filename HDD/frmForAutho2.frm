VERSION 5.00
Begin VB.Form frmForAutho2 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Forgot Autho (Step 2)"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
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
   Icon            =   "frmForAutho2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblGetAuthoCODESupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblGetAuthoCODE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Get Autho CODE"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Shape shapeGetAuthoCODE 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2760
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   960
      TabIndex        =   12
      Top             =   2760
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer:"
      Height          =   225
      Left            =   840
      TabIndex        =   10
      Top             =   3600
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question:"
      Height          =   225
      Left            =   720
      TabIndex        =   8
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   225
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   930
   End
   Begin VB.Line Line4 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   5040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmForAutho2.frx":030A
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5295
      X2              =   5295
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Forgot Autho (Step 2)"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5295
   End
End
Attribute VB_Name = "frmForAutho2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeCancel.BackColor = RGB(145, 155, 100)
shapeGetAuthoCODE.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
txtUserName.BackColor = RGB(145, 155, 100)
txtPassword.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtQuestion.BackColor = RGB(145, 155, 100)
txtAnswer.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
frmForAutho1.Show
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

Private Sub lblGetAuthoCODESupport_Click()
If txtUserName.Text = "" Then
    HDDMsgBox "Please enter your UserName."
    Exit Sub
End If
If InStr(1, txtUserName.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
End If
If txtPassword.Text = "" Then
    HDDMsgBox "Please enter your Password."
    Exit Sub
End If
If InStr(1, txtPassword.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
End If
If txtEMail.Text = "" Then
    HDDMsgBox "Please enter your E-Mail Address."
    Exit Sub
End If
If InStr(1, txtEMail.Text, " ") Then
    HDDMsgBox "Sorry, Spaces are not allowed."
    Exit Sub
End If
If txtQuestion.Text = "" Then
    HDDMsgBox "Please enter your Qyestion."
    Exit Sub
End If
If txtAnswer.Text = "" Then
    HDDMsgBox "Please enter your Answer."
    Exit Sub
End If

Dim flags As Long
Dim result As Boolean

result = InternetGetConnectedState(flags, 0)

If result Then
    frmForAutho3.Show
    Me.Hide
    Exit Sub
Else
    HDDMsgBox "You are not connected to the Internet. Please connect to the Internet."
    Exit Sub
End If
End Sub

Private Sub lblGetAuthoCODESupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGetAuthoCODE.ForeColor = RGB(145, 155, 100)
shapeGetAuthoCODE.BackColor = vbBlack
End Sub

Private Sub lblGetAuthoCODESupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeGetAuthoCODE.BackColor = RGB(145, 155, 100)
lblGetAuthoCODE.ForeColor = vbBlack
End Sub

Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEMail_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub
