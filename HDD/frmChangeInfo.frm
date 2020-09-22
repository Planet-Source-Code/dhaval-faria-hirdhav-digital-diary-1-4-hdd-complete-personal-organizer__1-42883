VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmChangeInfo 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Change Information"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
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
   Icon            =   "frmChangeInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   2520
      Width           =   2055
   End
   Begin MSMask.MaskEdBox txtBDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   225
      Left            =   1040
      TabIndex        =   18
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblChangeSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   720
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3375
      TabIndex        =   14
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label lblChange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      Height          =   225
      Left            =   1080
      TabIndex        =   13
      Top             =   3120
      Width           =   645
   End
   Begin VB.Shape shapeChange 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   720
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3000
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD / MM / YYYY"
      Height          =   225
      Left            =   2280
      TabIndex        =   12
      Top             =   2280
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      Height          =   225
      Left            =   1080
      TabIndex        =   10
      Top             =   1920
      Width           =   885
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5060
      X2              =   5060
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   225
      Left            =   1030
      TabIndex        =   7
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   225
      Left            =   1030
      TabIndex        =   4
      Top             =   840
      Width           =   960
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here is Username."
      Height          =   225
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Change Information"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3600
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5055
   End
End
Attribute VB_Name = "frmChangeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lblUsername.Caption = frmMain.lblUsername.Caption
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
txtFName.BackColor = RGB(145, 155, 100)
txtLName.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtBDate.BackColor = RGB(145, 155, 100)
shapeChange.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
txtPassword.BackColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

5 If ReS("Username") = lblUsername.Caption Then
    txtFName.Text = ReS("FirstName")
    txtLName.Text = ReS("LastName")
    txtEMail.Text = ReS("EMail")
    txtBDate.Text = ReS("BDate")
    
    ReS.Close
    db.Close
    
    Set ReS = Nothing
    Set db = Nothing
Else
    GoTo 5
End If
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
shapeCancel.BackColor = RGB(145, 155, 100)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblChangeSupport_Click()
If txtPassword.Text = "" Then
    HDDMsgBox "Please enter your password."
    Exit Sub
End If
HDDYesNoBox "Are you sure? do you want to change the information?"
    If Yes Then
        GoTo 10
    Else
        Exit Sub
    End If
10 Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

If ReS("Password") = txtPassword.Text Then
    
    ReS.Edit
    ReS("FirstName") = txtFName.Text
    ReS("LastName") = txtLName.Text
    ReS("EMail") = txtEMail.Text
    ReS("BDate") = txtBDate.Text
    ReS.Update
    
    ReS.Close
    db.Close
    
    Set ReS = Nothing
    Set db = Nothing
    
    HDDMsgBox "Information is successfully changed."
    frmAccountPref.Show
    Unload Me
    Exit Sub
Else
    HDDMsgBox "Sorry, Password is not matching, Please try again."
    
    ReS.Close
    db.Close
    
    Set ReS = Nothing
    Set db = Nothing
    
    Exit Sub
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

Private Sub txtBDate_GotFocus()
txtBDate.SelStart = 0
txtBDate.SelLength = Len(txtBDate.Text)
End Sub

Private Sub txtEMail_GotFocus()
txtEMail.SelStart = 0
txtEMail.SelLength = Len(txtEMail.Text)
End Sub

Private Sub txtFName_GotFocus()
txtFName.SelStart = 0
txtFName.SelLength = Len(txtFName.Text)
End Sub

Private Sub txtLName_GotFocus()
txtLName.SelStart = 0
txtLName.SelLength = Len(txtLName.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub
