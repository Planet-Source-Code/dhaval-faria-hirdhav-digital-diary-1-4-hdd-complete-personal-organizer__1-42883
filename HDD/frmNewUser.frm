VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmNewUser 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - New User"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
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
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtBirthDate 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      ToolTipText     =   "Incase if you forgot your Password, This will help you in remembering your Password."
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      TabIndex        =   15
      ToolTipText     =   "Incase if you forgot your Password, This will help you in remembering your Password."
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      ToolTipText     =   "Incase if you forgot your Password, This will help you in remembering your Password."
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtLastName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblAddSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   840
      TabIndex        =   20
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3120
      TabIndex        =   19
      Top             =   3615
      Width           =   585
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2760
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      Height          =   225
      Left            =   1300
      TabIndex        =   18
      Top             =   3615
      Width           =   330
   End
   Begin VB.Shape shapeAdd 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   840
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD/MM/YYYY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3825
      TabIndex        =   17
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      Height          =   225
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer:"
      Height          =   225
      Left            =   600
      TabIndex        =   13
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question:"
      Height          =   225
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   720
      TabIndex        =   9
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   225
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   960
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  New User"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2760
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4920
      X2              =   4920
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4920
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
txtFirstName.BackColor = RGB(145, 155, 100)
txtLastName.BackColor = RGB(145, 155, 100)
txtUserName.BackColor = RGB(145, 155, 100)
txtPassword.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtBirthDate.BackColor = RGB(145, 155, 100)
txtQuestion.BackColor = RGB(145, 155, 100)
txtAnswer.BackColor = RGB(145, 155, 100)
shapeAdd.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblAddSupport_Click()
'Check username is entered or blank?
'If blank show err message.
If txtUserName.Text = "" Then
    MsgBox "Please enter username."
    Exit Sub
ElseIf txtUserName.Text = " " Then
    MsgBox "Please enter username."
    Exit Sub
End If

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

'First Check whether Username is already exists in
'database or not?
'On Error GoTo ErrHan
Do While Not ReS.EOF
'If Username is already exist show err message
'else add user
        If txtUserName.Text = ReS("Username") Then
            HDDMsgBox "Sorry, Username is Already taken, please select another Username."
            Exit Sub
        End If
ReS.MoveNext
Loop

        ReS.AddNew
        ReS("FirstName") = txtFirstName.Text
        ReS("LastName") = txtLastName.Text
        ReS("Username") = txtUserName.Text
        ReS("Password") = txtPassword.Text
        ReS("Question") = txtQuestion.Text
        ReS("Answer") = txtAnswer.Text
        ReS("BDate") = txtBirthDate.Text
        ReS("EMail") = txtEMail.Text
        ReS.Update
        ChDir App.Path
        ChDir "Data"
        On Error Resume Next
        MkDir txtUserName.Text
        ChDir txtUserName.Text
        'Copy all the Data Tables to the Users Directory.
s:
        ChDir App.Path
        s1 = App.Path + "\Personal.dat"
        s2 = App.Path + "\Memo.dat"
        s3 = App.Path + "\Sch.dat"
        
r:
        ChDir App.Path
        ChDir "Data"
        ChDir txtUserName.Text
        r1 = "Personal.dat"
        r2 = "Memo.dat"
        r3 = "Sch.dat"
        
        FileCopy s1, r1
        FileCopy s2, r2
        FileCopy s3, r3
        
        Set s1 = Nothing
        Set r1 = Nothing
        Set s2 = Nothing
        Set r2 = Nothing
        Set s3 = Nothing
        Set r3 = Nothing
        
        ChDir App.Path
    
    HDDMsgBox "Congratulation, Account is created."
    frmLogin.Show
    Unload Me
    
End Sub

Private Sub lblAddSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = RGB(145, 155, 100)
shapeAdd.BackColor = vbBlack
End Sub

Private Sub lblAddSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = vbBlack
shapeAdd.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
frmLogin.Show
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

Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
'All letters will be in Upercase
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
'All letters will be in Upercase
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
If KeyAscii = 32 Then
HDDMsgBox "No spaces are allowed in Username."
SendKeys "{BACKSPACE}"
End If
End Sub
