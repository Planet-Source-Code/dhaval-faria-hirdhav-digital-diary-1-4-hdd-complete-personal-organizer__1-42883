VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmForgotPass 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Forgot Password?"
   ClientHeight    =   5535
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
   Icon            =   "frmForgotPass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   330
      Left            =   5000
      TabIndex        =   23
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   3120
      Width           =   2415
   End
   Begin MSMask.MaskEdBox txtBDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   1920
      TabIndex        =   9
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblYourPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Pass"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   480
      Left            =   1680
      TabIndex        =   22
      Top             =   4920
      Width           =   2010
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   4920
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Password:"
      Height          =   225
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Width           =   1350
   End
   Begin VB.Line Line4 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   120
      X2              =   4800
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel / Ok"
      Height          =   225
      Left            =   2985
      TabIndex        =   19
      Top             =   4080
      Width           =   975
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2760
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblRPSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblRP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retrive Password"
      Height          =   465
      Left            =   480
      TabIndex        =   17
      Top             =   3980
      Width           =   1395
   End
   Begin VB.Shape shapeRP 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   480
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer:"
      Height          =   225
      Left            =   960
      TabIndex        =   15
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question:"
      Height          =   225
      Left            =   960
      TabIndex        =   13
      Top             =   3120
      Width           =   810
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD / MM / YYYY"
      Height          =   225
      Left            =   2505
      TabIndex        =   12
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      Height          =   225
      Left            =   900
      TabIndex        =   10
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   1240
      TabIndex        =   8
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   225
      Left            =   840
      TabIndex        =   6
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   225
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   960
   End
   Begin VB.Line Line3 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   120
      X2              =   4800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please provide some of the information, so we can retrive your password."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4920
      X2              =   4920
      Y1              =   240
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   5520
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Forgot Password?"
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
      Width           =   4920
   End
End
Attribute VB_Name = "frmForgotPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public MyPass As String

Private Sub Form_Load()
shapeCancel.BackColor = RGB(145, 155, 100)
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
txtUsername.BackColor = RGB(145, 155, 100)
txtFName.BackColor = RGB(145, 155, 100)
txtLName.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtBDate.BackColor = RGB(145, 155, 100)
txtQuestion.BackColor = RGB(145, 155, 100)
txtAnswer.BackColor = RGB(145, 155, 100)
shapeRP.BackColor = RGB(145, 155, 100)
lblYourPass.Caption = ""
txtUsername.Text = frmLogin.txtUsername.Text
txtUsername.SelLength = Len(txtUsername.Text)
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
shapeCancel.BackColor = RGB(145, 155, 100)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblRPSupport_Click()
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")
On Error GoTo HanErr:
Do
    If ReS("Username") = txtUsername.Text And ReS("FirstName") = txtFName.Text And _
        ReS("LastName") = txtLName.Text And ReS("BDate") = txtBDate.Text And _
        ReS("EMail") = txtEMail.Text And ReS("Question") = txtQuestion.Text And _
        ReS("Answer") = txtAnswer.Text Then
        MyPass = ReS("Password")
        txtPass.Text = MyPass
        ShowPass
        Exit Sub
    End If
    ReS.MoveNext
Loop

HanErr:
    If Err.Number = 3021 Then
        lblYourPass.Caption = "N/A"
        HDDMsgBox "Sorry, informations are not matching."
    End If
End Sub

Private Sub lblRPSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRP.ForeColor = RGB(145, 155, 100)
shapeRP.BackColor = vbBlack
End Sub

Private Sub lblRPSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRP.ForeColor = vbBlack
shapeRP.BackColor = RGB(145, 155, 100)
End Sub

Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Function ShowPass()
Dim EachLetter
Dim SelStart
Dim SelLen
Dim AllLetters
SelStart = 0
SelLen = 1
Do
    For I = 97 To 122
        txtPass.SetFocus
        txtPass.SelStart = SelStart
        txtPass.SelLength = SelLen
        EachLetter = Chr(I)
        If SelStart = Len(txtPass.Text) Then
            Exit Function
        End If
        If EachLetter = txtPass.SelText Then
            txtPass.SetFocus
            txtPass.SelStart = SelStart
            txtPass.SelLength = SelLen
            SelStart = SelStart + 1
            AllLetters = AllLetters + EachLetter
            lblYourPass.Caption = AllLetters
        End If
    Next I
Loop
End Function
