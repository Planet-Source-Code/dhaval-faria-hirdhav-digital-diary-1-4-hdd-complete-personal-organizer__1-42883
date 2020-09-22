VERSION 5.00
Begin VB.Form frmChangeQA 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Change Q/A"
   ClientHeight    =   3615
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
   Icon            =   "frmChangeQA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNewAnswer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtNewQuestion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblChangeSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   720
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblChange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      Height          =   225
      Left            =   1080
      TabIndex        =   16
      Top             =   3140
      Width           =   645
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3000
      TabIndex        =   14
      Top             =   3140
      Width           =   585
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2640
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Answer:"
      Height          =   225
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Question:"
      Height          =   225
      Left            =   765
      TabIndex        =   10
      Top             =   2040
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   225
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label lblOldAnswer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here is Old Answer."
      Height          =   225
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Answer:"
      Height          =   225
      Left            =   975
      TabIndex        =   6
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblOldQuestion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here is Old Question."
      Height          =   225
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Question:"
      Height          =   225
      Left            =   885
      TabIndex        =   4
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here is Username."
      Height          =   225
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   930
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Change Q / A"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3015
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5040
   End
   Begin VB.Shape shapeChange 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   720
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "frmChangeQA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblUsername.Caption = frmMain.lblUsername.Caption
txtPassword.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeChange.BackColor = RGB(145, 155, 100)
txtNewQuestion.BackColor = RGB(145, 155, 100)
txtNewAnswer.BackColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

5 If ReS("Username") = lblUsername.Caption Then
    lblOldQuestion.Caption = ReS("Question")
    lblOldAnswer.Caption = ReS("Answer")
    
    ReS.Close
    db.Close
    
    Set ReS = Nothing
    Set db = Nothing
    Exit Sub
Else
    ReS.MoveNext
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
HDDYesNoBox "Are you sure do you want to change Q/A?"
    If Yes Then
        GoTo 12
    Else
        Exit Sub
    End If
12 Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

10 If ReS("Username") = lblUsername.Caption Then
        If ReS("Password") = txtPassword.Text Then
            ReS.Edit
            ReS("Question") = txtNewQuestion.Text
            ReS("Answer") = txtNewAnswer.Text
            ReS.Update
            
            ReS.Close
            db.Close
            
            Set ReS = Nothing
            Set db = Nothing
            
            HDDMsgBox "Q / A changed successfully."
            frmAccountPref.Show
            Unload Me
        Else
            HDDMsgBox "Sorry, Password is not matching, Please try again."
            
            ReS.Close
            db.Close
            
            Set ReS = Nothing
            Set db = Nothing
        End If
Else
    ReS.MoveNext
    GoTo 10
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

Private Sub txtNewAnswer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNewQuestion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub
