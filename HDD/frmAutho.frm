VERSION 5.00
Begin VB.Form frmAutho 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Authentication"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
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
   Icon            =   "frmAutho.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   1560
   End
   Begin VB.TextBox txtAutho 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblHelpSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblExitSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4005
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2895
      TabIndex        =   9
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1245
      TabIndex        =   8
      Top             =   1680
      Width           =   240
   End
   Begin VB.Shape shapeHelp 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   4080
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Shape shapeExit 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2400
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   720
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblGetAuthoSupport 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblGetAutho 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Get your Authent. CODE"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   800
      Width           =   1335
   End
   Begin VB.Shape shapeGetAutho 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   4560
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Authenti CODE:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your name and your authentication CODE to go further."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Authentication"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3165
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   6240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6255
      X2              =   6240
      Y1              =   240
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   2160
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   6255
   End
End
Attribute VB_Name = "frmAutho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public CorrectCode As Boolean
Public LenOfText
Public LenOfASC
Public MakeTotal
Public HaLfCoDe
Public FiNaLcOdE
Public CuStOmEcOdE

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
txtName.BackColor = RGB(145, 155, 100)
txtAutho.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeGetAutho.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
shapeExit.BackColor = RGB(145, 155, 100)
shapeHelp.BackColor = RGB(145, 155, 100)
CorrectCode = False
On Error GoTo AA
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Installation.dat")
Set ReS = db.OpenRecordset("Installation")

If ReS("Autho") = "Yes" Then
    frmWelcome.Show
    ReS.Close
    db.Close
    Set db = Nothing
    Set ReS = Nothing
    Unload Me
Exit Sub
End If

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing
SendKeys "{TAB}"

AA:
If Err.Number = 3021 Then
Exit Sub
End If
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblExitSupport_Click()
End
End Sub

Private Sub lblExitSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeExit.BackColor = vbBlack
lblExit.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblExitSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeExit.BackColor = RGB(145, 155, 100)
lblExit.ForeColor = vbBlack
End Sub

Private Sub lblGetAuthoSupport_Click()
frmAuthoHelp.Show
frmAutho.Hide
End Sub

Private Sub lblGetAuthoSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeGetAutho.BackColor = vbBlack
lblGetAutho.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblGetAuthoSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeGetAutho.BackColor = RGB(145, 155, 100)
lblGetAutho.ForeColor = vbBlack
End Sub

Private Sub lblHelpSupport_Click()
frmAuthoHelp.Show
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

Private Sub lblOkSupport_Click()
If txtName.Text = "" Then
    HDDMsgBox "Please enter your name"
    Exit Sub
End If
CheckCode
If CorrectCode = True Then
    Dim db As Database
    Dim ReS As Recordset
    
    Set db = OpenDatabase(App.Path + "\Installation.dat")
    Set ReS = db.OpenRecordset("Installation")
    
    ReS.Edit
    ReS("Autho") = "Yes"
    ReS.Update
    
    ReS.Close
    db.Close
    
    Set ReS = Nothing
    Set db = Nothing
    
    frmWelcome.Show
    Unload frmAutho
    Exit Sub
Else
    HDDMsgBox "Please Enter Correct Authentication CODE."
End If
End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = vbBlack
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = vbBlack
shapeOk.BackColor = RGB(145, 155, 100)
End Sub

Private Sub Timer1_Timer()
If txtName.Text = "Show me my AUTHO." Then
    txtAutho.Enabled = True
    txtName.Text = ""
    Exit Sub
End If
End Sub

Private Sub txtAutho_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Function CheckCode()
HaLfCoDe = ""
FiNaLcOdE = ""
LenOfText = Len(txtName.Text)
For j = 0 To LenOfText - 1
    txtName.SetFocus
    txtName.SelStart = j
    txtName.SelLength = 1
    HaLfCoDe = HaLfCoDe & Asc(txtName.SelText)
Next j
CuStOmEcOdE = txtAutho.Text
txtAutho.Text = HaLfCoDe

txtAutho.SetFocus
txtAutho.SelStart = 0
txtAutho.SelLength = 1
FiNaLcOdE = txtAutho.SelText
LenOfASC = Len(txtAutho.Text)
For k = 0 To LenOfASC - 1
    txtAutho.SetFocus
    txtAutho.SelStart = k
    txtAutho.SelLength = 1
    MakeTotal = txtAutho.SelText
    k = k + 1
    txtAutho.SetFocus
    txtAutho.SelStart = k
    txtAutho.SelLength = 1
    FiNaLcOdE = FiNaLcOdE & Val(MakeTotal) + Val(txtAutho.SelText)
    k = k - 1
Next k
txtAutho.Text = CuStOmEcOdE
If FiNaLcOdE = CuStOmEcOdE Then
    CorrectCode = True
End If
End Function
