VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmGetAutho3 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Get Autho (Step 3)"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
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
   Icon            =   "frmGetAutho3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   2040
   End
   Begin VB.TextBox txtAutho 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3360
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2160
      TabIndex        =   11
      Text            =   "Dhaval Faria"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblDoAgainSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblDoAgain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Do Again"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape shapeDoAgain 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1800
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblShowAuthoSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblShowAutho 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Show Authentication CODE"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Shape shapeShowAutho 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3240
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "HERE IS STATUS..."
      Height          =   705
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   4530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetAutho3.frx":030A
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5760
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5780
      X2              =   5780
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
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Get Autho (Step 3)"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3465
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5775
   End
End
Attribute VB_Name = "frmGetAutho3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public LenOfText
Public LenOfASC
Public MakeTotal
Public HaLfCoDe
Public FiNaLcOdE
Public CuStOmEcOdE

Private Sub Form_Load()
lblShowAutho.Enabled = False
lblShowAuthoSupport.Enabled = False
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeShowAutho.BackColor = RGB(145, 155, 100)
shapeDoAgain.BackColor = RGB(145, 155, 100)
lblDoAgain.Enabled = False
lblDoAgainSupport.Enabled = False
lblStatus.Caption = "Please Wait... Checking Internet Connection..."
txtName.Text = frmGetAutho2.txtFName.Text + " " + frmGetAutho2.txtLName.Text
End Sub

Private Sub lblCancelSupport_Click()
frmGetAutho2.Show
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

Public Function DoAll()
DoEvents
Dim flags As Long
Dim result As Boolean

result = InternetGetConnectedState(flags, 0)

If result Then
    lblStatus.Caption = "You are connected to the Internet..."
Else
    lblINetStatus.Caption = "You are not connected to the internet.. make sure that you are connected to the internet and than click on Do Again Button."
    lblDoAgain.Enabled = True
    lblDoAgainSupport.Enabled = True
    Exit Function
End If

Call GenCODE

lblStatus.Caption = "Sending Data to the Web..."

With frmGetAutho2

'Inet1.OpenURL ("http://localhost/Put.asp?FirstName=" + .txtFName.Text + "=LastName=" + .txtLName.Text + "=Gender=" + .txtGender.Text + "=UName=" + .txtUName.Text + "=Pass=" + .txtPassword.Text + "=EMail=" + .txtEMail.Text + "=City=" & _
              .txtCity.Text + "=State=" + .txtState.Text + "=Country=" + .txtCountry.Text + "=HQ=" + .txtQuestion.Text + "=HA=" + .txtAnswer.Text + "=Autho=" + FiNaLcOdE + "=Ver=" + "1.3")

'http://localhost/HDD/GetCode/Put.asp?FirstName=Dhaval=LastName=Faria=Gender=M=UName=dhavalhirdhav=Pass=faria=EMail=dhavalhirdhav@yahoo.com=City=Mumbai=State=Maharashtra=Country=INDIA=HQ=HOW ARE YOU?=HA=I AM FINE.=Autho=16456=Ver=1.3

'MsgBox "http://localhost/HDD/GetAutho/Put.asp?FirstName=" + .txtFName.Text + "=LastName=" + .txtLName.Text + "=Gender=" + .txtGender.Text + "=UName=" + .txtUName.Text + "=Pass=" + .txtPassword.Text + "=EMail=" + .txtEMail.Text + "=City=" & _
              .txtCity.Text + "=State=" + .txtState.Text + "=Country=" + .txtCountry.Text + "=HQ=" + .txtQuestion.Text + "=HA=" + .txtAnswer.Text + "=Autho=" + FiNaLcOdE + "=Ver=" + "1.3"

'strAnswer = Inet1.OpenURL("http://localhost/HDD/GetCode/Put.asp?FirstName=Dhaval=LastName=Faria=Gender=M=UName=dhavalhirdhav=Pass=faria=EMail=dhavalhirdhav@yahoo.com=City=Mumbai=State=Maharashtra=Country=INDIA=HQ=HOW ARE YOU?=HA=I AM FINE.=Autho=16456=Ver=1.3")

strAnswer = Inet1.OpenURL("http://www.hirdhav.com/HDD/GetCode/Put.asp?FirstName=" + .txtFName.Text + "=LastName=" + .txtLName.Text + "=Gender=" + .txtGender.Text + "=UName=" + .txtUName.Text + "=Pass=" + .txtPassword.Text + "=EMail=" + .txtEMail.Text + "=City=" & _
              .txtCity.Text + "=State=" + .txtState.Text + "=Country=" + .txtCountry.Text + "=HQ=" + .txtQuestion.Text + "=HA=" + .txtAnswer.Text + "=Autho=" + FiNaLcOdE + "=Ver=" + "1.3")

End With

DoEvents

If strAnswer = "NO" Then
    frmGetAutho2.Show
    Unload Me
    HDDMsgBox "Sorry, Username is already taken, Please choose another Username."
    Exit Function
Else
    DoEvents
    lblStatus.Caption = "All the procudure are completed.. Please Click on Show Authentication CODE to get Get your Autho CODE."
    lblShowAutho.Enabled = True
    lblShowAuthoSupport.Enabled = True
End If

End Function

Private Sub lblDoAgainSupport_Click()
lblStatus.Caption = "Please wait.. Checking Internet Connection..."
Call DoAll
End Sub

Private Sub lblDoAgainSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDoAgain.ForeColor = RGB(145, 155, 100)
shapeDoAgain.BackColor = vbBlack
End Sub

Private Sub lblDoAgainSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeDoAgain.BackColor = RGB(145, 155, 100)
lblDoAgain.ForeColor = vbBlack
End Sub

Private Function GenCODE()
'Generate Authentication CODE...
lblStatus.Caption = "Generating Authentication CODE..."
HaLfCoDe = ""
FiNaLcOdE = ""
LenOfText = Len(txtName.Text)
For j = 0 To LenOfText - 1
    txtName.SelStart = j
    txtName.SelLength = 1
    HaLfCoDe = HaLfCoDe & Asc(txtName.SelText)
Next j
txtAutho.Text = HaLfCoDe

txtAutho.SelStart = 0
txtAutho.SelLength = 1
FiNaLcOdE = txtAutho.SelText
LenOfASC = Len(txtAutho.Text)
For k = 0 To LenOfASC - 1
    txtAutho.SelStart = k
    txtAutho.SelLength = 1
    MakeTotal = txtAutho.SelText
    k = k + 1
    txtAutho.SelStart = k
    txtAutho.SelLength = 1
    FiNaLcOdE = FiNaLcOdE & Val(MakeTotal) + Val(txtAutho.SelText)
    k = k - 1
Next k
txtAutho.Text = FiNaLcOdE
End Function

Private Sub lblShowAuthoSupport_Click()
frmGetAutho4.Show
Me.Hide
End Sub

Private Sub lblShowAuthoSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblShowAutho.ForeColor = RGB(145, 155, 100)
shapeShowAutho.BackColor = vbBlack
End Sub

Private Sub lblShowAuthoSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeShowAutho.BackColor = RGB(145, 155, 100)
lblShowAutho.ForeColor = vbBlack
End Sub

Private Sub Timer1_Timer()
Call DoAll
Timer1.Enabled = False
End Sub
