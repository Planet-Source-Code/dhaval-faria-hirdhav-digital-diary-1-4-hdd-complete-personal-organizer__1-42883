VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - About Us"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3000
      TabIndex        =   19
      Top             =   360
      Width           =   1905
   End
   Begin VB.Label lblNextVersionSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2160
      TabIndex        =   18
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblNextVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Version"
      Height          =   225
      Left            =   2340
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblContactUsSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblContactUs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Us"
      Height          =   225
      Left            =   4080
      TabIndex        =   15
      Top             =   3960
      Width           =   930
   End
   Begin VB.Label lblHDDHistorySupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblHDDHistory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary History"
      Height          =   465
      Left            =   480
      TabIndex        =   13
      Top             =   3855
      Width           =   1500
   End
   Begin VB.Shape shapeHDDHistory 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   480
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Shape shapeNextVersion 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2160
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      Height          =   225
      Left            =   480
      TabIndex        =   11
      Top             =   4635
      Width           =   4875
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   480
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Shape shapeContactUs 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3840
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   0
      X2              =   5760
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   5760
      X2              =   5760
      Y1              =   240
      Y2              =   5040
   End
   Begin VB.Label lblAbtCompanySupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblAbtCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Hirdhav"
      Height          =   495
      Left            =   4185
      TabIndex        =   9
      Top             =   3255
      Width           =   780
   End
   Begin VB.Shape shapeAbtCompany 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3840
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblAbtDeveloperSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblAbtDeveloper 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Developer"
      Height          =   495
      Left            =   2475
      TabIndex        =   7
      Top             =   3255
      Width           =   855
   End
   Begin VB.Shape shapeAbtDeveloper 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2160
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblCreditsSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblCredits 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      Height          =   225
      Left            =   885
      TabIndex        =   5
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape shapeCredits 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   480
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   120
      X2              =   5640
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   120
      X2              =   5640
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblLicense 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To explore more about Hirdhav or Hirdhav Digital Diary click on the below given buttons."
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2595
      Width           =   4695
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label lblSoftName 
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Diary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   5040
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  About Us"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2700
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5760
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public UserName

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeCredits.BackColor = RGB(145, 155, 100)
shapeAbtDeveloper.BackColor = RGB(145, 155, 100)
shapeAbtCompany.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
shapeContactUs.BackColor = RGB(145, 155, 100)
shapeHDDHistory.BackColor = RGB(145, 155, 100)
shapeNextVersion.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
lblLicense.BackColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\HDD.dat")
Set ReS = db.OpenRecordset("Users")

FullName = ReS("FirstName") & " " & ReS("LastName")

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

lblLicense.Caption = "This product is licensed to" & vbCrLf & FullName

End Sub

Private Sub lblAbtCompanySupport_Click()
frmAbtHir.Show
Unload Me
End Sub

Private Sub lblAbtCompanySupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbtCompany.ForeColor = RGB(145, 155, 100)
shapeAbtCompany.BackColor = vbBlack
End Sub

Private Sub lblAbtCompanySupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbtCompany.ForeColor = vbBlack
shapeAbtCompany.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblAbtDeveloperSupport_Click()
frmAbtDev.Show
Unload Me
End Sub

Private Sub lblAbtDeveloperSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbtDeveloper.ForeColor = RGB(145, 155, 100)
shapeAbtDeveloper.BackColor = vbBlack
End Sub

Private Sub lblAbtDeveloperSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbtDeveloper.ForeColor = vbBlack
shapeAbtDeveloper.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblContactUsSupport_Click()
frmContactUs.Show
Unload Me
End Sub

Private Sub lblContactUsSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblContactUs.ForeColor = RGB(145, 155, 100)
shapeContactUs.BackColor = vbBlack
End Sub

Private Sub lblContactUsSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblContactUs.ForeColor = vbBlack
shapeContactUs.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCreditsSupport_Click()
frmCredits.Show
Me.Hide
End Sub

Private Sub lblCreditsSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCredits.ForeColor = RGB(145, 155, 100)
shapeCredits.BackColor = vbBlack
End Sub

Private Sub lblCreditsSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCredits.BackColor = RGB(145, 155, 100)
lblCredits.ForeColor = vbBlack
End Sub

Private Sub lblHDDHistorySupport_Click()
frmHDDHistory.Show
Me.Hide
End Sub

Private Sub lblHDDHistorySupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHDDHistory.ForeColor = RGB(145, 155, 100)
shapeHDDHistory.BackColor = vbBlack
End Sub

Private Sub lblHDDHistorySupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHDDHistory.ForeColor = vbBlack
shapeHDDHistory.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblNextVersionSupport_Click()
frmNextVer.Show
Me.Hide
End Sub

Private Sub lblNextVersionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNextVersion.ForeColor = RGB(145, 155, 100)
shapeNextVersion.BackColor = vbBlack
End Sub

Private Sub lblNextVersionSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNextVersion.ForeColor = vbBlack
shapeNextVersion.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblOkSupport_Click()
frmMain.Show
Unload Me
End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = vbBlack
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = vbBlack
shapeOk.BackColor = RGB(145, 155, 100)
End Sub
