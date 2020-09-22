VERSION 5.00
Begin VB.Form frmAbtDev 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - About Developers"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
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
   Icon            =   "frmAbtDev.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   840
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   4200
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dhavalf@hirdhav.com"
      Height          =   225
      Left            =   3240
      TabIndex        =   13
      Top             =   1920
      Width           =   1845
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2475
      TabIndex        =   9
      Top             =   4320
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1800
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label lblDevLan 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbtDev.frx":030A
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programming Language Known:"
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   2730
   End
   Begin VB.Label lblDevBDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14 March 1985"
      Height          =   225
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      Height          =   225
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label lblDevAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      Height          =   225
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      Height          =   225
      Left            =   3030
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shapeTop 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   120
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape shapeDown 
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   120
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Shape shapeRight 
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   2160
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shapeLeft 
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   120
      Top             =   480
      Width           =   135
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   240
      Y2              =   4800
   End
   Begin VB.Label lblDevName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   225
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   540
   End
   Begin VB.Image DevPic 
      Height          =   2295
      Left            =   240
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  About Developer"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   4800
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5160
   End
End
Attribute VB_Name = "frmAbtDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a As String
Dim t As String
Dim b As Integer
Dim I As Integer

Public FontColorR
Public FontColorG
Public FontColorB

Private Sub DevPic_Click()
If vbKeyRButton = True Then
    MsgBox "HI"
Else
    MsgBox "Bye"
End If
End Sub

Private Sub Form_Load()
FontColorR = 145
FontColorG = 155
FontColorB = 100
lblDevAge.Caption = "17"
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeOk.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
shapeLeft.BackColor = vbBlack
shapeRight.BackColor = vbBlack
shapeDown.BackColor = vbBlack
shapeTop.BackColor = vbblalck
DevPic.Picture = LoadPicture(App.Path + "\HDDP.dha")
'lblDevName.ForeColor = RGB(145, 155, 100)
lblDevAge.ForeColor = RGB(145, 155, 100)
lblDevBDate.ForeColor = RGB(145, 155, 100)
lblDevLan.ForeColor = RGB(145, 155, 100)
lblEMail.ForeColor = RGB(145, 155, 100)
DevPic.Height = 0
Animate
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblOkSupport_Click()
frmAbout.Show
Unload Me
End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = vbBlack
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeOk.BackColor = RGB(145, 155, 100)
lblOk.ForeColor = vbBlack
End Sub

Private Sub Timer1_Timer()
If FontColorR <> 0 Then
    FontColorR = FontColorR - 5
End If
If FontColorG <> 0 Then
    FontColorG = FontColorG - 5
End If
If FontColorB <> 0 Then
    FontColorB = FontColorB - 5
End If
'lblDevName.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblDevAge.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblDevBDate.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblDevLan.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblEMail.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
End Sub

Private Sub Timer2_Timer()
If DevPic.Height = 2295 Then
    Timer2.Enabled = False
End If
DevPic.Height = DevPic.Height + 10
End Sub

Sub Animate()
a = "Dhaval Faria"
I = Len(a)
b = 0
End Sub

Private Sub Timer3_Timer()
If lblDevName.Caption = a Then
    Timer3.Enabled = False
    Exit Sub
End If
t = Left(a, b)
lblDevName.Caption = t
b = b + 1
If b > I Then
    b = 0
End If
End Sub
