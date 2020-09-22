VERSION 5.00
Begin VB.Form frmMI 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - More Info"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
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
   Icon            =   "frmMI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMPAni 
      Interval        =   1
      Left            =   120
      Top             =   360
   End
   Begin VB.Timer tmrMP 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Top             =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      Height          =   225
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
      Height          =   225
      Left            =   2565
      TabIndex        =   4
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dhaval Faria"
      Height          =   225
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   225
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   4560
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Credits (More Info)"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3510
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4800
   End
End
Attribute VB_Name = "frmMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SupportNaMe
Public TotalLength
Public CurrentChars

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblName.Caption = ""
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub tmrMP_Timer()
If lblName.Caption = SupportNaMe Then
    tmrMP.Enabled = False
    Exit Sub
End If
t = Left(SupportNaMe, CurrentChars)
lblName.Caption = t
CurrentChars = CurrentChars + 1
If CurrentChars > TotalLength Then
    CurrentChars = 0
End If
End Sub

Private Sub tmrMPAni_Timer()
SupportNaMe = "Dhaval Faria"
TotalLength = Len(SupportNaMe)
CurrentChars = 0
tmrMP.Enabled = True
tmrMPAni.Enabled = False
End Sub
