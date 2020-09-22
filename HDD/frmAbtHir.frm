VERSION 5.00
Begin VB.Form frmAbtHir 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - About Hirdhav"
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
   Icon            =   "frmAbtHir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   240
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   5175
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1680
      TabIndex        =   20
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2520
      TabIndex        =   19
      Top             =   4320
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1680
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblITax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Virtual Desktop - v1.0(Shareware)"
      Height          =   225
      Left            =   600
      TabIndex        =   18
      Top             =   3240
      Width           =   3555
   End
   Begin VB.Label lblVoice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav HD FTP Client - v1.0 (Adware)"
      Height          =   225
      Left            =   600
      TabIndex        =   17
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Feature Developments:"
      Height          =   225
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   1950
   End
   Begin VB.Label lblAc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Joke - v1.0 (Freeware)"
      Height          =   225
      Left            =   600
      TabIndex        =   15
      Top             =   3000
      Width           =   2580
   End
   Begin VB.Label lblHDDF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary - v1.5 (Freeware)"
      Height          =   225
      Left            =   600
      TabIndex        =   14
      Top             =   2760
      Width           =   3165
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Developments:"
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1950
   End
   Begin VB.Label lblHDD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary - v1.4 (Freeware)"
      Height          =   225
      Left            =   600
      TabIndex        =   12
      Top             =   2160
      Width           =   3165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Past Developments:"
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblWebSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.hirdhav.com"
      Height          =   225
      Left            =   3000
      TabIndex        =   10
      Top             =   1560
      Width           =   2010
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site:"
      Height          =   225
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label lblFouName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dhaval Faria"
      Height          =   225
      Left            =   3600
      TabIndex        =   8
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Founder:"
      Height          =   225
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblEstaDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "23 Feb. 2002"
      Height          =   225
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Established At:"
      Height          =   225
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   1260
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   240
      Y2              =   4800
   End
   Begin VB.Label lblHirdhavName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav"
      Height          =   225
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   225
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   540
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      Height          =   225
      Left            =   1800
      TabIndex        =   2
      Top             =   780
      Width           =   135
   End
   Begin VB.Label lblHirdhav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      TabIndex        =   1
      Top             =   750
      Width           =   1740
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  About Hirdhav"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3120
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
Attribute VB_Name = "frmAbtHir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public FontColorR
Public FontColorG
Public FontColorB

Private Sub Form_Load()
FontColorR = 145
FontColorG = 155
FontColorB = 100
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblReg.ForeColor = RGB(145, 155, 100)
lblHirdhav.ForeColor = RGB(145, 155, 100)
lblHirdhavName.ForeColor = RGB(145, 155, 100)
lblEstaDate.ForeColor = RGB(145, 155, 100)
lblFouName.ForeColor = RGB(145, 155, 100)
lblWebSite.ForeColor = RGB(145, 155, 100)
lblHDD.ForeColor = RGB(145, 155, 100)
lblAc.ForeColor = RGB(145, 155, 100)
lblITax.ForeColor = RGB(145, 155, 100)
lblVoice.ForeColor = RGB(145, 155, 100)
lblHDDF.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblOkSupport_Click()
frmAbout.Show
Unload Me
End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeOk.BackColor = vbBlack
lblOk.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = vbBlack
shapeOk.BackColor = RGB(145, 155, 100)
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
lblHirdhav.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblReg.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblHirdhavName.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblEstaDate.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblFouName.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblWebSite.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblAc.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblHDD.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblHDDF.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblVoice.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblITax.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
End Sub
