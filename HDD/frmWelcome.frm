VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
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
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   360
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.4 Released On: 31 March 2002.      "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Released On: 20 May 2002."
      Height          =   225
      Left            =   1200
      TabIndex        =   12
      Top             =   2880
      Width           =   2565
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.3 Released On: 31 March 2002.       "
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.2 Released On: 26 January 2002.    "
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1 Released On: 15 December 2001."
      Height          =   225
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   4680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hddcontact@hirdhav.com"
      Height          =   225
      Left            =   1080
      TabIndex        =   8
      Top             =   3840
      Width           =   2190
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Us:"
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0 Released On: 01 November 2001."
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait.... Hirdhav Digital Diary is Loading..."
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   3960
   End
   Begin VB.Image imgMain 
      Height          =   1200
      Left            =   240
      Picture         =   "frmWelcome.frx":030A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblHirdhav 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1905
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label lblDD 
      AutoSize        =   -1  'True
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
      Height          =   480
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   2430
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  v 1.4"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   15
      Width           =   2325
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4680
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4695
      X2              =   4680
      Y1              =   240
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   5400
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   4695
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Timer1_Timer()
frmLogin.Show
Unload Me
End Sub
