VERSION 5.00
Begin VB.Form frmContactUs 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Contact Us"
   ClientHeight    =   5175
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
   Icon            =   "frmContactUs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5055
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
      TabIndex        =   19
      Top             =   0
      Width           =   5055
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2400
      TabIndex        =   17
      Top             =   4680
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1680
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblAdd6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INDIA."
      Height          =   225
      Left            =   1920
      TabIndex        =   16
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lblAdd5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maharashtra"
      Height          =   225
      Left            =   1920
      TabIndex        =   15
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblAdd4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mumbai : 400 062."
      Height          =   225
      Left            =   1920
      TabIndex        =   14
      Top             =   3720
      Width           =   1530
   End
   Begin VB.Label lblAdd3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goregaon (West),"
      Height          =   225
      Left            =   1920
      TabIndex        =   13
      Top             =   3480
      Width           =   1485
   End
   Begin VB.Label lblAdd2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M.G. Road,"
      Height          =   225
      Left            =   1920
      TabIndex        =   12
      Top             =   3240
      Width           =   885
   End
   Begin VB.Label lblAdd1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "95/6, Laxmi Niwas,"
      Height          =   225
      Left            =   1920
      TabIndex        =   11
      Top             =   3000
      Width           =   1590
   End
   Begin VB.Label lblHirdhavAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav"
      Height          =   225
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   225
      Left            =   1080
      TabIndex        =   9
      Top             =   2760
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Contact Us By Mail"
      Height          =   225
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "contactus@hirdhav.com"
      Height          =   225
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Contact Us By E-Mail"
      Height          =   225
      Left            =   2520
      TabIndex        =   6
      Top             =   1680
      Width           =   1845
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   5160
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "®"
      Height          =   225
      Left            =   1920
      TabIndex        =   5
      Top             =   480
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
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1740
   End
   Begin VB.Label lblNo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+(91) (022) 875 51 05"
      Height          =   225
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Width           =   1785
   End
   Begin VB.Label lblNo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+(91) (022) 873 56 89"
      Height          =   225
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Contact Us By Phone"
      Height          =   225
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Contact Us"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2865
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   5160
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5040
   End
End
Attribute VB_Name = "frmContactUs"
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
lblHirdhav.ForeColor = RGB(145, 155, 100)
lblReg.ForeColor = RGB(145, 155, 100)
lblNo1.ForeColor = RGB(145, 155, 100)
lblNo2.ForeColor = RGB(145, 155, 100)
lblEMail.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
lblHirdhavAdd.ForeColor = RGB(145, 155, 100)
lblAdd1.ForeColor = RGB(145, 155, 100)
lblAdd2.ForeColor = RGB(145, 155, 100)
lblAdd3.ForeColor = RGB(145, 155, 100)
lblAdd4.ForeColor = RGB(145, 155, 100)
lblAdd5.ForeColor = RGB(145, 155, 100)
lblAdd6.ForeColor = RGB(145, 155, 100)
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
lblHirdhav.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblReg.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblNo1.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblNo2.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblEMail.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblHirdhavAdd.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblAdd1.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblAdd2.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblAdd3.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblAdd4.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblAdd5.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
lblAdd6.ForeColor = RGB(FontColorR, FontColorG, FontColorB)
End Sub
