VERSION 5.00
Begin VB.Form frmForAutho1 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Forgot Autho (Step 1)"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
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
   Icon            =   "frmForAutho1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblNextSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblNext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next >"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape shapeNext 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3240
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Please click on Next button to check the Internet Status and to go further to Step 2."
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Status:"
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmForAutho1.frx":030A
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5055
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5300
      X2              =   5300
      Y1              =   240
      Y2              =   4440
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Dairy  -  Forgot Autho (Step 1)"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   20
      Width           =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   4440
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5295
   End
End
Attribute VB_Name = "frmForAutho1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeNext.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
frmAuthoHelp.Show
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

Private Sub lblNextSupport_Click()

lblStatus.Caption = "Please wait... Checking Internet Status..."

Dim flags As Long
Dim result As Boolean

result = InternetGetConnectedState(flags, 0)

If result Then
    lblStatus.Caption = "You are connected to the Internet... Please wait..."
    frmForAutho2.Show
    Unload Me
    Exit Sub
Else
    lblStatus.Caption = "Not Connected to the Internet. Please make sure that you are connected to the Internet and then click on Next Button."
    Exit Sub
End If
End Sub

Private Sub lblNextSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = RGB(145, 155, 100)
shapeNext.BackColor = vbBlack
End Sub

Private Sub lblNextSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeNext.BackColor = RGB(145, 155, 100)
lblNext.ForeColor = vbBlack
End Sub
