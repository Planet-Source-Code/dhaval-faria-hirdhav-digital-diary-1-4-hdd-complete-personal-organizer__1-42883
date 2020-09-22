VERSION 5.00
Begin VB.Form frmGetAutho1 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Get Autho (Step 1)"
   ClientHeight    =   3495
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
   Icon            =   "frmGetAutho1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblNextSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblNext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next >"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Shape shapeNext 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3240
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblINetStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on Next button to check the status of Internet Connection."
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Status:"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetAutho1.frx":030A
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5300
      X2              =   5300
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Get Autho (Step 1)"
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   15
      Width           =   3465
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
Attribute VB_Name = "frmGetAutho1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeNext.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
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

lblINetStatus.Caption = "Please wait... Checking Internet Status..."

Dim flags As Long
Dim result As Boolean

result = InternetGetConnectedState(flags, 0)

If result Then
    lblINetStatus.Caption = "You are connected to the Internet... Please wait..."
    frmGetAutho2.Show
    Unload Me
    Exit Sub
Else
    lblINetStatus.Caption = "Not Connected to the Internet. Please make sure that you are connected to the Internet and then click on Next Button."
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
