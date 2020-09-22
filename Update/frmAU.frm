VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmAU 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Auto Update"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
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
   Icon            =   "frmAU.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet net 
      Left            =   120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame FrameMessage 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblUpdateSupport 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   3720
         TabIndex        =   33
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblUpdate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update"
         Height          =   225
         Left            =   4160
         TabIndex        =   32
         Top             =   3960
         Width           =   600
      End
      Begin VB.Label lblCancel2Support 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1920
         TabIndex        =   31
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblCancel2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   225
         Left            =   1920
         TabIndex        =   30
         Top             =   3960
         Width           =   1545
      End
      Begin VB.Shape shapeUpdate 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   3720
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Shape shapeCancel2 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   1920
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "If you want to update your Hirdhav Digital Diary Software with this new version click on Update otherwise click on Cancel to Exit."
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   5655
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         Height          =   1575
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label lblNewVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Version is"
         Height          =   225
         Left            =   600
         TabIndex        =   27
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update Information:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblCancel1Support 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1920
         TabIndex        =   24
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblNext2Support 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   3720
         TabIndex        =   23
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblNext2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next >"
         Height          =   225
         Left            =   4200
         TabIndex        =   22
         Top             =   3960
         Width           =   540
      End
      Begin VB.Shape shapeNext2 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   3720
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblCancel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   225
         Left            =   2400
         TabIndex        =   21
         Top             =   3960
         Width           =   585
      End
      Begin VB.Shape shapeCancel1 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   1920
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Idle........."
         Height          =   615
         Left            =   480
         TabIndex        =   20
         Top             =   2520
         Width           =   4575
      End
      Begin VB.Label lblStartSupport 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   480
         TabIndex        =   19
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   225
         Left            =   1080
         TabIndex        =   18
         Top             =   1470
         Width           =   420
      End
      Begin VB.Shape shapeStart 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   480
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAU.frx":030A
         Height          =   735
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   540
      End
   End
   Begin VB.Frame FrameWelcome 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5775
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Check the Internet Connection."
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   2715
      End
      Begin VB.Label lblNext1Support 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   3720
         TabIndex        =   13
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblNext1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next >"
         Height          =   225
         Left            =   4240
         TabIndex        =   12
         Top             =   3960
         Width           =   540
      End
      Begin VB.Shape shapeNext1 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   3720
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblCancelSupport 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   225
         Left            =   2400
         TabIndex        =   10
         Top             =   3960
         Width           =   585
      End
      Begin VB.Shape shapeCancel 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   1920
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAU.frx":03A1
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data to Retrive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This wizard will get the following information:"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Check for the update."
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1920
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- A message to explain to user about new version."
         Height          =   225
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   4230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- Whether user wants to update the software or not?"
         Height          =   225
         Left            =   480
         TabIndex        =   4
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "If you want to update HDD click on Next button otherwise click on Cancel."
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   3240
         Width           =   5295
      End
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6000
      X2              =   6000
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6000
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   4920
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Auto Update"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2970
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   6000
   End
End
Attribute VB_Name = "frmAU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
FrameWelcome.BackColor = RGB(145, 155, 100)
Frame2.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
shapeStart.BackColor = RGB(145, 155, 100)
shapeCancel2.BackColor = RGB(145, 155, 100)
shapeUpdate.BackColor = RGB(145, 155, 100)
shapeNext2.BackColor = RGB(145, 155, 100)
shapeCancel1.BackColor = RGB(145, 155, 100)
shapeNext1.BackColor = RGB(145, 155, 100)
FrameMessage.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancel1Support_Click()
End
End Sub

Private Sub lblCancel1Support_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel1.ForeColor = RGB(145, 155, 100)
shapeCancel1.BackColor = vbBlack
End Sub

Private Sub lblCancel1Support_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel1.BackColor = RGB(145, 155, 100)
lblCancel1.ForeColor = vbBlack
End Sub

Private Sub lblCancel2Support_Click()
If lblCancel2.Caption = "Exit" Then
    Shell App.Path + nEXE
    End
Else
    End
End If
End Sub

Private Sub lblCancel2Support_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel2.ForeColor = RGB(145, 155, 100)
shapeCancel2.BackColor = vbBlack
End Sub

Private Sub lblCancel2Support_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel2.BackColor = RGB(145, 155, 100)
lblCancel2.ForeColor = vbBlack
End Sub

Private Sub lblCancelSupport_Click()
End
End Sub

Private Sub lblCancelSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = vbBlack
End Sub

Private Sub lblCancelSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel.BackColor = RGB(145, 155, 100)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblNext1Support_Click()
FrameWelcome.Visible = False
Frame2.Visible = True
lblNext2.Enabled = False
lblNext2Support.Enabled = False
End Sub

Private Sub lblNext1Support_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext1.ForeColor = RGB(145, 155, 100)
shapeNext1.BackColor = vbBlack
End Sub

Private Sub lblNext1Support_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeNext1.BackColor = RGB(145, 155, 100)
lblNext1.ForeColor = vbBlack
End Sub

Private Sub lblNext2Support_Click()
Frame2.Visible = False
FrameMessage.Visible = True
lblNewVer.Caption = "New Version of Hirdhav Digital Diary is " & nVer
lblMessage.Caption = nMsg
End Sub

Private Sub lblNext2Support_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext2.ForeColor = RGB(145, 155, 100)
shapeNext2.BackColor = vbBlack
End Sub

Private Sub lblNext2Support_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext2.ForeColor = vbBlack
shapeNext2.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblStartSupport_Click()
lblStart.Enabled = False
lblStartSupport.Enabled = False
lblStatus.Caption = "Please wait... Checking Internet Connection..."
CheckINET
End Sub

Private Sub lblStartSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStart.ForeColor = RGB(145, 155, 100)
shapeStart.BackColor = vbBlack
End Sub

Private Sub lblStartSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStart.ForeColor = vbBlack
shapeStart.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblUpdateSupport_Click()
DownloadFile
End Sub

Private Sub lblUpdateSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUpdate.ForeColor = RGB(145, 155, 100)
shapeUpdate.BackColor = vbBlack
End Sub

Private Sub lblUpdateSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeUpdate.BackColor = RGB(145, 155, 100)
lblUpdate.ForeColor = vbBlack
End Sub
