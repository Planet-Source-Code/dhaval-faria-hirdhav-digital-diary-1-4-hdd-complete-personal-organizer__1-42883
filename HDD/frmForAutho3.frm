VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmForAutho3 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Forgot Autho (Step 3)"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
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
   Icon            =   "frmForAutho3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4200
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   2760
   End
   Begin VB.Label lblFinishSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblFinish 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Finish"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Shape shapeFinish 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1680
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Status is here."
      Height          =   705
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   4320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmForAutho3.frx":030A
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5180
      X2              =   5180
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Forgot Autho (Step 3)"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5175
   End
End
Attribute VB_Name = "frmForAutho3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeFinish.BackColor = RGB(145, 155, 100)
lblFinish.Enabled = False
lblFinishSupport.Enabled = False
lblStatus.Caption = "Please wait.. Checking Internet Connection..."
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblFinishSupport_Click()
If lblFinish.Caption = "Try Again" Then
    HDDMsgBox "Sorry, Some of your information is incorrect."
    Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Dim flags As Long
Dim result As Boolean

result = InternetGetConnectedState(flags, 0)

If result Then
    lblStatus.Caption = "Sending data to the web."
    With frmForAutho2
        Dim strAnswer
        lblStatus.Caption = "Checking your Username and Password..."
        strAnswer = Inet1.OpenURL("http://www.hirdhav.com/HDD/GetCode/for.asp?Username=" + .txtUserName.Text + "=Password=" + .txtPassword.Text + "=EMail=" + .txtEMail.Text + "=HQ=" + .txtQuestion.Text + "=HA=" + .txtAnswer.Text)
        If strAnswer = "Sorry" Then
            HDDMsgBox "Some given information is incorrect, or Username is not registered."
            frmForAutho2.Show
            Unload Me
            Exit Sub
        Else
            lblStatus.Caption = "Getting your Authentication CODE...."
            strAnswer = Split(strAnswer, ":")
            frmAutho.Show
            With frmAutho
                .txtName.Text = strAnswer(1)
                .txtAutho.Text = strAnswer(3)
            End With
            Unload frmForAutho3
        End If
    End With
Else
    HDDMsgBox "Sorry, Not connected to the internet, Please try again."
    lblFinish.Caption = "Try Again"
    Exit Sub
End If
End Sub
