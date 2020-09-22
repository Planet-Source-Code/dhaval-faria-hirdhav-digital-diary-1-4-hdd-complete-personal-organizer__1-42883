VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Credits"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
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
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCredits 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lblMISupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblMI 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "More Info"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Shape shapeMI 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3480
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "If you want to know about the individual person information listed here, then click on More Info button next to Ok button."
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   915
      TabIndex        =   1
      Top             =   4095
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Credits"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   240
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   4560
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5280
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeOk.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
txtCredits.BackColor = RGB(145, 155, 100)
shapeMI.BackColor = RGB(145, 155, 100)
InsertCreditText
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblMISupport_Click()
'HDDMsgBox "This feature is not there, We will add it soon..."
frmMI.Show
Me.Hide
End Sub

Private Sub lblMISupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMI.ForeColor = RGB(145, 155, 100)
shapeMI.BackColor = vbBlack
End Sub

Private Sub lblMISupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeMI.BackColor = RGB(145, 155, 100)
lblMI.ForeColor = vbBlack
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
lblOk.ForeColor = vbBlack
shapeOk.BackColor = RGB(145, 155, 100)
End Sub

Private Function InsertCreditText()
txtCredits.Text = "CREDITS" & vbCrLf & _
                  "------------" & vbCrLf & vbCrLf & _
                  "Here is the list of those people who helps in developing this wonderfull diary." & vbCrLf & vbCrLf & _
                  "Original Concept By:" & vbCrLf & _
                  "----------------------------" & vbCrLf & _
                  "       Dhaval Faria" & vbCrLf & _
                  "       Mukesh Parikh" & vbCrLf & vbCrLf & _
                  "Developed & Designed By:" & vbCrLf & _
                  "-------------------------------------" & vbCrLf & _
                  "       Dhaval Faria" & vbCrLf & vbCrLf & _
                  "Concept & Developed of Contacts By:" & vbCrLf & _
                  "----------------------------------------------------" & vbCrLf & _
                  "       Dhaval Faria" & vbCrLf & vbCrLf & _
                  "Concept of Calender By:" & vbCrLf & _
                  "----------------------------------" & vbCrLf & _
                  "       Dhaval Faria" & vbCrLf & vbCrLf & _
                  "Development of Calender By:" & vbCrLf & _
                  "----------------------------------------" & vbCrLf & _
                  "       Dhaval Faria" & vbCrLf & vbCrLf & _
                  "Development and Concept of Whole Scheduler By:" & vbCrLf & _
                  "----------------------------------------------------------------------" & vbCrLf & _
                  "       Dhaval Faria" & vbCrLf & _
                  "       John Couture" & vbCrLf & vbCrLf & _
                  "Development and Concept of Whole Memo By:" & vbCrLf
MoreCredits1
End Function

Private Function MoreCredits1()
txtCredits.Text = txtCredits.Text & "-----------------------------------------------------------------" & vbCrLf & _
                                    "       Dhaval Faria" & vbCrLf & vbCrLf & _
                                    "Concept & Development of Account Editor By:" & vbCrLf & _
                                    "----------------------------------------------------------------" & vbCrLf & _
                                    "       Dhaval Faria" & vbCrLf & vbCrLf & _
                                    "Concept of Anivarsary & Birthday Reminder By:" & vbCrLf & _
                                    "------------------------------------------------------------------" & vbCrLf & _
                                    "       Vighnesh Prabhu" & vbCrLf & _
                                    "       Dhaval Faria" & vbCrLf & vbCrLf & _
                                    "Development of Anivarsary & Birthday Reminder By:" & vbCrLf & _
                                    "-------------------------------------------------------------------------" & vbCrLf & _
                                    "       Dhaval Faria" & vbCrLf & vbCrLf & _
                                    "Special Thanks To:" & vbCrLf & _
                                    "--------------------------" & vbCrLf & _
                                    "Ramnik Faria" & vbCrLf & _
                                    "Manjula Faria" & vbCrLf & _
                                    "Hiren Faria" & vbCrLf & _
                                    "Mukesh Parikh" & vbCrLf & _
                                    "John Couture" & vbCrLf & _
                                    "Vighnesh Prabhu" & vbCrLf & _
                                    "Kaustubh Gujar" & vbCrLf & _
                                    ""
End Function
