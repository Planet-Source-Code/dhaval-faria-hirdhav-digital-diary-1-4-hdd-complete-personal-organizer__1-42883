VERSION 5.00
Begin VB.Form frmNextVer 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary  -  Next Version"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
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
   Icon            =   "frmNextVer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNextVer 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2640
      TabIndex        =   3
      Top             =   3720
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1920
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5400
      X2              =   5400
      Y1              =   240
      Y2              =   4320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Here is information on what will be there in our Next Version of Hirdhav Digital Diary (HDD)."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5520
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   4320
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Next Version"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3030
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5400
   End
End
Attribute VB_Name = "frmNextVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
Me.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
txtNextVer.BackColor = RGB(145, 155, 100)
InsertNextVersionTEXT
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

Private Function InsertNextVersionTEXT()
txtNextVer.Text = "Next Version:" & vbCrLf & _
                  "------------------" & vbCrLf & vbCrLf & _
                  "    Our Next Version Of Hirdhav Digital Diary (HDD) will contains following things:" & vbCrLf & vbCrLf & _
                  "1)  Appoinment in Scheduler" & vbCrLf & vbCrLf & _
                  "More features will be also added.. This is just a first version so we can's decide what we will add in next version, Wel will add lots of things in our Next Version." & vbCrLf & vbCrLf & _
                  "This above new features are going to be add in the next version of Hirdhav Digital Diary (HDD)." & _
                  " And that version is version 1.4"
End Function
