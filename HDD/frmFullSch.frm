VERSION 5.00
Begin VB.Form frmFullSch 
   BorderStyle     =   0  'None
   Caption         =   "Hrdhav Digital Diary - Full Detail"
   ClientHeight    =   3975
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
   Icon            =   "frmFullSch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDetail 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1800
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5180
      X2              =   5180
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14/Mar/1985"
      Height          =   225
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   225
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   435
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
      Caption         =   "Hirdhav Digital Diary  -  Full Detail"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2745
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
Attribute VB_Name = "frmFullSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername
Public NeedDate

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
lstDetail.BackColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption
lblDate.Caption = SchDate & "/" & SchMonth & "/" & SchYear

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(SchMonth)

On Error GoTo ErrHan

Do
    NeedDate = ReS("Date")
    If NeedDate = lblDate.Caption Then
        lstDetail.AddItem ReS("TF") + ReS("AP1") + "  " + ReS("Description")
        ReS.MoveNext
    Else
        ReS.MoveNext
    End If
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

ErrHan:
    If Err.Number = 3021 Then
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        Exit Sub
    End If
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblOkSupport_Click()
frmCalender.Show
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
