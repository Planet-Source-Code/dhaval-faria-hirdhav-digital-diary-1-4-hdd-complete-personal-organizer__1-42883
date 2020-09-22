VERSION 5.00
Begin VB.Form frmDetailSch 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Full Scheduler Detail"
   ClientHeight    =   3855
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
   Icon            =   "frmDetailSch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2520
      TabIndex        =   13
      Top             =   3360
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1800
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblDesc 
      Height          =   855
      Left            =   1800
      TabIndex        =   12
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   225
      Left            =   600
      TabIndex        =   11
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label lblAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:10 AM"
      Height          =   225
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Time:"
      Height          =   225
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Label lblTT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:10 AM"
      Height          =   225
      Left            =   3360
      TabIndex        =   8
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   225
      Left            =   2880
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblTF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:10 AM"
      Height          =   225
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time From:"
      Height          =   225
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14/Mar/1985"
      Height          =   225
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   225
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   435
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   5160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anivarsary Reminder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5025
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
      Caption         =   "Hirdhav Digital Diary  -  Full Scheduler Detail"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3660
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5280
   End
End
Attribute VB_Name = "frmDetailSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername
Public CurrentMonth
Public CompInfo

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblDesc.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)

If Len(frmCalender.txtDate.Text) = 11 Then
    frmCalender.txtDate.SelStart = 3
    frmCalender.txtDate.SelLength = 3
Else
    frmCalender.txtDate.SelStart = 2
    frmCalender.txtDate.SelLength = 3
End If

CurrentMonth = frmCalender.txtDate.SelText

Dim db As Database
Dim ReS As Recordset

On Error GoTo ErrHan

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)


Do
CompInfo = ReS("TF") + ReS("AP1") + "  " + ReS("Description")
    If CompInfo = frmCalender.lstSchList.Text Then
        lblType.Caption = ReS("SchType")
        lblDate.Caption = ReS("Date")
        lblTF.Caption = ReS("TF") & " " & ReS("AP1")
        lblTT.Caption = ReS("TT") & " " & ReS("AP2")
        lblAT.Caption = ReS("AT") & " " & ReS("AP3")
        lblDesc.Caption = ReS("Description")
        Exit Sub
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
        frmCalender.Show
        'Unload Me
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
