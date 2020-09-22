VERSION 5.00
Begin VB.Form frmShowMe 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Reminder"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
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
   Icon            =   "frmShow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2760
      TabIndex        =   10
      Top             =   3000
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2040
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   5400
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblDesc 
      Height          =   735
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desc:"
      Height          =   225
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:10 AM"
      Height          =   225
      Left            =   3480
      TabIndex        =   7
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   225
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:10 AM"
      Height          =   225
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   225
      Left            =   1120
      TabIndex        =   4
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:10 AM"
      Height          =   225
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Time:"
      Height          =   225
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5420
      X2              =   5420
      Y1              =   240
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   5280
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblSchType 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5250
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Reminder"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   20
      X2              =   20
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5415
   End
End
Attribute VB_Name = "frmShowMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername
Public CurrentMonth
Public CurrentDate
Public CurrentYear
Public NeedDate
Public NeedTime

Private Sub Form_Load()
Beep
frmMain.Timer2.Enabled = False
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption
lblDesc.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)

CurrentYear = Year(Now)
CurrentDate = Day(Now)
CurrentMonth = Format(Date, "MMM")
NeedTime = Format(Time, "HH:MM AMPM")

NeedDate = CurrentDate & "/" & CurrentMonth & "/" & CurrentYear

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan:

Do
    If ReS("Date") & ReS("AT") + " " + ReS("AP3") = NeedDate & NeedTime Then
        lblSchType.Caption = ReS("SchType")
        lblAT.Caption = ReS("AT") + " " + ReS("AP3")
        lblFrom.Caption = ReS("TF") + " " + ReS("AP1")
        lblTo.Caption = ReS("TT") + " " + ReS("AP2")
        lblDesc.Caption = ReS("Description")
        ReS.Close
        db.Close
        Set ReS = Nothing
        Set db = Nothing
        frmMain.Timer2.Enabled = True
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
        ReS.Close
        db.Close
        Set ReS = Nothing
        Set db = Nothing
        frmMain.Timer4.Enabled = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub lblOkSupport_Click()
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan

Do
    If ReS("Date") & ReS("AT") + " " + ReS("AP3") = NeedDate & lblAT.Caption Then
        ReS.Delete
        ReS.Close
        db.Close
        Set ReS = Nothing
        Set db = Nothing
        frmMain.Timer2.Enabled = True
        Unload Me
    Else
        ReS.MoveNext
    End If
Loop

ErrHan:
If Err.Number = 3021 Then
    frmMain.Timer2.Enabled = True
    ReS.Close
    db.Close
    Set ReS = Nothing
    Set db = Nothing
    Unload Me
End If
End Sub

Private Sub lblOkSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOk.ForeColor = RGB(145, 155, 100)
shapeOk.BackColor = vbBlack
End Sub

Private Sub lblOkSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeOk.BackColor = RGB(145, 155, 100)
lblOk.ForeColor = vbBlack
End Sub
