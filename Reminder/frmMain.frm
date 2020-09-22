VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Reminder"
   ClientHeight    =   4095
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1560
      Top             =   360
   End
   Begin VB.Timer timerTime 
      Interval        =   1
      Left            =   1080
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   600
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   360
   End
   Begin VB.ListBox lstTime 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDates 
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstSch 
      Appearance      =   0  'Flat
      Height          =   2505
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:10:10 AM"
      Height          =   225
      Left            =   3840
      TabIndex        =   8
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      Height          =   225
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   465
   End
   Begin VB.Line Line3 
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   240
      X2              =   4920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5175
      X2              =   5160
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here is Username."
      Height          =   225
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   930
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
      Caption         =   "Hirdhav Digital Diary  -  Reminder"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5175
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   1920
      TabIndex        =   9
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1920
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public CurrentMonth
Public CurrentDate
Public CurrentYear

Private Sub Form_Load()
frmMain.Hide
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
lblUsername.Caption = frmLogin.txtUsername.Text
lstSch.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)

lstTime.Clear
lstSch.Clear

Unload frmLogin

CurrentMonth = Format(Date, "MMM")
CurrentDate = Day(Now)
CurrentYear = Year(Now)

frmMain.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + lblUsername.Caption + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan
frmMain.Hide
Do
    txtDates.Text = ReS("Date")
    If ReS("Date") = CurrentDate & "/" & CurrentMonth & "/" & CurrentYear Then
        lstSch.AddItem ReS("TF") & " " & ReS("AP1") & "    " & ReS("Description")
        lstTime.AddItem ReS("AT") & " " & ReS("AP3")
        ReS.MoveNext
    Else
        ReS.MoveNext
    End If
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing
frmMain.Hide

ErrHan:
    If Err.Number = 3021 Then
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        frmMain.Hide
        Exit Sub
    End If
        frmMain.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Sys As Long
    Sys = X / Screen.TwipsPerPixelX
Select Case Sys
Case WM_RBUTTONDOWN:
    PopupMenu frmMenu.mnuMain
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub lblOkSupport_Click()
Timer3.Enabled = True
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
lstTime.Clear
lstSch.Clear
Set db = OpenDatabase(App.Path + "\Data\" + lblUsername.Caption + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan

Do
    
    txtDates.Text = ReS("Date")
    If ReS("Date") = CurrentDate & "/" & CurrentMonth & "/" & CurrentYear Then
        lstSch.AddItem ReS("TF") & " " & ReS("AP1") & "    " & ReS("Description")
        lstTime.AddItem ReS("AT") & " " & ReS("AP3")
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

Private Sub Timer2_Timer()
Timer4.Enabled = False
lstTime.ListIndex = -1
For i = 1 To lstTime.ListCount
    lstTime.ListIndex = lstTime.ListIndex + 1
    If lstTime.Text = Format(Time, "HH:MM AMPM") Then
        frmShowMe.Show
        Timer2.Enabled = False
    End If
Next i
End Sub

Private Sub Timer3_Timer()
frmMain.Hide
End Sub

Private Sub Timer4_Timer()
Timer2.Enabled = True
End Sub

Private Sub timerTime_Timer()
lblTime.Caption = Format(Time, "HH:MM:SS AMPM")
CurrentMonth = Format(Date, "MMM")
CurrentDate = Day(Now)
CurrentYear = Year(Now)
End Sub
