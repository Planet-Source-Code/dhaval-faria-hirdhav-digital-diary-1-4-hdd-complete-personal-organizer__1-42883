VERSION 5.00
Begin VB.Form frmJumpDate 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Jump Date"
   ClientHeight    =   2415
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
   Icon            =   "frmJumpDate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMonth 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstDate 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "Enter Year"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtMonth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "Select Month"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Select Date"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblJumpDateSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblJumpDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jump Date"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Shape shapeJumpDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3000
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      Height          =   225
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label lblDownMonthSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblDownMonth 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month:"
      Height          =   225
      Left            =   1180
      TabIndex        =   6
      Top             =   840
      Width           =   570
   End
   Begin VB.Label lblDownDaySupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblDownDay 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   225
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   435
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5175
      X2              =   5175
      Y1              =   240
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   2400
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Jump Date"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2850
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   5175
   End
   Begin VB.Shape shapeDownDay 
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   3120
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape shapeDownMonth 
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   3120
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "frmJumpDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername
Public SelDate

Private Sub Form_Click()
lstMonth.Visible = False
lstDate.Visible = False
End Sub

Private Sub Form_Load()
strUsername = frmMain.lblUsername.Caption
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
txtDate.BackColor = RGB(145, 155, 100)
txtMonth.BackColor = RGB(145, 155, 100)
txtYear.BackColor = RGB(145, 155, 100)
lstDate.BackColor = RGB(145, 155, 100)
lstMonth.BackColor = RGB(145, 155, 100)
lblDownDay.ForeColor = RGB(145, 155, 100)
shapeDownDay.BackColor = vbBlack
shapeDownMonth.BackColor = vbBlack
lblDownMonth.ForeColor = RGB(145, 155, 100)
shapeJumpDate.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
lstDate.Height = 705
lstMonth.Height = 705
For I = 1 To 31
    lstDate.AddItem I
Next I
lstMonth.AddItem "Jan"
lstMonth.AddItem "Feb"
lstMonth.AddItem "Mar"
lstMonth.AddItem "Jun"
lstMonth.AddItem "Jul"
lstMonth.AddItem "Aug"
lstMonth.AddItem "Sep"
lstMonth.AddItem "Oct"
lstMonth.AddItem "Nov"
lstMonth.AddItem "Dec"
End Sub

Private Sub lblCancelSupport_Click()
Unload Me
End Sub

Private Sub lblCancelSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = vbBlack
End Sub

Private Sub lblCancelSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = vbBlack
shapeCancel.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblDownDaySupport_Click()
If lstDate.Visible = True Then
    lstDate.Visible = False
Else
    lstDate.Visible = True
    lstDate.SetFocus
End If
End Sub

Private Sub lblDownMonthSupport_Click()
If lstMonth.Visible = False Then
    lstMonth.Visible = True
    lstMonth.SetFocus
Else
    lstMonth.Visible = False
End If
End Sub

Private Sub lblJumpDateSupport_Click()

frmCalender.txtDate.Text = txtDate.Text
CurrentMonth = txtMonth.Text
CurrentYear = txtYear.Text

With frmCalender
    
    .lstSchList.Clear
    For j = 0 To 36
        .shapeAP1(j).Visible = False
        .shapeAP2(j).Visible = False
    Next j
    For I = 0 To 36
        .shapeDate(I).BackColor = RGB(145, 155, 100)
        .lblDate(I).ForeColor = vbBlack
    Next I
    
    If CurrentMonth = "Jan" Then
        .lblMonth.Caption = "January"
    ElseIf CurrentMonth = "Feb" Then
        .lblMonth.Caption = "February"
    ElseIf CurrentMonth = "Mar" Then
        .lblMonth.Caption = "March"
    ElseIf CurrentMonth = "Apr" Then
        .lblMonth.Caption = "April"
    ElseIf CurrentMonth = "May" Then
        .lblMonth.Caption = "May"
    ElseIf CurrentMonth = "Jun" Then
        .lblMonth.Caption = "June"
    ElseIf CurrentMonth = "Jul" Then
        .lblMonth.Caption = "July"
    ElseIf CurrentMonth = "Aug" Then
        .lblMonth.Caption = "August"
    ElseIf CurrentMonth = "Sep" Then
        .lblMonth.Caption = "September"
    ElseIf CurrentMonth = "Oct" Then
        .lblMonth.Caption = "October"
    ElseIf CurrentMonth = "Nov" Then
        .lblMonth.Caption = "November"
    ElseIf CurrentMonth = "Dec" Then
        .lblMonth.Caption = "December"
    End If
        .lblYear.Caption = CurrentYear

End With

IDs

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)
With frmCalender
On Error GoTo ErrHan
        .txtDate.Text = ReS("Date")
        If Len(.txtDate.Text) = 11 Then
            .txtDate.SelStart = 7
            .txtDate.SelLength = 4
        End If
        If Len(.txtDate.Text) = 10 Then
            .txtDate.SelStart = 6
            .txtDate.SelLength = 4
        End If
Do
    .txtDate.Text = ReS("Date")
    If Len(.txtDate.Text) = 11 Then
        .txtDate.SelStart = 0
        .txtDate.SelLength = 2
        SelDateTmp = .txtDate.SelText
        .txtDate.SelStart = 7
        .txtDate.SelLength = 4
        SelYearTmp = .txtDate.SelText
    Else
        .txtDate.SelStart = 0
        .txtDate.SelLength = 1
        SelDateTmp = .txtDate.SelText
        .txtDate.SelStart = 6
        .txtDate.SelLength = 4
        SelYearTmp = .txtDate.SelText
    End If
    AMPM1 = ReS("AP1")
    AMPM2 = ReS("AP2")
    
    If SelDateTmp & SelYearTmp = SelDate & CurrentYear Then
        .lstSchList.AddItem ReS("TF") & ReS("AP1") & "  " & ReS("Description")
    End If
    
    For I = 0 To 36
        If Len(txtDate.Text) = 11 Then
            .txtDate.SelStart = 7
            .txtDate.SelLength = 4
        Else
            .txtDate.SelStart = 6
            .txtDate.SelLength = 4
        End If
    If .txtDate.SelText = CurrentYear Then
        If .lblDate(I).Caption = SelDateTmp Then
            If AMPM1 = "AM" Then
                If .shapeDate(I).BackColor = vbBlack Then
                    .shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                .shapeAP1(I).Visible = True
            End If
            If AMPM1 = "PM" Then
                If .shapeDate(I).BackColor = vbBlack Then
                    .shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                .shapeAP2(I).Visible = True
            End If
            If AMPM2 = "AM" Then
                If .shapeDate(I).BackColor = vbBlack Then
                    .shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                .shapeAP1(I).Visible = True
            End If
            If AMPM2 = "PM" Then
                If .shapeDate(I).BackColor = vbBlack Then
                    .shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                .shapeAP2(I).Visible = True
            End If
        End If
    End If
    Next I
    ReS.MoveNext
Loop
End With

ReS.Close
db.Close

Set db = Nothing
Set ReS = Nothing

SelDate = txtDate.Text
Unload Me

ErrHan:
If Err.Number = 3021 Then
    SelDate = txtDate.Text
    Unload Me
    Exit Sub
End If
End Sub

Private Sub lblJumpDateSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblJumpDate.ForeColor = RGB(145, 155, 100)
shapeJumpDate.BackColor = vbBlack
End Sub

Private Sub lblJumpDateSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeJumpDate.BackColor = RGB(145, 155, 100)
lblJumpDate.ForeColor = vbBlack
End Sub

Private Sub lstDate_Click()
txtDate.Text = lstDate.Text
lstDate.Visible = False
End Sub

Private Sub lstDate_LostFocus()
lstDate.Visible = False
End Sub

Private Sub lstMonth_Click()
txtMonth.Text = lstMonth.Text
lstMonth.Visible = False
End Sub

Private Sub lstMonth_LostFocus()
lstMonth.Visible = False
End Sub

Private Sub txtYear_GotFocus()
txtYear.SelStart = 0
txtYear.SelLength = Len(txtYear.Text)
End Sub
