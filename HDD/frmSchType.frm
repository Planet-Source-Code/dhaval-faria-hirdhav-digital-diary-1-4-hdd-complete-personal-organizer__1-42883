VERSION 5.00
Begin VB.Form frmSchType 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Scheduler Type"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
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
   Icon            =   "frmSchType.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstSchType 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1110
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtSchType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Scheduler Type"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Text            =   "14/Mar/1985"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3045
      TabIndex        =   10
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   1380
      TabIndex        =   9
      Top             =   1560
      Width           =   240
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4920
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2640
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   840
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4920
      X2              =   4920
      Y1              =   240
      Y2              =   2280
   End
   Begin VB.Label lblDownSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblDown 
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
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape shapeDown 
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   4200
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduler Type:"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   225
      Left            =   1440
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
      Width           =   4935
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Scheduler Type"
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   15
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   2280
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4920
   End
End
Attribute VB_Name = "frmSchType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername
Public strSelMonth

Private Sub Form_Click()
lstSchType.Visible = False
txtSchType.SetFocus
End Sub

Private Sub Form_Load()
strUsername = frmMain.lblUsername.Caption
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
txtDate.BackColor = RGB(145, 155, 100)
shapeOk.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
txtDate.Text = SchDate & "/" & SchMonth & "/" & SchYear
selMonth = SchMonth
txtSchType.BackColor = RGB(145, 155, 100)
lstSchType.BackColor = RGB(145, 155, 100)
lstSchType.AddItem "Schedule"
lstSchType.AddItem "Reminder"
lstSchType.AddItem "Birthday Reminder"
lstSchType.AddItem "Aniversary Reminder"
lstSchType.AddItem "Appoinment"
lstSchType.AddItem "To Do"
shapeDown.BackColor = vbBlack
lblDown.ForeColor = RGB(145, 155, 100)
lstSchType.Height = 705
End Sub

Private Sub lblCancelSupport_Click()
Unload Me
End Sub

Private Sub lblCancelSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel.BackColor = vbBlack
lblCancel.ForeColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel.BackColor = RGB(145, 155, 100)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblDownSupport_Click()
If lstSchType.Visible = True Then
    txtSchType.SetFocus
    lstSchType.Visible = False
End If
lstSchType.Visible = True
lstSchType.SetFocus
End Sub

Private Sub lblOkSupport_Click()
If txtSchType.Text = "Scheduler Type" Then
    HDDMsgBox "Please select Scheduler Type"
    Exit Sub
End If
If lstSchType.ListIndex = "-1" Then
    HDDMsgBox "Please select Scheduler Type"
    Exit Sub
End If
For I = 0 To lstSchType.ListCount
    If txtSchType.Text <> lstSchType.Text Then
        HDDMsgBox "Please select Scheduler Type"
        Exit Sub
    End If
Next I
If txtSchType.Text = "Schedule" Then
    Unload frmCalender
    Me.Hide
    frmNewScheduler.Show
ElseIf txtSchType.Text = "Reminder" Then
    Unload frmCalender
    Me.Hide
    frmNewReminder.Show
ElseIf txtSchType.Text = "Birthday Reminder" Then
    Unload frmCalender
    Me.Hide
    frmNewBR.Show
ElseIf txtSchType.Text = "Aniversary Reminder" Then
    Unload frmCalender
    Me.Hide
    frmNewAR.Show
ElseIf txtSchType.Text = "Appoinment" Then
    HDDMsgBox "This feature will be available in Next version."
ElseIf txtSchType.Text = "To Do" Then
    Unload frmCalender
    Me.Hide
    frmNewToDo.Show
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

Private Sub lstSchType_Click()
txtSchType.Text = lstSchType.Text
lstSchType.Visible = False
txtSchType.SetFocus
End Sub

Private Sub lstSchType_LostFocus()
lstSchType.Visible = False
End Sub
