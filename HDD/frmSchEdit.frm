VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSchEdit 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Edit Scheduler"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
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
   Icon            =   "frmSchEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtAP3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Text            =   "AM"
      Top             =   1800
      Width           =   375
   End
   Begin MSMask.MaskEdBox txtAT 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtAP2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      TabIndex        =   12
      Text            =   "PM"
      Top             =   1440
      Width           =   375
   End
   Begin MSMask.MaskEdBox txtTT 
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   1440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtAP1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Text            =   "AM"
      Top             =   1440
      Width           =   375
   End
   Begin MSMask.MaskEdBox txtTF 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblEditSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   3480
      TabIndex        =   17
      Top             =   3600
      Width           =   585
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   225
      Left            =   1200
      TabIndex        =   16
      Top             =   3600
      Width           =   315
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   5280
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3000
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Shape shapeEdit 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   600
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   225
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Time:"
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1020
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time From:"
      Height          =   225
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14/Mar/1985"
      Height          =   225
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   225
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   435
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5295
      X2              =   5295
      Y1              =   240
      Y2              =   4200
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
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   4905
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
      Caption         =   "Hirdhav Digital Diary  -  Edit Scheduler"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   3165
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
      Width           =   5295
   End
End
Attribute VB_Name = "frmSchEdit"
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
txtTF.BackColor = RGB(145, 155, 100)
txtAP1.BackColor = RGB(145, 155, 100)
txtTT.BackColor = RGB(145, 155, 100)
txtAP2.BackColor = RGB(145, 155, 100)
txtAT.BackColor = RGB(145, 155, 100)
txtAP3.BackColor = RGB(145, 155, 100)
txtDesc.BackColor = RGB(145, 155, 100)
shapeEdit.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
strUsername = frmMain.lblUsername.Caption

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

Do
    NeedDate = ReS("TF") + ReS("AP1") + "  " + ReS("Description")
    If NeedDate = frmCalender.lstSchList.Text Then
        lblType.Caption = ReS("SchType")
        lblDate.Caption = ReS("Date")
        txtTF.Text = ReS("TF")
        txtAP1.Text = ReS("AP1")
        txtTT.Text = ReS("TT")
        txtAP2.Text = ReS("AP2")
        txtAT.Text = ReS("AT")
        txtAP3.Text = ReS("AP3")
        txtDesc.Text = ReS("Description")
        
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        Exit Sub
    Else
        ReS.MoveNext
    End If
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing
End Sub

Private Sub lblCancelSupport_Click()
frmCalender.Show
Unload Me
End Sub

Private Sub lblCancelSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = vbBlack
End Sub

Private Sub lblCancelSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCancel.BackColor = RGB(145, 155, 100)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblEditSupport_Click()
HDDYesNoBox "Are you sure? Do you want to Edit this?"

If Yes Then
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

Do
    If NeedDate = frmCalender.lstSchList.Text Then
        ReS.Edit
        ReS("AT") = txtAT.Text
        ReS("TF") = txtTF.Text
        ReS("TT") = txtTT.Text
        ReS("AP1") = txtAP1.Text
        ReS("AP2") = txtAP2.Text
        ReS("AP3") = txtAP3.Text
        ReS("Description") = txtDesc.Text
        ReS.Update
        
        ReS.Close
        db.Close
        
        Set ReS = Nothing
        Set db = Nothing
        HDDMsgBox "Record Edited Successfully."
        Unload frmCalender
        frmCalender.Show
        Unload Me
        Exit Sub
    Else
        ReS.MoveNext
    End If
Loop

End If
End Sub

Private Sub lblEditSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = RGB(145, 155, 100)
shapeEdit.BackColor = vbBlack
End Sub

Private Sub lblEditSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeEdit.BackColor = RGB(145, 155, 100)
lblEdit.ForeColor = vbBlack
End Sub
