VERSION 5.00
Begin VB.Form frmContacts 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Contacts"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
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
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comContacts 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   360
      TabIndex        =   35
      Text            =   "Personal"
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstNew 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3360
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstInfo 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   2385
      TabIndex        =   2
      Top             =   960
      Width           =   1860
   End
   Begin VB.ListBox lstName 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2025
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblDeleteSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2880
      TabIndex        =   39
      Top             =   4240
      Width           =   1335
   End
   Begin VB.Label lblEditSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      TabIndex        =   38
      Top             =   4240
      Width           =   1335
   End
   Begin VB.Label lblDelete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      Height          =   225
      Left            =   3280
      TabIndex        =   37
      Top             =   4380
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   225
      Left            =   840
      TabIndex        =   36
      Top             =   4380
      Width           =   315
   End
   Begin VB.Shape shapeDelete 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2880
      Top             =   4250
      Width           =   1335
   End
   Begin VB.Shape shapeEdit 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Top             =   4250
      Width           =   1335
   End
   Begin VB.Label lblNewSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1800
      TabIndex        =   33
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblNew 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      Height          =   225
      Left            =   2205
      TabIndex        =   32
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape shapeNew 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1800
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblMainMenuSupport 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   360
      TabIndex        =   31
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Label lblMainMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      Height          =   225
      Left            =   1800
      TabIndex        =   30
      Top             =   4875
      Width           =   915
   End
   Begin VB.Shape shapeMainMenu 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      Height          =   225
      Index           =   26
      Left            =   3675
      TabIndex        =   29
      Top             =   3940
      Width           =   75
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   225
      Index           =   25
      Left            =   3440
      TabIndex        =   28
      Top             =   3930
      Width           =   105
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   225
      Index           =   24
      Left            =   3195
      TabIndex        =   27
      Top             =   3930
      Width           =   105
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   225
      Index           =   23
      Left            =   2960
      TabIndex        =   26
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   225
      Index           =   22
      Left            =   2680
      TabIndex        =   25
      Top             =   3930
      Width           =   180
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   225
      Index           =   21
      Left            =   2480
      TabIndex        =   24
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   225
      Index           =   20
      Left            =   2240
      TabIndex        =   23
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   225
      Index           =   19
      Left            =   2000
      TabIndex        =   22
      Top             =   3930
      Width           =   105
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   225
      Index           =   18
      Left            =   1755
      TabIndex        =   21
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   225
      Index           =   17
      Left            =   1520
      TabIndex        =   20
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   225
      Index           =   16
      Left            =   1275
      TabIndex        =   19
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   225
      Index           =   15
      Left            =   1040
      TabIndex        =   18
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   225
      Index           =   14
      Left            =   795
      TabIndex        =   17
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   225
      Index           =   13
      Left            =   3800
      TabIndex        =   16
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   225
      Index           =   12
      Left            =   3540
      TabIndex        =   15
      Top             =   3690
      Width           =   150
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   225
      Index           =   11
      Left            =   3340
      TabIndex        =   14
      Top             =   3690
      Width           =   105
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   225
      Index           =   10
      Left            =   3080
      TabIndex        =   13
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   225
      Index           =   9
      Left            =   2840
      TabIndex        =   12
      Top             =   3690
      Width           =   105
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   225
      Index           =   8
      Left            =   2630
      TabIndex        =   11
      Top             =   3690
      Width           =   45
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   225
      Index           =   7
      Left            =   2360
      TabIndex        =   10
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   225
      Index           =   6
      Left            =   2120
      TabIndex        =   9
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   225
      Index           =   5
      Left            =   1900
      TabIndex        =   8
      Top             =   3690
      Width           =   90
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   225
      Index           =   4
      Left            =   1640
      TabIndex        =   7
      Top             =   3690
      Width           =   105
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   225
      Index           =   3
      Left            =   1400
      TabIndex        =   6
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   225
      Index           =   2
      Left            =   1160
      TabIndex        =   5
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   225
      Index           =   1
      Left            =   930
      TabIndex        =   4
      Top             =   3690
      Width           =   120
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   225
      Index           =   0
      Left            =   680
      TabIndex        =   3
      Top             =   3690
      Width           =   120
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   24
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   25
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   26
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   19
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   20
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   21
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   22
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   23
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   14
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   15
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   16
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   17
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   18
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   13
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   12
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   11
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   10
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   9
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   8
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   7
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   6
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   4
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Shape shapeAlpha 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3675
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Contacts"
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   15
      Width           =   2700
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   15
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   4560
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public tmpUsername As String
Public tmpAlpha As String

Private Sub comContacts_Click()
If comContacts.Text = "Buisness" Then
HDDMsgBox "Sorry, Not available It will be available very soon."
'lstName.Width = "3885"
'lstInfo.Visible = False
comContacts.ListIndex = 0
End If
If comContacts.Text = "Personal" Then
lstName.Width = "2025"
lstInfo.Width = "1860"
lstInfo.Visible = True
lstName.Visible = True
End If
End Sub

Private Sub Form_Load()
comContacts.AddItem "Personal"
comContacts.AddItem "Buisness"
comContacts.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
shapeEdit.BackColor = RGB(145, 155, 100)
shapeDelete.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
lstNew.BackColor = RGB(145, 155, 100)
lstNew.AddItem "Personal"
lstNew.AddItem "Buisness"
shapeNew.BackColor = RGB(145, 155, 100)
tmpUsername = frmMain.lblUsername.Caption
Me.BackColor = RGB(145, 155, 100)
lstName.BackColor = RGB(145, 155, 100)
lstInfo.BackColor = RGB(145, 155, 100)
shapeAlpha(0).BackColor = vbBlack
lblAlpha(0).ForeColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = RGB(145, 155, 100)
tmpAlpha = "A"
For I = 1 To 26
    shapeAlpha(I).BackColor = RGB(145, 155, 100)
Next I
For I = 1 To 13
    lblAlpha(I).ForeColor = vbBlack
Next I

If comContacts.Text = "Personal" Then
'Open database and add contacts names in to the list
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + tmpUsername + "\Personal.dat")
Set ReS = db.OpenRecordset("A")
On Error GoTo HanErr:
Do
    lstName.AddItem ReS("Name")
    lstInfo.AddItem ReS("Phone(H)")
    ReS.MoveNext
Loop
End If

HanErr:
    If Err.Number = 3021 Then
    
    End If
ReS.Close
db.Close
Set ReS = Nothing
Set db = Nothing
End Sub

Private Sub lblAlpha_Click(Index As Integer)
Dim AlphaNo As Integer
AlphaNo = lblAlpha(Index).Index
For I = 0 To 26
    shapeAlpha(I).BackColor = RGB(145, 155, 100)
    lblAlpha(I).ForeColor = vbBlack
Next I
shapeAlpha(AlphaNo).BackColor = vbBlack
lblAlpha(AlphaNo).ForeColor = RGB(145, 155, 100)

'Clear the list boxes
lstName.Clear
lstInfo.Clear

'Open database and add Contacts in list boxes.

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + tmpUsername + "\Personal.dat")
Set ReS = db.OpenRecordset(lblAlpha(AlphaNo).Caption)
tmpAlpha = lblAlpha(AlphaNo).Caption
On Error GoTo HanErr:
Do
    lstName.AddItem ReS("Name")
    lstInfo.AddItem ReS("Phone(H)")
    ReS.MoveNext
Loop

HanErr:
    If Err.Number = 3021 Then
    
    End If
lstName.SetFocus
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblDeleteSupport_Click()
If lstName.ListIndex = "-1" Then
    HDDMsgBox "Please selct the contact from the list."
Else
    HDDYesNoBox "Do you want to delete it?"
        
        If Yes Then
            
            Dim db As Database
            Dim ReS As Recordset
            
            Set db = OpenDatabase(App.Path + "\Data\" + tmpUsername + "\Personal.dat")
            Set ReS = db.OpenRecordset(tmpAlpha)
            
            ReS.Move (lstName.ListIndex)
            ReS.Delete
            
            ReS.Close
            db.Close
            
            Set ReS = Nothing
            Set db = Nothing
            
            HDDMsgBox "Contact deleted successfully."
            
            Unload Me
            frmContacts.Show
        
        End If

End If
End Sub

Private Sub lblDeleteSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDelete.ForeColor = RGB(145, 155, 100)
shapeDelete.BackColor = vbBlack
End Sub

Private Sub lblDeleteSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeDelete.BackColor = RGB(145, 155, 100)
lblDelete.ForeColor = vbBlack
End Sub

Private Sub lblEditSupport_Click()
If lstName.ListIndex = "-1" Then
    HDDMsgBox "Please select the contact from the list."
Else
    frmContactsEdit.Show
    Me.Hide
End If
End Sub

Private Sub lblEditSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = RGB(145, 155, 100)
shapeEdit.BackColor = vbBlack
End Sub

Private Sub lblEditSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = vbBlack
shapeEdit.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblMainMenuSupport_Click()
frmMain.Show
Unload Me
End Sub

Private Sub lblMainMenuSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMainMenu.ForeColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = vbBlack
End Sub

Private Sub lblMainMenuSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMainMenu.ForeColor = vbBlack
shapeMainMenu.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblNewSupport_Click()
frmNewPersonal.Show
Me.Hide
End Sub

Private Sub lblNewSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNew.ForeColor = RGB(145, 155, 100)
shapeNew.BackColor = vbBlack
End Sub

Private Sub lblNewSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNew.ForeColor = vbBlack
shapeNew.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lstInfo_Click()
lstName.ListIndex = lstInfo.ListIndex
End Sub

Private Sub lstInfo_DblClick()
If lstName.ListIndex = "-1" Then
    HDDMsgBox "Please select the contact from the list."
Else
    frmContactsDetails.Show
    Me.Hide
End If
End Sub

Private Sub lstName_Click()
lstInfo.ListIndex = lstName.ListIndex
End Sub

Private Sub lstName_DblClick()
If lstName.ListIndex = "-1" Then
    HDDMsgBox "Please select the contact from the list."
Else
    frmContactsDetails.Show
    Me.Hide
End If
End Sub

Private Sub lstNew_Click()
If lstNew.Text = "Personal" Then
frmNewPersonal.Show
Me.Hide
End If

If lstNew.Text = "Buisness" Then
HDDMsgBox "Sorry, Not available It will be available very soon."
End If
End Sub

Private Sub lstNew_LostFocus()
lstNew.Visible = False
End Sub
