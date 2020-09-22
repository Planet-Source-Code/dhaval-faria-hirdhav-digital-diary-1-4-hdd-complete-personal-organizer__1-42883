VERSION 5.00
Begin VB.Form frmNewPersonal 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - New Contacts"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
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
   Icon            =   "frmNewPersonal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWebSite 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Text            =   "http://"
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox txtAddressB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   20
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtFaxB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtPhoneB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtAddressH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtFaxH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtPhoneH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2520
      TabIndex        =   28
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblAddSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   720
      TabIndex        =   27
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   2860
      TabIndex        =   26
      Top             =   4800
      Width           =   585
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      Height          =   225
      Left            =   1170
      TabIndex        =   25
      Top             =   4800
      Width           =   330
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2520
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Shape shapeAdd 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   720
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4575
      X2              =   4560
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site:"
      Height          =   225
      Left            =   480
      TabIndex        =   23
      Top             =   4320
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      Height          =   225
      Left            =   840
      TabIndex        =   21
      Top             =   3960
      Width           =   435
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (B):"
      Height          =   225
      Left            =   240
      TabIndex        =   19
      Top             =   3600
      Width           =   1050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax (B):"
      Height          =   225
      Left            =   720
      TabIndex        =   17
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (B):"
      Height          =   225
      Left            =   480
      TabIndex        =   15
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      Height          =   225
      Left            =   480
      TabIndex        =   13
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   720
      TabIndex        =   11
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (H):"
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      Height          =   225
      Left            =   720
      TabIndex        =   7
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax (H):"
      Height          =   225
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (H):"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   225
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  New Contacts"
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   15
      Width           =   3120
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   4575
   End
End
Attribute VB_Name = "frmNewPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
lblCaption.ForeColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
txtName.BackColor = RGB(145, 155, 100)
txtPhoneH.BackColor = RGB(145, 155, 100)
txtFaxH.BackColor = RGB(145, 155, 100)
txtAddressH.BackColor = RGB(145, 155, 100)
txtMobile.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtCompany.BackColor = RGB(145, 155, 100)
txtWebSite.BackColor = RGB(145, 155, 100)
txtNote.BackColor = RGB(145, 155, 100)
txtPhoneB.BackColor = RGB(145, 155, 100)
txtFaxB.BackColor = RGB(145, 155, 100)
txtAddressB.BackColor = RGB(145, 155, 100)
shapeAdd.BackColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblAddSupport_Click()
Dim conType As String
conType = frmContacts.lstNew.Text
If txtName.Text = "" Then
    HDDMsgBox "Please enter Name."
    Exit Sub
End If
If txtName.Text = " " Then
    HDDMsgBox "Please enter Name."
    Exit Sub
End If
Dim strUsername As String
strUsername = frmMain.lblUsername.Caption
Me.Controls.Add "VB.textBox", "txtN"
With Me!txtN
    .Visible = False
    .MaxLength = 1
    .Text = txtName.Text
End With

Dim db As Database
Dim ReS As Recordset

On Error GoTo ErrHan:
Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Personal.dat")
Set ReS = db.OpenRecordset(Me!txtN.Text)

ReS.AddNew
ReS("Name") = txtName.Text
ReS("Phone(H)") = txtPhoneH.Text
ReS("Fax(H)") = txtFaxH.Text
ReS("Mobile") = txtMobile.Text
ReS("Address(H)") = txtAddressH.Text
ReS("EMail") = txtEMail.Text
ReS("Company") = txtCompany.Text
ReS("Phone(B)") = txtPhoneB.Text
ReS("Fax(B)") = txtFaxB.Text
ReS("Address(B)") = txtAddressB.Text
ReS("Note") = txtNote.Text
ReS("WebSite") = txtWebSite.Text
ReS.Update
HDDMsgBox "Record inserted successfully."
ReS.Close
db.Close
Set db = Nothing
Set ReS = Nothing
Unload frmContacts
frmContacts.Show
Unload Me
Exit Sub

ErrHan:
    
    If Err.Number = 3078 Then
        Me!txtN.Text = "*"
        'Dim db As Database
        'Dim ReS As Recordset

        Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Personal.dat")
        Set ReS = db.OpenRecordset(Me!txtN.Text)

        ReS.AddNew
        ReS("Name") = txtName.Text
        ReS("Phone(H)") = txtPhoneH.Text
        ReS("Fax(H)") = txtFaxH.Text
        ReS("Mobile") = txtMobile.Text
        ReS("Address(H)") = txtAddressH.Text
        ReS("EMail") = txtEMail.Text
        ReS("Company") = txtCompany.Text
        ReS("Phone(B)") = txtPhoneB.Text
        ReS("Fax(B)") = txtFaxB.Text
        ReS("Address(B)") = txtAddressB.Text
        ReS("Note") = txtNote.Text
        ReS("WebSite") = txtWebSite.Text
        ReS.Update
        HDDMsgBox "Record inserted successfully."
        ReS.Close
        db.Close
        Set db = Nothing
        Set ReS = Nothing
        Unload frmContacts
        frmContacts.Show
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub lblAddSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = RGB(145, 155, 100)
shapeAdd.BackColor = vbBlack
End Sub

Private Sub lblAddSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.ForeColor = vbBlack
shapeAdd.BackColor = RGB(145, 155, 100)
End Sub

Private Sub lblCancelSupport_Click()
frmContacts.Show
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

Private Sub txtWebSite_GotFocus()
SendKeys "{END}"
End Sub
