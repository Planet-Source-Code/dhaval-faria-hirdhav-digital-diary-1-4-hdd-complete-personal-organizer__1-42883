VERSION 5.00
Begin VB.Form frmContactsEdit 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Contects Edit"
   ClientHeight    =   5280
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
   Icon            =   "frmContactsEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReS 
      Height          =   330
      Left            =   120
      MaxLength       =   1
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Top             =   350
      Width           =   2655
   End
   Begin VB.TextBox txtPhoneH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   710
      Width           =   2655
   End
   Begin VB.TextBox txtFaxH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1070
      Width           =   2655
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1430
      Width           =   2655
   End
   Begin VB.TextBox txtAddressH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1790
      Width           =   2655
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2150
      Width           =   2655
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   2510
      Width           =   2655
   End
   Begin VB.TextBox txtPhoneB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2870
      Width           =   2655
   End
   Begin VB.TextBox txtFaxB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   3230
      Width           =   2655
   End
   Begin VB.TextBox txtAddressB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   3590
      Width           =   2655
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   3950
      Width           =   2655
   End
   Begin VB.TextBox txtWebSite 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "http://"
      Top             =   4310
      Width           =   2655
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblCancelSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   4670
      Width           =   1335
   End
   Begin VB.Label lblEditSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   4670
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      Height          =   225
      Left            =   1170
      TabIndex        =   15
      Top             =   4785
      Width           =   315
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Contacts Edit"
      Height          =   225
      Left            =   200
      TabIndex        =   28
      Top             =   15
      Width           =   3060
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   20
      Top             =   20
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   230
      Y2              =   5270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   225
      Left            =   720
      TabIndex        =   27
      Top             =   350
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (H):"
      Height          =   225
      Left            =   480
      TabIndex        =   26
      Top             =   710
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax (H):"
      Height          =   225
      Left            =   720
      TabIndex        =   25
      Top             =   1070
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      Height          =   225
      Left            =   720
      TabIndex        =   24
      Top             =   1430
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (H):"
      Height          =   225
      Left            =   240
      TabIndex        =   23
      Top             =   1790
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   720
      TabIndex        =   22
      Top             =   2150
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      Height          =   225
      Left            =   480
      TabIndex        =   21
      Top             =   2510
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (B):"
      Height          =   225
      Left            =   480
      TabIndex        =   20
      Top             =   2870
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax (B):"
      Height          =   225
      Left            =   720
      TabIndex        =   19
      Top             =   3230
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (B):"
      Height          =   225
      Left            =   240
      TabIndex        =   18
      Top             =   3590
      Width           =   1050
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      Height          =   225
      Left            =   840
      TabIndex        =   17
      Top             =   3950
      Width           =   435
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site:"
      Height          =   225
      Left            =   480
      TabIndex        =   16
      Top             =   4310
      Width           =   810
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4575
      X2              =   4560
      Y1              =   230
      Y2              =   5270
   End
   Begin VB.Shape shapeEdit 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   720
      Top             =   4670
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   5270
      Y2              =   5270
   End
   Begin VB.Label lblCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   2860
      TabIndex        =   14
      Top             =   4790
      Width           =   585
   End
   Begin VB.Shape shapeCancel 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2520
      Top             =   4670
      Width           =   1335
   End
End
Attribute VB_Name = "frmContactsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strUsername As String

Private Sub Form_Load()
Me.BackColor = RGB(145, 155, 100)
shapeEdit.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeCancel.BackColor = RGB(145, 155, 100)
txtName.BackColor = RGB(145, 155, 100)
txtPhoneH.BackColor = RGB(145, 155, 100)
txtFaxH.BackColor = RGB(145, 155, 100)
txtPhoneB.BackColor = RGB(145, 155, 100)
txtFaxB.BackColor = RGB(145, 155, 100)
txtAddressH.BackColor = RGB(145, 155, 100)
txtAddressB.BackColor = RGB(145, 155, 100)
txtWebSite.BackColor = RGB(145, 155, 100)
txtEMail.BackColor = RGB(145, 155, 100)
txtNote.BackColor = RGB(145, 155, 100)
txtMobile.BackColor = RGB(145, 155, 100)
txtCompany.BackColor = RGB(145, 155, 100)

strUsername = frmMain.lblUsername.Caption
txtName.Text = frmContacts.lstName.Text
txtReS.Text = txtName.Text

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Personal.dat")
Set ReS = db.OpenRecordset(txtReS.Text)

ReS.Move (frmContacts.lstName.ListIndex)
txtPhoneH.Text = ReS("Phone(H)")
txtFaxH.Text = ReS("Fax(H)")
txtMobile.Text = ReS("Mobile")
txtAddressH.Text = ReS("Address(H)")
txtEMail.Text = ReS("EMail")
txtCompany.Text = ReS("Company")
txtPhoneB.Text = ReS("Phone(B)")
txtFaxB.Text = ReS("Fax(B)")
txtAddressB.Text = ReS("Address(B)")
txtNote.Text = ReS("Note")
txtWebSite.Text = ReS("WebSite")

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

End Sub

Private Sub lblCancelSupport_Click()
Unload Me
frmContacts.Show
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

Private Sub lblEditSupport_Click()
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Personal.dat")
Set ReS = db.OpenRecordset(txtReS.Text)

ReS.Move (frmContacts.lstName.ListIndex)
ReS.Edit
ReS("Phone(H)") = txtPhoneH.Text
ReS("Fax(H)") = txtFaxH.Text
ReS("Address(H)") = txtAddressH.Text
ReS("Mobile") = txtMobile.Text
ReS("EMail") = txtEMail.Text
ReS("Phone(B)") = txtPhoneB.Text
ReS("Fax(B)") = txtFaxB.Text
ReS("Address(B)") = txtAddressB.Text
ReS("Note") = txtNote.Text
ReS("WebSite") = txtWebSite.Text
ReS("Company") = txtCompany.Text
ReS.Update
HDDMsgBox "Contact edited successfully."
frmContacts.Show
Unload Me
End Sub

Private Sub lblEditSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = RGB(145, 155, 100)
shapeEdit.BackColor = vbBlack
End Sub

Private Sub lblEditSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEdit.ForeColor = vbBlack
shapeEdit.BackColor = RGB(145, 155, 100)
End Sub
