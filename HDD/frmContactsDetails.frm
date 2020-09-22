VERSION 5.00
Begin VB.Form frmContactsDetails 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Contects Details"
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
   Icon            =   "frmContactsDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReS 
      Height          =   375
      Left            =   240
      MaxLength       =   1
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   180
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
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   710
      Width           =   2655
   End
   Begin VB.TextBox txtFaxH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   1070
      Width           =   2655
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   1430
      Width           =   2655
   End
   Begin VB.TextBox txtAddressH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   1790
      Width           =   2655
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   2150
      Width           =   2655
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   2510
      Width           =   2655
   End
   Begin VB.TextBox txtPhoneB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   2870
      Width           =   2655
   End
   Begin VB.TextBox txtFaxB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   3230
      Width           =   2655
   End
   Begin VB.TextBox txtAddressB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   3590
      Width           =   2655
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   3950
      Width           =   2655
   End
   Begin VB.TextBox txtWebSite 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "http://"
      Top             =   4310
      Width           =   2655
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblOkSupport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   4665
      Width           =   1335
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Contacts Details"
      Height          =   225
      Left            =   195
      TabIndex        =   26
      Top             =   15
      Width           =   3330
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
      TabIndex        =   25
      Top             =   350
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (H):"
      Height          =   225
      Left            =   480
      TabIndex        =   24
      Top             =   710
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax (H):"
      Height          =   225
      Left            =   720
      TabIndex        =   23
      Top             =   1070
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      Height          =   225
      Left            =   720
      TabIndex        =   22
      Top             =   1430
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (H):"
      Height          =   225
      Left            =   240
      TabIndex        =   21
      Top             =   1790
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   225
      Left            =   720
      TabIndex        =   20
      Top             =   2150
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      Height          =   225
      Left            =   480
      TabIndex        =   19
      Top             =   2510
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (B):"
      Height          =   225
      Left            =   480
      TabIndex        =   18
      Top             =   2870
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax (B):"
      Height          =   225
      Left            =   720
      TabIndex        =   17
      Top             =   3230
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (B):"
      Height          =   225
      Left            =   240
      TabIndex        =   16
      Top             =   3590
      Width           =   1050
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      Height          =   225
      Left            =   840
      TabIndex        =   15
      Top             =   3950
      Width           =   435
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site:"
      Height          =   225
      Left            =   480
      TabIndex        =   14
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
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   5270
      Y2              =   5270
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      Height          =   225
      Left            =   2220
      TabIndex        =   13
      Top             =   4785
      Width           =   240
   End
   Begin VB.Shape shapeOk 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1680
      Top             =   4665
      Width           =   1335
   End
End
Attribute VB_Name = "frmContactsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strUsername

Private Sub Form_Load()
strUsername = frmMain.lblUsername.Caption
txtName.Text = frmContacts.lstName.Text
txtReS.Text = txtName.Text
Me.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
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
shapeOk.BackColor = RGB(145, 155, 100)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Personal.dat")
Set ReS = db.OpenRecordset(txtReS.Text)

ReS.Move (frmContacts.lstName.ListIndex)
txtPhoneH.Text = ReS("Phone(H)")
txtFaxH.Text = ReS("Fax(H)")
txtMobile.Text = ReS("Mobile")
txtEMail.Text = ReS("EMail")
txtAddressH.Text = ReS("Address(H)")
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

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblOkSupport_Click()
frmContacts.Show
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
