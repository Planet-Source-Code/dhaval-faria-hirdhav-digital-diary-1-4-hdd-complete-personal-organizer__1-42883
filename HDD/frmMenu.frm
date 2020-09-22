VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuDates 
      Caption         =   "Dates"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFI 
         Caption         =   "Full Information"
      End
   End
   Begin VB.Menu mnuLB 
      Caption         =   "ListBox"
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strUsername
Public NeedDetail

Private Sub mnuDelete_Click()
strUsername = frmMain.lblUsername.Caption
If frmCalender.lstSchList.ListIndex = "-1" Then
    HDDMsgBox "Please select item from the list."
    Exit Sub
End If
HDDYesNoBox "Are you sure? Do you want to delete this?"
If Yes Then
    Dim db As Database
    Dim ReS As Recordset
    
    Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
    Set ReS = db.OpenRecordset(CurrentMonth)
    
    Do
        NeedDetail = ReS("TF") + ReS("AP1") + "  " + ReS("Description")
        If NeedDetail = frmCalender.lstSchList.Text Then
            ReS.Delete
            
            ReS.Close
            db.Close
            
            Set ReS = Nothing
            Set db = Nothing
            HDDMsgBox "Record deleted successfully."
            Unload frmCalender
            frmCalender.Show
            Unload Me
            Exit Sub
        Else
            ReS.MoveNext
        End If
    Loop
    ReS.Close
    db.Close
    
    Set ReS = Nothing
    Set db = Nothing
End If
End Sub

Private Sub mnuEdit_Click()
If frmCalender.lstSchList.ListIndex = "-1" Then
    HDDMsgBox "Please select item from the list."
    Exit Sub
End If
frmSchEdit.Show
frmCalender.Hide
End Sub

Private Sub mnuNew_Click()
frmSchType.Show
Me.Hide
End Sub
