VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   15
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4320
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Reminder"
      End
      Begin VB.Menu mnuSR 
         Caption         =   "Setup Reminder"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSHDD 
         Caption         =   "Start HDD"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuOpen_Click()
frmMain.Timer3.Enabled = False
frmMain.Show
End Sub

Private Sub mnuSHDD_Click()
Shell (App.Path + "\HDD.exe")
End Sub

Private Sub mnuSR_Click()
frmSetup.Show
End Sub
