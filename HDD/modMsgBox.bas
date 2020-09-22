Attribute VB_Name = "modMsgBox"

Public MyMessage As String
Public MyMessage1 As String
Public Yes As Boolean

Public Function HDDMsgBox(strMessage As String)
MyMessage = strMessage
Beep
frmMsgBox.Show vbModal
End Function

Public Function HDDYesNoBox(strMessage1 As String)
MyMessage1 = strMessage1
Beep
frmYesNoBox.Show vbModal
End Function
