Attribute VB_Name = "modDragForm"

Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "USER32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Sub DragForm(Who As Form)
On Local Error Resume Next
Call ReleaseCapture
Call SendMessage(Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

