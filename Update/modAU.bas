Attribute VB_Name = "modAU"

'We will save our Data in this Variables...
Global nVer As Single
Global nMsg As String
Global nURL As String
Global nEXE As String
Dim b() As Byte
Global Version As Single

'This are the Net Variables...
Global VersionURL As String
Global MessageURL As String
Global UpdateURL As String
Global ExeURL As String

Function CheckINET()
Dim flags As Long
Dim result As Boolean

    result = InternetGetConnectedState(flags, 0)
    If result Then
        frmAU.lblStatus.Caption = "Connected to the Internet... Please wait..."
        GetOrConnect
    Else
        frmAU.lblStatus.Caption = "Not Connected to the Internet... Please Connect to the Internet if you want to download the update..."
        GetOrConnect
        frmAU.lblStart.Enabled = True
        frmAU.lblStartSupport.Enabled = True
    End If
     
    'If flags And INTERNET_CONNECTION_MODEM Then Print "Connection Via Modem"
    'If flags And INTERNET_CONNECTION_LAN Then Print "Connecion Via LAN"
    'If flags And INTERNET_CONNECTION_PROXY Then Print "Connection uses a Proxy"
    'If flags And INTERNET_CONNECTION_MODEM_BUSY Then Print "Connection Via Modem but modem is busy"

End Function

Function GetOrConnect()
'Empty all the variables
nVer = 0
nMsg = ""
nURL = ""

'Show Checking Current Version in Status Bar
frmAU.lblStatus.Caption = "Checking the current version...."

'Get the current version
Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase("c:\dhaval\vb\hdd project\hdd\Installation.dat")
Set ReS = db.OpenRecordset("Installation")

Version = ReS("Version")

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

'Declare the URLs
VersionURL = "http://www26.brinkster.com/dhavalfa/Version1p3.txt"
MessageURL = "http://www26.brinkster.com/dhavalfa/Message1p3.txt"
UpdateURL = "http://www26.brinkster.com/dhavalfa/Update1p3.txt"
ExeURL = "http://www26.brinkster.com/dhavalfa/EXE1p3.txt"

On Error GoTo ErrHan:

'We will download the information...
frmAU.lblStatus.Caption = "Connecting to server...."
nVer = frmAU.net.OpenURL(VersionURL)
frmAU.lblStatus.Caption = "Getting version information...."
nMsg = frmAU.net.OpenURL(MessageURL)
frmAU.lblStatus.Caption = "Getting update message...."
nEXE = frmAU.net.OpenURL(ExeURL)

If nVer > Version Then
    b() = frmAU.net.OpenURL(nURL, icByteArray)
    frmAU.lblStatus.Caption = "Click on Next to update the software.... Or click on cancel to Exit...."
    frmAU.lblNext2.Enabled = True
    frmAU.lblNext2Support.Enabled = True
Else
    frmAU.lblStatus.Caption = nMsg
    Exit Function
End If

ErrHan:
    If Err.Number = "13" Then
        frmAU.lblStart.Enabled = True
        frmAU.lblStartSupport.Enabled = True
        frmAU.lblStatus.Caption = "Unable to connect to the server.... Please make sure you are connected to the internet."
    End If
End Function

Function DownloadUpdate()
nURL = frmAU.net.OpenURL(UpdateURL)
frmAU.lblStatus.Caption = "Downloading update file...."
frmAU.lblCancel2.Enabled = True
frmAU.lblCancel2Support.Enabled = True
frmAU.lblCancel2.Caption = "Exit"
End Function

Function DownloadFile()
Open App.Path + nEXE For Binary Access Write As #1
Put #1, , b()
Close #1
Erase b()

frmAU.lblMessage.Caption = "Please wait... Downloading File... When it completes Downloading the file you will see a button of Exit..."
frmAU.Label12.Visible = False
frmAU.Label13.Visible = False
frmAU.lblUpdate.Enabled = False
frmAU.lblUpdateSupport.Enabled = False
frmAU.lblNewVer.Visible = False
frmAU.lblCancel2.Enabled = False
frmAU.lblCancel2Support.Enabled = False
DownloadUpdate

Exit Function
End Function
