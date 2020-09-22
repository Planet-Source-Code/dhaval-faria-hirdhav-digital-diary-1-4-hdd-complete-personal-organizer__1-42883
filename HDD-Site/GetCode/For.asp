
<%
dim FullString
dim UserName
dim Password
dim EMail
dim HQ
dim HA
dim MyAnswer

FullString = Request.QueryString

FullString = split(FullString,"=")
for i = 0 to ubound(FullString)
	FullString(i) = replace(FullString(i),"%20"," ")
next

UserName = FullString(1)
Password = FullString(3)
EMail = FullString(5)
HQ = FullString(7)
HA = FullString(9)

dim cn
dim ReS

set cn = server.CreateObject("ADODB.Connection")
cn.Provider="Microsoft.Jet.OLEDB.4.0"
cn.Open Server.MapPath("WebData.mdb")

set ReS = server.CreateObject("ADODB.Recordset")
ReS.Open "SELECT * from MainData", cn,1,3

do while not ReS.EOF
	if UserName & Password & EMail & HQ & HA = ReS("UserName") & ReS("Password") & Res("EMail") & Res("Question") & ReS("Answer") then
		Response.Write "Name:" & res("FirstName") & " " & res("LastName")
		Response.Write ":CODE:" & res("AuthoCODE")
		MyAnswer=True
		dim objMail
		set objMail = createobject("CDONTS.NewMail")
		objMail.To = "hddcontact@hirdhav.com"
		objMail.From = "forgot@hirdhav.com"
		objMail.Subject = "Some one Forgot HDD Autho CODE."
		objMail.Body = res("FirstName") & " " & res("LastName") & " Forgot Autho CODE."
		objMail.send
		set objMail = nothing
		exit do
	else
		MyAnswer = false
		ReS.MoveNext
	end if
loop

if MyAnswer = false then
	Response.Write "Sorry"
end if

res.Close
cn.Close

set res=nothing
set cn=nothing
%>