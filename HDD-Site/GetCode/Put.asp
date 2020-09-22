
<%
dim FullString
dim MyAns
MyAns = "Yes"
FullString = Request.QueryString

FullString = split(FullString,"=")
for i = 0 to ubound(FullString)
	FullString(i) = replace(FullString(i),"%20"," ")
next

dim cn
dim ReS

set cn = server.CreateObject("ADODB.Connection")
cn.Provider="Microsoft.Jet.OLEDB.4.0"
cn.Open Server.MapPath("WebData.mdb")

set ReS = server.CreateObject("ADODB.Recordset")
ReS.Open "SELECT * from MainData", Cn,1,3

do while not ReS.EOF
	if FullString(7) = ReS("UserName") then
		Response.Write "NO"
		MyAns = "NO"
		res.Close
		cn.Close
		set res=nothing
		set cn=nothing
		exit do
	else
		ReS.MoveNext
	end if
loop

if MyAns = "Yes" then
res.AddNew
res("FirstName") = FullString(1)
res("LastName") = FullString(3)
res("Gender") = FullString(5)
res("UserName") = FullString(7)
res("Password") = FullString(9)
res("EMail") = FullString(11)
res("City") = FullString(13)
res("State") = FullString(15)
res("Country") = FullString(17)
res("Question") = FullString(19)
res("Answer") = FullString(21)
res("AuthoCODE") = FullString(23)
res("Version") = FullString(25)
res.update
res.Close
cn.close
Dim objMail
set objMail = CreateObject("CDONTS.NewMail")
objMail.From = FullString(11)
objMail.To = "hddcontact@hirdhav.com"
objMail.Subject = "Request of HDD Autho CODE"
objMail.Body = FullString(1) & " " & FullString(3) & " Needs AuthoCODE of HDD."
objMail.Send
set objMail = nothing
else
set res=nothing
set cn=nothing
end if

set res=nothing
set cn=nothing

%>