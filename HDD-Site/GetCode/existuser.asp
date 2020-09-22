
<%
Dim FullString
FullString = Request.QueryString
FullString = split(FullString, "=")

dim cn
dim ReS

set cn = server.createobject("ADODB.Connection")
cn.Provider="Microsoft.Jet.OLEDB.4.0"
cn.Open Server.MapPath("WebData.mdb")

set ReS = server.CreateObject("ADODB.Recordset")
ReS.Open "SELECT * from MainData", Cn,1,3

do while not ReS.EOF
	if res("Username") = FullString(1) then
		MyAns = true
		exit do
	else
		res.movenext
		MyAns = false
	end if
loop

if MyAns = true then
	Response.Write "ok"
else
	Response.Write "Sorry"
end if

res.Close
cn.Close

set res = nothing
set dn = nothing
%>