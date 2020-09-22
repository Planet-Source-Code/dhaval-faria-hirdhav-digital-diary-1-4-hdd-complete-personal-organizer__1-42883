<%
dim cn
dim ReS

set cn = createobject("ADODB.Connection")
set ReS = createobject("ADODB.Recordset")

set cn = server.CreateObject("ADODB.Connection")
cn.Provider="Microsoft.Jet.OLEDB.4.0"
cn.Open Server.MapPath("WebData.mdb")

set ReS = server.CreateObject("ADODB.Recordset")
ReS.Open "SELECT * from MainData", Cn,1,3

do while not res.eof
	Response.Write "First Name: " & res("FirstName") & "<br>"
	Response.Write "Last Name: " & res("LastName") & "<br>"
	Response.Write "Username: " & res("Username") & "<br>"
	Response.Write "Password: " & res("Password") & "<br>"
	Response.Write "City: " & res("city") & "<br>"
	Response.Write "State: " & res("state") & "<br>"
	Response.Write "Country: " & res("Country") & "<br>"
	Response.Write "E-Mail: " & res("EMail") & "<br>"
	Response.Write "AuthoCode: " & res("AuthoCODE") & "<br>"
	Response.Write "Question: " & res("Question") & "<br>"
	Response.Write "Answer: " & res("Answer") & "<br>"
	Response.Write "<hr>"
	res.movenext
loop

res.Close
cn.Close

set res=nothing
set cn=nothing
%>