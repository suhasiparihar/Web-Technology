<%@language=Vbscript%>

<%
Dim num,num1
response.cookies("count").Expires = date + 100
num1 = request.cookies("count")
if num1 = "" then
	response.cookies("count") = 1
	num1 = request.cookies("count")
else
	response.cookies("count") = num1 + 1
	num1 = request.cookies("count")
end if
Application.lock
num = Application("count")
Application.unlock
if num = "" then
	Application("count") = 1
	num = Application("count")
else
	Application("count") = num + 1
	num = Application("count")
end if
set conn = Server.CreateObject("ADODB.Connection")
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.Open "C:/inetpub/wwwroot/project/project.mdb"

set rs = Server.CreateObject("ADODB.RecordSet")
u = Request.form("username")
v = Request.form("password")
sql = "select Username,password from userdetails where Username='"&u&"' AND password='"&v&"'"
rs.Open sql,conn

if rs.EOF=True then
	server.transfer("loginfail.html")
else
	session("username")=u
	response.redirect("dashboard/index.html")
end if
rs.close
conn.close
%>