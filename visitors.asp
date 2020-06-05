<%@language=Vbscript%>
<html>
<body>
<%
Dim num
Application.lock
num = Application("count")
Application.unlock
if num = "" then
	Application("count") = 1
	num = Application("count")
	response.write("Welcome : Visits - "&num)
else
	Application("count") = num + 1
	num = Application("count")
	response.write("Visits - "&num)
end if
%>
</body>
</html>