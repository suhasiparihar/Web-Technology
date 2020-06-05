<%@Language="VBscript"%>
<html>
<body>
<%
  'dim con,rs,uname,fname,lname,email,pass
  uname=Request.form("username")
  fname=Request.form("firstname")
  lname=Request.form("lastname")
  pass=Request.form("password")
  email=Request.form("email")  
   set con=Server.CreateObject("ADODB.Connection")
  con.Provider="Microsoft.Jet.OLEDB.4.0"
  con.Open"C:/inetpub/wwwroot/project/project.mdb"
  set rs=Server.CreateObject("ADODB.RecordSet")
  rs.Open"userdetails",con,0,3,2
  rs.AddNew
  rs("Username")=uname
  rs("firstname")=fname
  rs("lastname")=lname
  rs("email")=email
  rs("password")=pass
  rs.Update
  rs.Movenext
  rs.close
  con.close
  set con=Nothing
  response.redirect("index.html")
%>
</body>
</html>