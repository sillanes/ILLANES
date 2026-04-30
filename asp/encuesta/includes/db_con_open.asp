<%
'Set dbCon = Nothing
'Response.Redirect "dummy.asp?code=1"
dbCon.Open "DRIVER=SQL Server;SERVER=SQLREPORTING01;APP=Web;DATABASE=Web", "webuser", "wwwpass"
%>
