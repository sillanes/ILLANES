<%

'pre_userid = Session("pre_userid")
sSQL = "EXEC wsp_Companies_combo"
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
counter = 0
While NOT dbRS.EOF
	counter = counter + 1
	dbRS.MoveNext
Wend
Set dbRS = Nothing
If counter > 0 Then
	Response.Write "<response_code>1</response_code>"
else
	Response.Write "<response_code>0</response_code>"
End If
%>