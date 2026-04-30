<%
sSQL = "EXEC usp_CountriesPerProduct_Sel " & CLng("0"&Company) & "," & CLng("0"&Product)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
comboList = ""
While NOT dbRS.EOF
	comboList = comboList & "<option value='" & dbRS("CountryID") & "'>" & dbRS("Description") & "</option>"
	dbRS.MoveNext
Wend
Set dbRS = Nothing
Response.Write comboList
%>