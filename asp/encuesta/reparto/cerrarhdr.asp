<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

Dim hojaRutaID
hojaRutaID = Request.QueryString("hdrid")
'DNI = Request.QueryString("dni")

If hojaRutaID = "" Or Not IsNumeric(hojaRutaID) Then
    Response.Write "<p style='color:red;'>ID de hoja de ruta inválido.</p>"
    Response.End
End If

On Error Resume Next

' Ejecutar stored procedure para cerrar hoja de ruta'
'conn.Execute "EXEC usp_Transportista_HojaDeRuta_Cerrar " & hojaRutaID & ",'" &DNI&"'"
 
'response.end
sSQL = "EXEC usp_Transportista_HojaDeRuta_Cerrar " & hojaRutaID 
Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, conn
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	Response.Redirect "cerrarhdr_error.asp?mensaje="
End If
 
hasError= dbRS("hasError")
ErrorMessage= dbRS("ErrorMessage") 

Set dbRS = Nothing 
If hasError = 1 Then
    Dim mensajeError
    mensajeError = Server.URLEncode(ErrorMessage)
    Response.Redirect "cerrarhdr_error.asp?mensaje=" & mensajeError
End If

Response.Redirect "hojaderutav3.asp?msg=ok"

conn.Close
Set conn = Nothing
%>
