<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001 
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Dim hdrOrigen, hdrDestino, clienteID, sql

hdrOrigen = Request.QueryString("HDROrigen")
hdrDestino = Request.QueryString("HDRDestino")
clienteID = Request.QueryString("ClienteID")

' Validar parámetros
If hdrOrigen = "" Or hdrDestino = "" Or clienteID = "" Then
    Response.Write "Faltan parámetros requeridos."
    Response.End
End If

' Ejecutar stored procedure para reasignar cliente
sql = "EXEC dbo.usp_Transportista_Reasignar_HDR " & hdrOrigen & ", " & hdrDestino & ", " & clienteID
response.write sql
On Error Resume Next
Conn.Execute sql

If Err.Number <> 0 Then
    Response.Write "Error al ejecutar la reasignación: " & Err.Description
    Err.Clear
    Conn.Close
    Set Conn = Nothing
    Response.End
End If

Conn.Close
Set Conn = Nothing

' Redirigir con mensaje de éxito
Response.Redirect "reasignar_cliente_hdr.asp?HojaDeRutaID=" & hdrOrigen & "&ClienteID=" & clienteID &"&msg=reasignado"
%>
