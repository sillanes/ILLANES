<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then Response.Redirect "/login.asp"

Dim HDR, ClienteID, nuevoEstado, observaciones, sql
HDR = Trim(Request.Form("hdr"))
ClienteID = Trim(Request.Form("clienteid"))
nuevoEstado = Replace(Trim(Request.Form("nuevoEstado")), "'", "''")
observaciones = Replace(Trim(Request.Form("observaciones")), "'", "''")

On Error Resume Next ' Inicia manejo de errores

Dim mensajeError, hasError

hasError = 1
' Ejecutar el nuevo stored procedure que incluye observaciones
sql = "EXEC [usp_Transportista_HojaDeRuta_Cliente_Estado_upd] " & _
      HDR & ", " & ClienteID & ", '" & nuevoEstado & "', '" & observaciones & "'"

response.write sql
conn.Execute sql
Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    ' Error de ejecución de SQL (por ejemplo, falta de permisos)
    mensajeError = "Ocurrió un error al intentar actualizar el estado : " & Err.Description
    Err.Clear
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    Response.Redirect "cambiarestadocliente_error.asp?msg=" & Server.URLEncode(mensajeError)
ElseIf Not rs Is Nothing And Not rs.EOF Then
    ' El stored procedure devolvió un recordset con un posible error lógico
    hasError = rs("hasError")
    mensajeError = rs("ErrorMessage")
    rs.Close
    Set rs = Nothing

 
    If hasError = 1 Then
        Response.Redirect "cambiarestadocliente_error.asp?msg=" & Server.URLEncode(mensajeError)
    Else
        Response.Redirect "cambiarestadocliente.asp?hdr=" & HDR & "&clienteid="  & ClienteID   & "&msg=ok"
    End If
Else
    ' No devolvió recordset, asumimos éxito
    Response.Redirect "cambiarestadocliente.asp?hdr=" & HDR  & "&clienteid="  & ClienteID & "&msg=ok"
End If
 
 


%>