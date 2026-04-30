<!--#include file="conexion.asp" -->
<%
If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

Dim hojaID
hojaID = Request.QueryString("hdrid")

If IsNumeric(hojaID) Then
    ' Ejecutar SP para cerrar la hoja de ruta
    Dim sql, cmd
    sql = "EXEC usp_Transportista_HojaDeRuta_Cerrar " & hojaID

    conn.Execute(sql)
End If

' Redirigir de nuevo al home
Response.Redirect "home.asp"
%>
