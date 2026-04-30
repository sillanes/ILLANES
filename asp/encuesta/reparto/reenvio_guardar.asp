<!--#include file="conexion.asp" -->
<%
Dim hdrid, clienteid, telefono
hdrid = Request.QueryString("HojaDeRutaID")
clienteid = Request.QueryString("ClienteID")
telefono = Request.QueryString("Telefono")

If IsNumeric(hdrid) And IsNumeric(clienteid) And telefono <> "" Then
    Dim sql
    sql = "EXEC usp_HojaDeRuta_Reenvio_ins " & hdrid & ", " & clienteid & ", '" & Replace(telefono, "'", "''") & "'"
    On Error Resume Next
    conn.Execute(sql)
    If Err.Number <> 0 Then
        Response.Write "❌ Error al guardar el reenvío: " & Err.Description
        Response.End
    End If
End If

Response.Redirect "hojaderutaV3.asp?hdrid=" & hdrid & "&msg=reenvio_ok"
%>
