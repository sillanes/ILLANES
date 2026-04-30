<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.ContentType = "application/json"
Response.CharSet = "UTF-8"
Response.Buffer = True
Response.Clear

Dim hojaderutaid
hojaderutaid = CLng(Request("hojaderutaid"))

If hojaderutaid = 0 Then
    Response.Write "[]"
    Response.End
End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "tracking.usp_GPS_Recorrido_Agrupado"

cmd.Parameters.Append cmd.CreateParameter("@HojaDeRutaID", 3, 1, , hojaderutaid)

Set rs = cmd.Execute

Dim first : first = True
Response.Write "["

Do Until rs.EOF

    If Not first Then Response.Write ","
    first = False

    Response.Write "{"
    Response.Write """orden"":" & rs("Orden") & ","
    Response.Write """grupo"":" & rs("GrupoID") & ","
    Response.Write """puntos"":" & rs("CantidadPuntos") & ","
    Response.Write """lat"":" & Replace(rs("Latitud"), ",", ".") & ","
    Response.Write """lon"":" & Replace(rs("Longitud"), ",", ".") & ","
    Response.Write """desde"":""" & rs("Inicio") & ""","
    Response.Write """hasta"":""" & rs("Fin") & ""","
    Response.Write """minutos"":" & rs("DuracionMin")
    Response.Write "}"

    rs.MoveNext
Loop

Response.Write "]"

rs.Close
Set rs = Nothing
Set cmd = Nothing
Response.End
%>
