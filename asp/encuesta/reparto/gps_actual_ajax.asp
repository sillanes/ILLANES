<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "application/json"

Dim rs, sql
sql = "
SELECT
    T.TransportistaID,
    T.Nombre,
    G.Latitud,
    G.Longitud,
    G.FechaHoraGPS
FROM tracking.Transportista_GPS_Actual G
JOIN Transportistas T ON T.TransportistaID = G.TransportistaID
"

Set rs = conn.Execute(sql)

Response.Write "["
Do While Not rs.EOF
    Response.Write "{""id"":" & rs("TransportistaID") & _
                   ",""nombre"":""" & rs("Nombre") & """" & _
                   ",""lat"":" & Replace(rs("Latitud"),",",".") & _
                   ",""lng"":" & Replace(rs("Longitud"),",",".") & _
                   ",""fecha"":""" & rs("FechaHoraGPS") & """}"
    rs.MoveNext
    If Not rs.EOF Then Response.Write ","
Loop
Response.Write "]"
%>
