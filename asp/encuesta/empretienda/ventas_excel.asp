<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

' Evitar cache
Response.Expires = 0
Response.Buffer = True
Response.Clear

' ContentType para Excel
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=DetalleVentas.xls"

Dim FechaInicio, FechaFin, orderBy, orderDir

If Request("inicio") <> "" Then
    FechaInicio = CDate(Request("inicio"))
Else
    FechaInicio = DateAdd("m",-1,Date())
End If

If Request("fin") <> "" Then
    FechaFin = CDate(Request("fin"))
Else
    FechaFin = Date()
End If

orderBy = Request("order")
If orderBy = "" Then orderBy = "Cliente"

orderDir = UCase(Request("dir"))
If orderDir <> "DESC" Then orderDir = "ASC"

Function toISO(fecha)
    Dim d
    d = CDate(fecha)
    toISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function

Dim rs, sql
sql = "EXEC report.usp_Ordenes_Ventas_Detalle_sel '" & toISO(FechaInicio) & "','" & toISO(FechaFin) & "','" & orderBy & "','" & orderDir & "'"
Set rs = conn.Execute(sql)

' Exportar como tabla HTML (Excel lo abre directo)
Response.Write "<table border='1'>"
Response.Write "<tr>"
For i=0 To rs.Fields.Count-1
    Response.Write "<th>" & rs.Fields(i).Name & "</th>"
Next
Response.Write "</tr>"

Do Until rs.EOF
    Response.Write "<tr>"
    For i=0 To rs.Fields.Count-1
        Response.Write "<td>" & rs.Fields(i).Value & "</td>"
    Next
    Response.Write "</tr>"
    rs.MoveNext
Loop

Response.Write "</table>"

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

Response.End
%>
