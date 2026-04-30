<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function SafeFileName(v)
    Dim s
    s = "" & Nz(v, "")
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    s = Replace(s, " ", "_")
    SafeFileName = s
End Function

Function Html(v)
    Html = Server.HTMLEncode("" & Nz(v, ""))
End Function

Function FormatDateTimeAR(v)
    If IsDate(v) Then
        FormatDateTimeAR = Right("0" & Day(v),2) & "/" & Right("0" & Month(v),2) & "/" & Year(v) & " " & _
                           Right("0" & Hour(v),2) & ":" & Right("0" & Minute(v),2)
    Else
        FormatDateTimeAR = ""
    End If
End Function

Dim cabeceraID, camionID, sql, rs
Dim nombreArchivo, numeroFactura, nroCarga

cabeceraID = CLng(0 & Request.QueryString("CabeceraID"))
camionID   = CLng(0 & Request.QueryString("CamionID"))

If cabeceraID <= 0 Then
    If camionID > 0 Then
        Response.Redirect "camiones_facturas.asp?CamionID=" & camionID & "&err=" & Server.URLEncode("CabeceraID invalido.")
    Else
        Response.Redirect "camiones_facturas.asp?err=" & Server.URLEncode("CabeceraID invalido.")
    End If
End If

sql = "EXEC dbo.usp_Camiones_Cabecera_Items_Exportar_Sel @CabeceraID=" & cabeceraID
Set rs = conn.Execute(sql)

If rs.EOF Then
    If camionID > 0 Then
        Response.Redirect "camiones_facturas.asp?CamionID=" & camionID & "&err=" & Server.URLEncode("No hay datos para exportar.")
    Else
        Response.Redirect "camiones_facturas.asp?err=" & Server.URLEncode("No hay datos para exportar.")
    End If
End If

numeroFactura = SafeFileName(rs("NumeroFactura"))
nroCarga      = SafeFileName(rs("NroCarga"))

If numeroFactura = "" Then numeroFactura = "sin_factura"
If nroCarga = "" Then nroCarga = "sin_carga"

nombreArchivo = "camion_" & nroCarga & "_factura_" & numeroFactura & ".xls"

Response.Clear
Response.Buffer = True
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=" & nombreArchivo
%>
<html>
<head>
<meta charset="utf-8">
<style>
body{font-family:Arial;font-size:12px}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #999;padding:6px;vertical-align:top}
th{background:#e9ecef}
.titulo{font-size:18px;font-weight:bold;margin-bottom:12px}
.sub{margin-bottom:10px}
</style>
</head>
<body>
<div class="titulo">Control de factura</div>
<div class="sub">
    <strong>Proveedor:</strong> <%=Html(rs("ProveedorNombre"))%><br>
    <strong>Factura:</strong> <%=Html(rs("NumeroFactura"))%><br>
    <strong>Remito:</strong> <%=Html(rs("NroRemito"))%><br>
    <strong>Orden de Compra:</strong> <%=Html(rs("OrdenCompra"))%><br>
    <strong>Nro Carga:</strong> <%=Html(rs("NroCarga"))%><br>
</div>

<table>
    <tr>
        <th>ItemID</th>
        <th>EAN</th>
        <th>Articulo</th>
        <th>Descripcion</th>
        <th>Unidad</th>
        <th>Cantidad Factura</th>
        <th>Cantidad Controlada</th>
        <th>Completo</th>
        <th>Observacion</th>
        <th>Usuario</th>
        <th>Fecha Control</th>
    </tr>
    <% Do Until rs.EOF %>
    <tr>
        <td><%=Html(rs("ItemID"))%></td>
        <td style="mso-number-format:'\@';"><%=Html(rs("CodigoEAN"))%></td>
        <td><%=Html(rs("Articulo"))%></td>
        <td><%=Html(rs("Descripcion"))%></td>
        <td><%=Html(rs("UnidadMedida"))%></td>
        <td><%=Html(rs("CantidadFactura"))%></td>
        <td><%=Html(rs("CantidadControlada"))%></td>
        <td><%=Html(rs("Completo"))%></td>
        <td><%=Html(rs("Observacion"))%></td>
        <td><%=Html(rs("Usuario"))%></td>
        <td><%=Html(FormatDateTimeAR(rs("FechaControl")))%></td>
    </tr>
    <%
        rs.MoveNext
       Loop
    %>
</table>
</body>
</html>
<%
If IsObject(rs) Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

conn.Close
Set conn = Nothing
%>