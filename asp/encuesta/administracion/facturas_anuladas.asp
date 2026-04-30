<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "login.asp"
End If

Dim msg, filtroTransportista, filtroHoja, filtroCliente

msg = Request("msg")
filtroTransportista = Trim(Request("transportista"))
filtroHoja          = Trim(Request("hojaderutaid"))
filtroCliente       = Trim(Request("clienteid"))

' ==========================
'  NORMALIZAR PARÁMETROS
' ==========================
If filtroTransportista = "" Then filtroTransportista = Null
If filtroHoja = "" Then filtroHoja = Null
If filtroCliente = "" Then filtroCliente = Null


' ==================================================
'  LLAMAR AL SP Y TRAER TODAS LAS FACTURAS ANULADAS
' ==================================================
Dim cmd, rs

Set cmd = Server.CreateObject("ADODB.Command")
With cmd
    .ActiveConnection = conn
    .CommandType = 4  'Stored Procedure
    .CommandText = "usp_Transportista_HojaDeRuta_Facturas_Anuladas_Sel"

    ' 3 filtros opcionales
    .Parameters.Append .CreateParameter("@TransportistaID", 3, 1, , filtroTransportista)
    .Parameters.Append .CreateParameter("@HojaDeRutaID", 3, 1, , filtroHoja)
    .Parameters.Append .CreateParameter("@ClienteID", 3, 1, , filtroCliente)
End With

Set rs = cmd.Execute()

%>
<html>
<head>
<meta charset="UTF-8">
<title>Facturas Anuladas</title>
<link rel="stylesheet" href="estilos.css">
<style>
    .filtros {
        background:#f8f9fa;
        padding:15px;
        border:1px solid #ddd;
        margin-bottom:20px;
        border-radius:5px;
    }
</style>
</head>

<body>

<h2>Facturas Anuladas por Hoja de Ruta</h2>

<% If msg <> "" Then %>
    <div style="padding:10px;background:#d4edda;color:#155724;border-radius:5px;margin-bottom:15px;">
        <%=msg%>
    </div>
<% End If %>

<!-- ==================================== -->
<!--        FORMULARIO DE FILTROS        -->
<!-- ==================================== -->

<div class="filtros">
    <form method="GET" action="facturas_anuladas.asp">
        <table cellpadding="5">
            <tr>
                <td><b>Transportista:</b></td>
                <td><input type="text" name="transportista" value="<%=Request("transportista")%>"></td>

                <td><b>Hoja de Ruta:</b></td>
                <td><input type="text" name="hojaderutaid" value="<%=Request("hojaderutaid")%>"></td>

                <td><b>Cliente:</b></td>
                <td><input type="text" name="clienteid" value="<%=Request("clienteid")%>"></td>

                <td>
                    <button type="submit" style="padding:6px 12px;background:#007bff;color:white;border:none;border-radius:4px;">
                        Filtrar
                    </button>
                </td>
            </tr>
        </table>
    </form>
</div>

<%
' ==================================================
'   CORTE DE CONTROL POR HojaDeRutaID
' ==================================================

Dim hojaActual
hojaActual = ""

If Not rs.EOF Then

    Do While Not rs.EOF

        ' Cambio de hoja → mostrar encabezado
        If hojaActual <> rs("HojaDeRutaID") Then
            
            ' Cerrar tabla anterior (si existía)
            If hojaActual <> "" Then
                Response.Write "</table><br>"
            End If

            hojaActual = rs("HojaDeRutaID")

            ' Encabezado de hoja
%>

            <h3 style="background:#f5f5f5;padding:10px;border-left:4px solid #007bff;">
                Hoja de Ruta <%=hojaActual%>
            </h3>

            <table border="1" cellpadding="6" cellspacing="0" width="100%">
                <tr style="background:#e9ecef;">
                    <th>Factura</th>
                    <th>Importe</th>
                    <th>Cliente</th>
                    <th>Estado</th>
                    <th>Acción</th>
                </tr>

<%
        End If  ' FIN corte de control

        ' ==========================
        '   FILA DE FACTURA
        ' ==========================
%>

        <tr>
            <td><%=rs("FacturaID")%></td>
            <td>$ <%=FormatNumber(rs("ImporteFactura"),2)%></td>
            <td><%=rs("Cliente")%></td>

            <td>
                <% If rs("Validado") = 1 Then %>
                    ✔ Verificado
                <% Else %>
                    ❌ Sin verificar
                <% End If %>
            </td>

            <td>
                <% If rs("Validado") = 0 Then %>
                    <a href="verificar_factura.asp?facturaid=<%=rs("FacturaID")%>&hojaderutaid=<%=hojaActual%>"
                       style="background:#007bff;color:white;padding:5px 10px;text-decoration:none;border-radius:4px;">
                       Verificar
                    </a>
                <% Else %>
                    —
                <% End If %>
            </td>
        </tr>

<%
        rs.MoveNext
    Loop

    ' cerrar última tabla
    Response.Write "</table>"

Else
%>
    <p>No hay facturas anuladas registradas.</p>
<%
End If

rs.Close : Set rs = Nothing
Set cmd = Nothing
%>

</body>
</html>
