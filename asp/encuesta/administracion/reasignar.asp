<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Dim fecha
fecha = Request("fecha")
 
If fecha = "" Then
    ' Fecha actual en formato yyyy-mm-dd
    fecha = Year(Date) & "-" & Right("0" & Month(Date), 2) & "-" & Right("0" & Day(Date), 2)


' Convertimos la fecha a tipo Date para poder operar

fechaDate = CDate(fecha)

' Restamos 3 días (una semana)
fechaDate = DateAdd("d", -3, fechaDate)

' Volvemos a convertir a string yyyy-mm-dd
fecha = Year(fechaDate) & "-" & Right("0" & Month(fechaDate), 2) & "-" & Right("0" & Day(fechaDate), 2)
End If
' Si no se seleccionó fecha, mostramos solo el calendario
If fecha <> "" Then
    ' Obtener hojas de ruta pendientes filtradas por fecha
    Dim sql, rs
    sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cerradas_Pendientes '" & fecha & "'"
	'response.write sql
    Set rs = Conn.Execute(sql)
End If
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reasignar Cliente</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 10px; border: 1px solid #ccc; text-align: left; }
        .action-icon { color: #007bff; cursor: pointer; }
        .group-header { background-color: #f2f2f2; font-weight: bold; }
        form { display: flex; align-items: center; gap: 10px; }
    </style>
    <script>
        function reasignarCliente(hdrid, clienteid, fecha) {
            window.location.href = "reasignar_cliente_hdr.asp?HojaDeRutaID=" + hdrid + "&ClienteID=" + clienteid + "&fecha=" + fecha ;
        }
        function toggleSidebar() {
            document.querySelector('.sidebar').classList.toggle('open');
        }
    </script>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("currentUser")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="main-content">
	<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
		
		<h2>Hojas de Ruta Pendientes</h2>
   
		
		<a href="/menu.asp" 
		   class="icon-btn" 
		   title="Cerrar hoja de ruta"
		   style="margin-left: left; background-color: #27ae60; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
			<i class="fas fa-arrow-left"></i> Menu
		</a>
	</div> 

    <!-- Calendario para seleccionar fecha -->
    <div style="margin-bottom: 20px;">
        <form method="get" action="">
            <label for="fecha">Seleccionar fecha:</label>
            <input type="date" id="fecha" name="fecha" value="<%= fecha %>">
            <input type="submit" value="Filtrar">
        </form>
    </div>

<% If fecha = "" Then %>
    <p>Por favor seleccione una fecha para mostrar las hojas de ruta.</p>
<% Else %>
    <%
    ' Variables para subtotales
    Dim lastHDR, lastTransportista
    lastHDR = ""
    lastTransportista = ""

    Dim totalFactura, totalCobrar, totalFacturas, reasignaciones

    Do While Not rs.EOF
        If rs("HojaDeRutaID") <> lastHDR Or rs("Transportista") <> lastTransportista Then
            ' Mostrar subtotales si no es el primer grupo
            If lastHDR <> "" Then
    %>
        <tr class="group-header">
            <td colspan="2"><strong>Totales</strong></td>
            <td><strong>$ <%= FormatNumber(totalFactura, 2) %></strong></td>
            <td><strong>$ <%= FormatNumber(totalCobrar, 2) %></strong></td>
            <td><strong><%= totalFacturas %></strong></td>
            <td><strong><%= reasignaciones %></strong></td>
            <td></td>
        </tr>
    <%
            End If

            ' Cerrar tabla anterior
            If lastHDR <> "" Then Response.Write "</tbody></table>"

            ' Nueva tabla
    %>
    <table>
        <thead>
            <tr class="group-header">
                <td colspan="7">Fecha: <%=FormatDateTime(rs("Fecha"),2)%> -- Hoja de Ruta: <%= rs("HojaDeRutaID") %>  -- Transportista: <%= rs("Transportista") %>   </td>
            </tr>
            <tr>
                <th>Cliente</th>
                <th>Forma de Pago</th>
                <th>Importe Factura</th>
                <th>Importe a Cobrar</th>
                <th>Total Facturas</th>
                <th>Reasignaciones</th>
                <th>Acción</th>
            </tr>
        </thead>
        <tbody>
    <%
            totalFactura = 0
            totalCobrar = 0
            totalFacturas = 0
            reasignaciones = 0
            lastHDR = rs("HojaDeRutaID")
            lastTransportista = rs("Transportista")
        End If

        If Not IsNull(rs("ImporteFactura")) Then totalFactura = totalFactura + CDbl(rs("ImporteFactura"))
        If Not IsNull(rs("ImporteAPagar")) Then totalCobrar = totalCobrar + CDbl(rs("ImporteAPagar"))
        If Not IsNull(rs("TotalFacturas")) Then totalFacturas = totalFacturas + CLng(rs("TotalFacturas"))
        If Not IsNull(rs("Reasignaciones")) Then reasignaciones = reasignaciones + CLng(rs("Reasignaciones"))
    %>
        <tr>
            <td><%= rs("ClienteID") & " - " & rs("ClienteNombre") %></td>
            <td><%= rs("FormaPago") %></td>
            <td>$ <%= FormatNumber(rs("ImporteFactura"), 2) %></td>
            <td>$ <%= FormatNumber(rs("ImporteAPagar"), 2) %></td>
            <td><%= rs("TotalFacturas") %></td>
            <td><%= rs("Reasignaciones") %></td>
            <td>
                <% If rs("Reasignaciones") < rs("ReasignacionesAllowed") Then %>
                    <i class="fas fa-random action-icon" title="Reasignar" onclick="reasignarCliente(<%= rs("HojaDeRutaID") %>, <%= rs("ClienteID") %>,'<%=Fecha%>')"></i>
                <% End If %>
            </td>
        </tr>
    <%
        rs.MoveNext
    Loop

    ' Mostrar subtotales del último grupo
    If lastHDR <> "" Then
    %>
        <tr class="group-header">
            <td colspan="2"><strong>Totales</strong></td>
            <td><strong>$ <%= FormatNumber(totalFactura, 2) %></strong></td>
            <td><strong>$ <%= FormatNumber(totalCobrar, 2) %></strong></td>
            <td><strong><%= totalFacturas %></strong></td>
            <td><strong><%= reasignaciones %></strong></td>
            <td></td>
        </tr>
    </tbody>
    </table>
    <%
    End If
    rs.Close
    Set rs = Nothing
End If
%>
</div>

</body>
</html>
