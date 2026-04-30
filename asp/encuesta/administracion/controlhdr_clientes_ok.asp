<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Dim hdrid
hdrid = Request.Form("hdrid")
If hdrid = "" Then
    Response.Redirect "controlhdr.asp"
End If

%>
<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Clientes de Hoja de Ruta</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        .resumen-totales {
            margin: 20px 0;
            padding: 15px;
            background-color: #f1f1f1;
            border-radius: 8px;
        }
    </style>
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
		
		<h1 style="margin: 0;">Hoja de Ruta: <%= hdrid %></h1>

		<a href="controlhdr.asp" 
		   class="icon-btn" 
		   title="Volver"
		   style="background-color: #2e30c7; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none; margin-left: 10px;">
			<i class="fas fa-arrow-left"></i> Volver
		</a>

	</div> 

    <table class="table table-striped">
        <thead>
            <tr>
                <th rowspan="2">Cliente</th>
                <th rowspan="2">Facturas</th>
                <th rowspan="2">Total a Cobrar</th>
                <th colspan="4">Total Cobrado</th>
                <th rowspan="2">Descuentos</th>
                <th rowspan="2">Diferencia</th>
                <th rowspan="2">Detalle</th>
            </tr>
            <tr>
 
 
                <th>Total</th>
 
                <th>Efectivo</th>
                <th>Cheque</th>
                <th>Transferencia</th> 
 
            </tr>			
        </thead>
        <tbody>
            <% 
            Dim totalCobrar, totalCobrado
            totalCobrar = 0.00
            totalCobrado = 0.00
            Set rs = conn.Execute("EXEC [dbo].[usp_Transportista_HojaDeRuta_Cabecera_sel] " & hdrid)
            Do Until rs.EOF
                totalCobrar = totalCobrar + FormatNumber(rs("ImporteACobrar"),2)
                totalCobrado = totalCobrado + FormatNumber(rs("TotalCobrado"),2)
				diferencia = FormatNumber(rs("ImporteACobrar"),2) - FormatNumber(rs("TotalCobrado"),2) - FormatNumber(rs("Descuentos"), 2)
				color = rs("color")
            %>
            <tr style="background-color: <%=color%>">
                <td><%= rs("ClienteID") & " - " & rs("ClienteNombre") %></td>
                <td><%= rs("TotalFacturas") %></td>
                <td>$<%= FormatNumber(rs("ImporteACobrar"), 2) %></td>
                <td>$<%= FormatNumber(rs("TotalCobrado"), 2) %></td>
                <td>$<%= FormatNumber(rs("Efectivo"), 2) %></td>
                <td>$<%= FormatNumber(rs("Cheque"), 2) %></td>
                <td>$<%= FormatNumber(rs("Transferencia"), 2) %></td>
                <td>$<%= FormatNumber(rs("Descuentos"), 2) %></td>
                <td>$<%= FormatNumber(diferencia, 2) %></td>
                <td><%=rs("Detalles")%></td>
<!--				
                <td>
                    <form method="post" action="detalle_cliente.asp">
                        <input type="hidden" name="clienteid" value="<%= rs("ClienteID") %>">
                        <input type="hidden" name="hdrid" value="<%= hdrid %>">
                        <input type="submit" value="Ver Detalle">
                    </form>
                </td>
-->
            </tr>
            <% 
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
            %>
        </tbody>
    </table>

    <div class="resumen-totales">
        <strong>Total a Cobrar:</strong> $<%= FormatNumber(totalCobrar, 2) %> <br>
        <strong>Total Cobrado:</strong> $<%= FormatNumber(totalCobrado, 2) %>
    </div>
</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>
</body>
</html>
