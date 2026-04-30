<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel Transportista</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
	    .detalle-toggle {
            cursor: pointer;
            color: #007bff;
            text-decoration: underline;
        }
        .detalle-content {
            display: none;
            margin-top: 5px;
            white-space: pre-wrap;
            font-size: 0.9em;
        }
		
        .tooltip {
            position: relative;
            display: inline-block;
        }

        .tooltip .tooltiptext {
            visibility: hidden;
            width: 160px;
            background-color: #555;
            color: #fff;
            text-align: center;
            padding: 5px 0;
            border-radius: 6px;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            margin-left: -80px;
            opacity: 0;
            transition: opacity 0.3s;
        }

        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }
    </style>
    <script>
        function solicitarReenvio(hdrid, clienteid) {
            var telefono = prompt("Ingrese el número de teléfono para reenviar:");
            if (telefono && telefono.trim() !== "") {
                window.location.href = "reenvio_guardar.asp?HojaDeRutaID=" + hdrid + "&ClienteID=" + clienteid + "&Telefono=" + encodeURIComponent(telefono);
            }
        }

        function solicitarDNIyCerrar(hdrid) {
            var dni = prompt("Por favor, ingrese el DNI de la persona que recibe:");
            if (dni && dni.trim() !== "" && /^\d{7,10}$/.test(dni.trim())) {
                window.location.href = "cerrarhdr.asp?hdrid=" + hdrid + "&DNI=" + encodeURIComponent(dni);
            } else {
                alert("Debe ingresar un DNI válido (solo números, mínimo 7 dígitos).");
            }
        }
        function toggleDetalle(id) {
            var div = document.getElementById(id);
            div.style.display = div.style.display === 'none' || div.style.display === '' ? 'block' : 'none';
        }
		function mostrarClienteNombre(nombre) {
			alert("Cliente: " + nombre);
		}		
    </script>
</head>
<body>

<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<div class="main-content">
<%
If Request("msg") = "ok" Then
%>
    <div style="background-color: #d4edda; color: #155724; padding: 10px; margin-bottom: 10px; border: 1px solid #c3e6cb;">
        ✅ Hoja de ruta cerrada correctamente.
    </div>
<%
ElseIf Request("msg") = "reenvio_ok" Then
%>
    <div style="background-color: #d1ecf1; color: #0c5460; padding: 10px; margin-bottom: 10px; border: 1px solid #bee5eb;">
        📩 Reenvío registrado con éxito.
    </div>
<%

End If
 
Dim hojaRutaID
hojaRutaID = Request.QueryString("hdrid")

If hojaRutaID = "" Then
    ' Mostrar listado de hojas de ruta abiertas
    Set rs = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Abiertas")

%>
<h2>Hojas de Ruta Abiertas</h2>

<table class="table table-striped table-bordered align-middle">
    <thead class="table-light">
        <tr>
            <th>HDR</th>
            <th>Clientes</th> 
            <th>Pendintes</th>
            <th>Acciones</th>
        </tr>
    </thead>
    <tbody>
    <%
    Do Until rs.EOF
    %>
        <tr>
            <td><%= rs("HojaDeRutaID") %></td>
            <td><%= rs("TotalClientes") %></td> 
            <td><%= rs("ClientesPendientes") %></td>
            <td>
                <a href="hojaderutaV3.asp?hdrid=<%= rs("HojaDeRutaID") %>" title="Seleccionar hoja de ruta"> <i class="fa-solid fa-truck-fast"></i></a>
                <a href="#" title="Cerrar Hoja de Ruta" onclick="solicitarDNIyCerrar('<%= rs("HojaDeRutaID") %>'); return false;"><i class="fa-solid fa-lock-open"></i></a>
            </td>
        </tr>
    <%
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    %>
    </tbody>
</table>

<%
Else
    ' Mostrar detalle de hoja de ruta seleccionada
Set rs = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Cabecera_Sel " & hojaRutaID)
%>
 
<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 20px;">
    
    <h1 style="margin: 0;">Hoja de Ruta: <%= hojaRutaID %></h1>

    <a href="hojaderuta.asp" 
       class="icon-btn" 
       title="Volver"
       style="background-color: #2e30c7; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none; margin-left: 10px;">
        <i class="fas fa-arrow-left"></i> Volver
    </a>

    <a href="#" 
       class="icon-btn" 
       title="Cerrar hoja de ruta"
       onclick="solicitarDNIyCerrar('<%= hojaRutaID %>'); return false;"
       style="margin-left: left; background-color: #27ae60; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
        <i class="fas fa-lock"></i> Cerrar hoja de ruta
    </a>
</div> 
<div style="overflow-x: auto;">

<table class="table table-striped table-bordered align-middle">
    <thead class="table-light">
 
        <tr>
            <th>Cliente</th> 
            <th>Importe a Cobrar</th> 
            <th>Importe Cobrado</th> 
            <th>Acciones</th>
        </tr>		
    </thead>
    <tbody>
<%
Dim totalACobrar, totalCobrado
totalACobrar = 0.000
totalCobrado= 0.00


Do Until rs.EOF
    Dim estado
    estado = LCase(rs("estado"))

    Dim icono, colorEstado, tooltip
    Select Case estado
        Case 1
            icono = "<i class='fas fa-check-circle' style='color:green;'></i>"
            colorEstado = "green"
            tooltip = rs("EstadoEntrega")
        Case 2
            icono = "<i class='fas fa-exclamation-circle' style='color:red;'></i>"
            colorEstado = "red"
            tooltip = rs("EstadoEntrega")
        Case 3
            icono = "<i class='fas fa-times-circle' style='color:red;'></i>"
            colorEstado = "red"
            tooltip = rs("EstadoEntrega")
        Case Else
            icono = "<i class='fas fa-hourglass-half' style='color:gray;'></i>"
            colorEstado = "gray"
            tooltip = rs("EstadoEntrega")
    End Select
	totalACobrar = totalACobrar + FormatNumber(rs("ImporteACobrar"), 2)
	totalCobrado = totalCobrado + FormatNumber(rs("TotalCobrado"), 2)

	 
%>
        <tr>
            <td> 
				<span class="detalle-toggle" onclick="mostrarClienteNombre('<%= Replace(rs("ClienteNombre"), "'", "\'") %>')"><%= rs("ClienteID") %></span> 
			</td> 
            <td>$<%= FormatNumber(rs("ImporteACobrar"), 2) %></td>  
            <td>$<%= FormatNumber(rs("TotalCobrado"), 2) %></td>  

            <td>
			  <a href="#" class="icon-btn reenviar" title="Reenviar al cliente" onclick="solicitarReenvio('<%= hojaRutaID %>', '<%= rs("ClienteID") %>'); return false;"><i class="fas fa-share"></i></a>
                 
                <% If estado = "0" Then %>
				 <a href="entregar.asp?clienteid=<%= rs("ClienteID") %>&hdrid=<%= hojaRutaID %>" 
                       class="icon-btn entregar" 
                       title="Marcar como entregado"
                       onclick="return confirm('¿Confirmar entrega al cliente?')">
                        <i class="fa-solid fa-box-open"></i></a>

					       <a href="rechazar.asp?clienteid=<%= rs("ClienteID") %>&hdrid=<%= hojaRutaID %>" 
                       class="icon-btn rechazar" 
                       title="Rechazar entrega">
                        <i class="fas fa-ban"></i> 
                    </a>
					
                <% Else %>
                    <span class="tooltip"> 
                        <%=icono%><em style="color: <%= colorEstado %>;"> <%= tooltip %></em>
                        <% If tooltip <> "" Then %>
                        <span class="tooltiptext"><%= tooltip %></span>
                        <% End If %>
                    </span>
                <% End If %>
            </td>
        </tr>
		

<%
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>

<tfoot>
    <tr>
        <th>Total</th>
        <th>$<%= FormatNumber(totalACobrar, 2) %></th>
        <th>$<%= FormatNumber(totalCobrado, 2) %></th>
        <th></th>
    </tr>
</tfoot>


    </tbody>
</table>
</div>
<%
End If

conn.Close
Set conn = Nothing
%>
</div>

<script>
    function toggleSidebar() {
        document.querySelector('.sidebar').classList.toggle('open');
    }
</script>

</body>
</html>
