<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>
<!--#include file="sidebar.asp" -->
<%
If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If

response.redirect "validarpendientes.asp"
Set rsAbiertas = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Abiertas 0, " & Session("TransportistaID"))
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel Logistica</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
	
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
    <h2>Hojas de Ruta Abiertas</h2>
	<div class="cards-container">
    <div class="table-responsive">
        <table>
            <tr>
                <th>Hoja de Ruta</th>
                <th>Clientes</th>
                <th>Facturas</th>
                <th>Acciones</th>
            </tr>
            <% If rsAbiertas.EOF Then %>
                <tr><td colspan="4" style="text-align:center;">No hay hojas de ruta abiertas.</td></tr>
            <% Else
                Do Until rsAbiertas.EOF %>
            <tr>
                <td><%= rsAbiertas("HojaDeRutaID") %></td>
                <td><%= rsAbiertas("TotalClientes") %></td>
                <td><%= rsAbiertas("TotalFacturas") %></td>
                <td>
                    <a href="hojaderutaV3.asp?hdrid=<%= rsAbiertas("HojaDeRutaID") %>" title="Trabajar Hoja de Ruta"><i class="fa-solid fa-truck-fast"></i></a>
 
                    <a href="cerrar.asp?hdrid=<%= rsAbiertas("HojaDeRutaID") %>" title="Cerrar Hoja de Ruta"><i class="fa-solid fa-lock-open"></i></a>
                </td>
            </tr>
            <% rsAbiertas.MoveNext : Loop : End If %>
        </table>
    </div>
	</div>
</div>

<script>
function toggleSidebar() {
    document.querySelector('.sidebar').classList.toggle('open');
}
</script>

</body>
</html>

<%
rsAbiertas.Close
Set rsAbiertas = Nothing
conn.Close
Set conn = Nothing
%>