<!-- home.asp -->
<!--#include file="conexion.asp" -->
<!--#include file="estilos.asp" -->
<%
If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If


' Ejecutar SP para hojas abiertas
Set rsAbiertas = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Abiertas")
 

%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <title>Panel Transportista</title>
 
    
</head>
<body>

<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex: 1;">👤 <%= Server.HTMLEncode(Session("NombreTransportista")) %></strong>
    <form method="post" action="logout.asp" style="margin: 0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>

<!--#include file="sidebar.asp" -->

<div class="main-content">

<h2>Hojas de Ruta Abiertas</h2>
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
                <a href="hojaderuta.asp?ID=<%= rsAbiertas("HojaDeRutaID") %>" class="fa-truck-loading" title="Trabajar Hoja de Ruta">
                    <i class="fas fa-truck-loading"></i>
                </a>
                <a href="cerrar.asp?hdridid=<%= rsAbiertas("HojaDeRutaID") %>" class="icon-btn cerrar" title="Cerrar Hoja de Ruta">
                    <i class="fas fa-circle-check"></i>
                </a>
            </td>
        </tr>
        <% rsAbiertas.MoveNext : Loop : End If %>
    </table>
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

