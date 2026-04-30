<!--#include file="conexion.asp" -->
<%
' Proteger la página
If Session("nombre") = "" Then
    Response.Redirect "login.asp"
End If
 
Set rs = conn.Execute("EXEC usp_Dashboard_Resumen")
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Inicio Transportista</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        * { box-sizing: border-box; font-family: Arial, sans-serif; }
        body { margin: 0; background-color: #f4f4f4; }
        .sidebar {
            width: 220px; background-color: #2c3e50; color: white;
            position: fixed; height: 100%; top: 0; left: 0;
            padding-top: 20px; transition: left 0.3s ease;
        }
        .sidebar h2 { text-align: center; margin-bottom: 30px; }
        .sidebar a {
            display: block; color: white; padding: 12px 20px;
            text-decoration: none;
        }
        .sidebar a:hover { background-color: #34495e; }
        .sidebar i { margin-right: 10px; }

        .main-content {
            margin-left: 220px; padding: 20px;
        }

        header {
            background-color: #2c3e50; color: white;
            padding: 10px 20px; display: flex;
            justify-content: space-between; align-items: center;
        }
        header .logout {
            background-color: #e74c3c; color: white; border: none;
            padding: 8px 14px; cursor: pointer;
            border-radius: 4px; font-size: 14px;
        }

        table {
            width: 100%; border-collapse: collapse; background: white;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #ccc; padding: 10px; text-align: left;
        }
        th {
            background-color: #ecf0f1;
        }
        .icon-btn {
            background: none; border: none; cursor: pointer;
            font-size: 18px; margin-right: 8px;
        }
        .icon-btn.entregar { color: green; }
        .icon-btn.rechazar { color: red; }

        .menu-toggle {
            display: none; background: none; border: none;
            color: white; font-size: 22px;
        }

        @media (max-width: 768px) {
            .sidebar {
                left: -220px;
            }
            .sidebar.open {
                left: 0;
            }
            .main-content {
                margin-left: 0; padding: 15px;
            }
            .menu-toggle {
                display: inline;
            }

            table, thead, tbody, th, td, tr {
                display: block;
                width: 100%;
            }

            thead {
                display: none !important;
            }

            tr {
                margin-bottom: 15px;
                background: white;
                border: 1px solid #ccc;
                padding: 10px;
                border-radius: 8px;
            }

            td {
                padding: 10px;
                text-align: right;
                position: relative;
            }

            td::before {
                content: attr(data-label);
                position: absolute;
                left: 10px;
                width: 50%;
                text-align: left;
                font-weight: bold;
            }
        }
    </style>
    <script>
        function toggleSidebar() {
            document.querySelector(".sidebar").classList.toggle("open");
        }
    </script>
</head>
<body>

<!-- Menú lateral -->
<div class="sidebar" id="sidebar">
    <h2>🚚 Transportista</h2>
    <a href="home.asp"><i class="fas fa-home"></i>Inicio</a>
    <a href="hojaderuta.asp"><i class="fas fa-map"></i>Hoja de Ruta</a>  
</div>

<div class="main-content">
    <header>
        <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
        <strong>👤 <%= Session("Nombre") %></strong>
        <form method="post" action="logout.asp" style="margin: 0;">
            <input type="submit" value="Cerrar sesión" class="logout">
        </form>
    </header>
 
    <h2>Resumen de Hojas de Ruta</h2>

    <table>
        <thead>
            <tr>
                <th>Hoja de Ruta</th>
                <th>Clientes</th>
                <th>Facturas</th>
                <th>Entregados</th>
                <th>Acciones</th>
            </tr>
        </thead>
        <tbody>
        <%
        Do Until rs.EOF
        %>
            <tr>
                <td data-label="Hoja de Ruta"><%= rs("HojaDeRutaID") %></td>
                <td data-label="Clientes"><%= rs("TotalClientes") %></td>
                <td data-label="Facturas"><%= rs("TotalFacturas") %></td>
                <td data-label="Entregados"><%= rs("ClientesEntregados") %></td>
                <td data-label="Acciones">
				<a href="hojaderuta.asp?ID=<%= rs("HojaDeRutaID") %>" class="icon-btn entregar" title="Entregar"><i class="fas fa-box"></i></a>
				<a href="rechazo.asp?ID=<%= rs("HojaDeRutaID") %>" class="icon-btn rechazar" title="Rechazar"><i class="fas fa-times"></i></a>
                </td>
            </tr>
        <%
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
        %>
        </tbody>
    </table>
</div>

</body>
</html>
