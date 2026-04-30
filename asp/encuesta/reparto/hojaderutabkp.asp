<!--#include file="conexion.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <title>Panel Transportista</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    
    <style>
        * { box-sizing: border-box; font-family: Arial, sans-serif; }
        body { margin: 0; background-color: #f4f4f4; }
        
        .sidebar {
            width: 220px;
            background-color: #2c3e50;
            color: white;
            position: fixed;
            height: 100%;
            top: 0;
            left: 0;
            padding-top: 60px; /* espacio para header fijo */
            transition: left 0.3s ease;
            overflow-y: auto;
            z-index: 1001;
        }
        .sidebar h2 {
            text-align: center;
            margin-bottom: 30px;
        }
        .sidebar a {
            display: block;
            color: white;
            padding: 12px 20px;
            text-decoration: none;
        }
        .sidebar a:hover {
            background-color: #34495e;
        }
        .sidebar i {
            margin-right: 10px;
        }

        header {
            background-color: #2c3e50;
            color: white;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: fixed;
            width: 100%;
            top: 0;
            left: 0;
            height: 60px;
            z-index: 1100;
        }
        .menu-toggle {
            display: none;
            background: none;
            border: none;
            color: white;
            font-size: 22px;
            cursor: pointer;
        }
        .logout {
            background-color: #e74c3c;
            color: white;
            border: none;
            padding: 8px 14px;
            cursor: pointer;
            border-radius: 4px;
            font-size: 14px;
        }

        .main-content {
            margin-left: 220px;
            padding: 80px 20px 40px 20px; /* padding top para header fijo */
            transition: margin-left 0.3s ease;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #ccc;
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #ecf0f1;
        }
        .icon-btn {
            background: none;
            border: none;
            cursor: pointer;
            font-size: 18px;
            margin-right: 8px;
            color: inherit;
            text-decoration: none;
        }
        .icon-btn.entregar { color: green; }
        .icon-btn.rechazar { color: red; }

        @media (max-width: 768px) {
            .sidebar {
                left: -220px; /* oculto por defecto */
            }
            .sidebar.open {
                left: 0;
            }
            .menu-toggle {
                display: inline;
            }
            .main-content {
                margin-left: 0 !important;
                padding: 80px 15px 40px 15px;
                overflow-x: auto; /* scroll horizontal en móvil si tabla ancha */
            }
            table {
                font-size: 14px;
                min-width: 600px; /* fuerza scroll horizontal si hace falta */
            }
        }
		
/* Contenedor con scroll horizontal para tabla */
.table-responsive {
    width: 100%;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch; /* para smooth scrolling en iOS */
}

/* Estilo base tabla (ya tienes) */
table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    margin-top: 20px;
    min-width: 600px; /* mínimo para que las columnas no queden muy chicas */
}

/* Ajustes móviles */
@media (max-width: 768px) {
    table {
        font-size: 14px;
        min-width: 0; /* para permitir que se reduzca */
    }
    th, td {
        padding: 8px 6px;
    }
}
		
    </style>
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
<%
Dim hojaRutaID
hojaRutaID = Request.QueryString("hdrid")

If hojaRutaID = "" Then
    ' Mostrar listado de hojas de ruta abiertas
    Set rs = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Abiertas")

%>
<h1>Hojas de Ruta Abiertas</h1>

<table class="table table-striped table-bordered align-middle">
    <thead class="table-light">
        <tr>
            <th>Hoja de Ruta</th>
            <th>Clientes</th>
            <th>Facturas</th>
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
            <td><%= rs("TotalFacturas") %></td>
            <td>
                <a href="hojaderuta.asp?hdrid=<%= rs("HojaDeRutaID") %>" class="icon-btn entregar" title="Seleccionar hoja de ruta">
                    <i class="fas fa-arrow-right"></i>
                </a>
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
    <a href="hojaderuta.asp" class="icon-btn" title="Volver" style="font-size: 18px; color: #2c3e50;">
        <i class="fas fa-arrow-left"></i> Volver
    </a>
    
    <h1 style="margin: 0;">Hoja de Ruta ID: <%= hojaRutaID %></h1>

    <a href="cerrarhdr.asp?hdrid=<%= hojaRutaID %>" 
       class="icon-btn" 
       title="Cerrar hoja de ruta"
       onclick="return confirm('¿Estás seguro de cerrar esta hoja de ruta?')"
       style="margin-left: auto; background-color: #27ae60; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
        <i class="fas fa-lock"></i> Cerrar hoja de ruta
    </a>
</div>

<table class="table table-striped table-bordered align-middle">
    <thead class="table-light">
        <tr>
            <th>Cliente</th>
            <th>Total Facturas</th>
            <th>Importe a Cobrar</th>
            <th>Estado</th>
            <th>Acciones</th>
        </tr>
    </thead>
    <tbody>
<%
Do Until rs.EOF
    Dim estadoEntrega
    estadoEntrega = rs("EstadoEntrega") ' Suponiendo que viene en el recordset

    Dim icono, colorEstado
    Select Case LCase(estadoEntrega)
        Case "entregado"
            icono = "<i class='fas fa-check-circle' style='color:green;'></i>"
            colorEstado = "green"
        Case "rechazado"
            icono = "<i class='fas fa-times-circle' style='color:red;'></i>"
            colorEstado = "red"
        Case Else
            icono = "<i class='fas fa-hourglass-half' style='color:gray;'></i>"
            colorEstado = "gray"
    End Select
%>
        <tr>
            <td><%= rs("ClienteNombre") %></td>
            <td><%= rs("TotalFacturas") %></td>
            <td>$<%= FormatNumber(rs("ImporteACobrar"), 2) %></td>
            <td><%= icono %> <%= estadoEntrega %></td>
            <td>
                <% If LCase(estadoEntrega) = "pendiente" or LCase(estadoEntrega) = "no iniciada" Then %>
                    <a href="entregar.asp?clienteid=<%= rs("ClienteID") %>&hdrid=<%= hojaRutaID %>" 
                       class="icon-btn entregar" 
                       title="Marcar como entregado"
                       onclick="return confirm('¿Confirmar entrega al cliente?')">
                        <i class="fas fa-check"></i>
                    </a>

                    <a href="rechazar.asp?clienteid=<%= rs("ClienteID") %>&hdrid=<%= hojaRutaID %>" 
                       class="icon-btn rechazar" 
                       title="Rechazar entrega"
                       onclick="return confirm('¿Confirmar rechazo de entrega?')">
                        <i class="fas fa-times"></i>
                    </a>
                <% Else %>
                    <em style="color: <%= colorEstado %>;">Sin acciones</em>
                <% End If %>
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
