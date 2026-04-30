<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "/login.asp"
End If

Dim hdrid, clienteid, rs, sql
hdrid = Request.QueryString("hdrid")
clienteid = Request.QueryString("clienteid")
cond = Request.QueryString("cond")

Set rs = conn.Execute("EXEC cobranza.usp_Transportista_HojaDeRuta_Cliente_Detalle " & hdrid & ", " & clienteid & ", '" & Replace(cond,"'","''") & "'" )
%>
<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Detalle Cliente</title>  
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link rel="stylesheet" href="estilos.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    <style>
     
        .container {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
        }
        .card {
            background: #fff;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        }
        .card h3 {
            margin-top: 0;
            color: #333;
            font-size: 18px;
            border-bottom: 1px solid #eee;
            padding-bottom: 8px;
        }
        .card p {
            margin: 8px 0;
            font-size: 15px;
        }
        .label {
            font-weight: bold;
            color: #555;
        }
    </style>    
	<script>
    
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
<div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;">
    <div style="display: flex; align-items: center; gap: 10px;">
        <h1 style="margin: 0;">Detalle Cliente</h1>
        <a href="controlhdr_clientes.asp?hdrid=<%=hdrid%>" 
           class="icon-btn" 
           title="Volver"
           style="background-color: #2e30c7; color: white; padding: 8px 12px; border-radius: 4px; text-decoration: none;">
            <i class="fas fa-arrow-left"></i> Volver
        </a>
         
    </div>
     
</div>
 
    
    <div class="container">
        <% If Not rs.EOF Then %>
            <div class="card">
                <h3><i class="fa fa-user"></i> Cliente</h3>
                <p><span class="label">Nombre:</span> <%= rs("ClienteNombre") %></p>
            </div>

            <div class="card">
                <h3><i class="fa fa-file-invoice"></i> Facturas</h3>
                <p><span class="label">Cantidad:</span> <%= rs("CantidadFacturas") %></p>
            </div>

            <div class="card">
                <h3><i class="fa fa-money-bill-wave"></i> Cobros</h3>
                <p><span class="label">Total Cobrado:</span> $<%= FormatNumber(rs("ImporteTotalCobrado"),2) %></p>
                <p><span class="label">Total Efectivo:</span> $<%= FormatNumber(rs("Efectivo"),2) %></p>
                <p><span class="label">Total Cheque:</span> $<%= FormatNumber(rs("Cheque"),2) %></p>
                <p><span class="label">Total Transferencia:</span> $<%= FormatNumber(rs("Transferencia"),2) %></p>
            </div>

            <div class="card">
                <h3><i class="fa fa-clipboard-check"></i> Estado</h3>
                <p><span class="label">Estado:</span> <%= rs("EstadoEntrega") %></p>
            </div>

            <div class="card">
                <h3><i class="fa fa-comment-dots"></i> Observaciones</h3>
                <p><span class="label">Observación CtaCte:</span> <%= rs("ObservacionCC") %></p>
                <p><span class="label">Observación Reparto:</span> <%= rs("ObservacionReparto") %></p>
            </div>
        <% Else %>
            <p>No se encontraron datos para este cliente.</p>
        <% End If %>
    </div> 
</div>	
</body>
</html>
<%
rs.Close
Set rs = Nothing
%>
