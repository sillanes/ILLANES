<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8" 

If Session("currentUser") = "" Then Response.Redirect "../login.asp"

' Conectar y traer campañas pendientes  
Dim rst
Set rst = conn.Execute("EXEC usp_Campania_Pendientes_sel")
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8" /> 
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Campañas Pendientes</title>
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
    table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 20px;
        font-size: 14px;
    }
    table th, table td {
        border: 1px solid #ccc;
        padding: 10px;
        text-align: left;
    }
    table th {
        background-color: #3498db;
        color: white;
    }
    table tr:nth-child(even) {
        background-color: #f9f9f9;
    }
	.btn-delete {
		background-color: #e74c3c;
		border: none;
		padding: 6px 12px;
		color: #fff;
		border-radius: 4px;
		cursor: pointer;
		text-decoration: none;
		font-size: 13px;
		margin-left: 5px;
	}
	.btn-delete:hover {
		background-color: #c0392b;
	}

	
    .btn-action {
        background-color: #2ecc71;
        border: none;
        padding: 6px 12px;
        color: #fff;
        border-radius: 4px;
        cursor: pointer;
        text-decoration: none;
        font-size: 13px;
    }
    .btn-action:hover {
        background-color: #27ae60;
    }
    @media (max-width: 768px) {
        table, thead, tbody, th, td, tr {
            display: block;
        }
        table tr { margin-bottom: 15px; }
        table th {
            text-align: right;
        }
        table td {
            text-align: right;
            padding-left: 50%;
            position: relative;
        }
        table td::before {
            content: attr(data-label);
            position: absolute;
            left: 10px;
            font-weight: bold;
            text-align: left;
        }
    }
	
	
.alert {
    padding: 12px 15px;
    border-radius: 6px;
    margin-bottom: 20px;
    font-size: 14px;
    font-weight: 500;
}
.alert.success {
    background-color: #d4edda;
    color: #155724;
    border: 1px solid #c3e6cb;
}
.alert.error {
    background-color: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
}
	
</style>
    <script>
		function toggleSidebar() {
			document.querySelector('.sidebar').classList.toggle('open');
		}

    </script>
</head>
<body>

<!--#include file="header.asp" -->

<div class="main-content">
    <h2>Campañas Pendientes</h2>
  
<%
Select Case Request.QueryString("success")
    Case "1"
        %><div class='alert success'>✅ Campaña eliminada correctamente.</div><%
End Select

Select Case Request.QueryString("error")
    Case "1"
        %><div class='alert error'>❌ Error al eliminar la campaña.</div><%
    Case "2"
        %><div class='alert error'>⚠️ ID de campaña inválido.</div><%
End Select
%> 
	
    <% If Not rst.EOF Then %>
        <table>
            <thead>
                <tr>
                    <th>Nombre</th>
                    <th>Canal</th>
                    <th>Descripción</th>
                    <th>Acción</th>
                </tr>
            </thead>
            <tbody>
            <% Do Until rst.EOF %>
                <tr>
                    <td data-label="Nombre"><%= Server.HTMLEncode(rst("NombreCampania")) %></td>
                    <td data-label="Canal"><%= Server.HTMLEncode(rst("Canal")) %></td>
                    <td data-label="Descripción"><%= Server.HTMLEncode(rst("Descripcion")) %></td>
					<td data-label="Acción">
						<% 
						Dim canal, urlDestino
						canal = LCase(Trim(rst("Canal")))
						If canal = "email" Then
							urlDestino = "campania_email_configurar.asp?CampaniaID=" & rst("CampaniaID")
						ElseIf canal = "whatsapp" Then
							urlDestino = "campania_whatsapp_configurar.asp?CampaniaID=" & rst("CampaniaID")
						Else
							urlDestino = "#"
						End If
						%>
						<a href="<%=urlDestino%>" class="btn-action">Seleccionar</a>
						<a href="campania_eliminar.asp?CampaniaID=<%=rst("CampaniaID")%>" 
						   class="btn-delete" 
						   onclick="return confirm('¿Seguro que desea eliminar esta campaña?');">
						   Eliminar
						</a>
					</td>

                </tr>
            <% 
                rst.MoveNext
            Loop %>
            </tbody>
        </table>
    <% Else %>
        <p>No hay campañas pendientes.</p>
    <% End If %>

    <%
    rst.Close
    Set rst = Nothing
    %>
</div>
</body>
</html>
