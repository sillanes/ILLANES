<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.AddHeader "Content-Type", "text/html; charset=UTF-8"

If Session("currentUser") = "" Then Response.Redirect "../login.asp"

' ============================
' Grilla 1: Campañas pendientes de activación
' ============================
Dim rsPendientes
Set rsPendientes = conn.Execute("EXEC usp_Campania_Pendientes_Activacion_Sel")

' ============================
' Grilla 2: Campañas en proceso/enviadas
' ============================
Dim rsEnProceso
Set rsEnProceso = conn.Execute("EXEC usp_Campania_EnProceso")
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Activación de Campañas</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
body { font-family: Arial, sans-serif; background:#f4f4f4; margin:0; padding:0; }
.main-content { max-width:1000px; margin:30px auto; padding:20px; background:#fff; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.1); }
h2 { text-align:center; margin-bottom:20px; }
table { width:100%; border-collapse:collapse; margin-top:20px; font-size:14px; }
table th, table td { border:1px solid #ddd; padding:10px; text-align:center; }
table th { background:#007bff; color:#fff; }
table tr:nth-child(even) { background:#f9f9f9; }
.btn-enviar {
    background: linear-gradient(135deg,#28a745,#218838);
    border: none;
    padding: 6px 12px;
    color: #fff;
    border-radius: 25px;
    cursor: pointer;
    font-size: 13px;
    text-decoration: none;
    display:inline-block;
}
.btn-enviar:hover { background: linear-gradient(135deg,#218838,#1e7e34); }
.icon { margin-right:5px; }
.status-inproceso { color:#ffc107; font-weight:bold; }
.status-enviado { color:#28a745; font-weight:bold; }
@media (max-width: 768px) {
    table, thead, tbody, th, td, tr { display: block; }
    table tr { margin-bottom: 15px; }
    table th { text-align: right; }
    table td { text-align: right; padding-left: 50%; position: relative; }
    table td::before { content: attr(data-label); position: absolute; left: 10px; font-weight: bold; text-align: left; }
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
    <h2>Campañas pendientes de activación</h2>

    <% If Not rsPendientes.EOF Then %>
        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Nombre</th>
                    <th>Canal</th>
                    <th>Total Destinatarios</th>
                    <th>Cantidad diaria</th>
                    <th>Acciones</th>
                </tr>
            </thead>
            <tbody>
            <% Do Until rsPendientes.EOF %>
                <tr>
                    <td data-label="ID"><%=rsPendientes("CampaniaID")%></td>
                    <td data-label="Nombre"><%=Server.HTMLEncode(rsPendientes("NombreCampania"))%></td>
                    <td data-label="Canal"><%=Server.HTMLEncode(rsPendientes("Canal"))%></td>
                    <td data-label="Cantidad diaria"><%=rsPendientes("TotalDestinatarios")%></td>
                    <td data-label="Cantidad diaria"><%=rsPendientes("CantidadDiaria")%></td>
                    <td data-label="Acciones">
                        <a href="campania_enviar.asp?CampaniaID=<%=rsPendientes("CampaniaID")%>" class="btn-enviar">
                            <i class="fa fa-paper-plane icon"></i> Enviar
                        </a>
                    </td>
                </tr>
            <% 
                rsPendientes.MoveNext
            Loop %>
            </tbody>
        </table>
    <% Else %>
        <p style="text-align:center; color:#666; margin:20px;">No hay campañas pendientes de activación.</p>
    <% End If %>

    <h2>Campañas en proceso / enviadas</h2>

    <% If Not rsEnProceso.EOF Then %>
        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Nombre</th>
                    <th>Canal</th>
                    <th>Cantidad diaria</th>
                    <th>Estado</th>
                    <th>Acciones</th>
                </tr>
            </thead>
            <tbody>
            <% Do Until rsEnProceso.EOF %>
                <tr>
                    <td data-label="ID"><%=rsEnProceso("CampaniaID")%></td>
                    <td data-label="Nombre"><%=Server.HTMLEncode(rsEnProceso("NombreCampania"))%></td>
                    <td data-label="Canal"><%=Server.HTMLEncode(rsEnProceso("Canal"))%></td>
                    <td data-label="Cantidad diaria"><%=rsEnProceso("CantidadDiaria")%></td>
                    <td data-label="Estado">
                        <% 
                        Dim enviado, total
                        enviado = rsEnProceso("Enviados") ' Número de envíos realizados
                        total = rsEnProceso("TotalDestinatarios") ' Total programado
                        If rsEnProceso("Estado") = "2" Then
                            Response.Write "<span class='status-inproceso'>" & enviado & " de " & total & "</span>"
                        Else
                            Response.Write "<span class='status-enviado'>Enviado " & enviado & " de " & total & "</span>"
                        End If
                        %>
                    </td>
                    <td data-label="Acciones">
                        <% If rsEnProceso("Estado") = 2 Then %>
                            <span style="color:#666;">Envío en curso</span>
                        <% Else %>
                            <span style="color:#999;">--</span>
                        <% End If %>
                    </td>
                </tr>
            <% 
                rsEnProceso.MoveNext
            Loop %>
            </tbody>
        </table>
    <% Else %>
        <p style="text-align:center; color:#666; margin:20px;">No hay campañas en proceso o enviadas.</p>
    <% End If %>

<%
rsPendientes.Close
Set rsPendientes = Nothing
rsEnProceso.Close
Set rsEnProceso = Nothing
%>

</div>
</body>
</html>
