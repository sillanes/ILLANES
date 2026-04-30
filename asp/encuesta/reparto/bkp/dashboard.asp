<!--#include file="conexion.asp" -->
<%
If IsEmpty(Session("TransportistaID")) Then
    Response.Redirect "login.asp"
End If

Dim totalClientes, entregados, pendientes
totalClientes = 0
entregados = 0
pendientes = 0
 

Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "usp_Dashboard_Resumen"
cmd.Parameters.Append cmd.CreateParameter("@HojaDeRutaID", 3, 1, , Session("HojaDeRutaID"))

Set rs = cmd.Execute
If Not rs.EOF Then
    totalClientes = rs("TotalClientes")
    entregados = rs("Entregados")
    pendientes = rs("Pendientes")
End If
rs.Close: conn.Close
%>
<!DOCTYPE html>
<html>
<head>
    <title>Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="style.css">
</head>
<body>
<div class="top-bar">
    <div class="user-info">
        <strong><%=Session("TransportistaNombre")%></strong>
        <a href="logout.asp">Cerrar sesión</a>
    </div>
</div>
<div class="dashboard">
    <h2>Dashboard de Reparto</h2>
    <canvas id="grafico1"></canvas>
    <canvas id="grafico2"></canvas>
</div>

<script>
const ctx1 = document.getElementById('grafico1');
new Chart(ctx1, {
    type: 'bar',
    data: {
        labels: ['Total Clientes'],
        datasets: [{
            label: 'Cantidad',
            data: [<%=totalClientes%>],
            backgroundColor: 'rgba(54, 162, 235, 0.6)'
        }]
    }
});

const ctx2 = document.getElementById('grafico2');
new Chart(ctx2, {
    type: 'doughnut',
    data: {
        labels: ['Entregados', 'Pendientes'],
        datasets: [{
            data: [<%=entregados%>, <%=pendientes%>],
            backgroundColor: ['#28a745', '#dc3545']
        }]
    }
});
</script>
</body>
</html>
