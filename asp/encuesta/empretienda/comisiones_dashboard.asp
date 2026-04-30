<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "login.asp"
End If

Dim rs, sql
sql = "EXEC whatsapp.usp_Dashboard_CampaniasWhatsApp"
Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Dashboard Campañas WhatsApp</title>
<link rel="stylesheet" href="estilos.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
body { font-family: Arial; background: #f7f7f7; margin: 0; padding: 20px; }
h1 { color: #333; }
table { border-collapse: collapse; width: 100%; background: #fff; margin-top: 15px; }
th, td { padding: 10px; text-align: center; border-bottom: 1px solid #ccc; }
th { background: #007BFF; color: white; }
tr:hover { background: #f1f1f1; }
.card { background: white; border-radius: 10px; padding: 15px; margin-bottom: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
canvas { width: 100%; max-width: 800px; margin: 30px auto; display: block; }
</style>
</head>
<body>
<h1>📊 Dashboard de Campañas WhatsApp</h1>

<div class="card">
    <table>
        <tr>
            <th>Campaña</th>
            <th>Total</th>
            <th>Enviados</th>
            <th>Entregados</th>
            <th>Leídos</th>
            <th>Respondidos</th>
            <th>% Leídos</th>
            <th>% Respondidos</th>
            <th>Última Actualización</th>
        </tr>
        <%
        Dim labels, enviados, leidos, respondidos
        labels = ""
        enviados = ""
        leidos = ""
        respondidos = ""

        Do Until rs.EOF
            Response.Write "<tr>"
            Response.Write "<td>" & rs("NombreCampania") & "</td>"
            Response.Write "<td>" & rs("TotalDestinatarios") & "</td>"
            Response.Write "<td>" & rs("Enviados") & "</td>"
            Response.Write "<td>" & rs("Entregados") & "</td>"
            Response.Write "<td>" & rs("Leidos") & "</td>"
            Response.Write "<td>" & rs("Respondidos") & "</td>"
            Response.Write "<td>" & rs("PorcentajeLeidos") & "%</td>"
            Response.Write "<td>" & rs("PorcentajeRespondidos") & "%</td>"
            Response.Write "<td>" & rs("UltimaActualizacion") & "</td>"
            Response.Write "</tr>"

            labels = labels & "'" & rs("NombreCampania") & "',"
            enviados = enviados & rs("Enviados") & ","
            leidos = leidos & rs("Leidos") & ","
            respondidos = respondidos & rs("Respondidos") & ","
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
    </table>
</div>

<div class="card">
    <canvas id="chartCampanias"></canvas>
</div>

<script>
const ctx = document.getElementById('chartCampanias');
new Chart(ctx, {
    type: 'bar',
    data: {
        labels: [<%=Left(labels, Len(labels)-1)%>],
        datasets: [
            {
                label: 'Enviados',
                data: [<%=Left(enviados, Len(enviados)-1)%>],
                backgroundColor: 'rgba(54, 162, 235, 0.6)'
            },
            {
                label: 'Leídos',
                data: [<%=Left(leidos, Len(leidos)-1)%>],
                backgroundColor: 'rgba(75, 192, 192, 0.6)'
            },
            {
                label: 'Respondidos',
                data: [<%=Left(respondidos, Len(respondidos)-1)%>],
                backgroundColor: 'rgba(255, 159, 64, 0.6)'
            }
        ]
    },
    options: {
        responsive: true,
        plugins: { legend: { position: 'bottom' } },
        scales: { y: { beginAtZero: true } }
    }
});
</script>
</body>
</html>
