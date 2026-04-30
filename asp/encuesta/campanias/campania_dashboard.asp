<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim rs, sql
sql = "EXEC  report.usp_Campania_WhastApp_Dashboard"
Set rs = conn.Execute(sql)
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Dashboard Campañas WhatsApp</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
body { font-family: Arial, sans-serif; background:#f4f4f4; margin:0; padding:0; }
.main-content { max-width:1000px; margin:30px auto; padding:20px; background:#fff; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.1); }
h1 { color: #333; text-align:center; }
h2 { text-align:center; margin-bottom:20px; }
table { border-collapse: collapse; width: 100%; background: #fff; margin-top: 15px; }
th, td { padding: 10px; text-align: center; border-bottom: 1px solid #ccc; }
th { background: #007BFF; color: white; }
tr:hover { background: #f1f1f1; }
.card { background: white; border-radius: 10px; padding: 15px; margin-bottom: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
canvas { width: 100%; max-width: 800px; margin: 30px auto; display: block; }
.btn-accion {
    background: #28a745;
    color: white;
    padding: 6px 10px;
    border-radius: 6px;
    text-decoration: none;
    margin: 0 4px;
    display: inline-flex;
    align-items: center;
    font-size: 13px;
    transition: all 0.2s ease-in-out;
}
.btn-accion:hover {
    background: #218838;
    transform: scale(1.05);
}
.btn-accion.error {
    background: #dc3545;
}
.btn-accion.error:hover {
    background: #b02a37;
}
.accion-texto {
    margin-left: 5px;
    font-size: 12px;
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
<h1>📊 Dashboard de Campañas WhatsApp</h1>

<div class="card">
    <table>
        <tr>
            <th>Campaña</th>
            <th>Total</th>
            <th>Enviados</th>
            <th>Entregados</th>
            <th>Leídos</th>
            <th>Errores</th>
            <th>Respondidos</th>
            <th>% Leídos</th>
            <th>% Respondidos</th>
            <th>Última Actualización</th>
			<th>Acción</th>
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
            Response.Write "<td>" & rs("Errores") & "</td>"
            Response.Write "<td>" & rs("Respondidos") & "</td>"
            Response.Write "<td>" & rs("PorcentajeLeidos") & "%</td>"
            Response.Write "<td>" & rs("PorcentajeRespondidos") & "%</td>"
            Response.Write "<td>" & rs("UltimaActualizacion") & "</td>"
            Response.Write "<td>"
            Response.Write "<a class='btn-accion' href='campania_respuestas.asp?campaniaid=" & rs("campaniaID") & "&source=1' title='Ver Respuestas'><i class='fa-solid fa-eye'></i> <span class='accion-texto'></span></a> "
            Response.Write "<a class='btn-accion error' href='campania_respuestas.asp?campaniaid=" & rs("campaniaID") & "&source=2' title='Ver Errores'><i class='fa-solid fa-triangle-exclamation'></i> <span class='accion-texto'></span></a>"
            Response.Write "</td>"
			Response.Write "</tr>"

            labels = labels & "'" & rs("NombreCampania") & "',"
            enviados = enviados & rs("Enviados") & ","
            leidos = leidos & rs("Leidos") & ","
            respondidos = respondidos & rs("Respondidos") & ","
            errores = errores & rs("errores") & ","
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
            },
            {
                label: 'Errores',
                data: [<%=Left(errores, Len(errores)-1)%>],
                backgroundColor: 'rgba(255, 0, 0, 0.6)'
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

</DIV>
</body>
</html>
