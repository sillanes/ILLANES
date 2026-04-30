<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

tipo = "Diario"

' ===========================================
' Datos para el gráfico resumen
' ===========================================
Set cmd = Server.CreateObject("ADODB.Command")
With cmd
    .ActiveConnection = conn
    .CommandType = 4
    .CommandText = "usp_Vehiculos_Checklist_Resumen"
    .Parameters.Append .CreateParameter("@Tipo", 202, 1, 20, tipo)
End With
Set rs = cmd.Execute()
If Not rs.EOF Then
    total = rs("TotalVehiculos")
    confallas = rs("ConFallas")
    sinfallas = rs("SinFallas")
Else
    total = 0 : confallas = 0 : sinfallas = 0
End If
rs.Close : Set rs = Nothing

' ===========================================
' Listado de ítems NO corregidos
' ===========================================
Set cmd2 = Server.CreateObject("ADODB.Command")
With cmd2
    .ActiveConnection = conn
    .CommandType = 4
    .CommandText = "usp_Vehiculos_Checklist_NoItems"
    .Parameters.Append .CreateParameter("@Tipo", 202, 1, 20, tipo)
End With
Set rs2 = cmd2.Execute()
%>
<!--#include file="sidebar.asp" -->
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0"> 
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<title>Dashboard Diario – Vehículos</title>
<link rel="stylesheet" href="estilos.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
body{font-family:Arial,Helvetica,sans-serif;background:#f4f6f8;color:#222;margin:0}
.container{max-width:1200px;margin:20px auto;padding:0 15px}
.flexbox{display:flex;gap:20px;flex-wrap:wrap;align-items:flex-start}
.card{background:#fff;border-radius:16px;box-shadow:0 10px 25px rgba(0,0,0,.08);padding:20px;margin-bottom:18px}
.card h2{margin-top:0;color:#1e3a8a;font-size:18px}
.tbl{width:100%;border-collapse:collapse}
.tbl th, .tbl td{padding:10px 8px;text-align:left;border-bottom:1px solid #e5e7eb;font-size:14px}
.tbl th{background:#eef2ff;color:#1e40af}
.badge-no{background:#fee2e2;color:#991b1b;padding:4px 8px;border-radius:6px;font-weight:600}
.btn-mini{cursor:pointer;padding:4px 8px;border-radius:6px;background:#2563eb;color:#fff;border:0;font-size:13px}
.chart-card{flex:0 0 40%;min-width:320px;text-align:center}
.table-card{flex:1}
.chart-container{width:100%;max-width:420px;margin:0 auto}
@media(max-width:900px){
  .flexbox{flex-direction:column;}
  .chart-card,.table-card{flex:1 1 100%;}
}
</style>
<script>
function marcarCorregido(checklistID, item){
  if(!confirm("¿Marcar el ítem '"+item+"' como corregido?")) return;
  fetch("vehiculos_dashboard_update.asp?ChecklistID="+checklistID+"&Item="+encodeURIComponent(item))
  .then(r=>r.text())
  .then(txt=>{
    alert(txt);
    location.reload();
  })
  .catch(e=>alert("Error: "+e));
}
</script>
</head>
<body>
<!--#include file="header.asp" -->
<div class="container">

  <div class="flexbox">
    <!-- Gráfico -->
    <div class="card chart-card">
      <h2>Resumen general – Checklist Diario</h2>
      <div class="chart-container">
        <canvas id="chartResumen"></canvas>
      </div>
      <p style="margin-top:10px;color:#555;font-size:14px;">
        Total vehículos: <strong><%=total%></strong>
      </p>
    </div>

    <!-- Tabla -->
    <div class="card table-card">
      <h2>Ítems “NO” pendientes</h2>
      <table class="tbl">
        <tr><th>Patente</th><th>Fecha</th><th>Item</th><th>Acción</th></tr>
<%
If Not rs2.EOF Then
  Do While Not rs2.EOF
    Response.Write "<tr>"
    Response.Write "<td>" & rs2("Patente") & "</td>"
    Response.Write "<td>" & rs2("Fecha") & "</td>"
    Response.Write "<td><span class='badge-no'>" & rs2("Item") & "</span></td>"
    Response.Write "<td><button class='btn-mini' onclick=""marcarCorregido(" & rs2("ChecklistID") & ",'" & Replace(rs2("Item"),"'","''") & "')"">Corregido ✓</button></td>"
    Response.Write "</tr>"
    rs2.MoveNext
  Loop
Else
  Response.Write "<tr><td colspan='4'>No hay ítems pendientes.</td></tr>"
End If
rs2.Close : Set rs2 = Nothing
conn.Close : Set conn = Nothing
%>
      </table>
    </div>
  </div>
</div>

<script>
const ctx = document.getElementById('chartResumen');
new Chart(ctx, {
  type: 'doughnut',
  data: {
    labels: ['Con fallas', 'Sin fallas'],
    datasets: [{
      data: [<%=confallas%>, <%=sinfallas%>],
      backgroundColor: ['#dc2626','#16a34a'],
      hoverOffset: 6
    }]
  },
  options: {
    plugins: {
      legend: { position: 'bottom' },
      title: { display: true, text: 'Estado general de vehículos' }
    },
    cutout: '60%'
  }
});
</script>
</body>
</html>
