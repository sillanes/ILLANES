<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

If Session("patente") = "" Then
    Response.Redirect "../login.asp"
End If

' ================================
' 1) Cargar Patentes
' ================================
sql = "EXEC dbo.usp_Vehiculos_Activos " '& Session("patente")
Set rs = conn.Execute(sql)

' ================================
' 2) Cargar Fechas y KM Servicio
' ================================
Dim rsServ, sqlServ
sqlServ = "EXEC dbo.usp_Vehiculos_KM_Service"
Set rsServ = conn.Execute(sqlServ)

Dim dictServ
Set dictServ = Server.CreateObject("Scripting.Dictionary")

Do While Not rsServ.EOF
    Dim vid, us, kms, rto, broma, seguro

    vid   = rsServ("VehiculoID")

    ' Convertir a string seguro
    us    = IIf(IsNull(rsServ("UltimoServicio")), "", CStr(rsServ("UltimoServicio")))
    kms   = IIf(IsNull(rsServ("KmServicio")), "", CStr(rsServ("KmServicio")))
    rto   = IIf(IsNull(rsServ("RTO")), "", CStr(rsServ("RTO")))
    broma = IIf(IsNull(rsServ("broma")), "", CStr(rsServ("broma")))
    seguro= IIf(IsNull(rsServ("SEGURO")), "", CStr(rsServ("SEGURO")))

    dictServ(vid) = us & "|" & kms & "|" & rto & "|" & broma & "|" & seguro

    rsServ.MoveNext
Loop
rsServ.Close : Set rsServ = Nothing

Dim opcionesVehiculo
opcionesVehiculo = ""

Do While Not rs.EOF
    Dim data, us1, km1, rto1, broma1, seguro1
    vid     = rs("VehiculoID")
    patente = Server.HTMLEncode(rs("Patente"))

    If dictServ.Exists(vid) Then
        data = Split(dictServ(vid), "|")
        us1     = data(0)
        km1     = data(1)
        rto1    = data(2)
        broma1  = data(3)
        seguro1 = data(4)
    Else
        us1     = ""
        km1     = ""
        rto1    = ""
        broma1  = ""
        seguro1 = ""
    End If

    opcionesVehiculo = opcionesVehiculo & _
        "<option value=""" & vid & """ " & _
        " data-ultimoserv=""" & us1 & """ " & _
        " data-kmserv=""" & km1 & """ " & _
        " data-rto=""" & rto1 & """ " & _
        " data-broma=""" & broma1 & """ " & _
        " data-seguro=""" & seguro1 & """>" & patente & "</option>"

    rs.MoveNext
Loop

rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing
%>

<!--#include file="sidebar.asp" -->

<!doctype html>
<html lang="es">
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<title>Vehículos – Iniciar Checklist</title>

<style>
  body{background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;color:#222;margin:0}
  .container{max-width:900px;margin:30px auto;padding:0 15px}
  .card{background:#fff;border-radius:16px;box-shadow:0 10px 25px rgba(0,0,0,.08);padding:24px;margin-bottom:18px}
  .row{display:flex;gap:16px;flex-wrap:wrap}
  .row .col{flex:1 1 280px}
  label{display:block;font-weight:600;margin-bottom:6px}
  input[type="date"],input[type="number"],select{
    width:100%;padding:12px 14px;border:1px solid #dce3ea;border-radius:12px;font-size:15px;background:#fff
  }
  .btns{display:flex;gap:12px;flex-wrap:wrap;margin-top:14px}
  .btn{border:0;border-radius:12px;padding:12px 18px;cursor:pointer;font-weight:600;transition:.2s}
  .btn-primary{background:#2563eb;color:#fff}
  .btn-soft{background:#eef2ff;color:#1e40af}
  .btn-soft.active{background:#1e40af;color:#fff}
  .title{font-size:22px;font-weight:800;margin:0 0 12px}
  .subtitle{color:#555;margin:0 0 16px}
</style>

</head>
<body>

<!--#include file="header.asp" -->

<div class="container">

  <div class="card">
    <h1 class="title">Checklist de Vehículos</h1>
    <p class="subtitle">Completá los datos base y elegí el tipo de control.</p>

    <form id="frmInicio" method="get" action="vehiculos_checklist.asp">

      <div class="row">
        <div class="col">
          <label for="fecha">Fecha</label>
          <input type="date" id="fecha" name="fecha"
      value="<%=Year(Date)%>-<%=Right("0"&Month(Date),2)%>-<%=Right("0"&Day(Date),2)%>"
      required />
        </div>

        <div class="col">
          <label for="km">KM Actual</label>
          <input type="number" id="km" name="km" min="0" required />
        </div>

        <div class="col">
          <label for="vehiculoid">Patente</label>
          <select id="vehiculoid" name="vehiculoid" required>
            <option value="">-- Seleccionar --</option>
            <%=opcionesVehiculo%>
          </select>
        </div>
      </div>

      <!-- Campos ocultos con datos del servicio -->
      <input type="hidden" id="rto" name="rto" />
      <input type="hidden" id="broma" name="broma" />
      <input type="hidden" id="seguro" name="seguro" />
      <input type="hidden" id="servicio" name="servicio" />
      <input type="hidden" id="kmserv" name="kmserv" />

      <div class="card" style="margin-top:16px">
        <label>Tipo de control</label>
        <div class="btns">
          <button type="button" class="btn btn-soft" id="btnD" onclick="setTipo(this,'Diario')">Diario</button>
          <button type="button" class="btn btn-soft" id="btnM" onclick="setTipo(this,'Mensual')">Mensual</button>
        </div>

        <input type="hidden" id="tipo" name="tipo">
      </div>

      <div class="btns">
        <button class="btn btn-primary" type="submit">Comenzar ▶</button>
      </div>

    </form>
  </div>

</div>

<script>
function toggleSidebar(){
  document.querySelector('.sidebar').classList.toggle('open');
}

function setTipo(btn, tipo){
  document.getElementById("tipo").value = tipo;

  // Visual marcar
  document.querySelectorAll(".btn-soft").forEach(b => b.classList.remove("active"));
  btn.classList.add("active");
}

document.getElementById("vehiculoid").addEventListener("change", function(){
    let opt = this.options[this.selectedIndex];

    document.getElementById("servicio").value = opt.dataset.ultimoserv || "";
    document.getElementById("kmserv").value   = opt.dataset.kmserv || "";
    document.getElementById("rto").value      = opt.dataset.rto || "";
    document.getElementById("broma").value    = opt.dataset.broma || "";
    document.getElementById("seguro").value   = opt.dataset.seguro || "";
});

document.getElementById("frmInicio").addEventListener("submit", function(e){
    if(!document.getElementById("tipo").value){
        e.preventDefault();
        alert("Seleccioná el tipo de control.");
    }
});
</script>

</body>
</html>
