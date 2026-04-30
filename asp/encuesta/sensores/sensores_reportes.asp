<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sensores_common.asp" -->
<!--#include file="sidebar.asp" -->
<%
' =========================
' Helpers locales
' =========================
Function IIf(b, t, f)
  If CBool(b) Then IIf = t Else IIf = f
End Function

Function ParseDateISO(s, defaultDate)
  ' espera yyyy-mm-dd
  On Error Resume Next
  If Trim(s) = "" Then
    ParseDateISO = defaultDate
  Else
    Dim y,m,d
    y = CInt(Left(s,4))
    m = CInt(Mid(s,6,2))
    d = CInt(Right(s,2))
    ParseDateISO = DateSerial(y,m,d)
  End If
  If Err.Number <> 0 Then
    ParseDateISO = defaultDate
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function ToISODate(dt)
  ToISODate = Year(dt) & "-" & Right("0"&Month(dt),2) & "-" & Right("0"&Day(dt),2)
End Function

Function SqlISODate(dt)
  SqlISODate = Year(dt) & "-" & Right("0"&Month(dt),2) & "-" & Right("0"&Day(dt),2)
End Function

Function RSIsEmpty(rs)
  ' True si rs es Nothing o EOF
  RSIsEmpty = True
  If rs Is Nothing Then Exit Function
  If rs.EOF Then Exit Function
  RSIsEmpty = False
End Function

' =========================
' Parámetros
' =========================
Dim tab
tab = LCase(Trim(Request("tab")))
If tab <> "vehiculo" Then tab = "deposito"

Dim depositoid, posicion, patente, desdeStr, hastaStr
depositoid = Trim(Request("depositoid"))
posicion   = UCase(Trim(Request("posicion")))
patente    = UCase(Trim(Request("patente")))
desdeStr   = Trim(Request("desde"))
hastaStr   = Trim(Request("hasta"))

' NUEVO: tipo temperatura para VEHICULO
Dim tipotemp
tipotemp = UCase(Trim(Request("tipotemp")))
If tipotemp <> "ABASTO" And tipotemp <> "CARGA" Then tipotemp = ""  ' "" => todos

Dim export
export = Trim(Request("export")) ' "1" => excel

Dim doBuscar
doBuscar = (Trim(Request("buscar")) = "1") Or (export = "1")

' defaults: últimos 7 días
Dim hoy, desdeDef, hastaDef
hoy = Date()
hastaDef = hoy
desdeDef = DateAdd("d",-7,hoy)

Dim desdeDt, hastaDt
desdeDt = ParseDateISO(desdeStr, desdeDef)
hastaDt = ParseDateISO(hastaStr, hastaDef)

' =========================
' Cargar depósitos (combo)
' =========================
Dim rsDep, errMsg
Set rsDep = Nothing
errMsg = ""

On Error Resume Next
Set rsDep = conn.Execute("EXEC dbo.usp_Sensores_Depositos_Sel")
If Err.Number <> 0 Then
  errMsg = "Error cargando depósitos: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
End If
On Error GoTo 0

' =========================
' Ejecutar reporte + resumen
' (CONCATENACIÓN SEGURA)
' =========================
Dim rsRep, rsSum
Set rsRep = Nothing
Set rsSum = Nothing

If errMsg = "" Then
  If doBuscar Then

    If tab = "deposito" Then

      If depositoid = "" Or Not IsNumeric(depositoid) Or CLng(depositoid) <= 0 Then
        errMsg = "Seleccioná un depósito."
      Else
        Dim posParam, sqlRepD, sqlSumD

        If posicion = "" Then
          posParam = "NULL"
        Else
          posParam = "'" & Replace(posicion,"'","''") & "'"
        End If

        sqlRepD = "EXEC dbo.usp_Sensores_Reporte_Deposito_Sel " & _
                  CLng(depositoid) & ", " & _
                  posParam & ", '" & _
                  SqlISODate(desdeDt) & "', '" & _
                  SqlISODate(hastaDt) & "'"

        sqlSumD = "EXEC dbo.usp_Sensores_Reporte_Deposito_Resumen_Sel " & _
                  CLng(depositoid) & ", " & _
                  posParam & ", '" & _
                  SqlISODate(desdeDt) & "', '" & _
                  SqlISODate(hastaDt) & "'"

        On Error Resume Next
        Set rsRep = conn.Execute(sqlRepD)
        If Err.Number <> 0 Then
          errMsg = "Error ejecutando reporte depósito: " & Err.Description & " | " & GetAdoErrors()
          Err.Clear
        End If

        If errMsg = "" Then
          Set rsSum = conn.Execute(sqlSumD)
          If Err.Number <> 0 Then
            errMsg = "Error ejecutando resumen depósito: " & Err.Description & " | " & GetAdoErrors()
            Err.Clear
          End If
        End If
        On Error GoTo 0

      End If

    Else ' vehiculo

      If patente = "" Then
        errMsg = "Ingresá la patente."
      Else
        Dim sqlRepV, sqlSumV, tipoParam

        If tipotemp = "" Then
          tipoParam = "NULL"
        Else
          tipoParam = "'" & Replace(tipotemp,"'","''") & "'"
        End If

        ' NUEVO: 4º parámetro tipotemp
        sqlRepV = "EXEC dbo.usp_Sensores_Reporte_Vehiculo_Sel '" & _
                  Replace(patente,"'","''") & "', '" & _
                  SqlISODate(desdeDt) & "', '" & _
                  SqlISODate(hastaDt) & "', " & _
                  tipoParam

        sqlSumV = "EXEC dbo.usp_Sensores_Reporte_Vehiculo_Resumen_Sel '" & _
                  Replace(patente,"'","''") & "', '" & _
                  SqlISODate(desdeDt) & "', '" & _
                  SqlISODate(hastaDt) & "', " & _
                  tipoParam

        On Error Resume Next
        Set rsRep = conn.Execute(sqlRepV)
        If Err.Number <> 0 Then
          errMsg = "Error ejecutando reporte vehículo: " & Err.Description & " | " & GetAdoErrors()
          Err.Clear
        End If

        If errMsg = "" Then
          Set rsSum = conn.Execute(sqlSumV)
          If Err.Number <> 0 Then
            errMsg = "Error ejecutando resumen vehículo: " & Err.Description & " | " & GetAdoErrors()
            Err.Clear
          End If
        End If
        On Error GoTo 0

      End If

    End If

  End If
End If

' =========================
' EXPORT EXCEL
' =========================
If export = "1" Then
  Response.Buffer = True
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "Content-Disposition", "attachment; filename=sensores_reporte.xls"

  Dim cE, tminE, tmaxE, tpromE
  cE = "" : tminE = "" : tmaxE = "" : tpromE = ""

  If errMsg = "" Then
    If Not rsSum Is Nothing Then
      If Not rsSum.EOF Then
        cE = Nz(rsSum("Cantidad"),0)
        tminE = rsSum("TempMin")
        tmaxE = rsSum("TempMax")
        tpromE = rsSum("TempProm")
      End If
    End If
  End If
%>
<meta charset="UTF-8">
<table border="1">
  <tr><th colspan="8">Resumen</th></tr>
  <tr>
    <td><b>Lecturas</b></td><td><%=HtmlEncode(cE)%></td>
    <td><b>Mín</b></td><td><%=HtmlEncode(Nz(tminE,""))%></td>
    <td><b>Máx</b></td><td><%=HtmlEncode(Nz(tmaxE,""))%></td>
    <td><b>Prom</b></td><td><%=HtmlEncode(Nz(tpromE,""))%></td>
  </tr>

  <tr><td colspan="8"></td></tr>

  <tr>
    <% If tab="deposito" Then %>
      <th>FechaHora</th><th>Deposito</th><th>Posicion</th><th>Temperatura</th><th>Usuario</th><th>Observacion</th><th></th><th></th>
    <% Else %>
      <th>FechaHora</th><th>Patente</th><th>TipoTemp</th><th>Temperatura</th><th>Repartidor</th><th>Ayudante</th><th>Usuario</th><th>Observacion</th>
    <% End If %>
  </tr>

  <%
    If errMsg <> "" Then
      Response.Write "<tr><td colspan='8'>" & HtmlEncode(errMsg) & "</td></tr>"
    ElseIf RSIsEmpty(rsRep) Then
      Response.Write "<tr><td colspan='8'>Sin datos</td></tr>"
    Else
      Do While Not rsRep.EOF
        If tab="deposito" Then
  %>
    <tr>
      <td><%=HtmlEncode(FormatoFechaHora(rsRep("FechaHora")))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Deposito"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Posicion"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Temperatura"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Usuario"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Observacion"),""))%></td>
      <td></td><td></td>
    </tr>
  <%
        Else
  %>
    <tr>
      <td><%=HtmlEncode(FormatoFechaHora(rsRep("FechaHora")))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Patente"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("TipoTemp"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Temperatura"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Repartidor"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Ayudante"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Usuario"),""))%></td>
      <td><%=HtmlEncode(Nz(rsRep("Observacion"),""))%></td>
    </tr>
  <%
        End If
        rsRep.MoveNext
      Loop
    End If
  %>
</table>
<%
  On Error Resume Next
  If Not rsDep Is Nothing Then If rsDep.State=1 Then rsDep.Close
  If Not rsRep Is Nothing Then If rsRep.State=1 Then rsRep.Close
  If Not rsSum Is Nothing Then If rsSum.State=1 Then rsSum.Close
  Set rsDep = Nothing
  Set rsRep = Nothing
  Set rsSum = Nothing
  On Error GoTo 0
  Response.End
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>ILLANES HNOS SRL</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

<style>
  .card-soft{border:1px solid #e6e8ef;border-radius:14px;box-shadow:0 2px 10px rgba(0,0,0,.04);}
  .table tbody tr:hover { background-color:#f8fafc; }

  .table thead th{
    font-size:.75rem;
    text-transform:uppercase;
    letter-spacing:.04em;
    color:#6b7280;
  }
  .table td, .table th{ vertical-align: middle !important; }

  .col-fecha{ text-align:left; }
  .col-texto{ text-align:left; }
  .col-badge{ text-align:center; }
  .col-num{ text-align:right; font-weight:600; }
  .col-user{ text-align:center; }
  .col-obs{ text-align:left; }

  .ellipsis{max-width:320px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
</style>

<script>
  function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}

  async function loadPosiciones(depositoId){
    const selPos = document.getElementById('posicion');
    const spinner = document.getElementById('posSpinner');
    const alertBox = document.getElementById('ajaxError');

    alertBox.classList.add('d-none');
    alertBox.innerText = '';
    selPos.innerHTML = '<option value="">(Todas)</option>';

    if(!depositoId || depositoId === '0'){
      selPos.disabled = true;
      return;
    }

    selPos.disabled = true;
    spinner.classList.remove('d-none');

    try{
      const url = 'sensores_deposito_posiciones_api.asp?depositoid=' + encodeURIComponent(depositoId);
      const resp = await fetch(url, { cache: 'no-store' });
      const data = await resp.json();
      if(!data.ok) throw new Error(data.error || 'Error');

      for(const it of (data.items || [])){
        const opt = document.createElement('option');
        opt.value = it.posicion || '';
        opt.textContent = it.posicion || '';
        selPos.appendChild(opt);
      }

      selPos.disabled = false;

    }catch(e){
      alertBox.innerText = 'No se pudieron cargar posiciones: ' + (e.message || e);
      alertBox.classList.remove('d-none');
      selPos.disabled = true;
    }finally{
      spinner.classList.add('d-none');
    }
  }

  window.addEventListener('DOMContentLoaded', () => {
    const tab = '<%=tab%>';
    if(tab === 'deposito'){
      const dep = document.getElementById('depositoid');
      if(dep && dep.value && dep.value !== '0'){
        loadPosiciones(dep.value).then(() => {
          const wanted = '<%=HtmlEncode(posicion)%>';
          if(wanted){
            const selPos = document.getElementById('posicion');
            for(const opt of selPos.options){
              if(opt.value === wanted){ opt.selected = true; break; }
            }
          }
        });
      }
    }
  });
</script>
</head>

<body class="bg-light">
<!--#include file="header.asp" -->

<div class="main-content">
  <div class="container my-4">

    <div class="d-flex align-items-center justify-content-between flex-wrap gap-2 mb-3">
      <div>
        <h4 class="mb-0"><i class="fa-solid fa-chart-column me-1"></i> Sensores · Reportes</h4>
        <div class="text-muted small">Filtrá por depósito o vehículo y exportá a Excel</div>
      </div>
      <div class="d-flex gap-2">
        <a href="sensores_dashboard.asp" class="btn btn-outline-secondary">
          <i class="fa-solid fa-arrow-left me-1"></i> Volver
        </a>
      </div>
    </div>

    <% If errMsg <> "" Then %>
      <div class="alert alert-danger card-soft">
        <b><i class="fa-solid fa-triangle-exclamation me-1"></i></b> <%=HtmlEncode(errMsg)%>
      </div>
    <% End If %>

    <div id="ajaxError" class="alert alert-warning card-soft d-none"></div>

    <ul class="nav nav-tabs mb-3">
      <li class="nav-item">
        <a class="nav-link <%=IIf(tab="deposito","active","")%>" href="sensores_reportes.asp?tab=deposito">Depósitos</a>
      </li>
      <li class="nav-item">
        <a class="nav-link <%=IIf(tab="vehiculo","active","")%>" href="sensores_reportes.asp?tab=vehiculo">Vehículos</a>
      </li>
    </ul>

    <div class="card card-soft mb-3">
      <div class="card-body">

        <form method="get" action="sensores_reportes.asp" class="row g-3">
          <input type="hidden" name="tab" value="<%=HtmlEncode(tab)%>">

          <% If tab="deposito" Then %>
            <div class="col-12 col-md-4">
              <label class="form-label">Depósito</label>
              <select id="depositoid" name="depositoid" class="form-select" onchange="loadPosiciones(this.value)" required>
                <option value="0">-- Seleccionar --</option>
                <%
                  If Not rsDep Is Nothing Then
                    If Not rsDep.EOF Then
                      Do While Not rsDep.EOF
                        Dim did, dnom, sel
                        did = CLng(rsDep("DepositoID"))
                        dnom = CStr(rsDep("Nombre"))
                        sel = ""
                        If IsNumeric(depositoid) Then
                          If CLng(depositoid)=did Then sel=" selected"
                        End If
                %>
                  <option value="<%=did%>"<%=sel%>><%=HtmlEncode(dnom)%></option>
                <%
                        rsDep.MoveNext
                      Loop
                    End If
                  End If
                %>
              </select>
            </div>

            <div class="col-12 col-md-3">
              <label class="form-label d-flex align-items-center gap-2">
                Posición
                <span id="posSpinner" class="d-none"><i class="fa-solid fa-spinner fa-spin text-muted"></i></span>
              </label>
              <select id="posicion" name="posicion" class="form-select" disabled>
                <option value="">(Todas)</option>
              </select>
            </div>
          <% Else %>
            <div class="col-12 col-md-3">
              <label class="form-label">Patente</label>
              <input type="text" name="patente" class="form-control" value="<%=HtmlEncode(patente)%>" maxlength="20" required>
            </div>

            <!-- NUEVO: filtro tipo temp -->
            <div class="col-12 col-md-3">
              <label class="form-label">Tipo</label>
              <select name="tipotemp" class="form-select">
                <option value="" <%=IIf(tipotemp="","selected","")%>>(Todos)</option>
                <option value="ABASTO" <%=IIf(tipotemp="ABASTO","selected","")%>>ABASTO</option>
                <option value="CARGA" <%=IIf(tipotemp="CARGA","selected","")%>>CARGA</option>
              </select>
            </div>
          <% End If %>

          <div class="col-6 col-md-2">
            <label class="form-label">Desde</label>
            <input type="date" name="desde" class="form-control" value="<%=HtmlEncode(ToISODate(desdeDt))%>" required>
          </div>

          <div class="col-6 col-md-2">
            <label class="form-label">Hasta</label>
            <input type="date" name="hasta" class="form-control" value="<%=HtmlEncode(ToISODate(hastaDt))%>" required>
          </div>

          <div class="col-12 d-flex gap-2">
            <button type="submit" name="buscar" value="1" class="btn btn-dark">
              <i class="fa-solid fa-magnifying-glass me-1"></i> Buscar
            </button>

            <button type="submit" name="export" value="1" class="btn btn-outline-success">
              <i class="fa-solid fa-file-excel me-1"></i> Exportar Excel
            </button>
          </div>
        </form>

      </div>
    </div>

    <%
      ' ===== RESUMEN =====
      Dim showResumen
      showResumen = False
      If doBuscar And errMsg = "" Then
        If Not rsSum Is Nothing Then
          If Not rsSum.EOF Then
            showResumen = True
          End If
        End If
      End If
    %>

    <% If showResumen Then %>
      <%
        Dim c2, tmin2, tmax2, tprom2
        c2 = Nz(rsSum("Cantidad"),0)
        tmin2 = rsSum("TempMin")
        tmax2 = rsSum("TempMax")
        tprom2 = rsSum("TempProm")
      %>

      <div class="row g-3 mb-3">
        <div class="col-6 col-lg-3">
          <div class="card card-soft"><div class="card-body">
            <div class="text-muted small">Lecturas</div>
            <div class="h4 mb-0 fw-bold"><%=c2%></div>
          </div></div>
        </div>
        <div class="col-6 col-lg-3">
          <div class="card card-soft"><div class="card-body">
            <div class="text-muted small">Mínima</div>
            <div class="h4 mb-0 fw-bold"><%=FormatoTemp(tmin2)%></div>
          </div></div>
        </div>
        <div class="col-6 col-lg-3">
          <div class="card card-soft"><div class="card-body">
            <div class="text-muted small">Máxima</div>
            <div class="h4 mb-0 fw-bold"><%=FormatoTemp(tmax2)%></div>
          </div></div>
        </div>
        <div class="col-6 col-lg-3">
          <div class="card card-soft"><div class="card-body">
            <div class="text-muted small">Promedio</div>
            <div class="h4 mb-0 fw-bold"><%=FormatoTemp(tprom2)%></div>
          </div></div>
        </div>
      </div>
    <% End If %>

    <div class="card card-soft">
      <div class="card-body">

        <%
          If Not doBuscar Then
        %>
            <div class="text-muted">Elegí filtros y hacé click en <b>Buscar</b>.</div>
        <%
          ElseIf errMsg <> "" Then
        %>
            <div class="text-muted">No se pudo generar el reporte.</div>
        <%
          ElseIf RSIsEmpty(rsRep) Then
        %>
            <div class="text-muted">Sin resultados para el filtro seleccionado.</div>
        <%
          Else
        %>

          <div class="table-responsive">
            <table class="table table-sm align-middle mb-0">
              <thead>
                <% If tab="deposito" Then %>
                <tr>
                  <th class="col-fecha">Fecha/Hora</th>
                  <th class="col-texto">Depósito</th>
                  <th class="col-badge">Posición</th>
                  <th class="col-num">Temp</th>
                  <th class="col-user">Usuario</th>
                  <th class="col-obs">Obs</th>
                </tr>
                <% Else %>
                <tr>
                  <th class="col-fecha">Fecha/Hora</th>
                  <th class="col-badge">Patente</th>
                  <th class="col-badge">Tipo</th>
                  <th class="col-num">Temp</th>
                  <th class="col-texto">Repartidor</th>
                  <th class="col-texto">Ayudante</th>
                  <th class="col-user">Usuario</th>
                  <th class="col-obs">Obs</th>
                </tr>
                <% End If %>
              </thead>
              <tbody>
                <%
                  Do While Not rsRep.EOF
                    If tab="deposito" Then
                      Dim pos3
                      pos3 = UCase(Trim(CStr(Nz(rsRep("Posicion"),""))))
                %>
                  <tr>
                    <td class="col-fecha"><%=FormatoFechaHora(rsRep("FechaHora"))%></td>
                    <td class="col-texto"><%=HtmlEncode(Nz(rsRep("Deposito"),""))%></td>
                    <td class="col-badge">
                      <span class="badge <%=PosBadgeClass(pos3)%>"><%=HtmlEncode(pos3)%></span>
                    </td>
                    <td class="col-num"><%=FormatoTemp(rsRep("Temperatura"))%></td>
                    <td class="col-user"><%=HtmlEncode(Nz(rsRep("Usuario"),"-"))%></td>
                    <td class="col-obs ellipsis" title="<%=HtmlEncode(Nz(rsRep("Observacion"),""))%>"><%=HtmlEncode(Nz(rsRep("Observacion"),""))%></td>
                  </tr>
                <%
                    Else
                      Dim tt
                      tt = UCase(Trim(CStr(Nz(rsRep("TipoTemp"),""))))
                %>
                  <tr>
                    <td class="col-fecha"><%=FormatoFechaHora(rsRep("FechaHora"))%></td>
                    <td class="col-badge">
                      <span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle"><%=HtmlEncode(Nz(rsRep("Patente"),""))%></span>
                    </td>
                    <td class="col-badge">
                      <% If tt="ABASTO" Then %>
                        <span class="badge bg-primary-subtle text-primary border border-primary-subtle">ABASTO</span>
                      <% ElseIf tt="CARGA" Then %>
                        <span class="badge bg-success-subtle text-success border border-success-subtle">CARGA</span>
                      <% Else %>
                        <span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle">-</span>
                      <% End If %>
                    </td>
                    <td class="col-num"><%=FormatoTemp(rsRep("Temperatura"))%></td>
                    <td class="col-texto"><%=HtmlEncode(Nz(rsRep("Repartidor"),"-"))%></td>
                    <td class="col-texto"><%=HtmlEncode(Nz(rsRep("Ayudante"),"-"))%></td>
                    <td class="col-user"><%=HtmlEncode(Nz(rsRep("Usuario"),"-"))%></td>
                    <td class="col-obs ellipsis" title="<%=HtmlEncode(Nz(rsRep("Observacion"),""))%>"><%=HtmlEncode(Nz(rsRep("Observacion"),""))%></td>
                  </tr>
                <%
                    End If
                    rsRep.MoveNext
                  Loop
                %>
              </tbody>
            </table>
          </div>

        <%
          End If
        %>

      </div>
    </div>

  </div>
</div>

</body>
</html>

<%
On Error Resume Next
If Not rsDep Is Nothing Then If rsDep.State=1 Then rsDep.Close
If Not rsRep Is Nothing Then If rsRep.State=1 Then rsRep.Close
If Not rsSum Is Nothing Then If rsSum.State=1 Then rsSum.Close
Set rsDep = Nothing
Set rsRep = Nothing
Set rsSum = Nothing
On Error GoTo 0
%>