<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
%>
<!--#include file="conexion.asp" -->
<!--#include file="sensores_common.asp" -->
<!--#include file="sidebar.asp" -->
<%
Dim rsDep, rsVeh, rsHoy, errMsg
Set rsDep = Nothing
Set rsVeh = Nothing
Set rsHoy = Nothing
errMsg = ""

On Error Resume Next

Set rsDep = conn.Execute("EXEC dbo.usp_Sensores_Dashboard_Depositos_Sel")
If Err.Number <> 0 Then
  errMsg = "Error ejecutando usp_Sensores_Dashboard_Depositos_Sel: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
End If

Set rsVeh = conn.Execute("EXEC dbo.usp_Sensores_Dashboard_Vehiculos_Sel")
If errMsg = "" And Err.Number <> 0 Then
  errMsg = "Error ejecutando usp_Sensores_Dashboard_Vehiculos_Sel: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
End If

Set rsHoy = conn.Execute("EXEC dbo.usp_Sensores_Dashboard_Hoy_Sel")
If errMsg = "" And Err.Number <> 0 Then
  errMsg = "Error ejecutando usp_Sensores_Dashboard_Hoy_Sel: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
End If

Dim hoyTotal, hoyDep, hoyVeh
hoyTotal = 0 : hoyDep = 0 : hoyVeh = 0
If errMsg = "" Then
  If Not rsHoy Is Nothing Then
    If Not rsHoy.EOF Then
      hoyTotal = Nz(rsHoy("Total"),0)
      hoyDep   = Nz(rsHoy("Deposito"),0)
      hoyVeh   = Nz(rsHoy("Vehiculo"),0)
    End If
  End If
End If

On Error GoTo 0
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
  .table thead th{font-size:.75rem;text-transform:uppercase;letter-spacing:.04em;color:#6b7280;}
  .ellipsis{max-width:280px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}

  .table td, .table th{ vertical-align: middle !important; }
  .col-fecha{ text-align:left; }
  .col-texto{ text-align:left; }
  .col-badge{ text-align:center; }
  .col-num{ text-align:right; font-weight:600; }
  .col-user{ text-align:center; }
  .col-obs{ text-align:left; }

  .table tbody tr:hover{ background-color:#f8fafc; }
</style>

<script>
  function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}
</script>
</head>

<body class="bg-light">

<!--#include file="header.asp" -->

<div class="main-content">
  <div class="container my-4">

    <div class="d-flex align-items-center justify-content-between flex-wrap gap-2 mb-3">
      <div>
        <h4 class="mb-0"><i class="fa-solid fa-temperature-half me-1"></i> Sensores &middot; Dashboard</h4>
        <div class="text-muted small">&Uacute;ltimas lecturas por dep&oacute;sitos/veh&iacute;culos</div>
      </div>

      <div class="d-flex gap-2">
        <a href="sensores_deposito_nuevo.asp" class="btn btn-dark">
          <i class="fa-solid fa-warehouse me-1"></i> Nueva lectura Dep&oacute;sito
        </a>
        <a href="sensores_vehiculo_nuevo.asp" class="btn btn-dark">
          <i class="fa-solid fa-truck me-1"></i> Nueva lectura Veh&iacute;culo
        </a>
        <a href="sensores_reportes.asp" class="btn btn-outline-secondary">
          <i class="fa-solid fa-chart-column me-1"></i> Reportes
        </a>
        <a href="sensores_dashboard.asp" class="btn btn-outline-secondary">
          <i class="fa-solid fa-rotate me-1"></i> Refrescar
        </a>
      </div>
    </div>

    <% If errMsg <> "" Then %>
      <div class="alert alert-danger card-soft">
        <b><i class="fa-solid fa-triangle-exclamation me-1"></i></b>
        <%=HtmlEncode(errMsg)%>
        <div class="small mt-2 text-muted">
          Tip: verific&aacute; que <b>conexion.asp</b> est&eacute; apuntando a la DB <b>Sensores</b> (Initial Catalog=Sensores).
        </div>
      </div>
    <% End If %>

    <div class="row g-3 mb-3">
      <div class="col-12 col-lg-4">
        <div class="card card-soft">
          <div class="card-body">
            <div class="text-muted small mb-1">Lecturas hoy</div>
            <div class="display-6 fw-bold"><%=hoyTotal%></div>
            <div class="text-muted small">Dep&oacute;sitos: <b><%=hoyDep%></b> &middot; Veh&iacute;culos: <b><%=hoyVeh%></b></div>
          </div>
        </div>
      </div>

      <div class="col-12 col-lg-8">
        <div class="card card-soft">
          <div class="card-body">
            <div class="text-muted small mb-2">Tips</div>
            <ul class="mb-0 small text-muted">
              <li>Pod&eacute;s registrar lecturas durante todo el d&iacute;a.</li>
              <li>Dep&oacute;sitos: c&aacute;mara + posici&oacute;n (ALTA/MEDIA/BAJA).</li>
              <li>Veh&iacute;culos: patente + <b>Temperatura Abasto</b> y <b>Temperatura Carga</b>.</li>
            </ul>
          </div>
        </div>
      </div>
    </div>

    <div class="row g-3">
      <!-- Depositos -->
      <div class="col-12 col-lg-6">
        <div class="card card-soft">
          <div class="card-body">
            <h6 class="mb-2">&Uacute;ltimas lecturas &middot; Dep&oacute;sitos</h6>

            <div class="table-responsive">
              <table class="table table-sm align-middle mb-0">
                <thead>
                  <tr>
                    <th class="col-texto">Dep&oacute;sito</th>
                    <th class="col-badge">Posici&oacute;n</th>
                    <th class="col-num">Temp</th>
                    <th class="col-fecha">Fecha/Hora</th>
                    <th class="col-texto">Responsable</th>
                    <th class="col-user">Usuario</th>
                    <th class="col-obs">Obs</th>
                  </tr>
                </thead>
                <tbody>
                <%
                  If rsDep Is Nothing Or rsDep.EOF Then
                %>
                  <tr><td colspan="7" class="text-muted">No hay lecturas de dep&oacute;sitos a&uacute;n.</td></tr>
                <%
                  Else
                    Do While Not rsDep.EOF
                      Dim pos, responsableTxt
                      pos = UCase(Trim(CStr(Nz(rsDep("Posicion"),""))))

                      responsableTxt = "-"
                      On Error Resume Next
                      responsableTxt = Nz(rsDep("Responsable"), "")
                      If Err.Number <> 0 Then
                        Err.Clear
                        responsableTxt = "-"
                      End If
                      On Error GoTo 0

                      If Trim(CStr(responsableTxt)) = "" Then responsableTxt = "-"
                %>
                  <tr>
                    <td class="col-texto"><%=HtmlEncode(Nz(rsDep("Deposito"),""))%></td>
                    <td class="col-badge"><span class="badge <%=PosBadgeClass(pos)%>"><%=HtmlEncode(pos)%></span></td>
                    <td class="col-num"><%=FormatoTemp(rsDep("Temperatura"))%></td>
                    <td class="col-fecha"><%=FormatoFechaHora(rsDep("FechaHora"))%></td>
                    <td class="col-texto"><%=HtmlEncode(responsableTxt)%></td>
                    <td class="col-user"><%=HtmlEncode(Nz(rsDep("Usuario"),"-"))%></td>
                    <td class="col-obs ellipsis" title="<%=HtmlEncode(Nz(rsDep("Observacion"),""))%>"><%=HtmlEncode(Nz(rsDep("Observacion"),""))%></td>
                  </tr>
                <%
                      rsDep.MoveNext
                    Loop
                  End If
                %>
                </tbody>
              </table>
            </div>

          </div>
        </div>
      </div>

      <!-- Vehiculos (1 fila por patente) -->
      <div class="col-12 col-lg-6">
        <div class="card card-soft">
          <div class="card-body">
            <h6 class="mb-2">Veh&iacute;culos &middot; &Uacute;ltima lectura por patente</h6>

            <div class="table-responsive">
              <table class="table table-sm align-middle mb-0">
                <thead>
                  <tr>
                    <th class="col-badge">Patente</th>
                    <th class="col-num">Temp Abasto</th>
                    <th class="col-fecha">Fecha Abasto</th>
                    <th class="col-num">Temp Carga</th>
                    <th class="col-fecha">Fecha Carga</th>
                    <th class="col-texto">Repartidor</th>
                    <th class="col-texto">Ayudante</th>
                    <th class="col-user">Usuario</th>
                  </tr>
                </thead>
                <tbody>
                <%
                  If rsVeh Is Nothing Or rsVeh.EOF Then
                %>
                  <tr><td colspan="8" class="text-muted">No hay lecturas de veh&iacute;culos a&uacute;n.</td></tr>
                <%
                  Else
                    Do While Not rsVeh.EOF
                %>
                  <tr>
                    <td class="col-badge">
                      <span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle">
                        <%=HtmlEncode(UCase(Trim(CStr(Nz(rsVeh("Patente"),"")))))%>
                      </span>
                    </td>

                    <td class="col-num"><%=FormatoTemp(rsVeh("TempAbasto"))%></td>
                    <td class="col-fecha"><%=FormatoFechaHora(rsVeh("FechaAbasto"))%></td>

                    <td class="col-num"><%=FormatoTemp(rsVeh("TempCarga"))%></td>
                    <td class="col-fecha"><%=FormatoFechaHora(rsVeh("FechaCarga"))%></td>

                    <td class="col-texto"><%=HtmlEncode(Nz(rsVeh("Repartidor"),"-"))%></td>
                    <td class="col-texto"><%=HtmlEncode(Nz(rsVeh("Ayudante"),"-"))%></td>
                    <td class="col-user"><%=HtmlEncode(Nz(rsVeh("Usuario"),"-"))%></td>
                  </tr>
                <%
                      rsVeh.MoveNext
                    Loop
                  End If
                %>
                </tbody>
              </table>
            </div>

            <div class="small text-muted mt-2">
              Se muestran hasta 25 patentes (ordenado por &uacute;ltima lectura).
            </div>

          </div>
        </div>
      </div>
    </div>

  </div>
</div>

</body>
</html>

<%
On Error Resume Next
If Not rsDep Is Nothing Then If rsDep.State = 1 Then rsDep.Close
If Not rsVeh Is Nothing Then If rsVeh.State = 1 Then rsVeh.Close
If Not rsHoy Is Nothing Then If rsHoy.State = 1 Then rsHoy.Close
Set rsDep = Nothing
Set rsVeh = Nothing
Set rsHoy = Nothing
On Error GoTo 0
%>