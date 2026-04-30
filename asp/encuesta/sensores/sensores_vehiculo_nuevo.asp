<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sensores_common.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("currentUser") = "" Then
  Response.Redirect "login.asp"
End If

Dim msg
msg = Trim(Request.QueryString("msg"))

' ===== Cargar transportistas (Repartidor / Ayudante) =====
Dim rsT, errMsg
Set rsT = Nothing
errMsg = ""

On Error Resume Next
Set rsT = conn.Execute("EXEC dbo.usp_Transportista_sel")
If Err.Number <> 0 Then
  errMsg = "Error cargando transportistas: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
End If
On Error GoTo 0

Set rsT2 = conn.Execute("EXEC dbo.usp_Transportista_sel")
If Err.Number <> 0 Then
  errMsg = "Error cargando transportistas: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
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
</style>

<script>
  function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}

  function normalizePatente(){
    var el = document.getElementById('patente');
    if(!el) return;
    el.value = (el.value || '').toUpperCase().trim().replace(/\s+/g,'');
  }

  function validarTemps(){
    // HTML ya valida required, pero agrego una ayudita visual
    normalizePatente();
    return true;
  }
</script>
</head>

<body class="bg-light">
<!--#include file="header.asp" -->

<div class="main-content">
  <div class="container my-4">

    <div class="d-flex align-items-center justify-content-between flex-wrap gap-2 mb-3">
      <div>
        <h4 class="mb-0"><i class="fa-solid fa-truck me-1"></i> Nueva lectura · Vehículo</h4>
        <div class="text-muted small">Ingresá patente, repartidor/ayudante y temperaturas</div>
      </div>

      <div class="d-flex gap-2">
        <a href="sensores_dashboard.asp" class="btn btn-outline-secondary">
          <i class="fa-solid fa-arrow-left me-1"></i> Volver
        </a>
      </div>
    </div>

    <% If msg <> "" Then %>
      <div class="alert alert-success card-soft"><%=HtmlEncode(msg)%></div>
    <% End If %>

    <% If errMsg <> "" Then %>
      <div class="alert alert-danger card-soft"><%=HtmlEncode(errMsg)%></div>
    <% End If %>

    <div class="card card-soft">
      <div class="card-body">

        <form method="post" action="sensores_vehiculo_guardar.asp" class="row g-3" onsubmit="return validarTemps();">

          <div class="col-12 col-md-4">
            <label class="form-label">Patente</label>
            <input type="text" id="patente" name="patente" class="form-control" maxlength="20"
                   placeholder="Ej: AA123BB" required onblur="normalizePatente();">
            <div class="form-text">Se guarda en mayúsculas (sin espacios).</div>
          </div>

          <div class="col-12 col-md-4">
            <label class="form-label">Repartidor</label>
            <select name="repartidorid" class="form-select" required>
              <option value="">-- Seleccionar --</option>
              <%
                If Not rsT Is Nothing Then
                  If Not rsT.EOF Then
                    rsT.MoveFirst
                    Do While Not rsT.EOF
                      ' Ajustá nombres de columnas si tu SP devuelve distinto
                      Dim tid, tnom
                      tid  = Nz(rsT("TransportistaID"), "")
                      tnom = Nz(rsT("Nombre"), "")
              %>
                <option value="<%=HtmlEncode(tid)%>"><%=HtmlEncode(tnom)%></option>
              <%
                      rsT.MoveNext
                    Loop
                  End If
                End If
              %>
            </select>
          </div>

          <div class="col-12 col-md-4">
            <label class="form-label">Ayudante</label>
            <select name="ayudanteid" class="form-select">
              <option value="">(Opcional)</option>
              <%
                If Not rsT2 Is Nothing Then
                  If Not rsT2.EOF Then
                    rsT2.MoveFirst
                    Do While Not rsT2.EOF
                      Dim tid2, tnom2
                      tid2  = Nz(rsT2("TransportistaID"), "")
                      tnom2 = Nz(rsT2("Nombre"), "")
              %>
                <option value="<%=HtmlEncode(tid2)%>"><%=HtmlEncode(tnom2)%></option>
              <%
                      rsT2.MoveNext
                    Loop
                  End If
                End If
              %>
            </select>
          </div>

          <!-- NUEVO: Temperatura Abasto -->
          <div class="col-12 col-md-4">
            <label class="form-label">Temperatura Abasto (°C)</label>
            <input type="number" step="0.01" name="temp_abasto" class="form-control" required>
            <div class="form-text">Ej: 7.20</div>
          </div>

          <!-- RENOMBRADO: Temperatura Carga -->
          <div class="col-12 col-md-4">
            <label class="form-label">Temperatura Carga (°C)</label>
            <input type="number" step="0.01" name="temp_carga" class="form-control" required>
            <div class="form-text">Ej: 7.20</div>
          </div>

          <div class="col-12 col-md-8">
            <label class="form-label">Observación (opcional)</label>
            <input type="text" name="observacion" maxlength="300" class="form-control"
                   placeholder="Ej: control previo a salida / regreso / etc.">
          </div>

          <div class="col-12 d-flex gap-2">
            <button type="submit" class="btn btn-dark">
              <i class="fa-solid fa-floppy-disk me-1"></i> Guardar lecturas
            </button>
            <a href="sensores_dashboard.asp" class="btn btn-outline-secondary">Cancelar</a>
          </div>

        </form>

      </div>
    </div>

  </div>
</div>

</body>
</html>

<%
On Error Resume Next
If Not rsT Is Nothing Then If rsT.State = 1 Then rsT.Close
Set rsT = Nothing
On Error GoTo 0

On Error Resume Next
If Not rsT2 Is Nothing Then If rsT2.State = 1 Then rsT2.Close
Set rsT2 = Nothing
On Error GoTo 0

%>