<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sensores_common.asp" -->
<!--#include file="sidebar.asp" -->
<%
Function IIf(b, t, f)
  If CBool(b) Then IIf = t Else IIf = f
End Function

' Lee un campo del recordset sin romper si no existe
Function RSField(rs, fieldName, defaultValue)
  On Error Resume Next
  Dim v
  v = rs(fieldName)
  If Err.Number <> 0 Then
    Err.Clear
    RSField = defaultValue
  Else
    If IsNull(v) Then
      RSField = defaultValue
    Else
      RSField = v
    End If
  End If
  On Error GoTo 0
End Function

Dim msg, errMsg
msg = Trim(Request.QueryString("msg"))
errMsg = ""

Dim rsDep, rsResp
Set rsDep  = Nothing
Set rsResp = Nothing

On Error Resume Next

Set rsDep = conn.Execute("EXEC dbo.usp_Sensores_Depositos_Sel")
If Err.Number <> 0 Then
  errMsg = "Error cargando depósitos: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
End If

If errMsg = "" Then
  Set rsResp = conn.Execute("EXEC dbo.usp_Sensores_Responsables_Sel")
  If Err.Number <> 0 Then
    errMsg = "Error cargando responsables: " & Err.Description & " | " & GetAdoErrors()
    Err.Clear
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
</style>

<script>
  function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}

  async function loadPosiciones(depositoId){
    const selPos = document.getElementById('posicion');
    const alertBox = document.getElementById('ajaxError');
    const spinner = document.getElementById('posSpinner');

    alertBox.classList.add('d-none');
    alertBox.innerText = '';
    selPos.innerHTML = '<option value="">-- Seleccionar --</option>';

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

      if(!data.ok){
        throw new Error(data.error || 'Error desconocido');
      }

      if(Array.isArray(data.items)){
        for(const it of data.items){
          const opt = document.createElement('option');
          opt.value = it.posicion || '';
          opt.textContent = it.posicion || '';
          selPos.appendChild(opt);
        }
      }

      selPos.disabled = false;

    }catch(e){
      alertBox.innerText = 'No se pudieron cargar las posiciones: ' + (e.message || e);
      alertBox.classList.remove('d-none');
      selPos.disabled = true;
    }finally{
      spinner.classList.add('d-none');
    }
  }

  function onDepositoChange(sel){
    loadPosiciones(sel.value);
  }

  function normalizeTemp(){
    const el = document.getElementById('temperatura');
    if(!el) return true;

    let v = (el.value || '').trim();
    if(v === '') return true;

    v = v.replace(/[^\d\-\.,]/g, '');
    v = v.replace(/(?!^)-/g, '');
    v = v.replace(/,/g, '.');

    const parts = v.split('.');
    if(parts.length > 2){
      v = parts.shift() + '.' + parts.join('');
    }

    if(v === '.' || v === '-' || v === '-.'){
      el.value = '';
      return true;
    }

    el.value = v;
    return true;
  }

  window.addEventListener('DOMContentLoaded', () => {
    const dep = document.getElementById('depositoid');
    if(dep && dep.value && dep.value !== '0'){
      loadPosiciones(dep.value);
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
        <h4 class="mb-0"><i class="fa-solid fa-warehouse me-1"></i> Nueva lectura · Depósito</h4>
        <div class="text-muted small">Seleccioná cámara + posición y registrá la temperatura</div>
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
      <div class="alert alert-danger card-soft">
        <b><i class="fa-solid fa-triangle-exclamation me-1"></i></b>
        <%=HtmlEncode(errMsg)%>
      </div>
    <% End If %>

    <div id="ajaxError" class="alert alert-warning card-soft d-none"></div>

    <div class="card card-soft">
      <div class="card-body">

        <form method="post" action="sensores_deposito_guardar.asp" class="row g-3" onsubmit="return normalizeTemp();">

          <div class="col-12 col-md-6">
            <label class="form-label">Depósito</label>
            <select id="depositoid" name="depositoid" class="form-select" onchange="onDepositoChange(this)" required>
              <option value="0">-- Seleccionar --</option>
              <%
                If Not rsDep Is Nothing Then
                  If Not rsDep.EOF Then
                    Do While Not rsDep.EOF
                      Dim did, dnom
                      did  = CLng(rsDep("DepositoID"))
                      dnom = CStr(rsDep("Nombre"))
              %>
                <option value="<%=did%>"><%=HtmlEncode(dnom)%></option>
              <%
                      rsDep.MoveNext
                    Loop
                  End If
                End If
              %>
            </select>
            <div class="form-text">Ej: Cámara de Helados / Cámara de Chocolates</div>
          </div>

          <div class="col-12 col-md-6">
            <label class="form-label d-flex align-items-center gap-2">
              Posición
              <span id="posSpinner" class="d-none">
                <i class="fa-solid fa-spinner fa-spin text-muted"></i>
              </span>
            </label>

            <select id="posicion" name="posicion" class="form-select" required disabled>
              <option value="">-- Seleccionar --</option>
            </select>

            <div class="form-text">ALTA / MEDIA / BAJA (según depósito)</div>
          </div>

          <!-- Responsable -->
          <div class="col-12 col-md-6">
            <label class="form-label">Responsable</label>
            <select name="responsableid" class="form-select" required>
              <option value="0">-- Seleccionar --</option>
              <%
                If Not rsResp Is Nothing Then
                  If Not rsResp.EOF Then
                    Do While Not rsResp.EOF

                      Dim rid, rnom, tmpId
                      rid = ""
                      rnom = ""

                      ' Buscar ID en varios nombres posibles
                      tmpId = RSField(rsResp, "ResponsableID", "")
                      If tmpId = "" Then tmpId = RSField(rsResp, "ID", "")
                      If tmpId = "" Then tmpId = RSField(rsResp, "Codigo", "")
                      If tmpId = "" Then tmpId = RSField(rsResp, "EmpleadoID", "")

                      If IsNumeric(tmpId) Then rid = CStr(CLng(tmpId))

                      ' Nombre
                      rnom = RSField(rsResp, "Nombre", "")
                      If rnom = "" Then rnom = RSField(rsResp, "Transportista", "")
                      If rnom = "" Then rnom = RSField(rsResp, "Descripcion", "")
                      If rnom = "" Then rnom = rid

                      ' Si no hay ID, NO agregamos option (evita values vacíos)
                      If rid <> "" Then
              %>
                <option value="<%=HtmlEncode(rid)%>"><%=HtmlEncode(CStr(rnom))%></option>
              <%
                      End If

                      rsResp.MoveNext
                    Loop
                  End If
                End If
              %>
            </select>
          </div>

          <div class="col-12 col-md-4">
            <label class="form-label">Temperatura (°C)</label>
            <input
              type="text"
              id="temperatura"
              name="temperatura"
              class="form-control"
              required
              inputmode="decimal"
              placeholder="Ej: -18.50"
              onblur="normalizeTemp();"
            >
            <div class="form-text">Podés escribir con punto o coma: -18.5 / -18,5</div>
          </div>

          <div class="col-12 col-md-8">
            <label class="form-label">Observación (opcional)</label>
            <input type="text" name="observacion" maxlength="300" class="form-control" placeholder="Ej: Puerta abierta / control diario / etc.">
          </div>

          <div class="col-12 d-flex gap-2">
            <button type="submit" class="btn btn-dark">
              <i class="fa-solid fa-floppy-disk me-1"></i> Guardar lectura
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
If Not rsDep Is Nothing Then If rsDep.State = 1 Then rsDep.Close
If Not rsResp Is Nothing Then If rsResp.State = 1 Then rsResp.Close
Set rsDep = Nothing
Set rsResp = Nothing
On Error GoTo 0
%>