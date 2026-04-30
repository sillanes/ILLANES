<%@ Language="VBScript" CodePage="65001" %>
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

Function SafeHTML(v)
    If IsNull(v) Or v = "" Then
        SafeHTML = ""
    Else
        SafeHTML = Server.HTMLEncode(CStr(v))
    End If
End Function

If Session("currentUser") = "" Then
    Response.Redirect "login.asp"
End If

Function F(name)
  Dim v: v = Request.Form(name)
  If IsNull(v) Then F = "" Else F = Trim(CStr(v))
End Function

Function SqlQuote(v)
  v = Replace(CStr(v), "'", "''")
  SqlQuote = "'" & v & "'"
End Function

Dim accion, msg, errMsg
accion = LCase(F("accion"))
msg = ""
errMsg = ""

On Error Resume Next

' ===========================
' Acciones POST
' ===========================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    If accion = "nuevo" Then
        Dim patenteN, marcaN, modeloN
        patenteN = UCase(F("patente"))
        marcaN   = F("marca")
        modeloN  = F("modelo")

        If patenteN = "" Or marcaN = "" Or modeloN = "" Then
            errMsg = "Completá Patente, Marca y Modelo."
        Else
            Dim sqlIns
            sqlIns = "EXEC dbo.usp_Vehiculos_ABM_Insertar " & _
                     SqlQuote(patenteN) & "," & SqlQuote(marcaN) & "," & SqlQuote(modeloN)

            conn.Execute sqlIns
            If Err.Number <> 0 Then
                errMsg = "Error al insertar: " & Err.Description
                Err.Clear
            Else
                msg = "Vehículo creado correctamente."
            End If
        End If

    ElseIf accion = "editar" Then
        Dim vidE, marcaE, modeloE, activoE
        vidE   = F("vehiculoid")
        marcaE = F("marca")
        modeloE= F("modelo")
        activoE= F("activo") ' "1" o "0"

        If vidE = "" Then
            errMsg = "Falta VehiculoID."
        ElseIf marcaE = "" Or modeloE = "" Then
            errMsg = "Completá Marca y Modelo."
        Else
            Dim sqlUpd
            sqlUpd = "EXEC dbo.usp_Vehiculos_ABM_Actualizar " & _
                     CLng(vidE) & "," & SqlQuote(marcaE) & "," & SqlQuote(modeloE) & "," & CInt(activoE)

            conn.Execute sqlUpd
            If Err.Number <> 0 Then
                errMsg = "Error al actualizar: " & Err.Description
                Err.Clear
            Else
                msg = "Vehículo actualizado correctamente."
            End If
        End If

    ElseIf accion = "toggle" Then
        Dim vidT, actT
        vidT = F("vehiculoid")
        actT = F("activo")

        If vidT <> "" Then
            Dim sqlT
            sqlT = "EXEC dbo.usp_Vehiculos_ABM_SetActivo " & CLng(vidT) & "," & CInt(actT)
            conn.Execute sqlT
            If Err.Number <> 0 Then
                errMsg = "Error al cambiar activo: " & Err.Description
                Err.Clear
            Else
                msg = "Estado actualizado."
            End If
        End If
    End If

End If

On Error GoTo 0

' ===========================
' Listado
' ===========================
Dim rs, sql
sql = "EXEC dbo.usp_Vehiculos_ABM_Listar"
Set rs = conn.Execute(sql)
%>

<!--#include file="sidebar.asp" -->
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Vehículos – ABM</title>

<link rel="stylesheet" href="estilos.css" />
<link rel="stylesheet" href="ventas.css" />
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<style>
body{background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;margin:0;color:#111827}
.container{max-width:1100px;margin:20px auto;padding:0 16px}
.card{background:#fff;border-radius:16px;box-shadow:0 10px 25px rgba(0,0,0,.08);padding:18px;margin-bottom:14px}
.row{display:flex;gap:14px;flex-wrap:wrap}
.col{flex:1 1 360px}
.title{font-size:22px;font-weight:800;margin:0 0 10px;color:#1f2937}
.muted{color:#6b7280}

label{display:block;font-weight:700;margin:10px 0 6px}
input[type="text"], select{
  width:100%;padding:10px 12px;border:1px solid #d1d5db;border-radius:12px;font-size:14px;background:#fff
}

.btn{border:0;border-radius:12px;padding:10px 14px;cursor:pointer;font-weight:800;transition:.15s ease}
.btn-primary{background:#2563eb;color:#fff}
.btn-ghost{background:#fff;border:1px solid #d1d5db;color:#111827}
.btn-danger{background:#dc2626;color:#fff}
.btn-soft{background:#eef2ff;color:#1e40af}

.table{width:100%;border-collapse:collapse}
.table th,.table td{padding:10px;border-bottom:1px solid #e5e7eb;text-align:left;font-size:14px}
.badge{display:inline-block;padding:3px 10px;border-radius:999px;font-weight:800;font-size:12px}
.badge-on{background:#dcfce7;color:#166534}
.badge-off{background:#fee2e2;color:#991b1b}

.alert-ok{background:#dcfce7;color:#166534;padding:10px 12px;border-radius:12px;font-weight:700}
.alert-err{background:#fee2e2;color:#991b1b;padding:10px 12px;border-radius:12px;font-weight:700}

.actions{display:flex;gap:8px;flex-wrap:wrap}
small.note{display:block;color:#6b7280;margin-top:6px}
hr.sep{border:0;border-top:1px solid #e5e7eb;margin:14px 0}
</style>

<script>
function cargarEdicion(vid, patente, marca, modelo, activo){
  document.getElementById('edit_vehiculoid').value = vid;
  document.getElementById('edit_patente').value = patente;
  document.getElementById('edit_marca').value = marca;
  document.getElementById('edit_modelo').value = modelo;
  document.getElementById('edit_activo').value = activo;

  document.getElementById('cardEditar').scrollIntoView({behavior:'smooth', block:'start'});
}

function limpiarNuevo(){
  document.getElementById('n_patente').value = '';
  document.getElementById('n_marca').value = '';
  document.getElementById('n_modelo').value = '';
}
</script>

</head>
<body>
<!--#include file="header.asp" -->

<div class="container">

  <div class="card">
    <div class="row" style="align-items:flex-start">
      <div class="col" style="flex:1 1 520px">
        <h1 class="title">ABM Vehículos</h1>
        <div class="muted">Alta / Modificación / Activar–Desactivar. No se borra, se marca como inactivo.</div>
      </div>
      <div class="col" style="flex:1 1 320px">
        <% If msg <> "" Then %>
          <div class="alert-ok"><i class="fa-solid fa-circle-check"></i> <%=Server.HTMLEncode(msg)%></div>
        <% End If %>
        <% If errMsg <> "" Then %>
          <div class="alert-err"><i class="fa-solid fa-triangle-exclamation"></i> <%=Server.HTMLEncode(errMsg)%></div>
        <% End If %>
      </div>
    </div>
  </div>

  <div class="row">

    <!-- =========================
         LISTADO
         ========================= -->
    <div class="col">
      <div class="card">
        <div class="title" style="font-size:18px">Vehículos</div>

        <table class="table">
          <thead>
            <tr>
              <th>Patente</th>
              <th>Marca</th>
              <th>Modelo</th>
              <th>Activo</th>
              <th>Acciones</th>
            </tr>
          </thead>
          <tbody>
          <%
            If rs.EOF Then
              Response.Write "<tr><td colspan='5' class='muted'>No hay vehículos.</td></tr>"
            Else
              Do While Not rs.EOF
                Dim vid, pat, mar, modl, act
                vid = rs("VehiculoID")
                pat = rs("Patente")
                mar = rs("Marca")
                modl= rs("Modelo")
                act = rs("Activo")
          %>
            <tr>
              <td><strong><%=Server.HTMLEncode(pat)%></strong></td>
              <td><%=SafeHTML(mar)%></td>
              <td><%=SafeHTML(modl)%></td>
              <td>
                <% If act = True Or act = 1 Then %>
                  <span class="badge badge-on">SI</span>
                <% Else %>
                  <span class="badge badge-off">NO</span>
                <% End If %>
              </td>
              <td>
                <div class="actions">
                  <button type="button" class="btn btn-soft"
                    onclick="cargarEdicion('<%=vid%>','<%=Replace(pat,"'","\'")%>','<%=Replace("" & mar,"'","\'")%>','<%=Replace("" & modl,"'","\'")%>','<%=IIf(act=True Or act=1, "1","0")%>')">
                    <i class="fa-solid fa-pen-to-square"></i> Editar
                  </button>

                  <form method="post" style="margin:0">
                    <input type="hidden" name="accion" value="toggle" />
                    <input type="hidden" name="vehiculoid" value="<%=vid%>" />
                    <input type="hidden" name="activo" value="<%=IIf(act=True Or act=1, 0, 1)%>" />
                    <% If act = True Or act = 1 Then %>
                      <button class="btn btn-danger" type="submit"><i class="fa-solid fa-ban"></i> Desactivar</button>
                    <% Else %>
                      <button class="btn btn-primary" type="submit"><i class="fa-solid fa-check"></i> Activar</button>
                    <% End If %>
                  </form>
                </div>
              </td>
            </tr>
          <%
                rs.MoveNext
              Loop
            End If
            rs.Close : Set rs = Nothing
          %>
          </tbody>
        </table>

        <small class="note">Tip: “Desactivar” no borra registros. Solo quita el vehículo del uso operativo.</small>
      </div>
    </div>

    <!-- =========================
         ALTA + EDICIÓN
         ========================= -->
    <div class="col">

      <!-- ALTA -->
      <div class="card">
        <div class="title" style="font-size:18px">Nuevo vehículo</div>
        <form method="post" action="vehiculos_abm.asp">
          <input type="hidden" name="accion" value="nuevo" />

          <label>Patente</label>
          <input type="text" id="n_patente" name="patente" maxlength="20" placeholder="Ej: AD213AH" required />

          <label>Marca</label>
          <input type="text" id="n_marca" name="marca" maxlength="50" placeholder="Ej: Peugeot" required />

          <label>Modelo</label>
          <input type="text" id="n_modelo" name="modelo" maxlength="50" placeholder="Ej: Partner" required />

          <div class="actions" style="margin-top:12px">
            <button class="btn btn-primary" type="submit"><i class="fa-solid fa-plus"></i> Crear</button>
            <button class="btn btn-ghost" type="button" onclick="limpiarNuevo()">Limpiar</button>
          </div>
        </form>
      </div>

      <!-- EDICIÓN -->
      <div class="card" id="cardEditar">
        <div class="title" style="font-size:18px">Editar vehículo</div>
        <div class="muted">La patente no se modifica. Solo marca, modelo y activo.</div>
        <hr class="sep" />

        <form method="post" action="vehiculos_abm.asp">
          <input type="hidden" name="accion" value="editar" />
          <input type="hidden" id="edit_vehiculoid" name="vehiculoid" value="" />

          <label>Patente</label>
          <input type="text" id="edit_patente" value="" disabled />

          <label>Marca</label>
          <input type="text" id="edit_marca" name="marca" maxlength="50" required />

          <label>Modelo</label>
          <input type="text" id="edit_modelo" name="modelo" maxlength="50" required />

          <label>Activo</label>
          <select id="edit_activo" name="activo">
            <option value="1">SI</option>
            <option value="0">NO</option>
          </select>

          <div class="actions" style="margin-top:12px">
            <button class="btn btn-primary" type="submit"><i class="fa-solid fa-floppy-disk"></i> Guardar</button>
            <button class="btn btn-ghost" type="button" onclick="window.scrollTo({top:0,behavior:'smooth'})">Arriba</button>
          </div>

          <small class="note">Para editar, hacé click en “Editar” desde la grilla.</small>
        </form>
      </div>

    </div>
  </div>

</div>

<%
conn.Close : Set conn = Nothing
%>
</body>
</html>