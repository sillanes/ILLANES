<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Trim(Session("currentUser") & "") = "" Then
    Response.Redirect "../login.asp"
End If

If Trim(Session("currentUser") & "") <> "admin" And Trim(Session("currentUser") & "") <> "rrhh" Then
    Response.Redirect "login.asp"
End If

Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function SqlSafe(v)
    SqlSafe = Replace(Trim(CStr(v & "")), "'", "''")
End Function

Function UrlEncodeUtf8(txt)
    UrlEncodeUtf8 = Server.URLEncode(CStr(txt & ""))
End Function

Function FechaFormato(f)
    If IsNull(f) Or IsEmpty(f) Or Trim(CStr(f & "")) = "" Then
        FechaFormato = ""
    Else
        On Error Resume Next
        FechaFormato = Right("0" & Day(CDate(f)),2) & "-" & Right("0" & Month(CDate(f)),2) & "-" & Year(CDate(f))
        If Err.Number <> 0 Then
            Err.Clear
            FechaFormato = CStr(f)
        End If
        On Error GoTo 0
    End If
End Function

Function FechaHoraFormato(f)
    If IsNull(f) Or IsEmpty(f) Or Trim(CStr(f & "")) = "" Then
        FechaHoraFormato = ""
    Else
        On Error Resume Next
        FechaHoraFormato = Right("0" & Day(CDate(f)),2) & "-" & _
                           Right("0" & Month(CDate(f)),2) & "-" & _
                           Year(CDate(f)) & " " & _
                           Right("0" & Hour(CDate(f)),2) & ":" & _
                           Right("0" & Minute(CDate(f)),2)
        If Err.Number <> 0 Then
            Err.Clear
            FechaHoraFormato = CStr(f)
        End If
        On Error GoTo 0
    End If
End Function

Function EstadoClase(estado)
    Dim s
    s = UCase(Trim(CStr(estado & "")))

    Select Case s
        Case "APROBADO"
            EstadoClase = "ok"
        Case "RECHAZADO"
            EstadoClase = "err"
        Case "PENDIENTE"
            EstadoClase = "pen"
        Case "FIRMADO"
            EstadoClase = "ok"
        Case Else
            EstadoClase = "neu"
    End Select
End Function

Function CampoExiste(rsObj, campo)
    On Error Resume Next
    Dim tmp
    tmp = rsObj.Fields(campo).Name
    CampoExiste = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Function GetCampo(rsObj, campo, alt)
    On Error Resume Next
    Dim v
    v = rsObj.Fields(campo).Value
    If Err.Number <> 0 Then
        Err.Clear
        GetCampo = alt
    Else
        If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
            GetCampo = alt
        Else
            GetCampo = v
        End If
    End If
    On Error GoTo 0
End Function

Function ObtenerRutaPDFFirmado(rsObj)
    Dim ruta
    ruta = ""

    If CampoExiste(rsObj, "RutaArchivoFirmado") Then
        ruta = Trim(CStr(GetCampo(rsObj, "RutaArchivoFirmado", "") & ""))
    End If

    If ruta = "" And CampoExiste(rsObj, "RutaArchivo") Then
        ruta = Trim(CStr(GetCampo(rsObj, "RutaArchivo", "") & ""))
    End If

    If ruta = "" And CampoExiste(rsObj, "RutaPDFFirmado") Then
        ruta = Trim(CStr(GetCampo(rsObj, "RutaPDFFirmado", "") & ""))
    End If

    If ruta = "" And CampoExiste(rsObj, "RutaFirmado") Then
        ruta = Trim(CStr(GetCampo(rsObj, "RutaFirmado", "") & ""))
    End If

    If ruta = "" And CampoExiste(rsObj, "ArchivoFirmado") Then
        ruta = Trim(CStr(GetCampo(rsObj, "ArchivoFirmado", "") & ""))
    End If

    ObtenerRutaPDFFirmado = ruta
End Function

Function EstaFirmado(rsObj)
    Dim estadoFirma, fechaFirma
    estadoFirma = UCase(Trim(CStr(GetCampo(rsObj, "EstadoFirma", "") & "")))
    fechaFirma  = Trim(CStr(GetCampo(rsObj, "FechaFirma", "") & ""))

    If estadoFirma = "FIRMADO" Or fechaFirma <> "" Then
        EstaFirmado = True
    Else
        EstaFirmado = False
    End If
End Function

Dim fechaDesde, fechaHasta, estadoFiltro
Dim msgOk, msgErr
Dim rs, sql

fechaDesde   = Trim(Request("fechadesde"))
fechaHasta   = Trim(Request("fechahasta"))
estadoFiltro = Trim(Request("estadofiltro"))

msgOk  = Trim(Request.QueryString("msgok"))
msgErr = Trim(Request.QueryString("msgerr"))

Dim fechaDesdeSQL, fechaHastaSQL, estadoFiltroSQL

If fechaDesde = "" Then
    fechaDesdeSQL = "NULL"
Else
    fechaDesdeSQL = "'" & SqlSafe(fechaDesde) & "'"
End If

If fechaHasta = "" Then
    fechaHastaSQL = "NULL"
Else
    fechaHastaSQL = "'" & SqlSafe(fechaHasta) & "'"
End If

If Trim(estadoFiltro) = "" Then
    estadoFiltroSQL = "NULL"
Else
    estadoFiltroSQL = "'" & SqlSafe(estadoFiltro) & "'"
End If

sql = "EXEC rrhh.usp_Vacaciones_Resumen_Sel " & _
      "@FechaDesde=" & fechaDesdeSQL & ", " & _
      "@FechaHasta=" & fechaHastaSQL & ", " & _
      "@Estado=" & estadoFiltroSQL

Set rs = conn.Execute(sql)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Vacaciones</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<style>
.alert-ok{background:#d4edda;color:#155724;border-radius:8px;padding:12px 14px;margin-bottom:14px}
.alert-error{background:#f8d7da;color:#721c24;border-radius:8px;padding:12px 14px;margin-bottom:14px}
.badge{display:inline-block;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:bold}
.badge.ok{background:#d4edda;color:#155724}
.badge.err{background:#f8d7da;color:#721c24}
.badge.pen{background:#fff3cd;color:#856404}
.badge.neu{background:#e2e3e5;color:#383d41}
.btn-info{background:#0d6efd;color:#fff;text-decoration:none;padding:8px 12px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.acciones-cell{text-align:center; white-space:nowrap;}
.icon-btn{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    width:34px;
    height:34px;
    border-radius:8px;
    text-decoration:none;
    background:#dc3545;
    color:#fff;
    transition:all .2s ease;
}
.icon-btn:hover{
    background:#bb2d3b;
    color:#fff;
    transform:translateY(-1px);
}
.icon-btn.disabled{
    background:#adb5bd;
    cursor:default;
    pointer-events:none;
}
</style>

<script>
function toggleSidebar(){
    document.querySelector('.sidebar').classList.toggle('open');
}
</script>
</head>

<!--#include file="header.asp" -->

<div class="main-content"> 
    <div class="header-row">
        <h2 class="page-title">Vacaciones</h2>
        <div style="display:flex; gap:10px; flex-wrap:wrap;">
            <a href="rrhh_vacaciones_dashboard.asp" class="btn-info">
                <i class="fa-solid fa-calendar-days"></i> Dashboard calendario
            </a>
        </div>
    </div>

    <% If msgOk <> "" Then %>
        <div class="alert-ok"><%=Server.HTMLEncode(msgOk)%></div>
    <% End If %>

    <% If msgErr <> "" Then %>
        <div class="alert-error"><%=Server.HTMLEncode(msgErr)%></div>
    <% End If %>

    <div class="card">
        <form method="get" action="rrhh_vacaciones.asp">
            <div class="form-grid">
                <div class="campo">
                    <label for="fechadesde">Fecha desde</label>
                    <input type="date" name="fechadesde" id="fechadesde" value="<%=Server.HTMLEncode(fechaDesde)%>">
                </div>
                <div class="campo">
                    <label for="fechahasta">Fecha hasta</label>
                    <input type="date" name="fechahasta" id="fechahasta" value="<%=Server.HTMLEncode(fechaHasta)%>">
                </div>
                <div class="campo">
                    <label for="estadofiltro">Estado</label>
                    <select name="estadofiltro" id="estadofiltro">
                        <option value="">Todos</option>
                        <option value="PENDIENTE" <% If estadoFiltro="PENDIENTE" Then Response.Write "selected" End If %>>Pendiente</option>
                        <option value="APROBADO" <% If estadoFiltro="APROBADO" Then Response.Write "selected" End If %>>Aprobado</option>
                        <option value="RECHAZADO" <% If estadoFiltro="RECHAZADO" Then Response.Write "selected" End If %>>Rechazado</option>
                    </select>
                </div>
            </div>

            <div class="filtros-acciones">
                <button type="submit" class="btn">Filtrar</button>
                <a href="rrhh_vacaciones.asp" class="btn-sec">Limpiar</a>
            </div>
        </form>
    </div>

    <div class="card">
        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Empleado</th>
                        <th>Legajo</th>
                        <th>CUIL</th>
                        <th>Teléfono</th>
                        <th>Desde</th>
                        <th>Hasta</th>
                        <th>Días</th>
                        <th>Estado</th>
                        <th>Solicitud</th>
                        <th>Fecha Firma</th>
                        <th>Estado Firma</th>
                        <th>PDF</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="13">No hay solicitudes para los filtros seleccionados.</td>
                    </tr>
                <%
                Else
                    Dim rutaPDFFirmado, mostrarPDF, urlPDF
                    Do Until rs.EOF
                        rutaPDFFirmado = ObtenerRutaPDFFirmado(rs)
                        mostrarPDF = (EstaFirmado(rs) And rutaPDFFirmado <> "")
                        urlPDF = "rrhh_vacaciones_pdf_ver.asp?ruta=" & UrlEncodeUtf8(rutaPDFFirmado)
                %>
                    <tr>
                        <td><%=Nz(rs("VacacionID"), "")%></td>
                        <td><%=Server.HTMLEncode(Nz(rs("ApellidoNombre"), ""))%></td>
                        <td><%=Server.HTMLEncode(Nz(rs("Legajo"), ""))%></td>
                        <td><%=Server.HTMLEncode(Nz(rs("CUIL"), ""))%></td>
                        <td><%=Server.HTMLEncode(Nz(rs("TelefonoWhatsApp"), ""))%></td>
                        <td><%=FechaFormato(rs("FechaDesde"))%></td>
                        <td><%=FechaFormato(rs("FechaHasta"))%></td>
                        <td><%=Nz(rs("CantidadDias"), "")%></td>
                        <td>
                            <span class="badge <%=EstadoClase(rs("Estado"))%>"><%=Server.HTMLEncode(Nz(rs("Estado"), ""))%></span>
                        </td>
                        <td><%=FechaHoraFormato(rs("FechaSolicitud"))%></td>
                        <td><%=FechaHoraFormato(rs("FechaFirma"))%></td>
                        <td>
                            <span class="badge <%=EstadoClase(rs("EstadoFirma"))%>"><%=Server.HTMLEncode(Nz(rs("EstadoFirma"), ""))%></span>
                        </td>
                        <td class="acciones-cell">
                            <% If mostrarPDF Then %>
                                <a href="<%=urlPDF%>" target="_blank" class="icon-btn" title="Ver PDF firmado">
                                    <i class="fa-solid fa-file-pdf"></i>
                                </a>
                            <% Else %>
                                <span class="icon-btn disabled" title="Sin PDF firmado disponible">
                                    <i class="fa-solid fa-file-pdf"></i>
                                </span>
                            <% End If %>
                        </td>
                    </tr>
                <%
                        rs.MoveNext
                    Loop
                End If
                %>
                </tbody>
            </table>
        </div>
    </div>
</div> 

<%
If IsObject(rs) Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>
</body>
</html>