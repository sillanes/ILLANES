<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Trim(Session("currentUser") & "") = "" Then
    Response.Redirect "login.asp"
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

Dim vacacionID, motivoRechazo, usuarioResolucionID
Dim fechaDesde, fechaHasta, estadoFiltro
Dim msgErr, sql, rs

vacacionID           = CLng("0" & Request("vacacionid"))
motivoRechazo        = Trim(Request("motivorechazo"))
usuarioResolucionID  = CLng("0" & Session("UsuarioID"))
fechaDesde           = Trim(Request("fechadesde"))
fechaHasta           = Trim(Request("fechahasta"))
estadoFiltro         = Trim(Request("estadofiltro"))
msgErr               = ""

If vacacionID <= 0 Then
    Response.Redirect "rrhh_vacaciones.asp"
End If

If LCase(Trim(Request("accion"))) = "guardar" Then
    If Trim(motivoRechazo) = "" Then
        msgErr = "Debés indicar el motivo del rechazo."
    Else
        On Error Resume Next

        Dim sqlRes
        sqlRes = "EXEC rrhh.usp_Vacaciones_Resolver " & _
                 "@VacacionID=" & vacacionID & ", " & _
                 "@Estado='RECHAZADO', " & _
                 "@MotivoRechazo='" & SqlSafe(motivoRechazo) & "', " & _
                 "@UsuarioResolucionID=" & usuarioResolucionID

        conn.Execute sqlRes

        If Err.Number <> 0 Then
            msgErr = "No se pudo rechazar la solicitud: " & Err.Description
            Err.Clear
        Else
            Response.Redirect "rrhh_vacaciones.asp?fechadesde=" & Server.URLEncode(fechaDesde) & _
                              "&fechahasta=" & Server.URLEncode(fechaHasta) & _
                              "&estadofiltro=" & Server.URLEncode(estadoFiltro) & _
                              "&msgok=" & Server.URLEncode("La solicitud fue rechazada y se envió la notificación por WhatsApp.")
        End If

        On Error GoTo 0
    End If
End If

sql = "EXEC rrhh.usp_Vacaciones_Resumen_Sel @FechaDesde=NULL, @FechaHasta=NULL, @Estado=NULL"
Set rs = conn.Execute(sql)

Do Until rs.EOF
    If CLng("0" & rs("VacacionID")) = vacacionID Then Exit Do
    rs.MoveNext
Loop

If rs.EOF Then
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    Response.Redirect "rrhh_vacaciones.asp?msgerr=" & Server.URLEncode("No se encontró la solicitud.")
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Rechazar vacaciones</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<style>
.alert-error{background:#f8d7da;color:#721c24;border-radius:8px;padding:12px 14px;margin-bottom:14px}
.info-grid{display:grid;grid-template-columns:repeat(2,minmax(220px,1fr));gap:12px}
.info-box{background:#fafafa;border:1px solid #ddd;border-radius:10px;padding:12px}
.info-lab{font-size:12px;color:#666;margin-bottom:4px}
.info-val{font-size:15px;color:#222;font-weight:bold;word-break:break-word}
.campo textarea{width:100%;min-height:140px;padding:10px;border:1px solid #ccc;border-radius:8px;box-sizing:border-box;resize:vertical}
@media (max-width:768px){
    .info-grid{grid-template-columns:1fr}
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
<div class="container">

    <div class="header-row">
        <h2 class="page-title">Rechazar solicitud de vacaciones</h2>
    </div>

    <% If msgErr <> "" Then %>
        <div class="alert-error"><%=Server.HTMLEncode(msgErr)%></div>
    <% End If %>

    <div class="card">
        <div class="info-grid">
            <div class="info-box">
                <div class="info-lab">Empleado</div>
                <div class="info-val"><%=Server.HTMLEncode(Nz(rs("ApellidoNombre"), ""))%></div>
            </div>
            <div class="info-box">
                <div class="info-lab">Legajo</div>
                <div class="info-val"><%=Server.HTMLEncode(Nz(rs("Legajo"), ""))%></div>
            </div>
            <div class="info-box">
                <div class="info-lab">Desde</div>
                <div class="info-val"><%=FechaFormato(rs("FechaDesde"))%></div>
            </div>
            <div class="info-box">
                <div class="info-lab">Hasta</div>
                <div class="info-val"><%=FechaFormato(rs("FechaHasta"))%></div>
            </div>
        </div>
    </div>

    <div class="card">
        <form method="post" action="rrhh_vacaciones_rechazar.asp">
            <input type="hidden" name="accion" value="guardar">
            <input type="hidden" name="vacacionid" value="<%=vacacionID%>">
            <input type="hidden" name="fechadesde" value="<%=Server.HTMLEncode(fechaDesde)%>">
            <input type="hidden" name="fechahasta" value="<%=Server.HTMLEncode(fechaHasta)%>">
            <input type="hidden" name="estadofiltro" value="<%=Server.HTMLEncode(estadoFiltro)%>">

            <div class="campo">
                <label for="motivorechazo">Motivo del rechazo</label>
                <textarea name="motivorechazo" id="motivorechazo"><%=Server.HTMLEncode(motivoRechazo)%></textarea>
            </div>

            <div class="filtros-acciones" style="margin-top:14px;">
                <button type="submit" class="btn-err">Confirmar rechazo</button>
                <a href="rrhh_vacaciones.asp?fechadesde=<%=Server.URLEncode(fechaDesde)%>&fechahasta=<%=Server.URLEncode(fechaHasta)%>&estadofiltro=<%=Server.URLEncode(estadoFiltro)%>" class="btn-sec">Cancelar</a>
            </div>
        </form>
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