<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar_empleados.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If CLng("0" & Session("Empleado_UsuarioID")) <= 0 Or CLng("0" & Session("Empleado_EmpleadoID")) <= 0 Then
    Response.Redirect "empleado_login.asp"
End If

Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function EstadoClase(estado)
    Dim s
    If IsNull(estado) Or IsEmpty(estado) Then
        s = ""
    Else
        s = UCase(Trim(CStr(estado)))
    End If

    Select Case s
        Case "OK", "PROCESADO", "DISPONIBLE", "NOTIFICADO", "LEIDO", "FIRMADO"
            EstadoClase = "ok"
        Case "ERROR"
            EstadoClase = "err"
        Case "PENDIENTE", "PROCESANDO"
            EstadoClase = "pen"
        Case "DUPLICADO"
            EstadoClase = "dup"
        Case Else
            EstadoClase = "neu"
    End Select
End Function

Function PeriodoFormato(p)
    If IsNull(p) Or Trim(CStr(p & "")) = "" Then
        PeriodoFormato = ""
    Else
        Dim a
        a = Split(CStr(p), "-")
        If UBound(a) = 1 Then
            PeriodoFormato = Right("0" & a(1), 2) & "-" & a(0)
        Else
            PeriodoFormato = CStr(p)
        End If
    End If
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

Dim empleadoID, reciboid
empleadoID = CLng("0" & Session("Empleado_EmpleadoID"))
reciboid   = CLng("0" & Request("reciboid"))

If reciboid <= 0 Then
    Response.Write "ReciboID inválido."
    Response.End
End If

Dim rs
Set rs = conn.Execute("EXEC empleado.usp_Empleado_Recibo_Detalle_Sel " & empleadoID & ", " & reciboid)

If rs.EOF Then
    Response.Write "No existe el recibo indicado o no pertenece al empleado."
    Response.End
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Portal del Empleado - Detalle recibo</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link rel="stylesheet" href="estilos.css">
<style>
body{margin:0;font-family:Arial, Helvetica, sans-serif;background:#f4f6f9}
.topbar{background:#fff;border-bottom:1px solid #ddd;padding:14px 18px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
.topbar .titulo{font-size:22px;font-weight:bold;color:#222}
.topbar .usuario{color:#555;font-size:14px}
.container{max-width:1100px;margin:20px auto;padding:0 12px;box-sizing:border-box}
.card{border:1px solid #ddd;border-radius:10px;padding:18px;background:#fff;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:12px}
.kpi{border:1px solid #e5e5e5;border-radius:10px;padding:12px;background:#fafafa;min-width:0}
.kpi .lab{font-size:12px;color:#666}
.kpi .val{font-size:18px;font-weight:bold;margin-top:4px;word-break:break-word}
.info-grid{display:grid;grid-template-columns:repeat(2,minmax(280px,1fr));gap:16px}
.info-box{border:1px solid #e5e5e5;border-radius:10px;padding:14px;background:#fff}
.info-box h3{margin:0 0 12px 0;font-size:18px}
.info-row{padding:8px 0;border-bottom:1px solid #f0f0f0;word-break:break-word}
.info-row:last-child{border-bottom:none}
.info-label{display:block;font-size:12px;color:#666;margin-bottom:3px}
.info-value{font-size:15px;color:#222}
.badge{display:inline-block;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:bold}
.badge.ok{background:#d4edda;color:#155724}
.badge.err{background:#f8d7da;color:#721c24}
.badge.pen{background:#fff3cd;color:#856404}
.badge.dup{background:#d1ecf1;color:#0c5460}
.badge.neu{background:#e2e3e5;color:#383d41}
.btn{background:#2b7cff;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.btn-sec{background:#6c757d;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}
@media (max-width: 900px){
    .grid{grid-template-columns:repeat(2,minmax(160px,1fr))}
    .info-grid{grid-template-columns:1fr}
}
@media (max-width: 768px){
    .container{margin:12px auto;padding:0 10px}
    .card{padding:14px}
    .grid{grid-template-columns:1fr}
    .kpi .val{font-size:16px}
    .actions a{width:100%;text-align:center;box-sizing:border-box}
    .topbar{padding:12px}
    .topbar .titulo{font-size:20px}
}
</style>
<script>
function toggleSidebar(){
    document.querySelector('.sidebar').classList.toggle('open');
}
</script>
</head>
<body>

<!--#include file="header_empleado.asp" --> 


<div class="main-content">

    <div class="card">
        <div class="grid">
      
            <div class="kpi">
                <div class="lab">Período</div>
                <div class="val"><%=PeriodoFormato(rs("Periodo"))%></div>
            </div>
            <div class="kpi">
                <div class="lab">Estado</div>
                <div class="val">
                    <span class="badge <%=EstadoClase(Nz(rs("Estado"), ""))%>"><%=Server.HTMLEncode(Nz(rs("Estado"), ""))%></span>
                </div>
            </div>
     
            <div class="kpi">
                <div class="lab">Fecha publicación</div>
                <div class="val"><%=Nz(rs("FechaPublicacion"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Fecha lectura</div>
                <div class="val"><%=FechaHoraFormato(rs("FechaLectura"))%></div>
            </div>
            <div class="kpi">
                <div class="lab">Fecha firma</div>
                <div class="val"><%=FechaHoraFormato(rs("FechaFirma"))%></div>
            </div>
        </div>
    </div>

    <div class="info-grid">
        <div class="card info-box">
            <h3>Datos del recibo</h3>

            <div class="info-row">
                <span class="info-label">Empleado</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("ApellidoNombre"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Legajo</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("Legajo"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">CUIL</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("CUIL"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Nombre archivo</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("NombreArchivo"), ""))%></div>
            </div>
        </div>

        <div class="card info-box">
            <h3>Acciones</h3>

            <div class="actions">
                <a href="empleado_recibo_pdf.asp?reciboid=<%=reciboid%>&modo=original" target="_blank" class="btn">Ver PDF</a>
            </div>

            <div class="actions">
                <a href="empleado_recibo_firmar.asp?reciboid=<%=reciboid%>" class="btn">Firmar recibo</a>
            </div>
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