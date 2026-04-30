<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Trim(Session("currentUser") & "") <> "admin" AND Trim(Session("currentUser") & "") <> "rrhh"  Then
    Response.Redirect "../login.asp"
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
    s = UCase(Trim(CStr(estado & "")))

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

Dim reciboid
reciboid = CLng("0" & Request.QueryString("reciboid"))

If reciboid <= 0 Then
    Response.Write "ReciboID inválido."
    Response.End
End If

Dim rs
Set rs = conn.Execute("EXEC rrhh.usp_Recibo_Detalle_Sel " & reciboid)

If rs.EOF Then
    Response.Write "No existe el recibo indicado."
    Response.End
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Detalle recibo</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
.container{max-width:1200px;margin:20px auto;padding:0 12px;box-sizing:border-box}
.card{border:1px solid #ddd;border-radius:10px;padding:18px;background:#fff;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:12px}
.kpi{border:1px solid #e5e5e5;border-radius:10px;padding:12px;background:#fafafa;min-width:0}
.kpi .lab{font-size:12px;color:#666}
.kpi .val{font-size:18px;font-weight:bold;margin-top:4px;word-break:break-word}
.badge{display:inline-block;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:bold}
.badge.ok{background:#d4edda;color:#155724}
.badge.err{background:#f8d7da;color:#721c24}
.badge.pen{background:#fff3cd;color:#856404}
.badge.dup{background:#d1ecf1;color:#0c5460}
.badge.neu{background:#e2e3e5;color:#383d41}
.btn{background:#2b7cff;color:#fff;text-decoration:none;padding:8px 12px;border-radius:8px;display:inline-block}
.btn-sec{background:#6c757d;color:#fff;text-decoration:none;padding:8px 12px;border-radius:8px;display:inline-block}
.header-row{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:18px}
.acciones-top{display:flex;gap:10px;flex-wrap:wrap}
.page-title{margin:0}
.info-grid{display:grid;grid-template-columns:repeat(2,minmax(280px,1fr));gap:16px}
.info-box{border:1px solid #e5e5e5;border-radius:10px;padding:14px;background:#fff}
.info-box h3{margin:0 0 12px 0;font-size:18px}
.info-row{padding:8px 0;border-bottom:1px solid #f0f0f0;word-break:break-word}
.info-row:last-child{border-bottom:none}
.info-label{display:block;font-size:12px;color:#666;margin-bottom:3px}
.info-value{font-size:15px;color:#222}
.small{font-size:12px;color:#666;word-break:break-word}
.link-pdf{display:inline-block;margin-top:8px;word-break:break-all}

@media (max-width: 900px){
    .grid{grid-template-columns:repeat(2,minmax(160px,1fr))}
    .info-grid{grid-template-columns:1fr}
}

@media (max-width: 768px){
    .container{margin:12px auto;padding:0 10px}
    .card{padding:14px}
    .page-title{font-size:22px;line-height:1.2}
    .acciones-top{width:100%}
    .acciones-top a{flex:1 1 100%;text-align:center;box-sizing:border-box}
    .btn,.btn-sec{width:100%;text-align:center;box-sizing:border-box}
    .grid{grid-template-columns:1fr}
    .kpi .val{font-size:16px}
}
</style>
<script>
function toggleSidebar(){
    var sb = document.querySelector('.sidebar');
    if(sb){ sb.classList.toggle('open'); }
}
</script>
</head>
<body>


<!--#include file="header.asp" -->
 


<div class="main-content">
    <div class="header-row">
        <h2 class="page-title">Detalle recibo #<%=rs("ReciboID")%></h2>
        <div class="acciones-top">
            <a href="rrhh_recibos.asp" class="btn-sec">Volver a recibos</a>
            <a href="rrhh_recibos_lote_detalle.asp?loteid=<%=rs("LoteID")%>" class="btn">Ver lote</a>
        </div>
    </div>

    <div class="card">
        <div class="grid">
            <div class="kpi">
                <div class="lab">ReciboID</div>
                <div class="val"><%=Nz(rs("ReciboID"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Período</div>
                <div class="val"><%=Nz(rs("Periodo"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Estado</div>
                <div class="val">
                    <span class="badge <%=EstadoClase(rs("Estado"))%>"><%=Server.HTMLEncode(Nz(rs("Estado"), ""))%></span>
                </div>
            </div>
            <div class="kpi">
                <div class="lab">Neto</div>
                <div class="val"><%=Nz(rs("Neto"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Fecha pago</div>
                <div class="val"><%=Nz(rs("FechaPago"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Fecha publicación</div>
                <div class="val"><%=Nz(rs("FechaPublicacion"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Fecha lectura</div>
                <div class="val"><%=Nz(rs("FechaLectura"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Fecha firma</div>
                <div class="val"><%=Nz(rs("FechaFirma"), "")%></div>
            </div>
        </div>
    </div>

    <div class="info-grid">
        <div class="card info-box">
            <h3>Empleado</h3>

            <div class="info-row">
                <span class="info-label">EmpleadoID</span>
                <div class="info-value"><%=Nz(rs("EmpleadoID"), "")%></div>
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
                <span class="info-label">DNI</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("DNI"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Apellido y nombre</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("ApellidoNombre"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Email</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("Email"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">WhatsApp</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("TelefonoWhatsApp"), ""))%></div>
            </div>
        </div>

        <div class="card info-box">
            <h3>Archivo / recibo</h3>

            <div class="info-row">
                <span class="info-label">Nombre archivo</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("NombreArchivo"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Ruta archivo</span>
                <div class="info-value small"><%=Server.HTMLEncode(Nz(rs("RutaArchivo"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Hash archivo</span>
                <div class="info-value small"><%=Server.HTMLEncode(Nz(rs("HashArchivo"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Observación firma</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("ObservacionEmpleado"), ""))%></div>
            </div>
        </div>

        <div class="card info-box">
            <h3>Lote origen</h3>

            <div class="info-row">
                <span class="info-label">LoteID</span>
                <div class="info-value"><%=Nz(rs("LoteID"), "")%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Archivo lote</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("NombreArchivoOriginal"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Ruta archivo original</span>
                <div class="info-value small"><%=Server.HTMLEncode(Nz(rs("RutaArchivoOriginal"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Estado lote</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("EstadoLote"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Fecha carga</span>
                <div class="info-value"><%=Nz(rs("FechaCarga"), "")%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Páginas / OK / Error</span>
                <div class="info-value">
                    <%=Nz(rs("CantidadPaginas"), 0)%> /
                    <%=Nz(rs("CantidadOK"), 0)%> /
                    <%=Nz(rs("CantidadError"), 0)%>
                </div>
            </div>
        </div>

        <div class="card info-box">
            <h3>Detalle de importación</h3>

            <div class="info-row">
                <span class="info-label">LoteDetalleID</span>
                <div class="info-value"><%=Nz(rs("LoteDetalleID"), "")%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Página</span>
                <div class="info-value"><%=Nz(rs("PaginaNumero"), "")%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Legajo detectado</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("LegajoDetectado"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">CUIL detectado</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("CUILDetectado"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Nombre detectado</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("ApellidoNombreDetectado"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Estado proceso</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("EstadoProceso"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Mensaje error</span>
                <div class="info-value"><%=Server.HTMLEncode(Nz(rs("MensajeError"), ""))%></div>
            </div>

            <div class="info-row">
                <span class="info-label">Ruta PDF individual</span>
                <div class="info-value small"><%=Server.HTMLEncode(Nz(rs("RutaPDFIndividual"), ""))%></div>
            </div>
        </div>
    </div>
</div>

<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
</body>
</html>