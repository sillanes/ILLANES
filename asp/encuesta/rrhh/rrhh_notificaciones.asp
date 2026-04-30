<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.Buffer   = True
Server.ScriptTimeout = 600

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

Function ToInt(v, alt)
    On Error Resume Next
    Dim n
    n = alt
    If Trim(CStr(v & "")) <> "" Then
        n = CLng(v)
        If Err.Number <> 0 Then
            n = alt
            Err.Clear
        End If
    End If
    ToInt = n
    On Error GoTo 0
End Function

Function Html(v)
    Html = Server.HTMLEncode(CStr(Nz(v, "")))
End Function

Function FormatearPeriodo(v)
    Dim s, a
    s = Trim(CStr(Nz(v, "")))

    If s = "" Then
        FormatearPeriodo = ""
        Exit Function
    End If

    If InStr(s, "-") > 0 Then
        a = Split(s, "-")
        If UBound(a) = 1 Then
            If Len(a(0)) = 4 Then
                FormatearPeriodo = Right("0" & a(1), 2) & "-" & a(0)
                Exit Function
            ElseIf Len(a(1)) = 4 Then
                FormatearPeriodo = Right("0" & a(0), 2) & "-" & a(1)
                Exit Function
            End If
        End If
    End If

    FormatearPeriodo = s
End Function

Function PeriodoParaSP(v)
    Dim s, a
    s = Trim(CStr(Nz(v, "")))

    If s = "" Then
        PeriodoParaSP = ""
        Exit Function
    End If

    If InStr(s, "-") > 0 Then
        a = Split(s, "-")
        If UBound(a) = 1 Then
            If Len(a(0)) = 2 And Len(a(1)) = 4 Then
                PeriodoParaSP = a(1) & "-" & Right("0" & a(0), 2)
                Exit Function
            End If
        End If
    End If

    PeriodoParaSP = s
End Function

Function MostrarFecha(v)
    If IsNull(v) Or Trim(CStr(v & "")) = "" Then
        MostrarFecha = ""
    Else
        On Error Resume Next
        MostrarFecha = Right("0" & Day(CDate(v)),2) & "/" & Right("0" & Month(CDate(v)),2) & "/" & Year(CDate(v)) & " " & Right("0" & Hour(CDate(v)),2) & ":" & Right("0" & Minute(CDate(v)),2)
        If Err.Number <> 0 Then
            MostrarFecha = CStr(v)
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

Function EstadoClase(estado)
    Dim s
    s = UCase(Trim(CStr(Nz(estado, ""))))

    Select Case s
        Case "FIRMADO"
            EstadoClase = "ok"
        Case "LEIDO", "LEÍDO"
            EstadoClase = "neu"
        Case "NOTIFICADO", "PUBLICADO"
            EstadoClase = "pen"
        Case Else
            EstadoClase = "neu"
    End Select
End Function

Function ObtenerUsuarioEnvioID()
    Dim x
    x = Trim(CStr(Session("UsuarioID") & ""))
    If x <> "" And IsNumeric(x) Then
        ObtenerUsuarioEnvioID = CLng(x)
        Exit Function
    End If

    x = Trim(CStr(Session("currentUserID") & ""))
    If x <> "" And IsNumeric(x) Then
        ObtenerUsuarioEnvioID = CLng(x)
        Exit Function
    End If

    ObtenerUsuarioEnvioID = Null
End Function

Dim mensajeOK, mensajeError
mensajeOK = ""
mensajeError = ""

Dim accion
accion = LCase(Trim(CStr(Request.Form("accion") & "")))

Dim filtroPeriodo, filtroLegajo, filtroNombre, filtroEstado, filtroSoloPendientes
filtroPeriodo        = Trim(CStr(Request("periodo") & ""))
filtroLegajo         = Trim(CStr(Request("legajo") & ""))
filtroNombre         = Trim(CStr(Request("nombre") & ""))
filtroEstado         = Trim(CStr(Request("estado") & ""))
filtroSoloPendientes = Trim(CStr(Request("solopendientes") & ""))

If filtroSoloPendientes = "" Then filtroSoloPendientes = "0"

'====================================================
' Registrar notificación WhatsApp
'====================================================
If accion = "notificar_whatsapp" Then
    Dim wReciboID, wEmpleadoID, wUsuarioEnvioID
    Dim cmdWsp

    wReciboID       = ToInt(Request.Form("reciboid"), 0)
    wEmpleadoID     = ToInt(Request.Form("empleadoid_hidden"), 0)
    wUsuarioEnvioID = ObtenerUsuarioEnvioID()

    If wReciboID <= 0 Or wEmpleadoID <= 0 Then
        mensajeError = "No se pudo identificar el recibo o el empleado."
    Else
        On Error Resume Next

        Set cmdWsp = Server.CreateObject("ADODB.Command")
        Set cmdWsp.ActiveConnection = conn
        cmdWsp.CommandType = 4
        cmdWsp.CommandText = "rrhh.usp_Recibo_Notificacion_WhatsApp_Ins"

        cmdWsp.Parameters.Append cmdWsp.CreateParameter("@ReciboID", 3, 1, , wReciboID)
        cmdWsp.Parameters.Append cmdWsp.CreateParameter("@EmpleadoID", 3, 1, , wEmpleadoID)
        cmdWsp.Parameters.Append cmdWsp.CreateParameter("@UsuarioEnvioID", 3, 1, , wUsuarioEnvioID)

        cmdWsp.Execute

        If Err.Number <> 0 Then
            mensajeError = "Error al registrar la notificación WhatsApp: " & Err.Description
            Err.Clear
        Else
            mensajeOK = "La notificación WhatsApp fue registrada correctamente."
        End If

        Set cmdWsp = Nothing
        On Error GoTo 0
    End If
End If

'====================================================
' Listado
'====================================================
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "rrhh.usp_Recibos_Notificaciones_Sel"

If filtroPeriodo = "" Then
    cmd.Parameters.Append cmd.CreateParameter("@Periodo", 200, 1, 20, Null)
Else
    cmd.Parameters.Append cmd.CreateParameter("@Periodo", 200, 1, 20, PeriodoParaSP(filtroPeriodo))
End If

cmd.Parameters.Append cmd.CreateParameter("@EmpleadoID", 3, 1, , Null)

If filtroLegajo = "" Then
    cmd.Parameters.Append cmd.CreateParameter("@Legajo", 200, 1, 50, Null)
Else
    cmd.Parameters.Append cmd.CreateParameter("@Legajo", 200, 1, 50, filtroLegajo)
End If



If filtroEstado = "" Then
    cmd.Parameters.Append cmd.CreateParameter("@Estado", 200, 1, 20, Null)
Else
    cmd.Parameters.Append cmd.CreateParameter("@Estado", 200, 1, 20, filtroEstado)
End If

cmd.Parameters.Append cmd.CreateParameter("@SoloPendientes", 11, 1, , CBool(filtroSoloPendientes = "1"))

If filtroNombre = "" Then
    cmd.Parameters.Append cmd.CreateParameter("@Nombre", 200, 1, 150, Null)
Else
    cmd.Parameters.Append cmd.CreateParameter("@Nombre", 200, 1, 150, filtroNombre)
End If

Set rs = cmd.Execute
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Notificaciones de recibos</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<style>
.filtros-grid{
    display:grid;
    grid-template-columns:repeat(6, minmax(150px,1fr));
    gap:14px;
    align-items:end;
}
.field label{
    display:block;
    font-size:13px;
    font-weight:600;
    margin-bottom:6px;
}
.field input[type=text],
.field select{
    width:100%;
    padding:10px 12px;
    border:1px solid #d1d5db;
    border-radius:10px;
    font-size:14px;
    background:#fff;
}
.field-check{
    display:flex;
    align-items:end;
    min-height:42px;
}
.field-check label{
    display:flex;
    align-items:center;
    gap:8px;
    margin:0;
    font-weight:400;
}
.alerta{
    padding:12px 14px;
    border-radius:10px;
    margin-bottom:14px;
}
.alerta.ok{
    background:#ecfdf5;
    color:#065f46;
    border:1px solid #a7f3d0;
}
.alerta.err{
    background:#fef2f2;
    color:#991b1b;
    border:1px solid #fecaca;
}
.leyenda-estados{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
    margin-bottom:10px;
}
.leyenda-item{
    display:flex;
    align-items:center;
    gap:8px;
    padding:8px 10px;
    background:#f8fafc;
    border:1px solid #e5e7eb;
    border-radius:10px;
    font-size:13px;
}
.leyenda-texto{
    margin-top:6px;
    font-size:13px;
    color:#6b7280;
    line-height:1.5;
}
.acciones-col{
    min-width:530px;
}
.acciones-inline{
    display:flex;
    align-items:center;
    gap:8px;
    flex-wrap:wrap;
}
.acciones-inline a,
.acciones-inline button{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    gap:6px;
    padding:7px 11px;
    font-size:13px;
    line-height:1.2;
    white-space:nowrap;
}
.acciones-inline form{
    margin:0;
    display:inline-flex;
}
.btn-gray{
    background:#6b7280;
    color:#fff;
    border:none;
    padding:8px 12px;
    border-radius:8px;
    cursor:pointer;
    text-decoration:none;
}
.btn-gray:hover{
    background:#5a6268;
}
.nowrap{
    white-space:nowrap;
}
@media (max-width: 1300px){
    .filtros-grid{ grid-template-columns:repeat(3, minmax(160px,1fr)); }
}
@media (max-width: 768px){
    .filtros-grid{ grid-template-columns:1fr; }
    .acciones-inline{
        flex-direction:column;
        align-items:stretch;
    }
    .acciones-inline a,
    .acciones-inline button,
    .acciones-inline form{
        width:100%;
    }
    .acciones-inline a,
    .acciones-inline button{
        justify-content:center;
    }
    .acciones-col{
        min-width:220px;
    }
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
        <h2 class="page-title">Notificaciones de recibos</h2>
    </div>

    <% If mensajeOK <> "" Then %>
        <div class="alerta ok"><%=Html(mensajeOK)%></div>
    <% End If %>

    <% If mensajeError <> "" Then %>
        <div class="alerta err"><%=Html(mensajeError)%></div>
    <% End If %>

    <div class="card">
        <div class="leyenda-estados">
            <div class="leyenda-item"><span class="badge neu">Pendiente</span> <span>Recibo aún no publicado</span></div>
            <div class="leyenda-item"><span class="badge pen">Publicado</span> <span>Publicado pero sin notificación</span></div>
            <div class="leyenda-item"><span class="badge pen">Notificado</span> <span>Se registró una notificación WhatsApp</span></div>
            <div class="leyenda-item"><span class="badge neu">Leído</span> <span>El empleado abrió el recibo</span></div>
            <div class="leyenda-item"><span class="badge ok">Firmado</span> <span>El recibo fue firmado</span></div>
        </div>
        <div class="leyenda-texto">
            El botón <strong>Notificar</strong> busca el teléfono del empleado y registra un envío en WhatsApp OnDemand usando el template <strong>EmpNotificacion</strong>.
            El botón <strong>Historial</strong> muestra las notificaciones registradas para ese recibo.
        </div>
    </div>

    <div class="card">
        <form method="get" action="rrhh_notificaciones.asp">
            <div class="filtros-grid">

                <div class="field">
                    <label>Período</label>
                    <input type="text" name="periodo" value="<%=Html(FormatearPeriodo(filtroPeriodo))%>" placeholder="MM-YYYY">
                </div>

                <div class="field">
                    <label>Legajo</label>
                    <input type="text" name="legajo" value="<%=Html(filtroLegajo)%>">
                </div>

                <div class="field">
                    <label>Nombre</label>
                    <input type="text" name="nombre" value="<%=Html(filtroNombre)%>" placeholder="Nombre o apellido">
                </div>

                <div class="field">
                    <label>Estado</label>
                    <select name="estado">
                        <option value="" <% If filtroEstado="" Then Response.Write("selected") End If %>>Todos</option>
                        <option value="Pendiente" <% If LCase(filtroEstado)="pendiente" Then Response.Write("selected") End If %>>Pendiente</option>
                        <option value="Publicado" <% If LCase(filtroEstado)="publicado" Then Response.Write("selected") End If %>>Publicado</option>
                        <option value="Notificado" <% If LCase(filtroEstado)="notificado" Then Response.Write("selected") End If %>>Notificado</option>
                        <option value="Leido" <% If LCase(filtroEstado)="leido" Then Response.Write("selected") End If %>>Leído</option>
                        <option value="Firmado" <% If LCase(filtroEstado)="firmado" Then Response.Write("selected") End If %>>Firmado</option>
                    </select>
                </div>

                <div class="field-check">
                    <label>
                        <input type="checkbox" name="solopendientes" value="1" <% If filtroSoloPendientes="1" Then Response.Write("checked") End If %>>
                        Solo faltante de firma
                    </label>
                </div>

                <div class="field">
                    <label>&nbsp;</label>
                    <button type="submit" class="btn"><i class="fa fa-search"></i> Buscar</button>
                </div>

            </div>
        </form>
    </div>

    <div class="card">
        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>Recibo</th>
                        <th>Legajo</th>
                        <th>Nombre</th>
                        <th>Período</th>
                        <th>Estado</th>
                        <th>Publicación</th>
                        <th>Lectura</th>
                        <th>Firma</th>
                        <th>Notif.</th>
                        <th class="acciones-col">Acción</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="10" class="mobile-full" data-label="Resultado">No hay resultados.</td>
                    </tr>
                <%
                Else
                    Do Until rs.EOF

                        Dim reciboID, empleadoID, legajo, nombreCompleto, periodoFmt
                        Dim estadoGeneral, fechaPublicacion, fechaLectura, fechaFirma
                        Dim cantNotif, rutaOriginal, rutaFirmado

                        reciboID         = Nz(rs("ReciboID"), "")
                        empleadoID       = Nz(rs("EmpleadoID"), "")
                        legajo           = Nz(rs("Legajo"), "")
                        nombreCompleto   = Nz(rs("Nombre"), "")
                        periodoFmt       = FormatearPeriodo(Nz(rs("Periodo"), ""))
                        estadoGeneral    = Nz(rs("EstadoGeneral"), "")
                        fechaPublicacion = MostrarFecha(rs("FechaPublicacion"))
                        fechaLectura     = MostrarFecha(rs("FechaLectura"))
                        fechaFirma       = MostrarFecha(rs("FechaFirma"))
                        cantNotif        = Nz(rs("CantidadNotificaciones"), 0)
                        rutaOriginal     = Trim(CStr(Nz(rs("RutaArchivo"), "")))
                        rutaFirmado      = Trim(CStr(Nz(rs("RutaArchivoFirmado"), "")))
                %>
                    <tr>
                        <td data-label="Recibo"><%=Html(reciboID)%></td>
                        <td data-label="Legajo"><%=Html(legajo)%></td>
                        <td data-label="Nombre"><%=Html(nombreCompleto)%></td>
                        <td data-label="Período" class="nowrap"><%=Html(periodoFmt)%></td>
                        <td data-label="Estado">
                            <span class="badge <%=EstadoClase(estadoGeneral)%>"><%=Html(estadoGeneral)%></span>
                        </td>
                        <td data-label="Publicación"><%=Html(fechaPublicacion)%></td>
                        <td data-label="Lectura"><%=Html(fechaLectura)%></td>
                        <td data-label="Firma"><%=Html(fechaFirma)%></td>
                        <td data-label="Notif."><%=Html(cantNotif)%></td>
                        <td data-label="Acción" class="acciones-col">
                            <div class="acciones-inline">

                                <form method="post" action="rrhh_notificaciones.asp">
                                    <input type="hidden" name="accion" value="notificar_whatsapp">
                                    <input type="hidden" name="reciboid" value="<%=Html(reciboID)%>">
                                    <input type="hidden" name="empleadoid_hidden" value="<%=Html(empleadoID)%>">

                                    <input type="hidden" name="periodo" value="<%=Html(FormatearPeriodo(filtroPeriodo))%>">
                                    <input type="hidden" name="legajo" value="<%=Html(filtroLegajo)%>">
                                    <input type="hidden" name="nombre" value="<%=Html(filtroNombre)%>">
                                    <input type="hidden" name="estado" value="<%=Html(filtroEstado)%>">
                                    <input type="hidden" name="solopendientes" value="<%=Html(filtroSoloPendientes)%>">

                                    <button type="submit" class="btn-sec">
                                        <i class="fa fa-bell"></i> Notificar
                                    </button>
                                </form>

                                <% If rutaOriginal <> "" Then %>
                                    <a class="btn-sec" href="rrhh_recibo_pdf.asp?reciboid=<%=Server.URLEncode(CStr(reciboID))%>&tipo=original" target="_blank">
                                        <i class="fa fa-file-pdf"></i> PDF original
                                    </a>
                                <% End If %>

                                <% If rutaFirmado <> "" Then %>
                                    <a class="btn-sec" href="rrhh_recibo_pdf.asp?reciboid=<%=Server.URLEncode(CStr(reciboID))%>&tipo=firmado" target="_blank">
                                        <i class="fa fa-signature"></i> PDF firmado
                                    </a>
                                <% End If %>

                                <a class="btn-gray" href="rrhh_recibo_historial.asp?reciboid=<%=Server.URLEncode(CStr(reciboID))%>">
                                    <i class="fa fa-clock-rotate-left"></i> Historial
                                </a>
                            </div>
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
On Error Resume Next
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
End If
Set rs = Nothing
Set cmd = Nothing
conn.Close
Set conn = Nothing
On Error GoTo 0
%>

</body>
</html>