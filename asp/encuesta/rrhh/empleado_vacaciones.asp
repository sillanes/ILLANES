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
        Case Else
            EstadoClase = "neu"
    End Select
End Function

Function FirmaClase(fechaFirma)
    If IsNull(fechaFirma) Or IsEmpty(fechaFirma) Or Trim(CStr(fechaFirma & "")) = "" Then
        FirmaClase = "pen"
    Else
        FirmaClase = "ok"
    End If
End Function

Function FirmaTexto(fechaFirma)
    If IsNull(fechaFirma) Or IsEmpty(fechaFirma) Or Trim(CStr(fechaFirma & "")) = "" Then
        FirmaTexto = "PENDIENTE"
    Else
        FirmaTexto = "FIRMADO"
    End If
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

Dim empleadoID
Dim msgOk, msgErr
Dim rsVac, sqlVac
Dim diasDisponibles

empleadoID = CLng("0" & Session("Empleado_EmpleadoID"))
diasDisponibles = 0

On Error Resume Next

Dim cmdDias
Set cmdDias = Server.CreateObject("ADODB.Command")
Set cmdDias.ActiveConnection = conn
cmdDias.CommandType = 4
cmdDias.CommandText = "rrhh.usp_Vacaciones_Empleado_Disponibles"

cmdDias.Parameters.Append cmdDias.CreateParameter("@EmpleadoID", 3, 1, , empleadoID)
cmdDias.Parameters.Append cmdDias.CreateParameter("@DiasDisponibles", 3, 2)

cmdDias.Execute

If Err.Number = 0 Then
    diasDisponibles = CLng("0" & cmdDias.Parameters("@DiasDisponibles").Value)
Else
    diasDisponibles = 0
    Err.Clear
End If

Set cmdDias = Nothing
On Error GoTo 0

msgOk  = Trim(Request.QueryString("msgok") & "")
msgErr = Trim(Request.QueryString("msgerr") & "")

sqlVac = "EXEC rrhh.usp_Vacaciones_Empleado_Sel @EmpleadoID=" & empleadoID
Set rsVac = conn.Execute(sqlVac)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Portal del Empleado - Vacaciones</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link rel="stylesheet" href="estilos.css">
<style>
body{
    margin:0;
    font-family:Arial, Helvetica, sans-serif;
    background:#f4f6f9;
}
.topbar{
    background:#fff;
    border-bottom:1px solid #ddd;
    padding:14px 18px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:12px;
    flex-wrap:wrap;
}
.topbar .titulo{
    font-size:22px;
    font-weight:bold;
    color:#222;
}
.topbar .usuario{
    color:#555;
    font-size:14px;
}
.container{
    max-width:1200px;
    margin:20px auto;
    padding:0 12px;
    box-sizing:border-box;
}
.card{
    border:1px solid #ddd;
    border-radius:10px;
    padding:18px;
    background:#fff;
    margin-bottom:20px;
}
.page-title{
    margin:0;
    font-size:22px;
}
.header-flex{
    display:flex;
    justify-content:space-between;
    align-items:center;
    flex-wrap:wrap;
    gap:10px;
    margin-bottom:14px;
}
.dias-disponibles{
    font-size:14px;
    color:#555;
    background:#eef4ff;
    border:1px solid #cfe0ff;
    padding:8px 12px;
    border-radius:8px;
}
.dias-disponibles strong{
    color:#2b7cff;
    font-size:16px;
}
.table-wrap{
    width:100%;
    overflow-x:auto;
    -webkit-overflow-scrolling:touch;
}
.table{
    width:100%;
    border-collapse:collapse;
    table-layout:auto;
}
.table th,.table td{
    border:1px solid #ddd;
    padding:9px;
    text-align:left;
    vertical-align:top;
    word-break:break-word;
}
.table th{
    background:#000;
    color:#fff;
}
.badge{
    display:inline-block;
    padding:6px 10px;
    border-radius:999px;
    font-size:12px;
    font-weight:bold;
}
.container{max-width:1200px;margin:20px auto;padding:0 12px;box-sizing:border-box}
.badge.ok{background:#d4edda;color:#155724}
.badge.err{background:#f8d7da;color:#721c24}
.badge.pen{background:#fff3cd;color:#856404}
.badge.neu{background:#e2e3e5;color:#383d41}
.btn{
    background:#2b7cff;
    color:#fff;
    text-decoration:none;
    padding:10px 14px;
    border-radius:8px;
    display:inline-block;
    border:none;
    cursor:pointer;
}
.btn:hover{
    background:#1f68d8;
}
.btn-sec{
    background:#6c757d;
    color:#fff;
    text-decoration:none;
    padding:10px 14px;
    border-radius:8px;
    display:inline-block;
    border:none;
    cursor:pointer;
}
.btn-pdf{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    width:38px;
    height:38px;
    border-radius:8px;
    text-decoration:none;
    background:#dc3545;
    color:#fff;
}
.btn-pdf:hover{
    background:#bb2d3b;
    color:#fff;
}
.alert-ok{
    background:#d4edda;
    color:#155724;
    border-radius:8px;
    padding:12px 14px;
    margin-bottom:14px;
}
.alert-error{
    background:#f8d7da;
    color:#721c24;
    border-radius:8px;
    padding:12px 14px;
    margin-bottom:14px;
}

@media (max-width: 768px){
    .container{margin:12px auto;padding:0 10px}
    .card{padding:14px}
    .topbar{padding:12px}
    .topbar .titulo{font-size:20px}
}

@media (max-width: 640px){
    .table thead{
        display:none;
    }

    .table,
    .table tbody,
    .table tr,
    .table td{
        display:block;
        width:100%;
        box-sizing:border-box;
    }

    .table tr{
        border:1px solid #ddd;
        border-radius:10px;
        margin-bottom:14px;
        overflow:hidden;
        background:#fff;
    }

    .table td{
        border:none;
        border-bottom:1px solid #eee;
        padding:10px 12px 10px 44%;
        position:relative;
        min-height:40px;
    }

    .table td:last-child{
        border-bottom:none;
    }

    .table td:before{
        content:attr(data-label);
        position:absolute;
        left:12px;
        top:10px;
        width:38%;
        font-weight:bold;
        color:#555;
        white-space:normal;
        line-height:1.2;
    }

    .table-wrap{
        overflow:visible;
    }

    .mobile-full{
        padding-left:12px !important;
    }

    .mobile-full:before{
        position:static !important;
        display:block;
        width:auto !important;
        margin-bottom:6px;
    }

    .btn,
    .btn-sec{
        width:100%;
        text-align:center;
        box-sizing:border-box;
    }

    .btn-pdf{
        width:100%;
        height:auto;
        padding:10px 14px;
    }
}

@media (max-width: 600px){
    .header-flex{
        flex-direction:column;
        align-items:flex-start;
    }

    .dias-disponibles{
        width:100%;
        text-align:left;
    }
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

    <% If msgOk <> "" Then %>
        <div class="alert-ok"><%=Server.HTMLEncode(msgOk)%></div>
    <% End If %>

    <% If msgErr <> "" Then %>
        <div class="alert-error"><%=Server.HTMLEncode(msgErr)%></div>
    <% End If %>

    <div class="card">
        <div class="header-flex">
            <h2 class="page-title">Mis solicitudes</h2>
            <div class="dias-disponibles">
                Días disponibles: <strong><%=diasDisponibles%></strong>
            </div>
        </div>

        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>VacacionID</th>
                        <th>Fecha desde</th>
                        <th>Fecha hasta</th>
                        <th>Días</th>
                        <th>Estado</th>
                        <th>Fecha solicitud</th>
                        <th>Fecha firma</th>
                        <th>Firma</th>
                        <th>Observaciones</th>
                        <th>Acción</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rsVac.EOF Then
                %>
                    <tr>
                        <td colspan="10" class="mobile-full" data-label="Resultado">Todavía no tenés solicitudes de vacaciones.</td>
                    </tr>
                <%
                Else
                    Dim rutaPDFFirmado, urlPDF, estadoActual, tieneFirma
                    Do Until rsVac.EOF
                        rutaPDFFirmado = ObtenerRutaPDFFirmado(rsVac)
                        urlPDF = "empleado_vacaciones_pdf_ver.asp?ruta=" & UrlEncodeUtf8(rutaPDFFirmado)
                        estadoActual = UCase(Trim(CStr(Nz(rsVac("Estado"), ""))))
                        tieneFirma = (Trim(CStr(Nz(rsVac("FechaFirma"), "") & "")) <> "")
                %>
                    <tr>
                        <td data-label="VacacionID"><%=Nz(rsVac("VacacionID"), "")%></td>
                        <td data-label="Fecha desde"><%=FechaFormato(rsVac("FechaDesde"))%></td>
                        <td data-label="Fecha hasta"><%=FechaFormato(rsVac("FechaHasta"))%></td>
                        <td data-label="Días"><%=Nz(rsVac("CantidadDias"), "")%></td>
                        <td data-label="Estado">
                            <span class="badge <%=EstadoClase(rsVac("Estado"))%>"><%=Server.HTMLEncode(Nz(rsVac("Estado"), ""))%></span>
                        </td>
                        <td data-label="Fecha solicitud"><%=FechaHoraFormato(rsVac("FechaSolicitud"))%></td>
                        <td data-label="Fecha firma"><%=FechaHoraFormato(rsVac("FechaFirma"))%></td>
                        <td data-label="Firma">
                            <span class="badge <%=FirmaClase(rsVac("FechaFirma"))%>"><%=FirmaTexto(rsVac("FechaFirma"))%></span>
                        </td>
                        <td data-label="Observaciones"><%=Server.HTMLEncode(Nz(rsVac("Observaciones"), ""))%></td>
                        <td data-label="Acción">
                            <%
                            If estadoActual = "APROBADO" And Not tieneFirma Then
                            %>
                                <a href="empleado_vacaciones_firmar.asp?vacacionid=<%=CLng(Nz(rsVac("VacacionID"),0))%>" class="btn">Firmar</a>
                            <%
                            ElseIf tieneFirma And rutaPDFFirmado <> "" Then
                            %>
                                <a href="<%=urlPDF%>" target="_blank" class="btn-pdf" title="Ver PDF firmado">
                                    <i class="fa-solid fa-file-pdf"></i>
                                </a>
                            <%
                            Else
                                Response.Write "-"
                            End If
                            %>
                        </td>
                    </tr>
                <%
                        rsVac.MoveNext
                    Loop
                End If
                %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<%
If IsObject(rsVac) Then
    If rsVac.State = 1 Then rsVac.Close
    Set rsVac = Nothing
End If

If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>
</body>
</html>