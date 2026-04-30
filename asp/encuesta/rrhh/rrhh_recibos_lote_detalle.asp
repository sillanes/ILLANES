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

Dim loteid
loteid = CLng("0" & Request.QueryString("loteid"))

If loteid <= 0 Then
    Response.Write "LoteID inválido."
    Response.End
End If

Dim rsCab, rsDet
Set rsCab = conn.Execute("EXEC rrhh.usp_RecibosLote_Cabecera_Sel " & loteid)

If rsCab.EOF Then
    Response.Write "No existe el lote indicado."
    Response.End
End If

Set rsDet = conn.Execute("EXEC rrhh.usp_RecibosLoteDetalle_Sel " & loteid)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Detalle lote</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

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
        <h2 class="page-title">Detalle lote #<%=rsCab("LoteID")%></h2>
        <div class="acciones-top">
            <a href="rrhh_recibos_carga.asp" class="btn">Nuevo lote</a>
            <a href="rrhh_recibos_lotes.asp" class="btn-sec">Volver</a>
        </div>
    </div>

    <div class="card">
        <div class="grid">
            <div class="kpi">
                <div class="lab">Período</div>
                <div class="val"><%=PeriodoFormato(Nz(rsCab("Periodo"), ""))%></div>
            </div>
            <div class="kpi">
                <div class="lab">Archivo</div>
                <div class="val" style="font-size:15px;"><%=Server.HTMLEncode(Nz(rsCab("NombreArchivoOriginal"), ""))%></div>
            </div>
            <div class="kpi">
                <div class="lab">Páginas</div>
                <div class="val"><%=Nz(rsCab("CantidadPaginas"), 0)%></div>
            </div>
            <div class="kpi">
                <div class="lab">OK</div>
                <div class="val"><%=Nz(rsCab("CantidadOK"), 0)%></div>
            </div>
            <div class="kpi">
                <div class="lab">Error</div>
                <div class="val"><%=Nz(rsCab("CantidadError"), 0)%></div>
            </div>
            <div class="kpi">
                <div class="lab">Estado</div>
                <div class="val" style="font-size:15px;">
                    <span class="badge <%=EstadoClase(rsCab("Estado"))%>"><%=Server.HTMLEncode(Nz(rsCab("Estado"), ""))%></span>
                </div>
            </div>
            <div class="kpi">
                <div class="lab">Fecha carga</div>
                <div class="val" style="font-size:15px;"><%=Nz(rsCab("FechaCarga"), "")%></div>
            </div>
            <div class="kpi">
                <div class="lab">Usuario</div>
                <div class="val" style="font-size:15px;"><%=Server.HTMLEncode(Nz(rsCab("Usuario"), ""))%></div>
            </div>
        </div>
        <div class="small ruta-box">
            Ruta archivo original: <%=Server.HTMLEncode(Nz(rsCab("RutaArchivoOriginal"), ""))%>
        </div>
    </div>

    <div class="card">
        <h3 style="margin-top:0;">Detalle por página</h3>

        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>Página</th>
                        <th>Legajo</th>
                        <th>CUIL</th>
                        <th>Nombre detectado</th>
                        <!-- <th>Empleado BD</th> -->
                        <th>Período</th>
                        <!-- <th>Fecha pago</th> -->
                        <!-- <th>Neto</th> -->
                        <th>Estado</th>
                        <th>Error</th>
                        <th>ReciboID</th>
                        <!-- <th>PDF</th> -->
                    </tr>
                </thead>
                <tbody>
                <%
                If rsDet.EOF Then
                %>
                    <tr>
                        <td colspan="12" class="mobile-full" data-label="Detalle">No hay detalle para este lote.</td>
                    </tr>
                <%
                Else
                    Do Until rsDet.EOF
                %>
                    <tr>
                        <td data-label="Página"><%=Nz(rsDet("PaginaNumero"), "")%></td>
                        <td data-label="Legajo"><%=Server.HTMLEncode(Nz(rsDet("LegajoDetectado"), ""))%></td>
                        <td data-label="CUIL"><%=Server.HTMLEncode(Nz(rsDet("CUILDetectado"), ""))%></td>
                        <td data-label="Nombre detectado"><%=Server.HTMLEncode(Nz(rsDet("ApellidoNombreDetectado"), ""))%></td>
                        <!-- <td data-label="Empleado BD"><%=Server.HTMLEncode(Nz(rsDet("EmpleadoBD"), ""))%></td> -->
                        <td data-label="Período"><%=PeriodoFormato(Nz(rsDet("PeriodoDetectado"), ""))%></td>
                        <!-- <td data-label="Fecha pago"><%=Nz(rsDet("FechaPagoDetectada"), "")%></td> -->
                        <!-- <td data-label="Neto"><%=Nz(rsDet("NetoDetectado"), "")%></td> -->
                        <td data-label="Estado">
                            <span class="badge <%=EstadoClase(rsDet("EstadoProceso"))%>"><%=Server.HTMLEncode(Nz(rsDet("EstadoProceso"), ""))%></span>
                        </td>
                        <td data-label="Error"><%=Server.HTMLEncode(Nz(rsDet("MensajeError"), ""))%></td>
                        <td data-label="ReciboID"><%=Nz(rsDet("ReciboID"), "")%></td>
                        <!-- <td data-label="PDF">
                            <%
                            If Trim(Nz(rsDet("RutaPDFIndividual"), "")) <> "" Then
                            %>
                                <span class="small"><%=Server.HTMLEncode(rsDet("RutaPDFIndividual"))%></span>
                            <%
                            End If
                            %>
                        </td> -->
                    </tr>
                <%
                        rsDet.MoveNext
                    Loop
                End If
                %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<%
rsDet.Close
Set rsDet = Nothing
rsCab.Close
Set rsCab = Nothing
conn.Close
Set conn = Nothing
%>
</body>
</html>