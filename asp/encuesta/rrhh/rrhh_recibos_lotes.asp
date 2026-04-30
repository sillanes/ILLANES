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
        Case "PROCESADO", "OK"
            EstadoClase = "ok"
        Case "ERROR"
            EstadoClase = "err"
        Case "PROCESANDO", "PENDIENTE"
            EstadoClase = "pen"
        Case Else
            EstadoClase = "neu"
    End Select
End Function

Dim rs
Set rs = conn.Execute("EXEC rrhh.usp_RecibosLote_Sel")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Lotes de recibos</title>

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
        <h2 class="page-title">Lotes de recibos</h2>
        <div class="acciones-top">
            <a href="rrhh_recibos_carga.asp" class="btn">Nuevo lote</a>
        </div>
    </div>

    <div class="card">
        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>LoteID</th>
                        <th>Período</th>
                        <th>Archivo</th>
                        <th>Páginas</th>
                        <th>OK</th>
                        <th>Error</th>
                        <th>Estado</th>
                        <th>Fecha carga</th>
                        <th>Usuario</th>
                        <th>Acción</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="10" class="mobile-full" data-label="Resultado">No hay lotes cargados.</td>
                    </tr>
                <%
                Else
                    Do Until rs.EOF
                %>
                    <tr>
                        <td data-label="LoteID"><%=rs("LoteID")%></td>
                        <td data-label="Período"><%=PeriodoFormato(Nz(rs("Periodo"), ""))%></td>
                        <td data-label="Archivo"><%=Server.HTMLEncode(Nz(rs("NombreArchivoOriginal"), ""))%></td>
                        <td data-label="Páginas"><%=Nz(rs("CantidadPaginas"), 0)%></td>
                        <td data-label="OK"><%=Nz(rs("CantidadOK"), 0)%></td>
                        <td data-label="Error"><%=Nz(rs("CantidadError"), 0)%></td>
                        <td data-label="Estado">
                            <span class="badge <%=EstadoClase(rs("Estado"))%>"><%=Server.HTMLEncode(Nz(rs("Estado"), ""))%></span>
                        </td>
                        <td data-label="Fecha carga"><%=Nz(rs("FechaCarga"), "")%></td>
                        <td data-label="Usuario"><%=Server.HTMLEncode(Nz(rs("Usuario"), ""))%></td>
                        <td data-label="Acción">
                            <a class="btn-sec" href="rrhh_recibos_lote_detalle.asp?loteid=<%=rs("LoteID")%>">Ver detalle</a>
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
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
</body>
</html>