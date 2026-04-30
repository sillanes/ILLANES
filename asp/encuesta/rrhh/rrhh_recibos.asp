<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Trim(Session("currentUser") & "") <> "admin" AND Trim(Session("currentUser") & "") <> "rrhh" Then
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
        Case  "OK", "PROCESADO"
            EstadoClase = "ok"
        Case "ERROR"
            EstadoClase = "err"
        Case "DISPONIBLE","PENDIENTE", "PROCESANDO"
            EstadoClase = "pen"
        Case "LEIDO"
            EstadoClase = "dup"
        Case "FIRMADO"
            EstadoClase = "ok"
        Case Else
            EstadoClase = "neu"
    End Select
End Function

Function SqlSafe(v)
    SqlSafe = Replace(Trim(CStr(v & "")), "'", "''")
End Function

Function FormatoPeriodo(p)
    Dim s, partes
    s = Trim(CStr(p & ""))

    If s = "" Then
        FormatoPeriodo = ""
        Exit Function
    End If

    partes = Split(s, "-")

    If UBound(partes) = 1 Then
        If Len(partes(0)) = 4 And Len(partes(1)) = 2 Then
            ' YYYY-MM -> MM-YYYY
            FormatoPeriodo = partes(1) & "-" & partes(0)
        Else
            FormatoPeriodo = s
        End If
    Else
        FormatoPeriodo = s
    End If
End Function

Function PeriodoToSQL(p)
    Dim s, partes
    s = Trim(CStr(p & ""))

    If s = "" Then
        PeriodoToSQL = ""
        Exit Function
    End If

    partes = Split(s, "-")

    If UBound(partes) = 1 Then
        If Len(partes(0)) = 2 And Len(partes(1)) = 4 Then
            ' MM-YYYY -> YYYY-MM
            PeriodoToSQL = partes(1) & "-" & partes(0)
        Else
            PeriodoToSQL = s
        End If
    Else
        PeriodoToSQL = s
    End If
End Function

Dim periodoInput, periodo, legajo, apellidonombre, estadoid
periodoInput    = Trim(Request("periodo"))
periodo         = PeriodoToSQL(periodoInput)
legajo          = Trim(Request("legajo"))
apellidonombre  = Trim(Request("apellidonombre"))
estadoid        = CLng("0" & Request("estadoid"))

Dim periodoSQL, legajoSQL, apellidoSQL, estadoSQL
If periodo = "" Then
    periodoSQL = "NULL"
Else
    periodoSQL = "'" & SqlSafe(periodo) & "'"
End If

If legajo = "" Then
    legajoSQL = "NULL"
Else
    legajoSQL = "'" & SqlSafe(legajo) & "'"
End If

If apellidonombre = "" Then
    apellidoSQL = "NULL"
Else
    apellidoSQL = "'" & SqlSafe(apellidonombre) & "'"
End If

If estadoid <= 0 Then
    estadoSQL = "NULL"
Else
    estadoSQL = CStr(estadoid)
End If

Dim rs, rsEstados, rsDashboard
Dim sqlRecibos, sqlEstados, sqlDashboard

sqlRecibos = "EXEC rrhh.usp_Recibos_Sel " & _
             "@Periodo=" & periodoSQL & ", " & _
             "@Legajo=" & legajoSQL & ", " & _
             "@ApellidoNombre=" & apellidoSQL & ", " & _
             "@EstadoID=" & estadoSQL

sqlDashboard = "EXEC rrhh.usp_Recibos_Dashboard " & _
               "@Periodo=" & periodoSQL & ", " & _
               "@Legajo=" & legajoSQL & ", " & _
               "@ApellidoNombre=" & apellidoSQL & ", " & _
               "@EstadoID=" & estadoSQL

sqlEstados = "EXEC rrhh.usp_ReciboEstado_Sel"
 
Dim t0, t1, t2, t3
t0 = Timer()
Set rs = conn.Execute(sqlRecibos)
t1 = Timer()

Set rsDashboard = conn.Execute(sqlDashboard) 
t2 = Timer()

Set rsEstados = conn.Execute(sqlEstados)
t3 = Timer()
 

Response.Write "<!-- sqlRecibos: " & FormatNumber(t1-t0, 3) & " s -->"
Response.Write "<!-- sqlDashboard: " & FormatNumber(t2-t1, 3) & " s -->"
Response.Write "<!-- sqlEstados: " & FormatNumber(t3-t2, 3) & " s -->"
 

%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Recibos</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<script>
function toggleSidebar(){
    document.querySelector('.sidebar').classList.toggle('open');
}
</script>

</head>
<body>

<!--#include file="header.asp" -->

<div class="main-content">
    <div class="header-row">
        <h2 class="page-title">Recibos</h2>
        <div class="acciones-top">
            <a href="rrhh_recibos_carga.asp" class="btn">Nuevo lote</a>
            <a href="rrhh_recibos_lotes.asp" class="btn-sec">Ver lotes</a>
        </div>
    </div>

    <div class="card">
        <form method="get" action="rrhh_recibos.asp">
            <div class="filtros-grid">
                <div class="campo">
                    <label for="periodo">Período</label>
                    <input type="text" name="periodo" id="periodo" value="<%=Server.HTMLEncode(periodoInput)%>" placeholder="MM-YYYY (ej: 01-2026)">
                </div>
                <div class="campo">
                    <label for="legajo">Legajo</label>
                    <input type="text" name="legajo" id="legajo" value="<%=Server.HTMLEncode(legajo)%>" placeholder="00034">
                </div>
                <div class="campo">
                    <label for="apellidonombre">Apellido y nombre</label>
                    <input type="text" name="apellidonombre" id="apellidonombre" value="<%=Server.HTMLEncode(apellidonombre)%>" placeholder="Cruciani">
                </div>
                <div class="campo">
                    <label for="estadoid">Estado</label>
                    <select name="estadoid" id="estadoid">
                        <option value="0">Todos</option>
                        <%
                        If Not rsEstados.EOF Then
                            Do Until rsEstados.EOF
                        %>
                            <option value="<%=rsEstados("EstadoID")%>" <% If CLng("0" & estadoid) = CLng("0" & rsEstados("EstadoID")) Then Response.Write "selected" End If %>>
                                <%=Server.HTMLEncode(Nz(rsEstados("Descripcion"), ""))%>
                            </option>
                        <%
                                rsEstados.MoveNext
                            Loop
                        End If
                        %>
                    </select>
                </div>
            </div>

            <div class="filtros-acciones">
                <button type="submit" class="btn">Filtrar</button>
                <a href="rrhh_recibos.asp" class="btn-sec">Limpiar</a>
            </div>
        </form>
    </div>

    <div class="card">
        <div class="header-row" style="margin-bottom:12px;">
            <h3 class="page-title" style="font-size:20px; margin:0;">Resumen por período</h3>
        </div>

        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>Período</th>
                        <th>Cantidad</th>
                        <th>Leídos</th>
                        <th>Firmados</th>
                        <th>Pendientes Firma</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rsDashboard.EOF Then
                %>
                    <tr>
                        <td colspan="5" class="mobile-full" data-label="Resultado">No hay datos de resumen para los filtros seleccionados.</td>
                    </tr>
                <%
                Else
                    Do Until rsDashboard.EOF
                %>
                    <tr>
                        <td data-label="Período"><%=FormatoPeriodo(Nz(rsDashboard("Periodo"), ""))%></td>
                        <td data-label="Cantidad"><%=Nz(rsDashboard("Cantidad"), 0)%></td>
                        <td data-label="Leídos"><%=Nz(rsDashboard("Leidos"), 0)%></td>
                        <td data-label="Firmados"><%=Nz(rsDashboard("Firmados"), 0)%></td>
                        <td data-label="Pendientes"><%=Nz(rsDashboard("Pendientes"), 0)%></td>
                    </tr>
                <%
                        rsDashboard.MoveNext
                    Loop
                End If
                %>
                </tbody>
            </table>
        </div>
    </div>

    <div class="card">
        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <th>ReciboID</th>
                        <th>Período</th>
                        <th>Legajo</th>
                        <th>Empleado</th>
                        <th>Fecha pago</th>
                        <th>Neto</th>
                        <th>Estado</th>
                        <th>Leído</th>
                        <th>Firmado</th>
                        <th>Lote</th>
                        <th>Acción</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="11" class="mobile-full" data-label="Resultado">No hay recibos para los filtros seleccionados.</td>
                    </tr>
                <%
                Else
                    Do Until rs.EOF
                %>
                    <tr>
                        <td data-label="ReciboID"><%=Nz(rs("ReciboID"), "")%></td>
                        <td data-label="Período"><%=FormatoPeriodo(Nz(rs("Periodo"), ""))%></td>
                        <td data-label="Legajo"><%=Server.HTMLEncode(Nz(rs("Legajo"), ""))%></td>
                        <td data-label="Empleado"><%=Server.HTMLEncode(Nz(rs("ApellidoNombre"), ""))%></td>
                        <td data-label="Fecha pago"><%=Nz(rs("FechaPago"), "")%></td>
                        <td data-label="Neto"><%=Nz(rs("Neto"), "")%></td>
                        <td data-label="Estado">
                            <span class="badge <%=EstadoClase(rs("Estado"))%>"><%=Server.HTMLEncode(Nz(rs("Estado"), ""))%></span>
                        </td>
                        <td data-label="Leído"><%=Nz(rs("FechaLectura"), "")%></td>
                        <td data-label="Firmado"><%=Nz(rs("FechaFirma"), "")%></td>
                        <td data-label="Lote"><%=Nz(rs("LoteID"), "")%></td>
                        <td data-label="Acción">
                            <a class="btn-sec" href="rrhh_recibo_detalle.asp?reciboid=<%=rs("ReciboID")%>">Ver</a>
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
    Set rs = Nothing
End If

If Not rsDashboard Is Nothing Then
    If rsDashboard.State = 1 Then rsDashboard.Close
    Set rsDashboard = Nothing
End If

If Not rsEstados Is Nothing Then
    If rsEstados.State = 1 Then rsEstados.Close
    Set rsEstados = Nothing
End If

If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
End If

On Error GoTo 0
%>

</body>
</html>