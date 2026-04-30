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
    Response.Redirect "../login.asp"
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

Function HtmlSafe(v)
    HtmlSafe = Server.HTMLEncode(CStr(v & ""))
End Function

Function FechaISO(d)
    FechaISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
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

Function EstadoClase(estado)
    Dim s
    s = UCase(Trim(CStr(estado & "")))

    Select Case s
        Case "APROBADO"
            EstadoClase = "evento-ok"
        Case "RECHAZADO"
            EstadoClase = "evento-err"
        Case "PENDIENTE"
            EstadoClase = "evento-pen"
        Case Else
            EstadoClase = "evento-neu"
    End Select
End Function

Function PrimerDiaMes(d)
    PrimerDiaMes = DateSerial(Year(d), Month(d), 1)
End Function

Function UltimoDiaMes(d)
    UltimoDiaMes = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Function QuitarEspaciosDobles(txt)
    Dim s
    s = Trim(CStr(txt & ""))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    QuitarEspaciosDobles = s
End Function

Function TresLetras(txt)
    txt = UCase(Trim(CStr(txt & "")))
    If Len(txt) >= 3 Then
        TresLetras = Left(txt, 3)
    ElseIf Len(txt) > 0 Then
        TresLetras = txt
    Else
        TresLetras = "---"
    End If
End Function

Function NombreEmpleadoCorto(nombreCompleto)
    Dim s, partes, apellido, nombre, bloques

    s = QuitarEspaciosDobles(nombreCompleto)

    If s = "" Then
        NombreEmpleadoCorto = "---"
        Exit Function
    End If

    apellido = ""
    nombre = ""

    If InStr(s, ",") > 0 Then
        bloques = Split(s, ",")
        apellido = Trim(bloques(0))

        If UBound(bloques) >= 1 Then
            partes = Split(QuitarEspaciosDobles(Trim(bloques(1))), " ")
            If UBound(partes) >= 0 Then
                nombre = Trim(partes(0))
            End If
        End If
    Else
        partes = Split(s, " ")
        If UBound(partes) >= 0 Then apellido = Trim(partes(0))
        If UBound(partes) >= 1 Then
            nombre = Trim(partes(1))
        Else
            nombre = Trim(partes(0))
        End If
    End If

    NombreEmpleadoCorto = TresLetras(apellido) & "-" & TresLetras(nombre)
End Function

Dim fechaDesde, fechaHasta, estadoFiltro, empleadoID
Dim fechaDesdeSQL, fechaHastaSQL, estadoFiltroSQL, empleadoIDSQL
Dim rs, rsEmp, sql, sqlEmp
Dim primerDiaVista, ultimoDiaVista, primerCalendario, ultimoCalendario
Dim wd, d

fechaDesde   = Trim(Request("fechadesde"))
fechaHasta   = Trim(Request("fechahasta"))
estadoFiltro = Trim(Request("estadofiltro"))
empleadoID   = CLng("0" & Request("empleadoid"))

If fechaDesde = "" Then
    primerDiaVista = PrimerDiaMes(Date())
    fechaDesde = FechaISO(primerDiaVista)
Else
    primerDiaVista = CDate(fechaDesde)
End If

If fechaHasta = "" Then
    ultimoDiaVista = UltimoDiaMes(primerDiaVista)
    fechaHasta = FechaISO(ultimoDiaVista)
Else
    ultimoDiaVista = CDate(fechaHasta)
End If

If ultimoDiaVista < primerDiaVista Then
    ultimoDiaVista = primerDiaVista
    fechaHasta = FechaISO(ultimoDiaVista)
End If

wd = Weekday(primerDiaVista, vbMonday)
primerCalendario = DateAdd("d", -(wd - 1), primerDiaVista)

wd = Weekday(ultimoDiaVista, vbMonday)
ultimoCalendario = DateAdd("d", 7 - wd, ultimoDiaVista)

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

If empleadoID <= 0 Then
    empleadoIDSQL = "NULL"
Else
    empleadoIDSQL = CStr(empleadoID)
End If

sqlEmp = "EXEC rrhh.usp_Empleados_Combo_Sel"
Set rsEmp = conn.Execute(sqlEmp)

sql = "EXEC rrhh.usp_Vacaciones_Resumen_Sel " & _
      "@FechaDesde=" & fechaDesdeSQL & ", " & _
      "@FechaHasta=" & fechaHastaSQL & ", " & _
      "@Estado=" & estadoFiltroSQL & ", " & _
      "@EmpleadoID=" & empleadoIDSQL

Set rs = conn.Execute(sql)

Dim eventos
Set eventos = Server.CreateObject("Scripting.Dictionary")

Do Until rs.EOF
    Dim fDesde, fHasta, actual, clave, htmlEvento, nombreMostrar

    fDesde = CDate(rs("FechaDesde"))
    fHasta = CDate(rs("FechaHasta"))

    If fDesde < primerDiaVista Then fDesde = primerDiaVista
    If fHasta > ultimoDiaVista Then fHasta = ultimoDiaVista

    nombreMostrar = NombreEmpleadoCorto(Nz(rs("ApellidoNombre"), ""))

    htmlEvento = "<div class=""evento " & EstadoClase(rs("Estado")) & """>" & _
                 "<span class=""evento-nombre"">" & HtmlSafe(nombreMostrar) & "</span>" & _
                 "</div>"

    actual = fDesde
    Do While actual <= fHasta
        clave = FechaISO(actual)

        If Not eventos.Exists(clave) Then
            eventos.Add clave, htmlEvento
        Else
            eventos(clave) = eventos(clave) & htmlEvento
        End If

        actual = DateAdd("d", 1, actual)
    Loop

    rs.MoveNext
Loop
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>RRHH - Dashboard Vacaciones</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="rrhh.css" />
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

<style>
.cal-filtros{
    display:grid;
    grid-template-columns:repeat(4,minmax(180px,1fr));
    gap:12px;
}
.cal-grid{
    display:grid;
    grid-template-columns:repeat(7,1fr);
    gap:8px;
}
.cal-head{
    background:#111;
    color:#fff;
    padding:10px;
    border-radius:8px;
    text-align:center;
    font-weight:bold;
}
.cal-day{
    background:#fff;
    border:1px solid #ddd;
    border-radius:10px;
    min-height:150px;
    padding:8px;
    box-sizing:border-box;
}
.cal-day.otro-mes{
    background:#f3f3f3;
}
.cal-day.fuera-rango{
    opacity:.55;
}
.cal-num{
    font-size:13px;
    font-weight:bold;
    margin-bottom:6px;
    color:#333;
}
.evento{
    padding:4px 6px;
    border-radius:6px;
    font-size:11px;
    margin-bottom:4px;
    word-break:break-word;
    line-height:1.25;
}
.evento-nombre{
    display:block;
    font-weight:bold;
    letter-spacing:.3px;
}
.evento-ok{
    background:#d4edda;
    color:#155724;
}
.evento-err{
    background:#f8d7da;
    color:#721c24;
}
.evento-pen{
    background:#fff3cd;
    color:#856404;
}
.evento-neu{
    background:#e2e3e5;
    color:#383d41;
}
.legend{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
    margin-top:12px;
}
.legend .evento{
    margin-bottom:0;
}
.resumen-rango{
    margin-top:10px;
    font-size:13px;
    color:#555;
}
@media (max-width: 1100px){
    .cal-filtros{
        grid-template-columns:repeat(2,minmax(180px,1fr));
    }
}
@media (max-width: 900px){
    .cal-grid{
        grid-template-columns:1fr;
    }
    .cal-head{
        display:none;
    }
    .cal-day{
        min-height:unset;
    }
}
@media (max-width: 768px){
    .cal-filtros{
        grid-template-columns:1fr;
    }
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
        <h2 class="page-title">Dashboard vacaciones</h2>
    </div>

    <div class="card">
        <form method="get" action="rrhh_vacaciones_dashboard.asp">
            <div class="cal-filtros">
                <div class="campo">
                    <label for="fechadesde">Fecha inicio</label>
                    <input type="date" name="fechadesde" id="fechadesde" value="<%=Server.HTMLEncode(fechaDesde)%>">
                </div>

                <div class="campo">
                    <label for="fechahasta">Fecha fin</label>
                    <input type="date" name="fechahasta" id="fechahasta" value="<%=Server.HTMLEncode(fechaHasta)%>">
                </div>

                <div class="campo">
                    <label for="empleadoid">Empleado</label>
                    <select name="empleadoid" id="empleadoid">
                        <option value="0">Todos</option>
                        <%
                        If Not rsEmp.EOF Then
                            Do Until rsEmp.EOF
                        %>
                            <option value="<%=rsEmp("EmpleadoID")%>" <% If CLng("0" & empleadoID)=CLng("0" & rsEmp("EmpleadoID")) Then Response.Write "selected" End If %>>
                                <%=Server.HTMLEncode(Nz(rsEmp("ApellidoNombre"), ""))%>
                            </option>
                        <%
                                rsEmp.MoveNext
                            Loop
                        End If
                        %>
                    </select>
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
                <button type="submit" class="btn">Ver calendario</button>
                <a href="rrhh_vacaciones_dashboard.asp" class="btn-sec">Limpiar</a>
                <a href="rrhh_vacaciones.asp" class="btn-sec">Volver al listado</a>
            </div>
        </form>

        <div class="legend">
            <div class="evento evento-pen">Pendiente</div>
            <div class="evento evento-ok">Aprobado</div>
            <div class="evento evento-err">Rechazado</div>
        </div>

        <div class="resumen-rango">
            Mostrando calendario desde <strong><%=FechaFormato(primerDiaVista)%></strong> hasta <strong><%=FechaFormato(ultimoDiaVista)%></strong>.
            <br>
            Código de empleado: <strong>3 letras del apellido + 3 letras del nombre</strong>.
        </div>
    </div>

    <div class="card">
        <div class="cal-grid">
            <div class="cal-head">Lunes</div>
            <div class="cal-head">Martes</div>
            <div class="cal-head">Miércoles</div>
            <div class="cal-head">Jueves</div>
            <div class="cal-head">Viernes</div>
            <div class="cal-head">Sábado</div>
            <div class="cal-head">Domingo</div>

            <%
            d = primerCalendario
            Do While d <= ultimoCalendario
                Dim cssExtra, claveDia
                cssExtra = ""
                If d < primerDiaVista Or d > ultimoDiaVista Then cssExtra = cssExtra & " fuera-rango"
                If Month(d) <> Month(primerDiaVista) And Month(d) <> Month(ultimoDiaVista) Then cssExtra = cssExtra & " otro-mes"
                claveDia = FechaISO(d)
            %>
                <div class="cal-day<%=cssExtra%>">
                    <div class="cal-num"><%=Day(d)%></div>
                    <% If eventos.Exists(claveDia) Then %>
                        <%=eventos(claveDia)%>
                    <% End If %>
                </div>
            <%
                d = DateAdd("d", 1, d)
            Loop
            %>
        </div>
    </div>

</div>
</div>

<%
If IsObject(rs) Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

If IsObject(rsEmp) Then
    If rsEmp.State = 1 Then rsEmp.Close
    Set rsEmp = Nothing
End If

Set eventos = Nothing

If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>
</body>
</html>