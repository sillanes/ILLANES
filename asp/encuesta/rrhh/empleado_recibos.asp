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

Function SqlSafe(v)
    SqlSafe = Replace(Trim(CStr(v & "")), "'", "''")
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

Function PeriodoParaSQL(p)
    Dim s, a
    s = Trim(CStr(p & ""))

    If s = "" Then
        PeriodoParaSQL = ""
        Exit Function
    End If

    a = Split(s, "-")

    If UBound(a) = 1 Then
        If Len(a(0)) = 2 And Len(a(1)) = 4 Then
            PeriodoParaSQL = a(1) & "-" & Right("0" & a(0), 2)
            Exit Function
        End If

        If Len(a(0)) = 4 And Len(a(1)) >= 1 Then
            PeriodoParaSQL = a(0) & "-" & Right("0" & a(1), 2)
            Exit Function
        End If
    End If

    PeriodoParaSQL = s
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

Dim empleadoID, periodo, periodoSP
empleadoID = CLng("0" & Session("Empleado_EmpleadoID"))
periodo    = Trim(Request("periodo"))
periodoSP  = PeriodoParaSQL(periodo)

Dim periodoSQL
If periodoSP = "" Then
    periodoSQL = "NULL"
Else
    periodoSQL = "'" & SqlSafe(periodoSP) & "'"
End If

Dim rs, sqlRecibos
sqlRecibos = "EXEC empleado.usp_Empleado_Recibos_Sel " & _
             "@EmpleadoID=" & empleadoID & ", " & _
             "@Periodo=" & periodoSQL & ", " & _
             "@EstadoID=NULL"

Set rs = conn.Execute(sqlRecibos)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Portal del Empleado - Mis Recibos</title>
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
.container{max-width:1200px;margin:20px auto;padding:0 12px;box-sizing:border-box}
.card{border:1px solid #ddd;border-radius:10px;padding:18px;background:#fff;margin-bottom:20px}
.table-wrap{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}
.table{width:100%;border-collapse:collapse;table-layout:auto}
.table th,.table td{border:1px solid #ddd;padding:9px;text-align:left;vertical-align:top;word-break:break-word}
.table th{background:#000000;color:#fff}
.badge{display:inline-block;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:bold}
.badge.ok{background:#d4edda;color:#155724}
.badge.err{background:#f8d7da;color:#721c24}
.badge.pen{background:#fff3cd;color:#856404}
.badge.dup{background:#d1ecf1;color:#0c5460}
.badge.neu{background:#e2e3e5;color:#383d41}
.btn{background:#2b7cff;color:#fff;text-decoration:none;padding:8px 12px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.btn-sec{background:#6c757d;color:#fff;text-decoration:none;padding:8px 12px;border-radius:8px;display:inline-block;border:none;cursor:pointer}
.filtros-grid{display:grid;grid-template-columns:minmax(220px,320px);gap:12px}
.campo label{display:block;font-size:12px;color:#555;margin-bottom:6px}
.campo input{width:100%;padding:10px;border:1px solid #ccc;border-radius:8px;box-sizing:border-box;background:#fff}
.filtros-acciones{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px}
.helper{font-size:12px;color:#666;margin-top:8px}

@media (max-width: 768px){
    .container{margin:12px auto;padding:0 10px}
    .card{padding:14px}
    .filtros-grid{grid-template-columns:1fr}
    .filtros-acciones .btn,
    .filtros-acciones .btn-sec{width:100%;text-align:center;box-sizing:border-box}
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
        <form method="get" action="empleado_recibos.asp">
            <div class="filtros-grid">
                <div class="campo">
                    <label for="periodo">Período</label>
                    <input type="text" name="periodo" id="periodo" value="<%=Server.HTMLEncode(periodo)%>" placeholder="04-2025">
                    <div class="helper">Ingresá el período en formato MM-YYYY.</div>
                </div>
            </div>

            <div class="filtros-acciones">
                <button type="submit" class="btn">Filtrar</button>
                <a href="empleado_recibos.asp" class="btn-sec">Limpiar</a>
            </div>
        </form>
    </div>

    <div class="card">
        <div class="table-wrap">
            <table class="table">
                <thead>
                    <tr>
                        <!--<th>ReciboID</th>  -->
                        <th>Período</th>
                        <!-- <th>Fecha pago</th> -->
                        <th>Estado</th>
                        <th>Leído</th>
                        <th>Firmado</th>
                        <th>Acción</th>
                    </tr>
                </thead>
                <tbody>
                <%
                If rs.EOF Then
                %>
                    <tr>
                        <td colspan="7" class="mobile-full" data-label="Resultado">No tenés recibos para los filtros seleccionados.</td>
                    </tr>
                <%
                Else
                    Do Until rs.EOF
                %>
                    <tr>
                        <!-- <td data-label="ReciboID"><%=Nz(rs("ReciboID"), "")%></td>  -->
                        <td data-label="Período"><%=PeriodoFormato(rs("Periodo"))%></td>
                        <!-- <td data-label="Fecha pago"><%=FechaFormato(rs("FechaPago"))%></td> -->
                        <td data-label="Estado">
                            <span class="badge <%=EstadoClase(rs("Estado"))%>"><%=Server.HTMLEncode(Nz(rs("Estado"), ""))%></span>
                        </td>
                        <td data-label="Leído"><%=FechaHoraFormato(rs("FechaLectura"))%></td>
                        <td data-label="Firmado"><%=FechaHoraFormato(rs("FechaFirma"))%></td>
                        <td data-label="Acción">
                            <a class="btn" href="empleado_recibo_detalle.asp?reciboid=<%=rs("ReciboID")%>">Ver</a>
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