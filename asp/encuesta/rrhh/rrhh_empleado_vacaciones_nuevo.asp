<%@LANGUAGE="VBScript" CODEPAGE="65001"%>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Session.LCID = 11274

If Trim(Session("currentUser") & "") = "" Then
    Response.Redirect "../login.asp"
End If

Function Nz(v, d)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = d
    Else
        Nz = v
    End If
End Function

Function Html(v)
    Html = Server.HTMLEncode(CStr(Nz(v, "")))
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

Dim EmpleadoID, accion, FechaDesde, FechaHasta, Observaciones
Dim msgError
Dim rsEmp, sqlEmp

EmpleadoID    = CLng("0" & Request("EmpleadoID"))
accion        = LCase(Trim(Request("accion")))
FechaDesde    = Trim(Request("FechaDesde"))
FechaHasta    = Trim(Request("FechaHasta"))
Observaciones = Trim(Request("Observaciones"))
msgError      = ""

If EmpleadoID <= 0 Then
    Response.Redirect "rrhh_empleados.asp?ok=0&msg=" & Server.URLEncode("Empleado inválido.")
End If

sqlEmp = "EXEC rrhh.usp_Empleado_sel @EmpleadoID=" & EmpleadoID
Set rsEmp = Conn.Execute(sqlEmp)

If rsEmp.EOF Then
    rsEmp.Close
    Set rsEmp = Nothing
    Response.Redirect "rrhh_empleados.asp?ok=0&msg=" & Server.URLEncode("No se encontró el empleado.")
End If

If accion = "guardar" Then
    If FechaDesde = "" Or FechaHasta = "" Then
        msgError = "Debes ingresar fecha desde y fecha hasta."
    Else
        On Error Resume Next

        Dim sqlIns, usuarioSQL, observacionesSQL

        If Trim(Session("UsuarioID") & "") = "" Then
            usuarioSQL = "NULL"
        Else
            usuarioSQL = CLng("0" & Session("UsuarioID"))
        End If

        If Observaciones = "" Then
            observacionesSQL = "NULL"
        Else
            observacionesSQL = "'" & SqlSafe(Observaciones) & "'"
        End If

        sqlIns = "EXEC rrhh.usp_Vacaciones_RRHH_Ins " & _
                 "@EmpleadoID=" & EmpleadoID & ", " & _
                 "@FechaDesde='" & SqlSafe(FechaDesde) & "', " & _
                 "@FechaHasta='" & SqlSafe(FechaHasta) & "', " & _
                 "@UsuarioResolucionID=" & usuarioSQL & ", " & _
                 "@Observaciones=" & observacionesSQL

        Conn.Execute sqlIns

        If Err.Number <> 0 Then
            msgError = "No se pudieron registrar las vacaciones: " & Err.Description
            Err.Clear
        Else
            Response.Redirect "rrhh_empleados.asp?ok=1&msg=" & Server.URLEncode("Vacaciones cargadas correctamente para el empleado.")
        End If

        On Error GoTo 0
    End If
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <title>Agregar Vacaciones</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="rrhh.css" />
    <link rel="stylesheet" href="ventas.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

    <style>
        body { background:#f4f6f9; }
        .main-content { margin-left:260px; padding:30px 20px; }
        .page-header {
            display:flex;
            justify-content:space-between;
            align-items:center;
            gap:12px;
            flex-wrap:wrap;
            margin-bottom:20px;
        }
        .page-title {
            margin:0;
            font-size:1.8rem;
            font-weight:700;
            color:#1f2937;
        }
        .page-subtitle {
            color:#6b7280;
            margin-top:6px;
        }
        .card-custom {
            background:#fff;
            border:0;
            border-radius:16px;
            box-shadow:0 8px 24px rgba(0,0,0,0.08);
            overflow:hidden;
        }
        .card-custom .card-header {
            background:linear-gradient(135deg, #6f42c1, #7c4dff);
            color:#fff;
            border:0;
            padding:16px 20px;
        }
        .form-grid {
            display:grid;
            grid-template-columns:repeat(2, minmax(240px,1fr));
            gap:16px;
        }
        .campo label {
            display:block;
            font-weight:600;
            margin-bottom:6px;
            color:#374151;
        }
        .campo input, .campo textarea {
            width:100%;
            padding:10px 12px;
            border:1px solid #d1d5db;
            border-radius:10px;
            box-sizing:border-box;
            background:#fff;
        }
        .campo textarea {
            min-height:100px;
            resize:vertical;
        }
        .info-empleado {
            display:grid;
            grid-template-columns:repeat(2, minmax(240px,1fr));
            gap:16px;
            margin-bottom:20px;
        }
        .info-box {
            background:#f8fafc;
            border:1px solid #e5e7eb;
            border-radius:12px;
            padding:12px 14px;
        }
        .info-label {
            font-size:12px;
            color:#6b7280;
            margin-bottom:4px;
        }
        .info-value {
            font-size:15px;
            font-weight:700;
            color:#1f2937;
        }
        .acciones {
            display:flex;
            gap:10px;
            flex-wrap:wrap;
            margin-top:20px;
        }
        .alert-custom {
            border-radius:12px;
            padding:14px 16px;
            margin-bottom:18px;
            border:1px solid transparent;
            font-weight:600;
            box-shadow:0 4px 14px rgba(0,0,0,.04);
        }
        .alert-danger-custom {
            background:#f8d7da;
            color:#842029;
            border-color:#f5c2c7;
        }
        @media (max-width: 991px) {
            .main-content { margin-left:0; padding:20px 12px; }
            .form-grid, .info-empleado { grid-template-columns:1fr; }
        }
    </style>
</head>
<body>
<!--#include file="header.asp" -->

<div class="main-content">
    <div class="page-header">
        <div>
            <h1 class="page-title"><i class="fa-solid fa-calendar-plus"></i> Agregar Vacaciones</h1>
            <div class="page-subtitle">Carga manual desde RRHH</div>
        </div>
    </div>

    <% If msgError <> "" Then %>
        <div class="alert-custom alert-danger-custom"><%=Html(msgError)%></div>
    <% End If %>

    <div class="card card-custom">
        <div class="card-header">
            <strong><i class="fa-solid fa-user"></i> Empleado</strong>
        </div>
        <div class="card-body">

            <div class="info-empleado">
                <div class="info-box">
                    <div class="info-label">Empleado</div>
                    <div class="info-value"><%=Html(rsEmp("Apellido"))%>, <%=Html(rsEmp("Nombre"))%></div>
                </div>

                <div class="info-box">
                    <div class="info-label">Legajo</div>
                    <div class="info-value"><%=Html(rsEmp("Legajo"))%></div>
                </div>

                <div class="info-box">
                    <div class="info-label">CUIL</div>
                    <div class="info-value"><%=Html(rsEmp("CUIL"))%></div>
                </div>

                <div class="info-box">
                    <div class="info-label">Fecha Ingreso</div>
                    <div class="info-value"><%=FechaFormato(rsEmp("FechaIngreso"))%></div>
                </div>
            </div>

            <form method="post" action="rrhh_empleado_vacaciones_nuevo.asp">
                <input type="hidden" name="accion" value="guardar">
                <input type="hidden" name="EmpleadoID" value="<%=EmpleadoID%>">

                <div class="form-grid">
                    <div class="campo">
                        <label for="FechaDesde">Fecha Desde</label>
                        <input type="date" name="FechaDesde" id="FechaDesde" value="<%=Html(FechaDesde)%>">
                    </div>

                    <div class="campo">
                        <label for="FechaHasta">Fecha Hasta</label>
                        <input type="date" name="FechaHasta" id="FechaHasta" value="<%=Html(FechaHasta)%>">
                    </div>
                </div>

                <div class="campo" style="margin-top:16px;">
                    <label for="Observaciones">Observaciones</label>
                    <textarea name="Observaciones" id="Observaciones"><%=Html(Observaciones)%></textarea>
                </div>

                <div class="acciones">
                    <button type="submit" class="btn btn-success">
                        <i class="fa-solid fa-floppy-disk"></i> Guardar Vacaciones
                    </button>

                    <a href="rrhh_empleados.asp" class="btn btn-secondary">
                        <i class="fa-solid fa-arrow-left"></i> Volver
                    </a>
                </div>
            </form>

        </div>
    </div>
</div>

</body>
</html>

<%
rsEmp.Close
Set rsEmp = Nothing

Conn.Close
Set Conn = Nothing
%>