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

Function Nz(v, defaultValue)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = defaultValue
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

Dim accion
accion = LCase(Trim(Request("accion")))

Dim Legajo, CUIL, DNI, Apellido, Nombre, TelefonoWhatsApp, Email, Activo, FechaIngreso
Dim msgError

Legajo           = Trim(Request.Form("Legajo"))
CUIL             = Trim(Request.Form("CUIL"))
DNI              = Trim(Request.Form("DNI"))
Apellido         = Trim(Request.Form("Apellido"))
Nombre           = Trim(Request.Form("Nombre"))
TelefonoWhatsApp = Trim(Request.Form("TelefonoWhatsApp"))
Email            = Trim(Request.Form("Email"))
Activo           = Trim(Request.Form("Activo"))
FechaIngreso     = Trim(Request.Form("FechaIngreso"))
msgError         = ""

If Activo = "" Then Activo = "1"

If accion = "guardar" Then

    If Legajo = "" Then
        msgError = "Debes ingresar el legajo."
    ElseIf Apellido = "" Then
        msgError = "Debes ingresar el apellido."
    ElseIf Nombre = "" Then
        msgError = "Debes ingresar el nombre."
    Else
        On Error Resume Next

        Dim sqlIns, fechaSQL, activoSQL, rsIns, nuevoEmpleadoID

        If FechaIngreso = "" Then
            fechaSQL = "NULL"
        Else
            fechaSQL = "'" & SqlSafe(FechaIngreso) & "'"
        End If

        If Activo = "1" Then
            activoSQL = "1"
        Else
            activoSQL = "0"
        End If

        sqlIns = "EXEC rrhh.usp_Empleado_Ins " & _
                 "@Legajo='" & SqlSafe(Legajo) & "', " & _
                 "@CUIL='" & SqlSafe(CUIL) & "', " & _
                 "@DNI='" & SqlSafe(DNI) & "', " & _
                 "@Apellido='" & SqlSafe(Apellido) & "', " & _
                 "@Nombre='" & SqlSafe(Nombre) & "', " & _
                 "@TelefonoWhatsApp='" & SqlSafe(TelefonoWhatsApp) & "', " & _
                 "@Email='" & SqlSafe(Email) & "', " & _
                 "@Activo=" & activoSQL & ", " & _
                 "@FechaIngreso=" & fechaSQL

        Set rsIns = Conn.Execute(sqlIns)

        If Err.Number <> 0 Then
            msgError = "No se pudo crear el empleado: " & Err.Description
            Err.Clear
        Else
            nuevoEmpleadoID = 0
            If Not rsIns Is Nothing Then
                If Not rsIns.EOF Then
                    nuevoEmpleadoID = CLng("0" & Nz(rsIns("EmpleadoID"),0))
                End If
                If rsIns.State = 1 Then rsIns.Close
                Set rsIns = Nothing
            End If

            Response.Redirect "rrhh_empleados.asp?ok=1&msg=" & Server.URLEncode("Empleado creado correctamente.")
        End If

        On Error GoTo 0
    End If
End If
%>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <title>Nuevo Empleado</title>
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
            background:linear-gradient(135deg, #1d4ed8, #2563eb);
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
        .campo input, .campo select {
            width:100%;
            padding:10px 12px;
            border:1px solid #d1d5db;
            border-radius:10px;
            box-sizing:border-box;
            background:#fff;
        }
        .acciones {
            display:flex;
            gap:10px;
            flex-wrap:wrap;
            margin-top:20px;
        }
        @media (max-width: 991px) {
            .main-content { margin-left:0; padding:20px 12px; }
            .form-grid { grid-template-columns:1fr; }
        }
    </style>
</head>
<body>
<!--#include file="header.asp" -->

<div class="main-content">
    <div class="page-header">
        <div>
            <h1 class="page-title"><i class="fa-solid fa-user-plus"></i> Nuevo Empleado</h1>
            <div class="page-subtitle">Alta de empleado</div>
        </div>
    </div>

    <% If msgError <> "" Then %>
        <div class="alert alert-danger"><%=Html(msgError)%></div>
    <% End If %>

    <div class="card card-custom">
        <div class="card-header">
            <strong><i class="fa-solid fa-address-card"></i> Datos del empleado</strong>
        </div>
        <div class="card-body">
            <form method="post" action="rrhh_empleado_nuevo.asp">
                <input type="hidden" name="accion" value="guardar">

                <div class="form-grid">
                    <div class="campo">
                        <label for="Legajo">Legajo</label>
                        <input type="text" name="Legajo" id="Legajo" value="<%=Html(Legajo)%>">
                    </div>

                    <div class="campo">
                        <label for="CUIL">CUIL</label>
                        <input type="text" name="CUIL" id="CUIL" value="<%=Html(CUIL)%>">
                    </div>

                    <div class="campo">
                        <label for="DNI">DNI</label>
                        <input type="text" name="DNI" id="DNI" value="<%=Html(DNI)%>">
                    </div>

                    <div class="campo">
                        <label for="FechaIngreso">Fecha Ingreso</label>
                        <input type="date" name="FechaIngreso" id="FechaIngreso" value="<%=Html(FechaIngreso)%>">
                    </div>

                    <div class="campo">
                        <label for="Apellido">Apellido</label>
                        <input type="text" name="Apellido" id="Apellido" value="<%=Html(Apellido)%>">
                    </div>

                    <div class="campo">
                        <label for="Nombre">Nombre</label>
                        <input type="text" name="Nombre" id="Nombre" value="<%=Html(Nombre)%>">
                    </div>

                    <div class="campo">
                        <label for="TelefonoWhatsApp">Teléfono WhatsApp</label>
                        <input type="text" name="TelefonoWhatsApp" id="TelefonoWhatsApp" value="<%=Html(TelefonoWhatsApp)%>">
                    </div>

                    <div class="campo">
                        <label for="Email">Email</label>
                        <input type="text" name="Email" id="Email" value="<%=Html(Email)%>">
                    </div>

                    <div class="campo">
                        <label for="Activo">Activo</label>
                        <select name="Activo" id="Activo">
                            <option value="1" <% If Activo = "1" Then Response.Write "selected" End If %>>Sí</option>
                            <option value="0" <% If Activo = "0" Then Response.Write "selected" End If %>>No</option>
                        </select>
                    </div>
                </div>

                <div class="acciones">
                    <button type="submit" class="btn btn-success">
                        <i class="fa-solid fa-floppy-disk"></i> Guardar
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
Conn.Close
Set Conn = Nothing
%>