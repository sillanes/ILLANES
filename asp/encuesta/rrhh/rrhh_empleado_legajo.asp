<%@LANGUAGE="VBScript" CODEPAGE="65001"%>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Session.LCID = 11274

If Trim(Session("currentUser") & "") = "" Then
    Response.Redirect "login.asp"
End If

Function Nz(v, defaultValue)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Function Html(v)
    Html = Server.HTMLEncode(CStr(Nz(v, "")))
End Function

Dim EmpleadoID
EmpleadoID = 0

If IsNumeric(Request.QueryString("EmpleadoID")) Then
    EmpleadoID = CLng(Request.QueryString("EmpleadoID"))
End If

If EmpleadoID <= 0 Then
    Response.Write "<h3 style='font-family:Arial;color:#b00020;'>EmpleadoID inválido.</h3>"
    Response.End
End If
%> 
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <title>Legajo del Empleado</title>
	<meta name="viewport" content="width=device-width, initial-scale=1.0">

	<link rel="stylesheet" href="estilos.css">
	<link rel="stylesheet" href="rrhh.css" />
	<link rel="stylesheet" href="ventas.css"> 


 
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" rel="stylesheet">
    <link href="estilos.css" rel="stylesheet">


    <style>
        body { background:#f4f6f9; }
        .main-content { margin-left:260px; padding:30px 20px; }
        .page-header {
            display:flex;
            justify-content:space-between;
            align-items:center;
            gap:12px;
            flex-wrap:wrap;
            margin-bottom:25px;
        }
        .page-title {
            font-size:1.8rem;
            font-weight:700;
            color:#1f2937;
            margin:0;
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
            background:linear-gradient(135deg, #0d6efd, #3b82f6);
            color:#fff;
            border:0;
            padding:18px 22px;
        }
        .card-custom .card-body {
            padding:25px;
        }
        .data-row {
            border-bottom:1px solid #eef2f7;
            padding:12px 0;
        }
        .data-row:last-child {
            border-bottom:none;
        }
        .label {
            color:#6b7280;
            font-weight:600;
            margin-bottom:4px;
        }
        .value {
            color:#111827;
            font-size:1rem;
        }
        .badge-activo {
            background:#198754;
            color:#fff;
            padding:6px 10px;
            border-radius:999px;
            font-size:.8rem;
            display:inline-block;
        }
        .badge-inactivo {
            background:#dc3545;
            color:#fff;
            padding:6px 10px;
            border-radius:999px;
            font-size:.8rem;
            display:inline-block;
        }
        .actions {
            display:flex;
            gap:10px;
            flex-wrap:wrap;
            margin-top:25px;
        }
        .btn {
            border-radius:10px;
            padding:10px 18px;
            font-weight:600;
        }
        @media (max-width: 991px) {
            .main-content { margin-left:0; padding:20px 12px; }
        }
    </style>
</head>
<body>

<!--#include file="header.asp" -->
 

<div class="main-content">
    <div class="page-header">
        <div>
            <h1 class="page-title"><i class="fa-solid fa-folder-open"></i> Legajo del empleado</h1>
            <div class="page-subtitle">Ficha del empleado</div>
        </div>
    </div>

    <%
    Dim cmd, rs
    Dim Legajo, CUIL, DNI, Apellido, Nombre, TelefonoWhatsApp, Email, Activo
    Dim msg, ok

    Legajo = ""
    CUIL = ""
    DNI = ""
    Apellido = ""
    Nombre = ""
    TelefonoWhatsApp = ""
    Email = ""
    Activo = 0

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = 4
    cmd.CommandText = "rrhh.usp_Empleado_sel"
    cmd.Parameters.Append cmd.CreateParameter("@EmpleadoID", 3, 1, , EmpleadoID)

    Set rs = cmd.Execute

    If rs.EOF Then
    %>
        <div class="alert alert-danger">No se encontró el empleado solicitado.</div>
        <a href="rrhh_empleados.asp" class="btn btn-secondary">
            <i class="fa-solid fa-arrow-left"></i> Volver
        </a>
    <%
        rs.Close
        Set rs = Nothing
        Set cmd = Nothing
        Conn.Close
        Set Conn = Nothing
        Response.End
    Else
        Legajo = CStr(Nz(rs("Legajo"), ""))
        CUIL = CStr(Nz(rs("CUIL"), ""))
        DNI = CStr(Nz(rs("DNI"), ""))
        Apellido = CStr(Nz(rs("Apellido"), ""))
        Nombre = CStr(Nz(rs("Nombre"), ""))
        TelefonoWhatsApp = CStr(Nz(rs("TelefonoWhatsApp"), ""))
        Email = CStr(Nz(rs("Email"), ""))
        Activo = CInt(Nz(rs("iActivo"), 0))
    End If

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    msg = Trim(Request.QueryString("msg") & "")
    ok  = Trim(Request.QueryString("ok") & "")
    %>

    <% If msg <> "" Then %>
        <% If ok = "1" Then %>
            <div class="alert alert-success"><%=Html(msg)%></div>
        <% Else %>
            <div class="alert alert-danger"><%=Html(msg)%></div>
        <% End If %>
    <% End If %>

    <div class="card card-custom">
        <div class="card-header">
            <h5 class="mb-0"><i class="fa-solid fa-id-badge"></i> Datos del empleado</h5>
        </div>
        <div class="card-body">

            <div class="row">
                <div class="col-md-6 data-row">
                    <div class="label">Legajo</div>
                    <div class="value"><%=Html(Legajo)%></div>
                </div>

                <div class="col-md-6 data-row">
                    <div class="label">Estado</div>
                    <div class="value">
                        <% If Activo = 1 Then %>
                            <span class="badge-activo">Activo</span>
                        <% Else %>
                            <span class="badge-inactivo">Inactivo</span>
                        <% End If %>
                    </div>
                </div>

                <div class="col-md-6 data-row">
                    <div class="label">Apellido</div>
                    <div class="value"><%=Html(Apellido)%></div>
                </div>

                <div class="col-md-6 data-row">
                    <div class="label">Nombre</div>
                    <div class="value"><%=Html(Nombre)%></div>
                </div>

                <div class="col-md-6 data-row">
                    <div class="label">CUIL</div>
                    <div class="value"><%=Html(CUIL)%></div>
                </div>

                <div class="col-md-6 data-row">
                    <div class="label">DNI</div>
                    <div class="value"><%=Html(DNI)%></div>
                </div>

                <div class="col-md-6 data-row">
                    <div class="label">Teléfono WhatsApp</div>
                    <div class="value"><%=Html(TelefonoWhatsApp)%></div>
                </div>

                <div class="col-md-6 data-row">
                    <div class="label">Email</div>
                    <div class="value"><%=Html(Email)%></div>
                </div>
            </div>

            <div class="actions">
                <a href="rrhh_empleados.asp" class="btn btn-secondary">
                    <i class="fa-solid fa-arrow-left"></i> Volver
                </a>

                <a href="rrhh_empleado_editar.asp?EmpleadoID=<%=EmpleadoID%>" class="btn btn-primary">
                    <i class="fa-solid fa-pen-to-square"></i> Editar datos
                </a>
            </div>

        </div>
    </div>
</div>

</body>
</html>
<%
Conn.Close
Set Conn = Nothing
%>