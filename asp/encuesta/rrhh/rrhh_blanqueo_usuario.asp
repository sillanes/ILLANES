<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Response.Buffer   = True

If Trim(Session("currentUser") & "") = "" Then
    Response.Redirect "../login.asp"
End If

Function Nz(v, defaultValue)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Function HtmlEncode(v)
    HtmlEncode = Server.HTMLEncode("" & v)
End Function

Function CLngNz(v)
    If IsNull(v) Or Trim("" & v) = "" Or Not IsNumeric(v) Then
        CLngNz = 0
    Else
        CLngNz = CLng(v)
    End If
End Function

Dim accion, empleadoID
Dim mensaje, tipoMensaje
Dim cmdAccion, rsAccion, resultadoSP

accion     = LCase(Trim(Request.Form("accion") & ""))
empleadoID = CLngNz(Request.Form("empleadoid"))

mensaje     = ""
tipoMensaje = ""

'========================================================
' ACCION: BLANQUEAR USUARIO
'========================================================
If accion = "blanquear" And empleadoID > 0 Then
    On Error Resume Next

    Set cmdAccion = Server.CreateObject("ADODB.Command")
    Set cmdAccion.ActiveConnection = conn
    cmdAccion.CommandType = 4 ' adCmdStoredProc
    cmdAccion.CommandText = "rrhh.usp_Usuario_Blanqueo"

    ' Ajustar este parámetro si tu SP recibe otro nombre, por ejemplo @UsuarioID
    cmdAccion.Parameters.Append cmdAccion.CreateParameter("@EmpleadoID", 3, 1, , empleadoID)

    Set rsAccion = cmdAccion.Execute

    If Err.Number <> 0 Then
        mensaje = "Error al ejecutar el blanqueo: " & Err.Description
        tipoMensaje = "danger"
        Err.Clear
    Else
        resultadoSP = 0

        If Not rsAccion Is Nothing Then
            If Not rsAccion.EOF Then
                If rsAccion.Fields.Count > 0 Then
                    resultadoSP = CLngNz(rsAccion.Fields(0).Value)
                End If
            End If
        End If

        If resultadoSP = 1 Then
            mensaje = "El usuario fue blanqueado correctamente."
            tipoMensaje = "success"
        Else
            mensaje = "No se pudo blanquear el usuario."
            tipoMensaje = "warning"
        End If
    End If

    On Error GoTo 0

    If Not rsAccion Is Nothing Then
        If rsAccion.State = 1 Then rsAccion.Close
        Set rsAccion = Nothing
    End If

    Set cmdAccion = Nothing
End If

'========================================================
' LISTADO DE EMPLEADOS ACTIVOS
'========================================================
Dim cmd, rs

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 4 ' adCmdStoredProc
cmd.CommandText = "rrhh.usp_Emplado_GetAll"

Set rs = cmd.Execute
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Blanqueo de Usuarios</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="rrhh.css" />
    <link rel="stylesheet" href="ventas.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
	<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">

    <style>
        body {
            background-color: #f4f6f9;
        }

        .content-wrapper {
            margin-left: 260px;
            padding: 24px;
        }

        .page-title {
            font-size: 1.6rem;
            font-weight: 700;
            color: #1f2937;
            margin-bottom: 20px;
        }

        .card-custom {
            border: 0;
            border-radius: 16px;
            box-shadow: 0 4px 18px rgba(0,0,0,.08);
            overflow: hidden;
        }

        .card-header-custom {
            background: linear-gradient(135deg, #0d6efd, #0b5ed7);
            color: #fff;
            padding: 16px 20px;
            font-size: 1.05rem;
            font-weight: 600;
        }

        .table thead th {
            vertical-align: middle;
            white-space: nowrap;
        }

        .table tbody td {
            vertical-align: middle;
        }

        .btn-blanquear {
            min-width: 110px;
        }

        .badge-activo {
            background-color: #198754;
            font-size: .8rem;
        }

        @media (max-width: 991.98px) {
            .content-wrapper {
                margin-left: 0;
                padding: 16px;
            }
        }
    </style>
</head>
<body>

<!--#include file="header.asp" -->

    <div class="content-wrapper">
        <div class="page-title">
            <i class="fa-solid fa-user-shield me-2"></i>Blanqueo de Usuarios
        </div>

        <% If Trim(mensaje) <> "" Then %>
            <div class="alert alert-<%=tipoMensaje%> alert-dismissible fade show" role="alert">
                <%=HtmlEncode(mensaje)%>
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Cerrar"></button>
            </div>
        <% End If %>

        <div class="card card-custom">
            <div class="card-header-custom">
                Nómina de empleados activos
            </div>

            <div class="card-body p-0">
                <div class="table-responsive">
                    <table class="table table-hover table-bordered align-middle mb-0">
                        <thead class="table-light">
                            <tr>
                                <th style="width:110px;">Legajo</th>
                                <th>Apellido</th>
                                <th>Nombre</th>
                                <th style="width:110px;">Estado</th>
                                <th style="width:140px;" class="text-center">Acción</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            Dim tieneDatos
                            tieneDatos = False

                            If Not rs Is Nothing Then
                                If Not rs.EOF Then
                                    Do Until rs.EOF
                                        If CLngNz(rs("Activo")) = 1 Then
                                            tieneDatos = True
                            %>
                                            <tr>
                                                <td><%=HtmlEncode(Nz(rs("Legajo"), ""))%></td>
                                                <td><%=HtmlEncode(Nz(rs("Apellido"), ""))%></td>
                                                <td><%=HtmlEncode(Nz(rs("Nombre"), ""))%></td>
                                                <td>
                                                    <span class="badge badge-activo">Activo</span>
                                                </td>
                                                <td class="text-center">
                                                    <form method="post" action="rrhh_blanqueo_usuario.asp" class="d-inline" onsubmit="return confirm('¿Desea blanquear este usuario?');">
                                                        <input type="hidden" name="accion" value="blanquear">
                                                        <input type="hidden" name="empleadoid" value="<%=CLngNz(rs("EmpleadoID"))%>">
                                                        <button type="submit" class="btn btn-sm btn-warning btn-blanquear">
                                                            <i class="fa-solid fa-eraser me-1"></i> Blanquear
                                                        </button>
                                                    </form>
                                                </td>
                                            </tr>
                            <%
                                        End If
                                        rs.MoveNext
                                    Loop
                                End If
                            End If

                            If Not tieneDatos Then
                            %>
                                <tr>
                                    <td colspan="5" class="text-center text-muted py-4">
                                        No se encontraron empleados activos.
                                    </td>
                                </tr>
                            <%
                            End If
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script src="../assets/bootstrap/js/bootstrap.bundle.min.js"></script>
</body>
</html>
<%
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

Set cmd = Nothing

If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>