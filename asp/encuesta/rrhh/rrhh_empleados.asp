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
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Function Html(v)
    Html = Server.HTMLEncode(CStr(Nz(v, "")))
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
%> 

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <title>Nómina de Empleados</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="rrhh.css" />
    <link rel="stylesheet" href="ventas.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

    <style>
        body { background:#f4f6f9; }

        .main-content {
            margin-left:260px;
            padding:30px 20px;
        }

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

        .table thead th {
            white-space:nowrap;
            vertical-align:middle;
            cursor:pointer;
            user-select:none;
            position:relative;
        }

        .table thead th.sortable:hover {
            background:#eef4ff;
        }

        .table thead th .sort-icon {
            margin-left:6px;
            font-size:.75rem;
            color:#6b7280;
        }

        .table td {
            vertical-align:middle;
        }

        .badge-activo {
            background:#198754;
            color:#fff;
            padding:6px 10px;
            border-radius:999px;
            font-size:.8rem;
        }

        .badge-inactivo {
            background:#dc3545;
            color:#fff;
            padding:6px 10px;
            border-radius:999px;
            font-size:.8rem;
        }

        .btn-action {
            border-radius:10px;
        }

        .btn-vacaciones {
            background:#6f42c1;
            color:#fff !important;
        }

        .btn-vacaciones:hover {
            background:#5b36a3;
            color:#fff !important;
        }

        .search-box {
            max-width:620px;
        }

        .actions-inline {
            display:flex;
            gap:6px;
            flex-wrap:nowrap;
        }

        .alert-custom {
            border-radius:12px;
            padding:14px 16px;
            margin-bottom:18px;
            border:1px solid transparent;
            font-weight:600;
            box-shadow:0 4px 14px rgba(0,0,0,.04);
        }

        .alert-success-custom {
            background:#d1e7dd;
            color:#0f5132;
            border-color:#badbcc;
        }

        .alert-danger-custom {
            background:#f8d7da;
            color:#842029;
            border-color:#f5c2c7;
        }

        @media (max-width: 991px) {
            .main-content {
                margin-left:0;
                padding:20px 12px;
            }

            .actions-inline {
                flex-wrap:wrap;
            }
        }
    </style>
</head>
<body>
<!--#include file="header.asp" -->

<div class="main-content">
    <div class="page-header">
        <div>
            <h1 class="page-title"><i class="fa-solid fa-users"></i> Nómina de Empleados</h1>
            <div class="page-subtitle">Listado completo de empleados</div>
        </div>

        <a href="rrhh_empleado_nuevo.asp" class="btn btn-success">
            <i class="fa-solid fa-user-plus"></i> Nuevo Empleado
        </a>
    </div>

    <%
    Dim msg, ok
    msg = Trim(Request.QueryString("msg") & "")
    ok  = Trim(Request.QueryString("ok") & "")

    If msg <> "" Then
        If ok = "1" Then
    %>
        <div class="alert-custom alert-success-custom"><%=Html(msg)%></div>
    <%
        Else
    %>
        <div class="alert-custom alert-danger-custom"><%=Html(msg)%></div>
    <%
        End If
    End If
    %>

    <div class="card card-custom">
        <div class="card-header">
            <strong><i class="fa-solid fa-address-card"></i> Empleados</strong>
        </div>
        <div class="card-body">

            <div class="mb-3 search-box">
                <input type="text" id="txtBuscar" class="form-control" placeholder="Buscar por legajo, CUIL, DNI, apellido, nombre, fecha ingreso, antigüedad, días disponibles, teléfono o email">
            </div>

            <div class="table-responsive">
                <table class="table table-bordered table-hover align-middle" id="tablaEmpleados">
                    <thead class="table-light">
                        <tr>
                            <th>Acciones</th>
                            <th class="sortable">Legajo <span class="sort-icon">↕</span></th>
                            <th class="sortable">CUIL <span class="sort-icon">↕</span></th>
                            <th class="sortable">DNI <span class="sort-icon">↕</span></th>
                            <th class="sortable">Apellido <span class="sort-icon">↕</span></th>
                            <th class="sortable">Nombre <span class="sort-icon">↕</span></th>
                            <th class="sortable">Fecha Ingreso <span class="sort-icon">↕</span></th>
                            <th class="sortable">Antigüedad <span class="sort-icon">↕</span></th>
                            <th class="sortable">Días Tomados <span class="sort-icon">↕</span></th>
                            <th class="sortable">Días Disponibles <span class="sort-icon">↕</span></th>
                            <th class="sortable">Teléfono WhatsApp <span class="sort-icon">↕</span></th>
                            <th class="sortable">Email <span class="sort-icon">↕</span></th>
                            <th class="sortable">Activo <span class="sort-icon">↕</span></th>
                        </tr>
                    </thead>
                    <tbody>
                    <%
                    Dim cmd, rs, activoValor

                    Set cmd = Server.CreateObject("ADODB.Command")
                    Set cmd.ActiveConnection = Conn
                    cmd.CommandType = 4
                    cmd.CommandText = "rrhh.usp_Emplado_GetAll"

                    Set rs = cmd.Execute

                    If rs.EOF Then
                    %>
                        <tr>
                            <td colspan="12" class="text-center text-muted">No hay empleados para mostrar.</td>
                        </tr>
                    <%
                    Else
                        Do Until rs.EOF
                            activoValor = CLng("0" & Nz(rs("Activo"), 0))
                    %>
                        <tr>
                            <td>
                                <div class="actions-inline">
                                    <a href="rrhh_empleado_editar.asp?EmpleadoID=<%=CLng(Nz(rs("EmpleadoID"),0))%>" class="btn btn-sm btn-primary btn-action" title="Editar empleado">
                                        <i class="fa-solid fa-pen-to-square"></i>
                                    </a>

                                    <a href="rrhh_empleado_legajo.asp?EmpleadoID=<%=CLng(Nz(rs("EmpleadoID"),0))%>" class="btn btn-sm btn-info btn-action" title="Ver legajo">
                                        <i class="fa-solid fa-folder-open"></i>
                                    </a>

                                    <a href="rrhh_empleado_vacaciones_nuevo.asp?EmpleadoID=<%=CLng(Nz(rs("EmpleadoID"),0))%>" class="btn btn-sm btn-vacaciones btn-action" title="Agregar vacaciones">
                                        <i class="fa-solid fa-calendar-plus"></i>
                                    </a>
                                </div>
                            </td>
                            <td><%=Html(rs("Legajo"))%></td>
                            <td><%=Html(rs("CUIL"))%></td>
                            <td><%=Html(rs("DNI"))%></td>
                            <td><%=Html(rs("Apellido"))%></td>
                            <td><%=Html(rs("Nombre"))%></td>
                            <td><%=FechaFormato(rs("FechaIngreso"))%></td>
                            <td><%=Html(rs("Antiguedad"))%></td>
                            <td><%=Html(rs("DiasTomados"))%></td>
                            <td><%=Html(rs("DiasDisponibles"))%></td>
                            <td><%=Html(rs("TelefonoWhatsApp"))%></td>
                            <td><%=Html(rs("Email"))%></td>
                            <td>
                                <% If activoValor <> 0 Then %>
                                    <span class="badge-activo">Sí</span>
                                <% Else %>
                                    <span class="badge-inactivo">No</span>
                                <% End If %>
                            </td>
                        </tr>
                    <%
                            rs.MoveNext
                        Loop
                    End If

                    rs.Close
                    Set rs = Nothing
                    Set cmd = Nothing
                    %>
                    </tbody>
                </table>
            </div>

        </div>
    </div>
</div>

<script>
document.getElementById('txtBuscar').addEventListener('keyup', function () {
    var filtro = this.value.toLowerCase();
    var filas = document.querySelectorAll('#tablaEmpleados tbody tr');

    filas.forEach(function (fila) {
        var texto = fila.innerText.toLowerCase();
        fila.style.display = (texto.indexOf(filtro) > -1) ? '' : 'none';
    });
});

(function () {
    const table = document.getElementById('tablaEmpleados');
    const headers = table.querySelectorAll('thead th.sortable');
    const tbody = table.querySelector('tbody');
    let sortState = {};

    function getCellValue(row, index) {
        const cell = row.children[index];
        return cell ? cell.innerText.trim() : '';
    }

    function parseValue(value) {
        const v = value.trim();

        if (/^\d{2}-\d{2}-\d{4}$/.test(v)) {
            const parts = v.split('-');
            return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
        }

        if (/^\d+\s+años$/i.test(v)) {
            return parseInt(v, 10);
        }

        if (/^\d+$/.test(v)) {
            return parseInt(v, 10);
        }

        return v.toLowerCase();
    }

    headers.forEach(function (header) {
        header.addEventListener('click', function () {
            const index = Array.prototype.indexOf.call(header.parentNode.children, header);
            const rows = Array.from(tbody.querySelectorAll('tr')).filter(function (r) {
                return r.children.length > 1;
            });

            const current = sortState[index] === 'asc' ? 'desc' : 'asc';
            sortState = {};
            sortState[index] = current;

            rows.sort(function (a, b) {
                const aVal = parseValue(getCellValue(a, index));
                const bVal = parseValue(getCellValue(b, index));

                if (aVal < bVal) return current === 'asc' ? -1 : 1;
                if (aVal > bVal) return current === 'asc' ? 1 : -1;
                return 0;
            });

            rows.forEach(function (row) {
                tbody.appendChild(row);
            });
        });
    });
})();
</script>

</body>
</html>
<%
Conn.Close
Set Conn = Nothing
%>