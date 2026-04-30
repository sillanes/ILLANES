<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->

<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else
        IIf = sFalse
    End If
End Function

If Session("currentUser") = "" Then
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
    HtmlEncode = Server.HTMLEncode("" & Nz(v, ""))
End Function

Function SafeSql(v)
    SafeSql = Replace(Trim("" & v), "'", "''")
End Function

Function FormatMoney2(v)
    If IsNull(v) Or v = "" Then
        FormatMoney2 = "0,00"
    Else
        FormatMoney2 = FormatNumber(CDbl(v), 2, -1, 0, -1)
    End If
End Function

Function FormatDateAR(v)
    If IsDate(v) Then
        FormatDateAR = Right("0" & Day(v),2) & "/" & Right("0" & Month(v),2) & "/" & Year(v)
    Else
        FormatDateAR = "-"
    End If
End Function

Function BadgeEstado(estadoTexto)
    Select Case UCase("" & estadoTexto)
        Case "CARGADA"
            BadgeEstado = "<span class='badge bg-secondary'>Cargada</span>"
        Case "EN_CONTROL"
            BadgeEstado = "<span class='badge bg-warning text-dark'>En control</span>"
        Case "CERRADA"
            BadgeEstado = "<span class='badge bg-success'>Cerrada</span>"
        Case Else
            BadgeEstado = "<span class='badge bg-light text-dark border'>" & HtmlEncode(estadoTexto) & "</span>"
    End Select
End Function

Dim estado, proveedor, numeroCarga, fechaDesde, fechaHasta
Dim sql, rs
Dim msg, errMsg
Dim totalFacturas, totalPendientes, totalConfirmadas

estado      = Trim("" & Request("estado"))
proveedor   = Trim("" & Request("proveedor"))
numeroCarga = Trim("" & Request("numeroCarga"))
fechaDesde  = Trim("" & Request("fechaDesde"))
fechaHasta  = Trim("" & Request("fechaHasta"))
msg         = Trim("" & Request("msg"))
errMsg      = Trim("" & Request("err"))

sql = "EXEC dbo.usp_Camiones_Facturas_Sel " & _
      "@Estado=" & IIf(estado = "", "NULL", "'" & SafeSql(estado) & "'") & ", " & _
      "@FechaDesde=" & IIf(fechaDesde = "", "NULL", "'" & SafeSql(fechaDesde) & "'") & ", " & _
      "@FechaHasta=" & IIf(fechaHasta = "", "NULL", "'" & SafeSql(fechaHasta) & "'") & ", " & _
      "@Proveedor=" & IIf(proveedor = "", "NULL", "'" & SafeSql(proveedor) & "'") & ", " & _
      "@NumeroCarga=" & IIf(numeroCarga = "", "NULL", "'" & SafeSql(numeroCarga) & "'")

Set rs = conn.Execute(sql)

totalFacturas = 0
totalPendientes = 0
totalConfirmadas = 0
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <title>Control de Camiones - Facturas</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <link rel="stylesheet" href="ventas.css">
    <link rel="stylesheet" href="estilos.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" rel="stylesheet">
    <style>
        body { background:#f5f7fb; }
        .main-content { padding:24px; }
        .page-card {
            background:#fff;
            border-radius:18px;
            box-shadow:0 6px 18px rgba(0,0,0,.08);
        }
        .summary-card {
            border:0;
            border-radius:16px;
            box-shadow:0 4px 14px rgba(0,0,0,.06);
        }
        .summary-number {
            font-size:1.8rem;
            font-weight:700;
        }
        .nowrap { white-space:nowrap; }
        .table thead th { white-space:nowrap; }
        .acciones-superiores {
            display:flex;
            flex-wrap:wrap;
            gap:.5rem;
        }
    </style>
</head>
<body>

<!--#include file="header.asp" --> 
<div class="main-content">
<div class="container-fluid main-content">
    <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
        <div>
            <h2 class="mb-1"><i class="fa-solid fa-truck-ramp-box me-2"></i>Control de Camiones</h2>
            <div class="text-muted">Listado de facturas cargadas</div>
        </div>
        <div class="acciones-superiores">
            <a href="camiones_facturas_carga.asp" class="btn btn-primary">
                <i class="fa-solid fa-file-arrow-up me-1"></i> Cargar PDF
            </a>
            <a href="camiones_facturas_carga_csv.asp" class="btn btn-success">
                <i class="fa-solid fa-file-csv me-1"></i> Cargar CSV
            </a>
        </div>
    </div>

    <% If msg <> "" Then %>
        <div class="alert alert-success"><%=HtmlEncode(msg)%></div>
    <% End If %>

    <% If errMsg <> "" Then %>
        <div class="alert alert-danger"><%=HtmlEncode(errMsg)%></div>
    <% End If %>

    <div class="page-card p-3 p-md-4 mb-4">
        <form method="get" action="camiones_facturas.asp" class="row g-3">
            <div class="col-md-2">
                <label class="form-label">Estado</label>
                <select name="estado" class="form-select">
                    <option value="" <% If estado = "" Then Response.Write("selected") End If %>>Todos</option>
                    <option value="CARGADA" <% If UCase(estado) = "CARGADA" Then Response.Write("selected") End If %>>Cargada</option>
                    <option value="EN_CONTROL" <% If UCase(estado) = "EN_CONTROL" Then Response.Write("selected") End If %>>En control</option>
                    <option value="CERRADA" <% If UCase(estado) = "CERRADA" Then Response.Write("selected") End If %>>Cerrada</option>
                </select>
            </div>

            <div class="col-md-3">
                <label class="form-label">Proveedor</label>
                <input type="text" name="proveedor" class="form-control" value="<%=HtmlEncode(proveedor)%>">
            </div>

            <div class="col-md-2">
                <label class="form-label">N° Carga</label>
                <input type="text" name="numeroCarga" class="form-control" value="<%=HtmlEncode(numeroCarga)%>">
            </div>

            <div class="col-md-2">
                <label class="form-label">Fecha desde</label>
                <input type="date" name="fechaDesde" class="form-control" value="<%=HtmlEncode(fechaDesde)%>">
            </div>

            <div class="col-md-2">
                <label class="form-label">Fecha hasta</label>
                <input type="date" name="fechaHasta" class="form-control" value="<%=HtmlEncode(fechaHasta)%>">
            </div>

            <div class="col-md-1 d-flex align-items-end">
                <button type="submit" class="btn btn-dark w-100">
                    <i class="fa-solid fa-magnifying-glass"></i>
                </button>
            </div>
        </form>
    </div>

    <%
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            totalFacturas = totalFacturas + 1
            totalPendientes = totalPendientes + CLng(Nz(rs("Pendientes"), 0))
            totalConfirmadas = totalConfirmadas + CLng(Nz(rs("Confirmados"), 0))
            rs.MoveNext
        Loop
        rs.MoveFirst
    End If
    %>

    <div class="row g-3 mb-4">
        <div class="col-md-4">
            <div class="card summary-card">
                <div class="card-body">
                    <div class="text-muted">Facturas</div>
                    <div class="summary-number"><%=totalFacturas%></div>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card summary-card">
                <div class="card-body">
                    <div class="text-muted">Ítems pendientes</div>
                    <div class="summary-number"><%=totalPendientes%></div>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card summary-card">
                <div class="card-body">
                    <div class="text-muted">Ítems confirmados</div>
                    <div class="summary-number"><%=totalConfirmadas%></div>
                </div>
            </div>
        </div>
    </div>

    <div class="page-card p-3 p-md-4">
        <div class="table-responsive">
            <table class="table table-hover align-middle">
                <thead class="table-light">
                    <tr>
                        <th>Fecha</th>
                        <th>Proveedor</th>
                        <th>Factura</th>
                        <th>N° Carga</th> 
                        <th class="text-center">Items</th>
                        <th class="text-center">Pend.</th>
                        <th class="text-center">Conf.</th>
                        <th class="text-center">Estado</th>
                        <th class="text-center">Acción</th>
                    </tr>
                </thead>
                <tbody>
                <% If rs.EOF Then %>
                    <tr>
                        <td colspan="9" class="text-center text-muted py-4">No se encontraron facturas.</td>
                    </tr>
                <% Else %>
                    <% Do Until rs.EOF %>
                        <tr>
                            <td class="nowrap"><%=FormatDateAR(rs("FechaFactura"))%></td>
                            <td><%=HtmlEncode(rs("ProveedorNombre"))%></td>
                            <td class="nowrap"><strong><%=HtmlEncode(rs("NumeroFactura"))%></strong></td>
                            <td class="nowrap"><%=HtmlEncode(rs("NumeroCarga"))%></td>
                            <!-- <td><%=HtmlEncode(rs("ArchivoNombre"))%></td> -->
                            <td class="text-center"><%=Nz(rs("TotalItems"),0)%></td>
                            <td class="text-center"><%=Nz(rs("Pendientes"),0)%></td>
                            <td class="text-center"><%=Nz(rs("Confirmados"),0)%></td>
                            <td class="text-center"><%=BadgeEstado(rs("Estado"))%></td>
                            <td class="text-center">
                                <a href="camiones_factura_control.asp?FacturaID=<%=rs("FacturaID")%>" class="btn btn-sm btn-primary" title="Controlar">
                                    <i class="fa-solid fa-clipboard-check"></i>
                                </a>
                            </td>
                        </tr>
                    <%
                        rs.MoveNext
                    Loop
                    %>
                <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>
 
</div>
 
</body>
</html>
<%
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
%>