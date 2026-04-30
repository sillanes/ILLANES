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

Dim cabeceraID, camionID, sqlCab, sqlDet, rsCab, rsDet
Dim msg, errMsg, volverUrl
Dim articuloBusqueda, articuloSP, tituloItems

Set rsCab = Nothing
Set rsDet = Nothing

cabeceraID       = CLng(0 & Request.QueryString("CabeceraID"))
camionID         = CLng(0 & Request.QueryString("CamionID"))
msg              = Trim("" & Request.QueryString("msg"))
errMsg           = Trim("" & Request.QueryString("err"))
articuloBusqueda = Trim("" & Request.QueryString("Articulo"))

If camionID > 0 Then
    volverUrl = "camiones_facturas.asp?CamionID=" & camionID
Else
    volverUrl = "camiones_facturas.asp"
End If

If cabeceraID <= 0 Then
    Response.Redirect volverUrl & "&err=" & Server.URLEncode("CabeceraID invalido.")
End If

sqlCab = "EXEC dbo.usp_Camiones_Cabecera_Control_Sel @CabeceraID=" & cabeceraID
Set rsCab = conn.Execute(sqlCab)

If rsCab.EOF Then
    Response.Redirect volverUrl & "&err=" & Server.URLEncode("La cabecera no existe.")
End If

If articuloBusqueda = "" Then
    articuloSP = "''"
    tituloItems = "Items pendientes"
Else
    articuloSP = "'" & SafeSql(articuloBusqueda) & "'"
    tituloItems = "Items pendientes filtrados por artículo: " & articuloBusqueda
End If

sqlDet = "EXEC dbo.usp_Camiones_Cabecera_Items_Pendientes_Sel " & _
         "@CabeceraID=" & cabeceraID & ", " & _
         "@Articulo=" & articuloSP
Set rsDet = conn.Execute(sqlDet)
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <title>Control de Factura</title>
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
        .item-card {
            border:1px solid #e9ecef;
            border-radius:16px;
            background:#fff;
        }
        .item-title {
            font-weight:700;
            font-size:1.05rem;
        }
        .item-meta {
            color:#6c757d;
            font-size:.92rem;
        }
        .busqueda-box {
            border:1px solid #dbe7f3;
            background:#f8fbff;
            border-radius:16px;
            padding:18px;
        }
    </style>
</head>
<body>

<!--#include file="header.asp" -->
<div class="main-content">
<div class="container-fluid main-content">
    <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
        <div>
            <h2 class="mb-1"><i class="fa-solid fa-clipboard-check me-2"></i>Control de factura</h2>
            <div class="text-muted">Confirmacion de items pendientes</div>
        </div>
        <div>
            <a href="<%=volverUrl%>" class="btn btn-outline-secondary">
                <i class="fa-solid fa-arrow-left me-1"></i> Volver
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
        <div class="row g-3">
            <div class="col-md-3">
                <div class="text-muted small">Proveedor</div>
                <div><strong><%=HtmlEncode(rsCab("ProveedorNombre"))%></strong></div>
            </div>
            <div class="col-md-2">
                <div class="text-muted small">Factura</div>
                <div><strong><%=HtmlEncode(rsCab("NumeroFactura"))%></strong></div>
            </div>
            <div class="col-md-2">
                <div class="text-muted small">Fecha</div>
                <div><%=FormatDateAR(rsCab("FechaFactura"))%></div>
            </div>
            <div class="col-md-2">
                <div class="text-muted small">Nro Carga</div>
                <div><%=HtmlEncode(rsCab("NumeroCarga"))%></div>
            </div>
            <div class="col-md-3">
                <div class="text-muted small">Estado</div>
                <div><%=BadgeEstado(rsCab("Estado"))%></div>
            </div>

            <div class="col-md-4">
                <div class="text-muted small">Orden de Compra</div>
                <div><%=HtmlEncode(rsCab("OrdenCompra"))%></div>
            </div>
            <div class="col-md-4">
                <div class="text-muted small">Remito</div>
                <div><%=HtmlEncode(rsCab("NroRemito"))%></div>
            </div>
            <div class="col-md-4">
                <div class="text-muted small">Chofer</div>
                <div><%=HtmlEncode(rsCab("Chofer"))%></div>
            </div>
        </div>
    </div>

    <div class="page-card p-3 p-md-4 mb-4">
        <div class="busqueda-box">
            <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
                <div>
                    <h5 class="mb-1"><i class="fa-solid fa-magnifying-glass me-2"></i>Buscar artículo dentro de la factura</h5>
                    <div class="text-muted">Ingresa el código de artículo para filtrar los ítems pendientes</div>
                </div>
            </div>

            <form method="get" action="camiones_factura_control.asp" class="row g-3 align-items-end">
                <input type="hidden" name="CabeceraID" value="<%=cabeceraID%>">
                <input type="hidden" name="CamionID" value="<%=camionID%>">

                <div class="col-md-5">
                    <label class="form-label">Artículo</label>
                    <input type="text" name="Articulo" class="form-control" value="<%=HtmlEncode(articuloBusqueda)%>" placeholder="Ej: 140203">
                </div>

                <div class="col-auto">
                    <button type="submit" class="btn btn-primary">
                        <i class="fa-solid fa-magnifying-glass me-1"></i> Buscar
                    </button>
                </div>

                <div class="col-auto">
                    <a href="camiones_factura_control.asp?CabeceraID=<%=cabeceraID%>&CamionID=<%=camionID%>" class="btn btn-outline-secondary">
                        <i class="fa-solid fa-eraser me-1"></i> Limpiar
                    </a>
                </div>
            </form>

            <% If articuloBusqueda <> "" Then %>
                <div class="alert alert-info mt-3 mb-0">
                    Mostrando solo los ítems pendientes que coinciden con el artículo <strong><%=HtmlEncode(articuloBusqueda)%></strong>.
                </div>
            <% End If %>
        </div>
    </div>

    <div class="page-card p-3 p-md-4">
        <h5 class="mb-3"><%=HtmlEncode(tituloItems)%></h5>

        <% If rsDet.EOF Then %>
            <% If articuloBusqueda <> "" Then %>
                <div class="alert alert-warning mb-0">
                    <i class="fa-solid fa-triangle-exclamation me-1"></i>
                    No se encontraron ítems pendientes para el artículo buscado.
                </div>
            <% Else %>
                <div class="alert alert-success mb-0">
                    <i class="fa-solid fa-circle-check me-1"></i>
                    No hay items pendientes. La factura quedo totalmente controlada.
                </div>
            <% End If %>
        <% Else %>
            <% Do Until rsDet.EOF %>
                <div class="item-card p-3 mb-3">
                    <form method="post" action="camiones_factura_item_guardar.asp" class="row g-3">
                        <input type="hidden" name="CabeceraID" value="<%=cabeceraID%>">
                        <input type="hidden" name="CamionID" value="<%=camionID%>">
                        <input type="hidden" name="ItemID" value="<%=rsDet("ItemID")%>">

                        <div class="col-md-8">
                            <div class="item-title">
                                Articulo <%=HtmlEncode(rsDet("Articulo"))%> - <%=HtmlEncode(rsDet("Descripcion"))%>
                            </div>
                            <div class="item-meta mt-1">
                                Cantidad factura: <strong><%=HtmlEncode(rsDet("CantidadTexto"))%></strong>
                                <% If Trim("" & Nz(rsDet("CodigoEAN"), "")) <> "" Then %>
                                    | EAN: <strong><%=HtmlEncode(rsDet("CodigoEAN"))%></strong>
                                <% End If %>
                            </div>
                        </div>

                        <div class="col-md-2">
                            <label class="form-label">Cantidad controlada</label>
                            <input type="text" name="CantidadControlada" class="form-control" placeholder="Ej: 16 BULTOS">
                        </div>

                        <div class="col-md-2 d-flex align-items-end">
                            <div class="form-check">
                                <input class="form-check-input" type="checkbox" name="Completo" id="completo_<%=rsDet("ItemID")%>" value="1">
                                <label class="form-check-label" for="completo_<%=rsDet("ItemID")%>">Completo</label>
                            </div>
                        </div>

                        <div class="col-md-10">
                            <label class="form-label">Observacion</label>
                            <input type="text" name="Observacion" class="form-control" maxlength="300">
                        </div>

                        <div class="col-md-2 d-flex align-items-end">
                            <button type="submit" class="btn btn-success w-100">
                                <i class="fa-solid fa-floppy-disk me-1"></i> Guardar
                            </button>
                        </div>
                    </form>
                </div>
            <%
                rsDet.MoveNext
            Loop
            %>
        <% End If %>
    </div>
</div>

</div>

</body>
</html>
<%
If IsObject(rsCab) Then
    If rsCab.State = 1 Then rsCab.Close
    Set rsCab = Nothing
End If

If IsObject(rsDet) Then
    If rsDet.State = 1 Then rsDet.Close
    Set rsDet = Nothing
End If
%>