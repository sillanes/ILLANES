<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->

<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Server.ScriptTimeout = 3600

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else
        IIf = sFalse
    End If
End Function

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

Function FormatDateAR(v)
    If IsDate(v) Then
        FormatDateAR = Right("0" & Day(v),2) & "/" & Right("0" & Month(v),2) & "/" & Year(v)
    Else
        FormatDateAR = "-"
    End If
End Function

Function FormatDateTimeAR(v)
    If IsDate(v) Then
        FormatDateTimeAR = Right("0" & Day(v),2) & "/" & Right("0" & Month(v),2) & "/" & Year(v) & " " & _
                           Right("0" & Hour(v),2) & ":" & Right("0" & Minute(v),2)
    Else
        FormatDateTimeAR = "-"
    End If
End Function

Function BadgeTemperatura(v)
    If Trim("" & Nz(v, "")) = "" Then
        BadgeTemperatura = "<span class='badge bg-danger'>Pendiente</span>"
    Else
        BadgeTemperatura = "<span class='badge bg-success'>" & Replace(FormatNumber(CDbl(v), 2, -1, 0, -1), ",", ".") & " &deg;C</span>"
    End If
End Function

Function BadgeEstadoCamion(v)
    Select Case CLng("0" & Nz(v, 0))
        Case 0
            BadgeEstadoCamion = "<span class='badge bg-secondary'>Pendiente</span>"
        Case 1
            BadgeEstadoCamion = "<span class='badge bg-warning text-dark'>En proceso</span>"
        Case 2
            BadgeEstadoCamion = "<span class='badge bg-success'>Finalizado</span>"
        Case Else
            BadgeEstadoCamion = "<span class='badge bg-light text-dark border'>Sin estado</span>"
    End Select
End Function

Function ExecScalarFromSP(sqlText, fieldName)
    Dim rsTmp, val
    val = ""

    Set rsTmp = Nothing
    Set rsTmp = conn.Execute(sqlText)

    If IsObject(rsTmp) Then
        If Not rsTmp.EOF Then
            val = rsTmp(fieldName)
        End If
        If rsTmp.State = 1 Then rsTmp.Close
    End If

    Set rsTmp = Nothing
    ExecScalarFromSP = val
End Function

Function FieldExists(rs, fieldName)
    Dim f
    FieldExists = False

    If Not IsObject(rs) Then Exit Function

    For Each f In rs.Fields
        If LCase(f.Name) = LCase(fieldName) Then
            FieldExists = True
            Exit Function
        End If
    Next
End Function

Function GetFieldValue(rs, fieldName, defaultValue)
    If FieldExists(rs, fieldName) Then
        GetFieldValue = Nz(rs(fieldName), defaultValue)
    Else
        GetFieldValue = defaultValue
    End If
End Function

Function IsCheckedValue(v)
    Dim s
    s = UCase(Trim("" & Nz(v, "")))

    If s = "1" Or s = "TRUE" Or s = "SI" Or s = "S" Then
        IsCheckedValue = True
    Else
        IsCheckedValue = False
    End If
End Function

Dim action, msg, errMsg
Dim camionIDSel, temperaturaControl
Dim rsCamiones, rsCamionSel, rsItems
Dim sqlCamiones, sqlCamionSel, sqlItems, sqlTemp, sqlGuardarItem, sqlFinalizarCamion
Dim camionSeleccionado, camionTieneTemperatura
Dim numeroCargaCamion, tituloCamion
Dim filasAfectadas
Dim hayRSCamiones, hayRSCamionSel, hayRSItems
Dim buscarArticulo, usandoBusquedaArticulo, textoBusquedaInfo
Dim totalItemsPendientes
Dim itemID, cantidadControlada, completo, observacion
Dim camionIDFinalizar

Set rsCamiones = Nothing
Set rsCamionSel = Nothing
Set rsItems = Nothing

hayRSCamiones = False
hayRSCamionSel = False
hayRSItems = False

action                 = LCase(Trim("" & Request("action")))
camionIDSel            = CLng("0" & Request("CamionID"))
temperaturaControl     = Replace(Trim("" & Request("TemperaturaControl")), ",", ".")
buscarArticulo         = Trim("" & Request("BuscarArticulo"))
usandoBusquedaArticulo = False
textoBusquedaInfo      = ""

msg    = Trim("" & Request("msg"))
errMsg = Trim("" & Request("err"))

camionSeleccionado     = False
camionTieneTemperatura = False
numeroCargaCamion      = ""
tituloCamion           = ""
filasAfectadas         = 0
totalItemsPendientes   = 0

If action = "finalizar_camion" Then
    camionIDFinalizar = CLng("0" & Request.Form("CamionIDFinalizar"))

    If camionIDFinalizar <= 0 Then
        errMsg = "No se recibio un camion valido para finalizar."
    Else
        On Error Resume Next

        sqlFinalizarCamion = "EXEC dbo.usp_Camiones_upd " & _
                             "@CamionID=" & camionIDFinalizar & ", " & _
                             "@Estado=2"

        conn.Execute sqlFinalizarCamion

        If Err.Number <> 0 Then
            errMsg = "No se pudo finalizar el camion: " & Err.Description
            Err.Clear
        Else
            msg = "Camion finalizado correctamente."
        End If

        On Error GoTo 0
    End If
End If

If action = "guardar_temperatura" Then
    If camionIDSel <= 0 Then
        errMsg = "No se recibio un camion valido."
    ElseIf Trim(temperaturaControl) = "" Then
        errMsg = "Debes ingresar la temperatura tomada."
    ElseIf Not IsNumeric(temperaturaControl) Then
        errMsg = "La temperatura ingresada no es valida."
    Else
        On Error Resume Next

        sqlTemp = "EXEC dbo.usp_Camiones_Temperatura_Upd " & _
                  "@CamionID=" & camionIDSel & ", " & _
                  "@TemperaturaControl=" & Replace(CStr(CDbl(temperaturaControl)), ",", ".") & ", " & _
                  "@UsuarioTemperatura='" & SafeSql(Session("currentUser")) & "'"

        filasAfectadas = CLng("0" & ExecScalarFromSP(sqlTemp, "FilasAfectadas"))

        If Err.Number <> 0 Then
            errMsg = "No se pudo guardar la temperatura: " & Err.Description
            Err.Clear
        ElseIf filasAfectadas > 0 Then
            msg = "Temperatura registrada correctamente."
        Else
            errMsg = "No se encontro el camion para actualizar."
        End If

        On Error GoTo 0
    End If
End If

If action = "guardar_item" Then
    itemID              = CLng("0" & Request.Form("ItemID"))
    cantidadControlada  = Trim("" & Request.Form("CantidadControlada"))
    observacion         = Trim("" & Request.Form("Observacion"))
    completo            = 0

    If Trim("" & Request.Form("Completo")) <> "" Then
        completo = 1
    End If

    If camionIDSel <= 0 Then
        errMsg = "No se recibio un camion valido."
    ElseIf itemID <= 0 Then
        errMsg = "No se recibio un item valido."
    ElseIf completo = 0 And Trim(cantidadControlada) = "" Then
        errMsg = "Debes marcar Completo o ingresar la cantidad controlada."
    Else
        On Error Resume Next

        sqlGuardarItem = "EXEC dbo.usp_Camiones_Cabecera_Item_Control_Ins " & _
                         "@ItemID=" & itemID & ", " & _
                         "@CantidadControlada='" & SafeSql(cantidadControlada) & "', " & _
                         "@Completo=" & completo & ", " & _
                         "@Observacion='" & SafeSql(observacion) & "', " & _
                         "@Usuario='" & SafeSql(Session("currentUser")) & "'"

        conn.Execute sqlGuardarItem

        If Err.Number <> 0 Then
            errMsg = "No se pudo guardar el control del item: " & Err.Description
            Err.Clear
        Else
            msg = "Control guardado correctamente."
        End If

        On Error GoTo 0
    End If
End If

If camionIDSel > 0 Then
    On Error Resume Next
    sqlCamionSel = "EXEC dbo.usp_Camiones_Sel @CamionID=" & camionIDSel & ", @NumeroCarga=NULL"
    Set rsCamionSel = conn.Execute(sqlCamionSel)

    If Err.Number <> 0 Then
        errMsg = "Error al consultar el camion seleccionado: " & Err.Description
        Err.Clear
    ElseIf IsObject(rsCamionSel) Then
        hayRSCamionSel = True

        If Not rsCamionSel.EOF Then
            camionSeleccionado = True
            numeroCargaCamion = Trim("" & Nz(rsCamionSel("NroCarga"), ""))
            tituloCamion = "Camion #" & rsCamionSel("CamionID") & " - Carga " & numeroCargaCamion

            If Trim("" & Nz(rsCamionSel("TemperaturaControl"), "")) <> "" Then
                camionTieneTemperatura = True
            End If
        End If
    End If
    On Error GoTo 0
End If

If Not camionSeleccionado Then
    On Error Resume Next
    sqlCamiones = "EXEC dbo.usp_Camiones_Sel @CamionID=NULL, @NumeroCarga=NULL"
    Set rsCamiones = conn.Execute(sqlCamiones)

    If Err.Number <> 0 Then
        errMsg = "Error al consultar camiones: " & Err.Description
        Err.Clear
    ElseIf IsObject(rsCamiones) Then
        hayRSCamiones = True
    End If
    On Error GoTo 0
End If

If camionSeleccionado And camionTieneTemperatura Then
    On Error Resume Next

    If Trim(buscarArticulo) <> "" Then
        usandoBusquedaArticulo = True
        textoBusquedaInfo = "Buscando por codigo o descripcion: " & buscarArticulo
    End If

    sqlItems = "EXEC dbo.usp_Camiones_Items_Pendientes_PorCamion_Sel " & _
               "@CamionID=" & camionIDSel & ", " & _
               "@Articulo='" & SafeSql(buscarArticulo) & "'"

    Set rsItems = conn.Execute(sqlItems)

    If Err.Number <> 0 Then
        errMsg = "Error al consultar los items pendientes del camion: " & Err.Description
        Err.Clear
    ElseIf IsObject(rsItems) Then
        hayRSItems = True

        If Not rsItems.EOF Then
            rsItems.MoveFirst
            Do Until rsItems.EOF
                totalItemsPendientes = totalItemsPendientes + 1
                rsItems.MoveNext
            Loop
            rsItems.MoveFirst
        End If
    End If

    On Error GoTo 0
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <title>Control de Camiones - Items</title>
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

        .compact-card {
            padding:12px 14px !important;
            border-radius:14px;
        }

        .compact-card h4,
        .compact-card h5 {
            font-size:1rem;
            margin-bottom:4px;
        }

        .compact-card .text-muted {
            font-size:.82rem;
        }

        .compact-card .data-grid {
            gap:6px 12px;
        }

        .compact-card .data-label {
            font-size:.70rem;
        }

        .compact-card .data-value {
            font-size:.85rem;
        }

        .compact-card .badge {
            font-size:.72rem;
        }

        .acciones-superiores {
            display:flex;
            flex-wrap:wrap;
            gap:.5rem;
        }

        .temperatura-box,
        .busqueda-articulo-box {
            border:1px solid #dbe7f3;
            background:#f8fbff;
            border-radius:14px;
            padding:12px 14px;
        }

        .camion-card,
        .item-card {
            border:1px solid #e9ecef;
            border-radius:18px;
            background:#fff;
            box-shadow:0 3px 10px rgba(0,0,0,.04);
            height:100%;
        }

        .card-title-main {
            font-size:1.05rem;
            font-weight:700;
            color:#212529;
            margin-bottom:.15rem;
        }

        .compact-card .card-title-main {
            font-size:.95rem;
        }

        .card-subtitle {
            color:#6c757d;
            font-size:.92rem;
        }

        .compact-card .card-subtitle {
            font-size:.80rem;
        }

        .data-grid {
            display:grid;
            grid-template-columns:repeat(2, minmax(0, 1fr));
            gap:12px 18px;
        }

        .data-item {
            min-width:0;
        }

        .data-label {
            font-size:.78rem;
            color:#6c757d;
            margin-bottom:2px;
        }

        .data-value {
            font-weight:600;
            color:#212529;
            word-break:break-word;
        }

        .card-actions {
            display:flex;
            flex-wrap:wrap;
            gap:.5rem;
        }

        .camion-actions {
            display:grid;
            grid-template-columns:1fr 1fr 1fr;
            gap:8px;
            margin-top:12px;
        }

        .btn-camion-trabajar,
        .btn-camion-finalizar,
        .btn-camion-excel {
            border:0;
            width:100%;
            min-height:44px;
            border-radius:14px;
            font-weight:700;
            font-size:.88rem;
            display:flex;
            align-items:center;
            justify-content:center;
            gap:7px;
            text-decoration:none;
            transition:all .15s ease;
            white-space:nowrap;
        }

        .btn-camion-trabajar {
            background:#1f2937;
            color:#fff;
        }

        .btn-camion-trabajar:hover {
            background:#111827;
            color:#fff;
            transform:translateY(-1px);
        }

        .btn-camion-excel {
            background:#ecfdf5;
            color:#047857;
            border:1px solid #bbf7d0;
        }

        .btn-camion-excel:hover {
            background:#dcfce7;
            color:#065f46;
            transform:translateY(-1px);
        }

        .btn-camion-finalizar {
            background:#fff1f2;
            color:#be123c;
            border:1px solid #fecdd3;
        }

        .btn-camion-finalizar:hover {
            background:#ffe4e6;
            color:#9f1239;
            transform:translateY(-1px);
        }

        .camion-header,
        .item-header {
            display:flex;
            justify-content:space-between;
            align-items:flex-start;
            gap:12px;
        }

        .item-divider {
            border-top:1px solid #eef1f4;
            margin:12px 0;
        }

        .btn-volver-mobile {
            border-radius:14px;
            font-weight:600;
            padding:12px 14px;
        }

        .paso-actual {
            font-size:.85rem;
            color:#6c757d;
        }

        .busqueda-articulo-box .form-label,
        .temperatura-box .form-label {
            font-size:.78rem;
            margin-bottom:3px;
        }

        .busqueda-articulo-box .form-control,
        .temperatura-box .form-control,
        .control-item-box .form-control,
        .control-item-box textarea {
            padding:.40rem .65rem;
            font-size:.90rem;
        }

        .busqueda-articulo-box .btn,
        .temperatura-box .btn,
        .control-item-box .btn {
            padding:.40rem .75rem;
            font-size:.90rem;
        }

        .item-code {
            font-size:1rem;
            font-weight:800;
        }

        .item-desc {
            font-size:.92rem;
            color:#495057;
        }

        .item-cantidad {
            font-size:1.05rem;
            font-weight:800;
        }

        .control-item-box {
            background:#f8fbff;
            border:1px solid #dbe7f3;
            border-radius:14px;
            padding:12px;
        }

        .control-item-title {
            font-size:.86rem;
            font-weight:700;
            color:#495057;
            margin-bottom:8px;
        }

        .form-check-label {
            font-size:.90rem;
            font-weight:600;
        }

        @media (max-width: 767.98px) {
            .main-content { padding:14px; }

            .data-grid {
                grid-template-columns:1fr;
                gap:10px;
            }

            .compact-card .data-grid {
                grid-template-columns:repeat(2, minmax(0, 1fr));
                gap:6px 10px;
            }

            .camion-header,
            .item-header {
                flex-direction:column;
                align-items:flex-start;
            }

            .compact-card .camion-header {
                flex-direction:row;
                align-items:flex-start;
            }

            .card-actions .btn {
                width:100%;
            }

            .acciones-superiores {
                width:100%;
            }

            .acciones-superiores .btn {
                width:100%;
            }

            .camion-actions {
                grid-template-columns:1fr 1fr 1fr;
            }

            .btn-camion-trabajar,
            .btn-camion-finalizar,
            .btn-camion-excel {
                font-size:.82rem;
                gap:5px;
                min-height:42px;
            }
        }
    </style>

    <script>
      function toggleSidebar(){
          document.querySelector('.sidebar').classList.toggle('open');
      }

      function validarControlItem(formulario) {
          var cantidad = formulario.CantidadControlada.value.trim();
          var completo = formulario.Completo.checked;

          if (!completo && cantidad === '') {
              alert('Debes marcar Completo o ingresar la cantidad controlada.');
              formulario.CantidadControlada.focus();
              return false;
          }

          return true;
      }

      function confirmarFinalizarCamion() {
          return confirm('¿Seguro que deseas finalizar este camión? Una vez finalizado no aparecerá en el listado principal.');
      }
    </script>
</head>
<body>

<!--#include file="header.asp" -->

<div class="main-content">
<div class="container-fluid main-content">

    <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
        <div>
            <h2 class="mb-1"><i class="fa-solid fa-truck-ramp-box me-2"></i>Control de Camiones</h2>

            <% If camionSeleccionado Then %>
                <div class="paso-actual">Trabajando con <%=HtmlEncode(tituloCamion)%></div>
            <% Else %>
                <div class="text-muted">Selecciona un camión para comenzar</div>
            <% End If %>
        </div>

        <% If Not camionSeleccionado Then %>
            <div class="acciones-superiores">
                <a href="camiones_facturas_carga_csv.asp" class="btn btn-success">
                    <i class="fa-solid fa-file-csv me-1"></i> Cargar CSV
                </a>
            </div>
        <% End If %>
    </div>

    <% If msg <> "" Then %>
        <div class="alert alert-success"><%=HtmlEncode(msg)%></div>
    <% End If %>

    <% If errMsg <> "" Then %>
        <div class="alert alert-danger"><%=HtmlEncode(errMsg)%></div>
    <% End If %>

    <% If camionSeleccionado Then %>
        <div class="mb-3">
            <a href="camiones_facturas.asp" class="btn btn-outline-dark w-100 btn-volver-mobile">
                <i class="fa-solid fa-arrow-left me-1"></i> Volver a camiones
            </a>
        </div>
    <% End If %>

    <% If Not camionSeleccionado Then %>

        <div class="page-card p-3 p-md-4 mb-4">
            <div class="d-flex justify-content-between align-items-center mb-3">
                <div>
                    <h4 class="mb-1">Camiones cargados</h4>
                    <div class="text-muted">Elige el camión con el que quieres trabajar</div>
                </div>
            </div>

            <% If Not hayRSCamiones Then %>
                <div class="text-center text-muted py-4">No se pudieron consultar los camiones.</div>
            <% ElseIf rsCamiones.EOF Then %>
                <div class="text-center text-muted py-4">No se encontraron camiones cargados.</div>
            <% Else %>
                <div class="row g-3">
                    <% Do Until rsCamiones.EOF %>

                         
                        <div class="col-12 col-md-6 col-xl-4">
                            <div class="camion-card p-3">
                                <div class="camion-header mb-3">
                                    <div>
                                        <div class="card-title-main">
                                            Camión #<%=rsCamiones("CamionID")%>
                                        </div>
                                        <div class="card-subtitle">
                                            Carga <%=HtmlEncode(rsCamiones("NroCarga"))%>
                                        </div>
                                    </div>
                                    <div>
                                        <%=BadgeTemperatura(rsCamiones("TemperaturaControl"))%>
                                    </div>
                                </div>

                                <div class="data-grid mb-3">
                                    <div class="data-item">
                                        <div class="data-label">Fecha</div>
                                        <div class="data-value"><%=FormatDateTimeAR(rsCamiones("FechaImportacion"))%></div>
                                    </div>
                                    <div class="data-item">
                                        <div class="data-label">Estado</div>
                                        <div class="data-value"><%=BadgeEstadoCamion(GetFieldValue(rsCamiones, "Estado", 0))%></div>
                                    </div>
                                    <div class="data-item">
                                        <div class="data-label">Pallets</div>
                                        <div class="data-value"><%=HtmlEncode(rsCamiones("CantidadPallets"))%></div>
                                    </div>
                                    <div class="data-item">
                                        <div class="data-label">Empresa transporte</div>
                                        <div class="data-value"><%=HtmlEncode(rsCamiones("EmpresaTransporte"))%></div>
                                    </div>
                                    <div class="data-item">
                                        <div class="data-label">Chofer</div>
                                        <div class="data-value"><%=HtmlEncode(rsCamiones("Chofer"))%></div>
                                    </div>
                                    <div class="data-item">
                                        <div class="data-label">Teléfono</div>
                                        <div class="data-value"><%=HtmlEncode(rsCamiones("Telefono"))%></div>
                                    </div>
                                </div>

                                <div class="camion-actions">
									
									<% If CLng("0" & GetFieldValue(rsCamiones, "Estado", 0)) <> 2 Then %>
                                    <a href="camiones_facturas.asp?CamionID=<%=rsCamiones("CamionID")%>" class="btn-camion-trabajar" title="Trabajar con este camión">
                                        <i class="fa-solid fa-truck-fast"></i>
                                        <span>Trabajar</span>
                                    </a>

				
                                    <form method="post" action="camiones_facturas.asp" class="m-0" onsubmit="return confirmarFinalizarCamion();">
                                        <input type="hidden" name="action" value="finalizar_camion">
                                        <input type="hidden" name="CamionIDFinalizar" value="<%=rsCamiones("CamionID")%>">
                                        <button type="submit" class="btn-camion-finalizar">
                                            <i class="fa-solid fa-flag-checkered"></i>
                                            <span>Finalizar</span>
                                        </button>
                                    </form>
									

									<% else %>
										<a href="camiones_items_exportar_excel.asp?CamionID=<%=rsCamiones("CamionID")%>" class="btn-camion-excel" title="Descargar Excel">
											<i class="fa-solid fa-file-excel"></i>
											<span>Excel</span>
										</a>
									<% End If %>

									
                                </div>
                            </div>
                        </div> 

                    <%
                        rsCamiones.MoveNext
                    Loop
                    %>
                </div>
            <% End If %>
        </div>

    <% Else %>

        <div class="page-card compact-card mb-3">
            <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-2">
                <div>
                    <h4 class="mb-1">Camión seleccionado</h4>
                    <div class="text-muted"><%=HtmlEncode(tituloCamion)%></div>
                </div>
                <div>
                    <% If hayRSCamionSel Then %>
                        <%=BadgeTemperatura(rsCamionSel("TemperaturaControl"))%>
                    <% Else %>
                        <span class="badge bg-secondary">Sin datos</span>
                    <% End If %>
                </div>
            </div>

            <% If hayRSCamionSel Then %>
            <div class="data-grid">
                <div class="data-item">
                    <div class="data-label">Nro Cuenta</div>
                    <div class="data-value"><%=HtmlEncode(rsCamionSel("NroCuenta"))%></div>
                </div>
                <div class="data-item">
                    <div class="data-label">Fecha carga</div>
                    <div class="data-value"><%=HtmlEncode(rsCamionSel("FechaCarga"))%></div>
                </div>
                <div class="data-item">
                    <div class="data-label">Chofer</div>
                    <div class="data-value"><%=HtmlEncode(rsCamionSel("Chofer"))%></div>
                </div>
                <div class="data-item">
                    <div class="data-label">Pallets</div>
                    <div class="data-value"><%=HtmlEncode(rsCamionSel("CantidadPallets"))%></div>
                </div>
            </div>
            <% End If %>
        </div>

        <% If Not camionTieneTemperatura Then %>

            <div class="page-card compact-card mb-3">
                <div class="temperatura-box">
                    <h4 class="mb-1"><i class="fa-solid fa-temperature-half me-2"></i>Registrar temperatura</h4>
                    <p class="text-muted mb-2">Registrar la temperatura tomada del camión.</p>

                    <form method="post" action="camiones_facturas.asp">
                        <input type="hidden" name="action" value="guardar_temperatura">
                        <input type="hidden" name="CamionID" value="<%=camionIDSel%>">

                        <div class="row g-2 align-items-end">
                            <div class="col-12 col-md-4">
                                <label class="form-label">Temperatura tomada (&deg;C)</label>
                                <input type="number" step="0.01" name="TemperaturaControl" class="form-control" required>
                            </div>
                            <div class="col-12 col-md-4">
                                <button type="submit" class="btn btn-success w-100">
                                    <i class="fa-solid fa-floppy-disk me-1"></i> Guardar y continuar
                                </button>
                            </div>
                        </div>
                    </form>
                </div>
            </div>

        <% Else %>

            <div class="page-card compact-card mb-3">
                <div class="busqueda-articulo-box">
                    <h5 class="mb-2"><i class="fa-solid fa-magnifying-glass me-2"></i>Buscar en ítems pendientes</h5>

                    <form method="get" action="camiones_facturas.asp" class="row g-2 align-items-end">
                        <input type="hidden" name="CamionID" value="<%=camionIDSel%>">

                        <div class="col-12 col-md-6">
                            <label class="form-label">Código o nombre del artículo</label>
                            <input type="text" name="BuscarArticulo" class="form-control" value="<%=HtmlEncode(buscarArticulo)%>" placeholder="Ej: 140203 o BAGLEY">
                        </div>

                        <div class="col-12 col-md-auto">
                            <button type="submit" class="btn btn-primary w-100">
                                <i class="fa-solid fa-magnifying-glass me-1"></i> Buscar
                            </button>
                        </div>

                        <div class="col-12 col-md-auto">
                            <a href="camiones_facturas.asp?CamionID=<%=camionIDSel%>" class="btn btn-outline-secondary w-100">
                                <i class="fa-solid fa-eraser me-1"></i> Limpiar
                            </a>
                        </div>
                    </form>

                    <% If usandoBusquedaArticulo Then %>
                        <div class="alert alert-info mt-2 mb-0 py-2">
                            <%=HtmlEncode(textoBusquedaInfo)%>
                        </div>
                    <% End If %>
                </div>
            </div>

            <div class="page-card p-3 p-md-4">
                <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
                    <div>
                        <h4 class="mb-1">Ítems pendientes del camión</h4>
                        <div class="text-muted">
                            <% If usandoBusquedaArticulo Then %>
                                Resultado filtrado por código o descripción
                            <% Else %>
                                Todos los artículos pendientes de todas las facturas
                            <% End If %>
                        </div>
                    </div>

                    <div>
                        <span class="badge bg-primary fs-6"><%=totalItemsPendientes%> pendientes</span>
                    </div>
                </div>

                <% If Not hayRSItems Then %>
                    <div class="text-center text-muted py-4">No se pudieron consultar los ítems pendientes del camión.</div>
                <% ElseIf rsItems.EOF Then %>
                    <div class="text-center text-muted py-4">
                        <% If usandoBusquedaArticulo Then %>
                            No se encontraron ítems pendientes que coincidan con la búsqueda.
                        <% Else %>
                            No se encontraron ítems pendientes para este camión.
                        <% End If %>
                    </div>
                <% Else %>

                    <div class="row g-3">
                        <% Do Until rsItems.EOF %>
                            <div class="col-12 col-xl-6">
                                <div class="item-card p-3">
                                    <div class="item-header mb-2">
                                        <div>
                                            <div class="item-code">
                                                <i class="fa-solid fa-barcode me-1"></i>
                                                <%=HtmlEncode(rsItems("Articulo"))%>
                                            </div>
                                            <div class="item-desc">
                                                <%=HtmlEncode(rsItems("Descripcion"))%>
                                            </div>
                                        </div>

                                        <div class="text-end">
                                            <div class="data-label">Cantidad sistema</div>
                                            <div class="item-cantidad"><%=HtmlEncode(rsItems("Cantidad"))%></div>
                                        </div>
                                    </div>

                                    <div class="item-divider"></div>

                                    <div class="data-grid">
                                        <div class="data-item">
                                            <div class="data-label">Factura</div>
                                            <div class="data-value"><%=HtmlEncode(rsItems("NumeroFactura"))%></div>
                                        </div>

                                        <div class="data-item">
                                            <div class="data-label">Proveedor</div>
                                            <div class="data-value"><%=HtmlEncode(rsItems("ProveedorNombre"))%></div>
                                        </div>

                                        <div class="data-item">
                                            <div class="data-label">Fecha factura</div>
                                            <div class="data-value"><%=FormatDateAR(rsItems("FechaFactura"))%></div>
                                        </div>

                                        <div class="data-item">
                                            <div class="data-label">Remito</div>
                                            <div class="data-value"><%=HtmlEncode(rsItems("NroRemito"))%></div>
                                        </div>

                                        <div class="data-item">
                                            <div class="data-label">Orden compra</div>
                                            <div class="data-value"><%=HtmlEncode(rsItems("OrdenCompra"))%></div>
                                        </div>

                                        <div class="data-item">
                                            <div class="data-label">Unidad</div>
                                            <div class="data-value"><%=HtmlEncode(rsItems("UnidadMedida"))%></div>
                                        </div>
                                    </div>

                                    <div class="item-divider"></div>

                                    <form method="post" action="camiones_facturas.asp" class="control-item-box" onsubmit="return validarControlItem(this);">
                                        <input type="hidden" name="action" value="guardar_item">
                                        <input type="hidden" name="CamionID" value="<%=camionIDSel%>">
                                        <input type="hidden" name="BuscarArticulo" value="<%=HtmlEncode(buscarArticulo)%>">
                                        <input type="hidden" name="ItemID" value="<%=HtmlEncode(rsItems("ItemID"))%>">

                                        <div class="control-item-title">
                                            <i class="fa-solid fa-clipboard-check me-1"></i> Control del artículo
                                        </div>

                                        <div class="row g-2 align-items-end">
                                            <div class="col-12 col-md-4">
                                                <label class="form-label">Cantidad controlada</label>
                                                <input type="text"
                                                       name="CantidadControlada"
                                                       class="form-control"
                                                       value="<%=HtmlEncode(GetFieldValue(rsItems, "CantidadControlada", ""))%>"
                                                       placeholder="Ej: 10">
                                            </div>

                                            <div class="col-12 col-md-3">
                                                <div class="form-check mt-2">
                                                    <input class="form-check-input"
                                                           type="checkbox"
                                                           name="Completo"
                                                           value="1"
                                                           id="Completo_<%=HtmlEncode(rsItems("ItemID"))%>"
                                                           <% If IsCheckedValue(GetFieldValue(rsItems, "Completo", 0)) Then Response.Write("checked") End If %>>
                                                    <label class="form-check-label" for="Completo_<%=HtmlEncode(rsItems("ItemID"))%>">
                                                        Completo
                                                    </label>
                                                </div>
                                            </div>

                                            <div class="col-12 col-md-5">
                                                <label class="form-label">Observación</label>
                                                <textarea name="Observacion"
                                                          class="form-control"
                                                          rows="1"
                                                          placeholder="Opcional"><%=HtmlEncode(GetFieldValue(rsItems, "Observacion", ""))%></textarea>
                                            </div>

                                            <div class="col-12">
                                                <button type="submit" class="btn btn-success w-100">
                                                    <i class="fa-solid fa-floppy-disk me-1"></i> Guardar control
                                                </button>
                                            </div>
                                        </div>
                                    </form>
                                </div>
                            </div>
                        <%
                            rsItems.MoveNext
                        Loop
                        %>
                    </div>

                <% End If %>
            </div>

        <% End If %>

    <% End If %>

</div>
</div>

</body>
</html>
<%
If IsObject(rsItems) Then
    On Error Resume Next
    If rsItems.State = 1 Then rsItems.Close
    Set rsItems = Nothing
    On Error GoTo 0
End If

If IsObject(rsCamionSel) Then
    On Error Resume Next
    If rsCamionSel.State = 1 Then rsCamionSel.Close
    Set rsCamionSel = Nothing
    On Error GoTo 0
End If

If IsObject(rsCamiones) Then
    On Error Resume Next
    If rsCamiones.State = 1 Then rsCamiones.Close
    Set rsCamiones = Nothing
    On Error GoTo 0
End If

conn.Close
Set conn = Nothing
%>