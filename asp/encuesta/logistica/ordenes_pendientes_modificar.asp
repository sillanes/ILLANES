<%@ Language="VBScript" %>
<!--#include file="../empretienda/conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim BlackList, ErrorPage
BlackList = Array("/*", "*/", "@@",_
                  "cursor","exec","execute",_
				  "CREATE","truncate","delete",_
                  "nchar", "varchar", "nvarchar", "iframe"_
                  )
'Note: We can include following keyword to make a stronger scan but it will also 
'protect users to input these words even those are valid input
'  "!", "char", "alter", "begin", "cast", "create",  
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")
 

%><!--#include virtual="./includes/sql-check.asp"--><%

For Each s in Request.Form 
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		PrepareReport("Post Varibale")
		' Redirect to an error page
		Response.Redirect(ErrorPage)
	End If
Next


Dim OrdenID, FileID
FileID = Request("FileID")
OrdenID = Request("OrdenID")

' ==========================================
' Tomar parámetros de fechas (opcional)
' ==========================================
Dim FechaInicio, FechaFin
If Request("FechaInicio") <> "" Then
    FechaInicio = CDate(Request("FechaInicio"))
Else
    FechaInicio = DateAdd("m",-1,Date()) ' default 1 mes atrás
End If

If Request("FechaFin") <> "" Then
    FechaFin = CDate(Request("FechaFin"))
Else
    FechaFin = Date()
End If

mensaje = Request("mensaje")
Function toISO(fecha)
    Dim d
    d = CDate(fecha)
    toISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function


' Traer datos de la orden
Dim rs, sql
sql = "EXEC report.usp_Ordenes_Ventas_Detalle_sel_ID " & FileID & "," & OrdenID
Set rs = conn.Execute(sql)

If rs.EOF Then
    Response.Write "<p>No se encontró la orden.</p>"
    Response.End
End If

' Traer detalle de cajas
Dim rsCajas, sqlCajas
sqlCajas = "EXEC report.usp_Ordenes_Ventas_Cajas_sel " & FileID & "," & OrdenID
'response.write sqlCajas
Set rsCajas = conn.Execute(sqlCajas)
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Modificar Orden</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<link rel="stylesheet" href="estilos.css"> 
<link rel="stylesheet" href="ventas.css">
<style>
    .form-container { max-width:800px; margin:0 auto; background:#fff; padding:20px; border-radius:8px; box-shadow:0 2px 5px rgba(0,0,0,0.1); }
    .form-group { margin-bottom:15px; }
    .form-group label { display:flex; align-items:center; font-weight:600; margin-bottom:6px; }
    .form-group label i { margin-left:6px; font-size:14px; color:#007bff; }
    .form-group .readonly-text { padding:8px; background:#f9f9f9; border:1px solid #ddd; border-radius:6px; }
    .form-actions { display:flex; gap:10px; margin-top:20px; }
    .btn { display:inline-flex; align-items:center; gap:6px; padding:10px 18px; border-radius:6px; text-decoration:none; font-weight:500; cursor:pointer; }
    .btn-primary { background:#007bff; color:#fff; border:none; }
    .btn-primary:hover { background:#0056b3; }
    .btn-secondary { background:#6c757d; color:#fff; text-decoration:none; }
    .btn-secondary:hover { background:#545b62; }

    /* Header con botón volver */
    .header { display:flex; justify-content:space-between; align-items:center; padding:15px 20px; background:#f8f9fa; border-bottom:1px solid #ddd; }
    .header h2 { margin:0; font-size:20px; font-weight:bold; text-align:center; flex:1; }
    .btn-volver { display:inline-flex; align-items:center; gap:6px; padding:8px 16px; font-size:14px; color:#fff; background:linear-gradient(135deg, #007bff, #0056b3); border-radius:25px; text-decoration:none; font-weight:500; box-shadow:0 2px 6px rgba(0,0,0,0.15); transition:all 0.3s ease; }
    .btn-volver:hover { background:linear-gradient(135deg, #0056b3, #003f88); transform:translateY(-2px); box-shadow:0 4px 10px rgba(0,0,0,0.2); }

    /* Tabla de cajas */
    .table-container { margin-top:20px; overflow-x:auto; }
    table { width:100%; border-collapse:collapse; background:#fff; box-shadow:0 2px 5px rgba(0,0,0,0.1); }
    th, td { padding:8px 12px; border:1px solid #ddd; text-align:center; }
    th { background:#007bff; color:#fff; }
    td:first-child { text-align:left; }
	
    .textarea-container {
        width: 100%; /* Ensure the container takes full width */
        /* Add any desired padding or margin to the container */
        padding: 0px;
    }

    #myTextArea {
        width: 100%; /* Make the textarea fill its parent's width */
        box-sizing: border-box; /* Include padding and border in the 100% width calculation */
        /* Add any desired padding, border, or margin to the textarea itself */
        padding: 5px;
        border: 1px solid #ccc;
    }
	
</style>
</head>
<body>


<!--#include file="header.asp" -->
<div class="main-content">
	<% If mensaje <> "" Then %>
        <div style="background:#d4edda; color:#155724; padding:10px; margin-bottom:15px; border:1px solid #c3e6cb; border-radius:5px;">
            ✅ <%=mensaje%>.
        </div>
    <% End If %>
	
    <div class="header">
        <h2>Modificar Orden</h2>
        <a href="ordenes_pendientes_detalle.asp?FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>" class="btn-volver"><i class="fas fa-arrow-left"></i> Volver</a>
    </div>

    <div class="form-container">
        <form method="post" action="ordenes_pendientes_modificar_guardar.asp">
            <input type="hidden" name="FileID" value="<%=FileID%>">
            <input type="hidden" name="OrdenID" value="<%=OrdenID%>">
            <input type="hidden" name="FechaInicio" value="<%=toISO(FechaInicio)%>">
            <input type="hidden" name="FechaFin" value="<%=toISO(FechaFin)%>">

			<!-- Traer lista de vendedores -->
			<%
			Dim rsVendedores, sqlVendedores
			sqlVendedores = "EXEC report.Vendedores_Combo"
			Set rsVendedores = conn.Execute(sqlVendedores)
			%>

			<div class="form-group">
				<label>Vendedor <i class="fas fa-check" title="Modificable"></i></label>
				<select name="VendedorID" style="width:100%; padding:8px; border-radius:6px; border:1px solid #ddd;">
					<% Do While Not rsVendedores.EOF %>
						<option value="<%=rsVendedores("VendedorID")%>" <% If rsVendedores("VendedorID") = rs("VendedorID") Then Response.Write "selected" %>>
							<%=rsVendedores("Nombre")%>
						</option>
						<% rsVendedores.MoveNext
					Loop %>
				</select>
			</div>

			<%
			rsVendedores.Close
			Set rsVendedores = Nothing
			%>

            <div class="form-group">
                <label>OrdenID <i class="fas fa-lock" title="No modificable"></i></label>
                <div class="readonly-text"><%=rs("OrdenID")%></div>
            </div>
			
            <div class="form-group">
                <label>Cliente <i class="fas fa-lock" title="No modificable"></i></label>
                <div class="readonly-text"><%=rs("Cliente")%></div>
            </div>
			


            <div class="form-group">
                <label>Email <i class="fas fa-lock" title="No modificable"></i></label>
                <div class="readonly-text"><%=rs("email")%></div>
            </div>

            <div class="form-group">
                <label>Teléfono <i class="fas fa-lock" title="No modificable"></i></label>
                <div class="readonly-text"><%=rs("Telefono")%></div>
            </div>

            <div class="form-group">
                <label>Fecha Entrega <i class="fas fa-check" title="Modificable"></i></label>
                <input type="date" name="FechaEntrega" value="<%=Year(rs("FechaEntrega"))%>-<%=Right("0"&Month(rs("FechaEntrega")),2)%>-<%=Right("0"&Day(rs("FechaEntrega")),2)%>">
            </div>

            <div class="form-group">
                <label>Observaciones <i class="fas fa-lock" title="No modificable"></i></label>
                <div class="readonly-text"><%=rs("Observaciones")%></div>
            </div>
            
            <div class="form-group">
                <label>Observaciones Empretienda <i class="fas fa-lock" title="No modificable"></i></label>
							
				<div class="textarea-container">
					<textarea class="readonly-text" id="myTextArea" name="ObservacionesEmpretienda" rows="4" readonly><%=rs("ObservacionesEmpretienda")%></textarea>
				</div>
	 
            </div>
            <div class="form-group">
                <label>Observaciones Administracion <i class="fas fa-lock" title="No modificable"></i></label>
							
				<div class="textarea-container">
					<textarea class="readonly-text" id="myTextArea" name="ObservacionesAdicional" rows="4" readonly><%=rs("ObservacionADM")%></textarea>
				</div>
	 
            </div>
            <div class="form-group">
                <label>Observaciones Logistica <i class="fas fa-check" title="Modificable"></i></label>
							
				<div class="textarea-container">
					<textarea id="myTextArea" name="ObservacionesLogistica" rows="4"><%=rs("ObservacionLogistica")%></textarea>
				</div>
	 
            </div>

            <div class="form-group">
                <label>Calle <i class="fas fa-check" title="Modificable"></i></label>
                <input type="text" name="Envio_Calle" value="<%=rs("Envio_Calle")%>" 
                       style="width:100%; padding:8px; border-radius:6px; border:1px solid #ddd;">
            </div>

            <div class="form-group">
                <label>Número <i class="fas fa-check" title="Modificable"></i></label>
                <input type="text" name="Envio_Numero" value="<%=rs("Envio_Numero")%>" 
                       style="width:100%; padding:8px; border-radius:6px; border:1px solid #ddd;">
            </div>

            <div class="form-group">
                <label>Piso <i class="fas fa-check" title="Modificable"></i></label>
                <input type="text" name="Envio_Piso" value="<%=rs("Envio_Piso")%>" 
                       style="width:100%; padding:8px; border-radius:6px; border:1px solid #ddd;">
            </div>

            <div class="form-group">
                <label>Dpto <i class="fas fa-check" title="Modificable"></i></label>
                <input type="text" name="Envio_Dpto" value="<%=rs("Envio_Dpto")%>" 
                       style="width:100%; padding:8px; border-radius:6px; border:1px solid #ddd;">
            </div>

            <div class="form-group">
                <label>Ciudad <i class="fas fa-check" title="Modificable"></i></label>
                <input type="text" name="Envio_Ciudad" value="<%=rs("Envio_Ciudad")%>" 
                       style="width:100%; padding:8px; border-radius:6px; border:1px solid #ddd;">
            </div>

            <!-- Tabla de Cajas -->
            <div class="table-container">
                <h3>Detalle de Cajas</h3>
                <table>
                    <tr>
                        <th>Tipo de Caja</th>
                        <th>Cantidad</th>
                        <th>Modificado</th>
                    </tr>
                    <% Do While Not rsCajas.EOF %>
                        <tr>
                            <td><%=rsCajas("TipoCaja")%></td>
                            <td><%=rsCajas("Cantidad")%></td>
							<td>
							<% If rsCajas("Modificado") = 1 Then %>
								<span style="color:green; font-weight:bold;">Sí</span>
							<% Else %>
								<span style="color:red; font-weight:bold;">No</span>
							<% End If %>
							</td>
                        </tr>
                        <% rsCajas.MoveNext
                    Loop %>
                </table>
            </div>

			<!-- Marcar como entregada -->
			<div class="form-group" style="margin-top:25px; border-top:1px solid #ddd; padding-top:15px;">
				<label>¿Orden Entregada? <i class="fas fa-check" title="Modificable"></i></label>
				<div style="display:flex; align-items:center; gap:20px; padding:8px 0;">
					<label style="display:flex; align-items:center; gap:6px;">
						<input type="radio" name="OrdenEntregada" value="1" 
							<% If Not rs("Entregado") = 0 Then Response.Write "checked" %> 
							onclick="toggleFechaEntrega(true)"> Sí
					</label>
					<label style="display:flex; align-items:center; gap:6px;">
						<input type="radio" name="OrdenEntregada" value="0" 
							<% If rs("Entregado") = 0 Then Response.Write "checked" %> 
							onclick="toggleFechaEntrega(false)"> No
					</label>
				</div>

				<div id="fechaEntregadaContainer" style="margin-top:10px; display:none;">
					<label>Fecha de Entrega Real <i class="fas fa-calendar-day" title="Seleccionar fecha de entrega real"></i></label>
					<input type="date" name="FechaEntregada" id="FechaEntregada" 
						   value="<% If Not IsNull(rs("FechaEntregado")) Then Response.Write(Year(rs("FechaEntregado")) & "-" & Right("0"&Month(rs("FechaEntregado")),2) & "-" & Right("0"&Day(rs("FechaEntregado")),2)) %>"
						   style="padding:8px; border-radius:6px; border:1px solid #ddd; width:100%;">
				</div>
			</div>

	 
        </form>
    </div>
</div>

<script> 
function toggleSidebar(){
    document.querySelector('.sidebar').classList.toggle('open');
}
function toggleFechaEntrega(mostrar) {
    const cont = document.getElementById('fechaEntregadaContainer');
    if (mostrar) {
        cont.style.display = 'block';
    } else {
        cont.style.display = 'none';
        document.getElementById('FechaEntregada').value = '';
    }
}

// Mostrar automáticamente el campo si ya estaba marcado como entregado
window.addEventListener('DOMContentLoaded', function() {
    <% If Not rs("Entregado") = 0 Then %>
        document.getElementById('fechaEntregadaContainer').style.display = 'block';
    <% End If %>
});
</script>

</body>
</html>
<%
rs.Close
Set rs = Nothing
rsCajas.Close
Set rsCajas = Nothing
%>
