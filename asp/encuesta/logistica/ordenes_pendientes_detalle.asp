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

' ==========================================
' Ordenamiento
' ==========================================
Dim orderBy, orderDir, allowedColumns
orderBy = Request("order")
orderDir = UCase(Request("dir"))

' Columnas permitidas
allowedColumns = "Cliente,VendedorID,Zona,OrdenID,Monto,Pagado,FechaEntrega,Envio_Calle"

If orderBy = "" Or InStr(1, allowedColumns, orderBy, vbTextCompare) = 0 Then
    orderBy = "Cliente" ' default
End If

If orderDir <> "DESC" Then
    orderDir = "ASC"
End If

' ==========================================
' Función para normalizar números
' ==========================================
Function parseNumber(v)
    Dim s, i, dotCount, posLastDot, posLastComma
    If IsNull(v) Then parseNumber = 0 : Exit Function

    ' si es numérico y no contiene coma -> usar directo
    If IsNumeric(v) And InStr(CStr(v), ",") = 0 Then
        On Error Resume Next
        parseNumber = CDbl(v)
        If Err.Number <> 0 Then Err.Clear : parseNumber = 0
        On Error GoTo 0
        Exit Function
    End If

    ' tratar como string
    s = Trim(CStr(v))
    If s = "" Then parseNumber = 0 : Exit Function

    ' limpiar símbolos comunes
    s = Replace(s, "$", "")
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")

    ' contar puntos
    dotCount = 0
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "." Then dotCount = dotCount + 1
    Next

    posLastDot = InStrRev(s, ".")
    posLastComma = InStrRev(s, ",")

    If InStr(s, ",") > 0 Then
        If InStr(s, ".") > 0 Then
            If posLastComma > posLastDot Then
                ' formato 1.234.567,89
                s = Replace(s, ".", "")
                s = Replace(s, ",", ".")
            Else
                ' formato 1,234,567.89
                s = Replace(s, ",", "")
            End If
        Else
            ' solo coma -> coma decimal
            s = Replace(s, ".", ",")
        End If
    Else
        ' no hay coma
        If dotCount > 1 Then
            s = Replace(s, ",", ".")
        End If
    End If

    On Error Resume Next
    parseNumber = CDbl(s)
    If Err.Number <> 0 Then Err.Clear : parseNumber = 0
    On Error GoTo 0
End Function

Function IIf(expr, truePart, falsePart)
    If expr Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function
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
 
' ==========================================
' Ejecutar SP principal
' ==========================================
Dim rs, sql
'sql = "EXEC report.usp_Ordenes_Ventas_Detalle_sel '" & toISO(FechaInicio) & "','" & toISO(FechaFin) & "'"
sql = "EXEC report.usp_Ordenes_Ventas_Detalle_sel '" & toISO(FechaInicio) & "','" & toISO(FechaFin) & "','" & orderBy & "','" & orderDir & "'" & ",'LogisticaPendientes'"
'response.write sql
Set rs = conn.Execute(sql)

' ==========================================
' Acumuladores para totales
' ==========================================
Dim totalBronce, totalPlata, totalOro, totalPlatino, totalEspecial,tot,RowID
totalBronce = 0
totalPlata = 0
totalOro = 0
totalPlatino = 0
totalEspecial = 0
tot=0
RowID =1
%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Detalle de Pedidos</title>
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
    h1 { text-align:center; margin-bottom:20px; color:#444; }
    th, td { padding:8px 12px; border:1px solid #ddd; text-align:center; white-space:nowrap; }
    th { background:#007bff; color:#fff; }
    td:first-child, th:first-child { text-align:left; }
	 
    .telefono { max-width:100px; }
    .observaciones { max-width:150px; }
    .icon-btn { text-decoration:none; margin:0 3px; font-size:16px; color:#007bff; }
    .icon-btn.delete { color:#c0392b; }
    .icon-btn:hover { opacity:0.8; }


	/* Ajuste de columna Cliente */
	.cliente {
		max-width: 200px;        /* ancho más pequeño */
		white-space: nowrap;
		overflow: hidden;
		text-overflow: ellipsis;
	}
	.cliente:hover {
		white-space: normal;     /* permite ver todo el contenido al pasar el mouse */
		background: #fff;
		position: absolute;
		z-index: 100;
		box-shadow: 0 2px 5px rgba(0,0,0,0.2);
		padding: 2px 6px;
		border-radius: 4px;
	}
.header {
    display: flex;
    justify-content: center;       /* centramos el título */
    align-items: center;
    padding: 15px 20px;
    background: #f8f9fa;
    border-bottom: 1px solid #ddd;
    position: relative;            /* para posicionar el botón */
}

.header {
    display: flex;
    justify-content: space-between; /* título a la izquierda, botón a la derecha */
    align-items: center;
    padding: 15px 20px;
    background: #f8f9fa;
    border-bottom: 1px solid #ddd;
}

.header h2 {
    margin: 0;
    font-size: 20px;
    font-weight: bold;
    text-align: center;
    flex: 1;
}

.btn-volver {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 8px 16px;
    font-size: 14px;
    color: #fff;
    background: linear-gradient(135deg, #007bff, #0056b3);
    border-radius: 25px; /* botón redondeado */
    text-decoration: none;
    font-weight: 500;
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    transition: all 0.3s ease;
}

.btn-volver:hover {
    background: linear-gradient(135deg, #0056b3, #003f88);
    transform: translateY(-2px);
    box-shadow: 0 4px 10px rgba(0,0,0,0.2);
}

.estado-rojo { background-color:#ff0000; }   /* verde claro */
.estado-verde { background-color:#d4edda; }   /* verde claro */
.estado-amarillo { background-color:#fff3cd; }/* amarillo claro */
.estado-azul { background-color:#cce5ff; }    /* azul claro */

    /* Modal */
    #obsModal { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); justify-content:center; align-items:center; z-index:1000; }
    #obsModalContent { background:#fff; padding:20px; max-width:400px; width:90%; border-radius:6px; position:relative; }
    #closeModal { position:absolute; top:10px; right:15px; cursor:pointer; font-size:20px; }

    /* Responsive */
    @media(max-width:768px){
        th, td { font-size:12px; padding:6px; }
         
        .mail { max-width:150px; }
        .observaciones { max-width:100px; }
    }

th a {
    color: #fff;
    text-decoration: none;
}
th a:hover {
    text-decoration: underline;
}

/* Resaltado al pasar el mouse */
.table-container table tr:hover {
    background-color: #f0f8ff; /* azul muy claro */
    cursor: pointer;
}
.table-container table tbody tr:hover {
    background-color: #f0f8ff;
    cursor: pointer;
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
    <h2>Detalle de Pedidos</h2>
    <div style="display:flex; gap:10px;">
        <a href="ordenespendientes.asp?txtFechaInicio=<%=toISO(FechaInicio)%>&txtFechaFin=<%=toISO(FechaFin)%>" class="btn-volver">
            <i class="fas fa-arrow-left"></i> Volver
        </a>

    </div>
</div>

<div class="table-container">
<table>
<tr>
    <th rowspan="2"></th>
    <th rowspan="2">Acción</th>
    <th rowspan="2">
	<a href="?order=FechaEntrega&dir=<%=IIf(orderBy="FechaEntrega" And orderDir="ASC","DESC","ASC")%>&FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>">
		Fecha <br/> Entrega <% If orderBy="FechaEntrega" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
	</a>
	
	</th>
    <th rowspan="2">
	<a href="?order=OrdenID&dir=<%=IIf(orderBy="OrdenID" And orderDir="ASC","DESC","ASC")%>&FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>">
		OrdenID <% If orderBy="OrdenID" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
	</a>	
	</th>
	<th rowspan="2">
		<a href="?order=Cliente&dir=<%=IIf(orderBy="Cliente" And orderDir="ASC","DESC","ASC")%>&FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>">
			Cliente <% If orderBy="Cliente" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>

	</th>

	<th rowspan="2">
		<a href="?FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>&order=Vendedor&dir=<%=IIf(orderBy="Vendedor" And orderDir="ASC","DESC","ASC")%>">
			Vendedor <% If orderBy="Vendedor" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>

	<th rowspan="2">
		<a href="?FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>&order=Zona&dir=<%=IIf(orderBy="Zona" And orderDir="ASC","DESC","ASC")%>">
			Zona <% If orderBy="Zona" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
 
    <th colspan="6">Direccion Entrega</th>
  
</tr>
<tr>
    
    <th>Recibe</th>
    <th>
		<a href="?FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>&order=Envio_Calle&dir=<%=IIf(orderBy="Envio_Calle" And orderDir="ASC","DESC","ASC")%>">
			Calle <% If orderBy="Envio_Calle" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
    <th>Nro</th>
    <th>Piso</th>
    <th>Dpto</th>
    <th>Ciudad</th>
</tr>

<%
Do While Not rs.EOF
    Dim Zona, Monto, OrdenID,FileID, Cliente, VendedorID, Oro, Plata, Platino, Especial, Bronce, Pagado, FechaEntrega, Telefono, Mail, Obs, ObsTrunc,TotalCajas, Estado, Modificado
    OrdenID = rs("OrdenID")
    FileID = rs("FileID")
    Cliente = rs("Cliente") 							  
    VendedorID = rs("VendedorID") 							  
    Vendedor = rs("Vendedor") 
	Zona =  rs("Zona") 
    Monto = parseNumber(rs("Monto"))
    Oro = parseNumber(rs("CajasOro"))
    Plata = parseNumber(rs("CajasPlata"))
    Platino = parseNumber(rs("CajasPlatino"))
    Especial = parseNumber(rs("CajasEspecial"))
    Bronce = parseNumber(rs("CajasBronce"))
    TotalCajas = parseNumber(rs("TotalCajas"))
    Pagado = rs("Pagado")
    FechaEntrega = rs("FechaEntrega")
    Telefono = rs("Telefono")
    Mail = rs("Mail")
    Obs = rs("Observaciones")
    Estado = rs("Estado")        ' <-- nuevo
    'NewEstado = rs("NewEstado")        ' <-- nuevo
    Modificado = rs("Modificado") ' <-- 
	' Datos entrega
	NombreRecibe = rs("Envio_Nombre")
	Calle = rs("Envio_Calle")
	Nro = rs("Envio_Numero")
	Piso = rs("Envio_Piso")
	Dpto= rs("Envio_Dpto")
	Ciudad= rs("Envio_Ciudad")
	' Fin Datos entrega

    FacturaArchivo = rs("FacturaArchivo") ' ej: 2025-09\B00090031910406FB0000051171.PDF

    ' Extraer solo el nombre del archivo
    If Len(Trim(FacturaArchivo)) > 0 Then
        Dim pos
        pos = InStrRev(FacturaArchivo, "\")
        If pos > 0 Then
            fileName = Mid(FacturaArchivo, pos + 1)
        Else
            fileName = FacturaArchivo
        End If
    Else
        fileName = ""
    End If
	
	itemFecha = (rs("FechaEntrega"))

	If IsDate(itemFecha) Then
		fechaFormateada = Right("0" & Day(itemFecha),2) & "-" & Right("0" & Month(itemFecha),2) & "-" & Year(itemFecha)
	Else
		fechaFormateada = ""
	End If

	FechaEntrega = fechaFormateada
    
    ' Truncar observaciones
    If Len(Obs) > 10 Then
        ObsTrunc = Left(Obs,10) & "..."
    Else
        ObsTrunc = Obs
    End If

    ' Determinar color de fila según Estado
    Dim rowClass
    Select Case Estado  
		' Verde Entregado
        'Case 0: rowClass = "estado-rojo"
        'Case 1: rowClass = "estado-rojo"
        Case 3: rowClass = "estado-verde"
        Case Else: rowClass = ""
    End Select

    totimporte = totimporte + Monto
    totalOro = totalOro + Oro
    totalPlata = totalPlata + Plata
    totalPlatino = totalPlatino + Platino
    totalEspecial = totalEspecial + Especial
    totalBronce = totalBronce + Bronce
    tot = tot + TotalCajas
%>
<tr class="<%=rowClass%>">
	<td>
		<%=RowID%>
		<% If Modificado = 1 Then %>
            <i class="fas fa-flag" style="color:red;" title="Modificado"></i>
        <% End If %>

		<% If UCase(Trim(Pagado)) = "SI" Then %>
			<i class="fas fa-check-circle" style="color:green; margin-left:5px;" title="Pagado"></i> 
		<% End If %>
		
	</td>
    <td>
        <a href="ordenes_pendientes_modificar.asp?FileID=<%=FileID%>&OrdenID=<%=OrdenID%>&FechaInicio=<%=toISO(FechaInicio)%>&FechaFin=<%=toISO(FechaFin)%>" class="icon-btn edit" title="Editar"><i class="fas fa-edit"></i></a>
    
        <% If fileName <> "" Then %>
            <a href="descargar_facturas.asp?file=<%=Server.URLEncode(fileName)%>" class="icon-btn" title="Descargar Factura">
                <i class="fas fa-file-pdf"></i>
            </a>
        <% End If %>
    </td>
	
    <td><%=Server.HTMLEncode(CStr((FechaEntrega)))%></td>
    <td><%=OrdenID%></td>
    <td class="cliente"><%=Cliente%></td>
    <td class="cliente"><%=Vendedor%></td>
    <td><%=Zona%></td>
  

    <td class="cliente"><%=NombreRecibe%></td>
    <td class="cliente"><%=Calle%></td>
    <td><%=Nro%></td>
    <td><%=Piso%></td>
    <td><%=Dpto%></td>
    <td><%=Ciudad%></td>
	 
 
</tr>
<%
	RowID = RowID +1
    rs.MoveNext
Loop 
rs.Close
Set rs = Nothing
%>

<tr style="font-weight:bold; background:#f1f1f1;">
    <td colspan="6">Total</td>  
    <td><%=totalOro%></td>
    <td><%=totalPlata%></td>
    <td><%=totalPlatino%></td>
    <td><%=totalEspecial%></td>
    <td><%=totalBronce%></td>
    <td><%=tot%></td>
    <td colspan="6"></td>
    <td colspan="5"></td>
</tr>
</table>
</div>

<!-- Modal Observaciones -->
<div id="obsModal">
    <div id="obsModalContent">
        <span id="closeModal">&times;</span>
        <h3>Observaciones</h3>
        <p id="modalContent"></p>
    </div>
</div>

<script>
document.querySelectorAll('.viewObs').forEach(function(link){
    link.addEventListener('click', function(e){
        e.preventDefault();
        var obs = this.dataset.obs;
        document.getElementById('modalContent').innerText = obs;
        document.getElementById('obsModal').style.display = 'flex';
    });
});

document.getElementById('closeModal').addEventListener('click', function(){
    document.getElementById('obsModal').style.display = 'none';
});

function toggleSidebar(){
    document.querySelector('.sidebar').classList.toggle('open');
}
</script>

</div>
</body>
</html>
