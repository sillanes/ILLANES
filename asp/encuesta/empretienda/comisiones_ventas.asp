<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
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
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")
%><!--#include virtual="./includes/sql-check.asp"--><%

For Each s in Request.Form 
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		Response.Redirect(ErrorPage)
	End If
Next

Dim orderBy, orderDir, allowedColumns
orderBy = Request("order")
orderDir = UCase(Request("dir"))
allowedColumns = "Cliente,VendedorID,Zona,OrdenID,Monto,Pagado"

If orderBy = "" Or InStr(1, allowedColumns, orderBy, vbTextCompare) = 0 Then orderBy = "Cliente"
If orderDir <> "DESC" Then orderDir = "ASC"

Function parseNumber(v)
    Dim s
    If IsNull(v) Or Trim(v) = "" Then parseNumber = 0 : Exit Function
    s = Replace(Replace(Replace(Trim(CStr(v)),"$","")," ",""),Chr(160),"")
    If IsNumeric(s) Then parseNumber = CDbl(s) Else parseNumber = 0
End Function

Function IIf(expr, truePart, falsePart)
    If expr Then IIf = truePart Else IIf = falsePart
End Function

Dim FechaInicio, FechaFin
If Request("inicio") <> "" Then
    FechaInicio = CDate(Request("inicio"))
Else
    'FechaInicio = DateAdd("m",-1,Date())
	FechaIncio = "NULL"
End If

If Request("fin") <> "" Then
    FechaFin = CDate(Request("fin"))
Else
    FechaFin = Date()
End If

mensaje = Request("mensaje")

Function toISO(fecha)
    Dim d
    d = CDate(fecha)
    toISO = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function

Dim rs, sql
sql = "EXEC report.[usp_Ordenes_Comisiones_Detalle_sel] '" & toISO(FechaInicio) & "','" & toISO(FechaFin) & "','" & orderBy & "','" & orderDir & "'"
'response.write sql
Set rs = conn.Execute(sql)

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
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Comisiones de Ventas</title>
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
h1 { text-align:center; margin-bottom:20px; color:#444; }
th, td { padding:8px 12px; border:1px solid #ddd; text-align:center; white-space:nowrap; }
th { background:#007bff; color:#fff; }
td:first-child, th:first-child { text-align:left; }
.icon-btn { text-decoration:none; margin:0 3px; font-size:16px; color:#007bff; }
.icon-btn:hover { opacity:0.8; }
.header { display:flex; justify-content:space-between; align-items:center; padding:15px 20px; background:#f8f9fa; border-bottom:1px solid #ddd; }
.header h2 { margin:0; font-size:20px; font-weight:bold; flex:1; text-align:center; }
.btn-volver { display:inline-flex; align-items:center; gap:6px; padding:8px 16px; font-size:14px; color:#fff; background:linear-gradient(135deg, #007bff, #0056b3); border-radius:25px; text-decoration:none; font-weight:500; box-shadow:0 2px 6px rgba(0,0,0,0.15); transition:all 0.3s ease; }
.btn-volver:hover { background:linear-gradient(135deg, #0056b3, #003f88); transform:translateY(-2px); box-shadow:0 4px 10px rgba(0,0,0,0.2); }
.estado-verde { background-color:#d4edda; }
.estado-amarillo { background-color:#fff3cd; }
.estado-azul { background-color:#cce5ff; }
th a { color:#fff; text-decoration:none; }
th a:hover { text-decoration:underline; }
.table-container table tr:hover { background-color:#f0f8ff; cursor:pointer; }

.estado-verde { background-color:#d4edda; }   /* verde claro */
.estado-amarillo { background-color:#fff3cd; }/* amarillo claro */
.estado-azul { background-color:#cce5ff; }    /* azul claro */

</style>
</head>
<body>

<!--#include file="header.asp" -->

<div class="main-content">
	<% If mensaje <> "" Then %>
        <div style="background:#d4edda; color:#155724; padding:10px; margin-bottom:15px; border:1px solid #c3e6cb; border-radius:5px;">
            ✅ <%=(mensaje)%>.
        </div>
    <% End If %>
	
<div class="header">
    <h2>Comisiones de Ventas</h2>
 
</div>

<div class="table-container">
<table>
<tr>
    <th></th>
    <th>Acción</th>
    <th>
		<a href="?order=OrdenID&dir=<%=IIf(orderBy="OrdenID" And orderDir="ASC","DESC","ASC")%>">
		OrdenID <% If orderBy="OrdenID" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
    <th>
		<a href="?order=Cliente&dir=<%=IIf(orderBy="Cliente" And orderDir="ASC","DESC","ASC")%>">
			Cliente <% If orderBy="Cliente" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
    <th>		
		<a href="?inicio=<%=toISO(FechaInicio)%>&fin=<%=toISO(FechaFin)%>&order=Vendedor&dir=<%=IIf(orderBy="Vendedor" And orderDir="ASC","DESC","ASC")%>">
			Vendedor <% If orderBy="Vendedor" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
    <th>
		<a href="?inicio=<%=toISO(FechaInicio)%>&fin=<%=toISO(FechaFin)%>&order=Zona&dir=<%=IIf(orderBy="Zona" And orderDir="ASC","DESC","ASC")%>">
			Zona <% If orderBy="Zona" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
    <th>
		<a href="?order=Monto&dir=<%=IIf(orderBy="Monto" And orderDir="ASC","DESC","ASC")%>">
			Importe <% If orderBy="Monto" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
    <th>Cajas Oro</th>
    <th>Cajas Plata</th>
    <th>Cajas Platino</th>
    <th>Cajas Especial</th>
    <th>Cajas Bronce</th>
    <th>Total Cajas</th>
    <th>		
		<a href="?order=Pagado&dir=<%=IIf(orderBy="Pagado" And orderDir="ASC","DESC","ASC")%>">
			Pagado <% If orderBy="Pagado" Then Response.Write(IIf(orderDir="ASC","&uarr;","&darr;")) End If %>
		</a>
	</th>
    <th>Fecha Entrega</th>
</tr>

<%
Do While Not rs.EOF
    Dim OrdenID, FileID, Cliente, Vendedor, Zona, Monto, Oro, Plata, Platino, Especial, Bronce, TotalCajas, Pagado, FechaEntrega
    OrdenID = rs("OrdenID")
    FileID = rs("FileID")
    Cliente = rs("Cliente")
    Vendedor = rs("Vendedor")
    Zona = rs("Zona")
    Monto = parseNumber(rs("Monto"))
    Oro = parseNumber(rs("CajasOro"))
    Plata = parseNumber(rs("CajasPlata"))
    Platino = parseNumber(rs("CajasPlatino"))
    Especial = parseNumber(rs("CajasEspecial"))
    Bronce = parseNumber(rs("CajasBronce"))
    TotalCajas = parseNumber(rs("TotalCajas"))
    Pagado = rs("Pagado")
	Estado = rs("Estado")


	
	itemFecha = (rs("FechaEntrega"))

	If IsDate(itemFecha) Then
		fechaFormateada = Right("0" & Day(itemFecha),2) & "-" & Right("0" & Month(itemFecha),2) & "-" & Year(itemFecha)
	Else
		fechaFormateada = ""
	End If

	FechaEntrega = fechaFormateada
	
    totimporte = totimporte + Monto
    totalOro = totalOro + Oro
    totalPlata = totalPlata + Plata
    totalPlatino = totalPlatino + Platino
    totalEspecial = totalEspecial + Especial
    totalBronce = totalBronce + Bronce
    tot = tot + TotalCajas
	
    Dim rowClass
    Select Case Estado
        Case 1: rowClass = "estado-verde"
        Case 2: rowClass = "estado-amarillo"
        Case 3: rowClass = "estado-azul"
        Case Else: rowClass = ""
    End Select
	
%>
<tr class="<%=rowClass%>">
    <td><%=RowID%></td>
    <td>
        <a href="comisiones_ventas_modificar.asp?FileID=<%=FileID%>&OrdenID=<%=OrdenID%>&inicio=<%=toISO(FechaInicio)%>&fin=<%=toISO(FechaFin)%>" 
           class="icon-btn" title="Editar comisión"><i class="fas fa-edit"></i></a>
    </td>
    <td><%=OrdenID%></td>
    <td><%=Cliente%></td>
    <td><%=Vendedor%></td>
    <td><%=Zona%></td>
    <td>$ <%=FormatNumber(Monto,2)%></td>
    <td><%=Oro%></td>
    <td><%=Plata%></td>
    <td><%=Platino%></td>
    <td><%=Especial%></td>
    <td><%=Bronce%></td>
    <td><%=TotalCajas%></td>
    <td><%=Pagado%></td>
    <td><%=Server.HTMLEncode(CStr(FechaEntrega))%></td>
</tr>
<%
    RowID = RowID + 1
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>

<tr style="font-weight:bold; background:#f1f1f1;">
    <td colspan="6">Total</td>
    <td>$ <%=FormatNumber(totimporte,2)%></td>
    <td><%=totalOro%></td>
    <td><%=totalPlata%></td>
    <td><%=totalPlatino%></td>
    <td><%=totalEspecial%></td>
    <td><%=totalBronce%></td>
    <td><%=tot%></td>
    <td colspan="2"></td>
</tr>
</table>
</div>
</div>
</body>
</html>
