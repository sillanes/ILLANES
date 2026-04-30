<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"


Dim hdrid
hdrid = Request.QueryString("hdrid")
 

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=hoja_de_ruta_" & hdrid & ".xls"

%>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        table {
            border-collapse: collapse;
            width: 100%;
        }
        th, td {
            border: 1px solid #000;
            padding: 5px;
            text-align: center;
        }
        th {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
<%
Dim rs
Set rs = conn.Execute("EXEC usp_Transportista_HojaDeRuta_Control_Resumen " & hdrid)
On Error Resume Next 
If Err.Number <> 0 Then
    Response.Write "Error: " & Err.Description
    Response.End
End If

%>

<h2>Clientes de la Hoja de Ruta <%= hdrid %></h2>
<table>
    <thead>
        <tr>
			<th>ClienteID</th>
            <th>Cliente</th>
            <th>Facturas</th> 
            <th>Total</th>
            <th>Total Cobrado</th>
            <th>Efectivo</th>
            <th>Cheque</th>
            <th>Transferencia</th>
            <th>Descuentos</th>
            <th>Diferencia</th>
            <th>Observacion</th> 
            <th>Acciones</th>
            <th>FormaPago</th>
        </tr>
    </thead>
    <tbody>
<%
Do Until rs.EOF
%>
        <tr>	
			<td><%= rs("ClienteID") %></td>
            <td><%= rs("ClienteNombre") %></td>
            <td><%= rs("TotalFacturas") %></td>
            <td><%= replace(replace(FormatNumber(rs("ImporteACobrar"), 2),".",""),",",".") %></td>
            <td><%= replace(replace(FormatNumber(rs("TotalCobrado"), 2),".",""),",",".") %></td>
            <td><%= replace(replace(FormatNumber(rs("Efectivo"), 2),".",""),",",".") %></td>
            <td><%= replace(replace(FormatNumber(rs("Cheque"), 2),".",""),",",".") %></td>
            <td><%= replace(replace(FormatNumber(rs("Transferencia"), 2),".",""),",",".") %></td>
            <td><%= replace(replace(FormatNumber(rs("Descuentos"), 2),".",""),",",".") %></td>
            <td><%= replace(replace(FormatNumber(rs("Diferencia"), 2),".",""),",",".") %></td>
            <td><%= rs("EstadoEntrega")%> &nbsp; <%= rs("Observacion") %></td> 
            <td><%= rs("Acciones") %></td>
            <td><%= rs("FormaPago") %></td>
        </tr>
<%
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
    </tbody>
</table>
</body>
</html>
