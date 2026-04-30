<!--#include file="header.asp" -->

<%
Dim hojaRutaID
hojaRutaID = Request.QueryString("ID")
If hojaRutaID = "" Then
    Response.Write "<div class='alert alert-warning'>No se especificó la Hoja de Ruta.</div>"
Else
    ' Consulta ejemplo, reemplaza con la tuya
    Set rs = conn.Execute("EXEC usp_HojaDeRuta_Detalle " & hojaRutaID)
%>

<h1>Hoja de Ruta ID: <%= hojaRutaID %></h1>

<table class="table table-striped table-bordered align-middle">
    <thead class="table-light">
        <tr>
            <th>Cliente</th>
            <th>Dirección</th>
            <th>Estado</th>
            <th>Fecha Entrega</th>
            <th>Acciones</th>
        </tr>
    </thead>
    <tbody>
        <%
        Do Until rs.EOF
        %>
        <tr>
            <td><%= rs("ClienteNombre") %></td>
            <td><%= rs("Direccion") %></td>
            <td><%= rs("Estado") %></td>
            <td><%= rs("FechaEntrega") %></td>
            <td>
                <!-- Aquí podrías poner botones para marcar entrega, editar, etc -->
                <a href="entregar.asp?ID=<%= rs("DetalleID") %>" class="btn btn-sm btn-primary" title="Marcar como entregado"><i class="fas fa-check"></i></a>
            </td>
        </tr>
        <%
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
    </tbody>
</table>

<%
End If
conn.Close
Set conn = Nothing
%>

<!--#include file="footer.asp" -->
