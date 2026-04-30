<%
'--- Trae las transferencias del período solicitado
sSQL = "EXEC [dbo].[usp_Transferencias_Pendientes_PorClientes_Sel] " & periodoID & "," & Session("PUNTO_VENTA_ID") 
Set dbRS = Server.CreateObject("ADODB.Recordset")
dbRS.Open sSQL, dbCon
IF Err.Number <> 0 THEN
    Response.Clear
    Response.Redirect "./error.asp"
End If
%>

<table border="0" align="center" cellpadding="0" cellspacing="0"
       class="contentTable table-width-lg"
       style="width:100%; margin-bottom:15px;">

  <!---- Título de la tabla ---->
  <tr>
    <td colspan="6" height="20">
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr class="tableHeader">
          <td align="left"><span class="formHeader">Transferencias Pendientes</span></td>
        </tr>
      </table>
    </td>
  </tr>

  <!---- Encabezados de columna ---->
  <tr align="center" class="columnTop">
    <td width="5%"><b>Fecha</b></td>
    <td width="15%"><b>Banco</b></td>
    <td width="35%"><b>Cliente</b></td> 
    <td width="15%"><b>Importe</b></td>
    <td width="10%"><b>Nro&nbsp;Cliente</b></td>
    <td width="15%"><b>Accion</b></td>
  </tr>

<% If Not dbRS.EOF Then
     Do Until dbRS.EOF
%>
  <tr class="itemTD11">
    <td><%= dbRS("Fecha") %></td>
    <td><%= dbRS("Nombre") %></td>
    <td><%= dbRS("Cliente") %></td> 
    <td>$ <%= FormatNumber(dbRS("Importe"),2) %></td>

    <!---- input Nº Cliente (máx 8 dígitos) ---->
    <td>
      <input type="text"
             name="NroCliente_<%= dbRS("RowID") %>"
             id="NroCliente_<%= dbRS("RowID") %>"
             size="8"  maxlength="8"
             pattern="\d{4,8}"
             title="4 a 8 dígitos"
             autocomplete="off">
    </td>

    <!---- Botón Guardar para esta fila ---->
    <td>
      <a href="javascript:guardarFila(<%= dbRS("RowID") %>)"
         title="Guardar cambios">
         <img src="../images/guardado.png" alt="Guardar">
      </a>
    </td>
  </tr>
<%
       dbRS.MoveNext
     Loop
   Else
%>
  <tr class="itemTD11">
    <td colspan="6">No hay transferencias pendientes.</td>
  </tr>
<% End If %>
</table>

<%
dbRS.Close
Set dbRS = Nothing
%>
