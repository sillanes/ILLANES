<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("NombreTransportista") = "" Then Response.End

Dim facturaID
facturaID = Request.QueryString("facturaid")
If facturaID = "" Then Response.End

Dim sql, rs
sql = "EXEC dbo.usp_Transportista_HojaDeRuta_Cliente_Facturas_Detalle_Sel " & facturaID
Set rs = conn.Execute(sql)

If rs.EOF Then
  Response.Write "<li class='list-group-item text-muted'>Sin ítems</li>"
  Response.End
End If

Function FormatearParaJS(ByVal valor)
  Dim s
  s = Trim(CStr(valor))
  s = Replace(s, ".", "")
  s = Replace(s, ",", ".")
  FormatearParaJS = s
End Function
%>

<%
Do Until rs.EOF
%>
<li class="list-group-item d-flex justify-content-between align-items-start item-factura"
    data-facturaid="<%= facturaID %>"
    data-cantorig="<%= rs("Cantidad") %>"
    data-totalimp="<%= FormatearParaJS(rs("TotalConImpuestos")) %>">

  <div class="ms-2 me-auto">
    <div class="fw-bold"><%= rs("Descripcion") %></div>

    <div>
      Artículo: <%= rs("Articulo") %>
      &nbsp;&nbsp;
      Cant:
      <span class="item-cantidad"><%= rs("Cantidad") %></span>

      <!-- 👇 CAMBIO CLAVE -->
      <input type="hidden"
             name="item_<%= facturaID %>_<%= rs("Articulo") %>"
             value="<%= rs("Cantidad") %>"
             class="input-cantidad-item">

      <% If FormatNumber(rs("Dto"),2) <> 100 Then %>
        <button type="button"
                class="btn btn-sm btn-outline-secondary ms-1 btn-restar">-</button>

        <% If FormatNumber(rs("Dto"),2) > 0 Then %>
          <span class="badge bg-warning text-dark ms-2">Dto</span>
        <% End If %>
      <% Else %>
        <span class="badge bg-warning text-dark ms-2">Dto</span>
      <% End If %>
    </div>

    <small>
      Unitario: $<%= FormatNumber(rs("PrecioUnitario"), 2) %><br>
      IVA: $<%= FormatNumber(rs("ImporteConIVA"), 2) %><br>
      IIBB: $<%= FormatNumber(rs("PercepcionIIBB"), 2) %><br>

      <strong>Total c/ Impuestos:
        $<span class="item-total-imp">
          <%= FormatNumber(rs("TotalConImpuestos"), 2) %>
        </span>
      </strong>
    </small>
  </div>
</li>
<%
  rs.MoveNext
Loop

rs.Close
Set rs = Nothing
%>
