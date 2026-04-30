<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("NombreTransportista") = "" Then
    Response.Redirect "login.asp"
End If


%>

<!--#include file="sidebar.asp" -->
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel Transportista</title>
    <link rel="stylesheet" href="estilos.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
    <style>
        /* ----  CARD LAYOUT  ----*/
        .cards-container{display:flex;flex-direction:column;gap:12px}
        .cliente-card{border:1px solid #ddd;border-radius:6px;padding:12px;display:flex;flex-direction:column;box-shadow:0 2px 3px rgba(0,0,0,.05);transition:background .3s}
        .card-row{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap}
        .card-row span{margin-right:8px;font-size:.9rem}
        .card-actions{margin-top:8px;display:flex;gap:8px}
        .total-bar{margin-top:16px;background:#f7f7f7;border:1px solid #ddd;border-radius:6px;padding:10px;font-weight:bold;display:flex;justify-content:space-between}
        @media(min-width:550px){.cliente-card{flex-direction:row;align-items:center}.card-row{flex:1;margin:0}}

        /* estados de entrega */
        .estado-entregado{background:#e8f7e4}
        .estado-pendiente{background:#fff9e6}
        .estado-anulada{background:#fdecec}
        .estado-otro{background:#f8f8f8}

        /* tooltip */
        .tooltip{position:relative;display:inline-block}
        .tooltip .tooltiptext{visibility:hidden;width:160px;background-color:#555;color:#fff;text-align:center;padding:5px 0;border-radius:6px;position:absolute;z-index:1;bottom:125%;left:50%;margin-left:-80px;opacity:0;transition:opacity .3s}
        .tooltip:hover .tooltiptext{visibility:visible;opacity:1}
        .detalle-toggle{cursor:pointer;color:#007bff;text-decoration:underline}
    </style>
    <script>
        function solicitarReenvio(hdrid,clienteid){var tel=prompt("Ingrese el número de teléfono para reenviar:");if(tel&&tel.trim()!=""){location.href="reenvio_guardar.asp?HojaDeRutaID="+hdrid+"&ClienteID="+clienteid+"&Telefono="+encodeURIComponent(tel)}}
        /*function solicitarDNIyCerrar(hdrid){var dni=prompt("Por favor, ingrese el DNI de la persona que recibe:");if(dni&&dni.trim()!==""&&/^\d{7,10}$/.test(dni.trim())){location.href="cerrarhdr.asp?hdrid="+hdrid+"&DNI="+encodeURIComponent(dni)}else{alert("Debe ingresar un DNI válido (solo números, mínimo 7 dígitos).")}}*/
        function mostrarClienteNombre(n,a){alert("Cliente: "+n + "\nDireccion: "+a)}
		function openMap(lat,lon){if(lat&&lon){window.open("https://www.google.com/maps?q="+lat+","+lon,"_blank");}}
        function toggleSidebar(){document.querySelector('.sidebar').classList.toggle('open');}
    </script>
</head>
<body>
<header>
    <button class="menu-toggle" onclick="toggleSidebar()"><i class="fas fa-bars"></i></button>
    <strong style="flex:1;">👤 <%=Server.HTMLEncode(Session("NombreTransportista"))%></strong>
    <form method="post" action="logout.asp" style="margin:0;">
        <input type="submit" value="Cerrar sesión" class="logout">
    </form>
</header>
<div class="main-content">
<% If Request("msg") = "ok" Then %>
<div style="background:#d4edda;color:#155724;padding:10px;margin-bottom:10px;border:1px solid #c3e6cb;">✅ Hoja de ruta cerrada correctamente.</div>
<% ElseIf Request("msg") = "reenvio_ok" Then %>
<div style="background:#d1ecf1;color:#0c5460;padding:10px;margin-bottom:10px;border:1px solid #bee5eb;">📩 Reenvío registrado con éxito.</div>
<% End If %>
<%
Dim hojaRutaID:hojaRutaID=Request.QueryString("hdrid")
If hojaRutaID="" Then
	'response.write "EXEC usp_Transportista_HojaDeRuta_Abiertas 0," & Session("TransportistaID")
  Set rs=conn.Execute("EXEC usp_Transportista_HojaDeRuta_Abiertas 0," & Session("TransportistaID"))
%>
<h2>Hojas de Ruta Abiertas</h2>
<div class="cards-container">
<% 
	If rs.EOF Then
%>
	<div class="cliente-card estado-otro">
    <div class="card-row"><span><strong>No tiene hojas de rutas pendientes</strong></span></div>
	</div>
<%
	Else
 
	Do Until rs.EOF 
%>
  <div class="cliente-card estado-otro">
    <div class="card-row"><span><strong>HDR:</strong> <%=rs("HojaDeRutaID")%></span><span><strong>Clientes:</strong> <%=rs("TotalClientes")%></span><span><strong>Clientes sin trabajar:</strong> <%=rs("ClientesPendientes")%></span></div>
    <div class="card-actions">
        <a href="hojaderutaV3.asp?hdrid=<%=rs("HojaDeRutaID")%>" title="Seleccionar hoja de ruta"><i class="fa-solid fa-truck-fast"></i></a>
        <a href="cerrarhdr.asp?hdrid=<%=rs("HojaDeRutaID")%>" title="Cerrar"><i class="fa-solid fa-lock-open"></i></a>
    </div>
  </div>
<% 
	rs.MoveNext:Loop:rs.Close:Set rs=Nothing 
	End If
%>
</div>
<% Else '--- detalle de HDR --- %>
<% Set rs=conn.Execute("EXEC usp_Transportista_HojaDeRuta_Cabecera_SelV3 " & hojaRutaID) %>
<div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
  <h1 style="margin:0;">Hoja de Ruta: <%=hojaRutaID%></h1>
  <a href="hojaderutaV3.asp" class="icon-btn" title="Volver" style="background:#2e30c7;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;margin-left:10px;"><i class="fas fa-arrow-left"></i> Volver</a>
  <a href="cerrarhdr.asp?hdrid=<%=hojaRutaID%>" class="icon-btn" title="Cerrar" style="background:#27ae60;color:#fff;padding:8px 12px;border-radius:4px;text-decoration:none;"><i class="fas fa-lock"></i> Cerrar</a>
</div>
<% Dim totalACobrar,totalCobrado:totalACobrar=0:totalCobrado=0 %>
<div class="cards-container">
<% Do Until rs.EOF
   Dim estado:estado=CInt(rs("estado"))
   Dim icono,tooltip,statusClass
   Select Case estado
     Case 1:icono="<i class='fas fa-check-circle' style='color:green;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-entregado"
     Case 2:icono="<i class='fas fa-exclamation-circle' style='color:#e69500;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-anulada"
     Case 3:icono="<i class='fas fa-times-circle' style='color:red;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-pendiente"
     Case Else:icono="<i class='fas fa-hourglass-half' style='color:gray;'></i>":tooltip=rs("EstadoEntrega"):statusClass="estado-otro"
   End Select
   totalACobrar=totalACobrar+CDbl(rs("ImporteACobrar"))
   totalCobrado=totalCobrado+CDbl(rs("TotalCobrado"))
   
   Dim lat,lon:lat="" : lon=""
   On Error Resume Next
   lat = rs("Latitud"):lon = rs("Longitud")
   
%>
   <div class="cliente-card <%=statusClass%>">
     <div class="card-row">
			 
        <span class="detalle-toggle" onclick="mostrarClienteNombre('<%=Replace(rs("ClienteNombre"),"'","\'")%>','<%=Replace(rs("Direccion"),"'","\'")%>')"><strong>Cliente:</strong> <%=rs("ClienteID")%></span>
        <span><strong>Cond:</strong> <%=rs("FormaPago")%></span>
        <span><strong>Facturas:</strong> <%=rs("TotalFacturas")%></span>
		<span><strong>Imp:</strong> $<%=FormatNumber(rs("ImporteACobrar"),2)%></span>
	
<% If UCase(Trim(rs("FormaPago"))) = "CONTADO" And estado = 1 Then %>
    <span>
        <i class="fas fa-dollar-sign" title="Efectivo"></i> $<%=FormatNumber(rs("efectivo"),2)%> &nbsp;
        <i class="fas fa-exchange-alt" title="Transferencia"></i> $<%=FormatNumber(rs("transferencia"),2)%>&nbsp; 		
        <i class="fas fa-money-check" title="Cheque"></i> $<%=FormatNumber(rs("cheque"),2)%> 
    </span>
<% Else %>
        <span><strong>Cobrado:</strong> $<%=FormatNumber(rs("TotalCobrado"),2)%></span>
<% End If %>



	
     </div>
     <div class="card-actions"> 
		<a href="#" class="icon-btn reenviar" title="Reenviar al cliente" onclick="solicitarReenvio('<%=hojaRutaID%>','<%=rs("ClienteID")%>');return false;"><i class="fas fa-share"></i></a>
	    <a href="entregarV3.asp?clienteid=<%=rs("ClienteID")%>&hdrid=<%=hojaRutaID%>&cond=<%= Server.URLEncode(rs("FormaPago"))%>" class="icon-btn" title="Marcar como entregado"><i class="fa-solid fa-box-open"></i></a>
		
        <% If estado=0 Then %>
           <a href="rechazar.asp?clienteid=<%=rs("ClienteID")%>&hdrid=<%=hojaRutaID%>&cond=<%= Server.URLEncode(rs("FormaPago"))%>" class="icon-btn" title="Rechazar entrega"><i class="fas fa-ban"></i></a>
        <% Else %>
           <span class="tooltip"><%=icono%><span class="tooltiptext"><%=tooltip%></span></span>
        <% End If %>

        <% If lat<>"" And lon<>"" Then %>
           <span><a href="#" title="Ver en mapa" onclick="openMap('<%=replace(lat,",",".")%>','<%=replace(lon,",",".")%>');return false;"><i class="fas fa-map-marked-alt"></i></a></span>
        <% End If %>		
		

     </div>

   </div>
<% rs.MoveNext:Loop %>
</div>
<%
set rs = rs.NextRecordset
%>
<div class="total-bar"> 
	<strong>Resumen</strong>
</div>
<%
Do Until rs.EOF
%>
     <div class="total-bar"><span><strong>Total Facturas</strong> $<%=FormatNumber(rs("ImporteACobrar"),2)%> </span></div>
     <div class="total-bar"><span><strong>Total Efectivo:</strong> $<%=FormatNumber(rs("efectivo"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total Cheque:</strong> $<%=FormatNumber(rs("cheque"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total Transferencia:</strong> $<%=FormatNumber(rs("transferencia"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total CC:</strong> $<%=FormatNumber(rs("CC"),2)%>  </span> </div>
     <div class="total-bar"><span><strong>Total Cobrado</strong> $<%=FormatNumber(rs("totalCobrado"),2)%> </span>  </div>

	
<% rs.MoveNext:Loop:rs.Close:Set rs=Nothing %>
 
<% End If %>
</div>
<%
conn.Close:Set conn=Nothing
%>
</body>
</html>
