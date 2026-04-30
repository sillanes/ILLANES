<%@ Language=VBScript %> 

<!--#include file="includes/punto_venta.asp" --> 


<!--#INCLUDE FILE="_upload.asp"-->
<%
'ON ERROR RESUME NEXT
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 
 

If Request.QueryString.Count>0 Then
	Set objRequest = Request.QueryString
Else
	Set objRequest = Request.Form
End If

doWhat = objRequest("doWhat")  
   
If doWhat = "-2"  Then 
	submit_logout = "Salir"
	Session("currentUser") = "" 
End If


'submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If 

Dim NombreMenu
if Session("currentUser") = "illanes"  OR Session("currentUser") = "admin" or Session("currentUser") = "hojaderuta" or Session("currentUser") = "pxp"  or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" Then
	NombreMenu = "SISTEMA DE GESTIÓN DE LOGÍSTICA"
ElseIf Session("currentUser") = "admin" or Session("currentUser") = "1001" Then
	NombreMenu = "SISTEMA DE GESTIÓN DE VENTAS"
ElseIF Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion"  or Session("currentUser") = "rillanes" or Session("currentUser") = "bancos"  or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "eillanes"    or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe"  or Session("currentUser") = "sbosio"  or Session("currentUser") = "mperla"  or Session("currentUser") = "cromero"  or Session("currentUser") = "ahenriquez" or Session("currentUser") = "ccordero" Then
	NombreMenu = "SISTEMA DE GESTIÓN ADMINISTRATIVO"
Else
	NombreMenu ="SISTEMA DE GESTIÓN DE LOGÍSTICA"
End IF
    

Set objRequest = Nothing
%>
<HTML>
<HEAD>
<TITLE>ILLANES HNOS SRL</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="../includes/style.css">  
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
<link rel="stylesheet" type="text/css" href="../includes/calendar_cool.css" media="all" />
<script type="text/javascript" src="../includes/calendar_cool.js"></script>
<script type="text/javascript" src="../includes/copy.js"></script>

 
  
</HEAD>
<body> 
	<form name="FF" method="post" action="menu.asp">  
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
    <div class="wrap" align="center">
  
		
 	
	<h2>   
		<%=NombreMenu%><br/> 	 
		Bienvenidx: <span class="textfail"> <%=UCASE(Session("currentUser")) %></span><br/> 	 
	</h2>


<% if  Session("currentUser") = "admin"  Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>Campanias<br/></td> 
		</tr> 	
		<%if Session("currentUser") = "admin" or Session("currentUser") = "rrhh"  Then %>	
		<tr>
			<td>Adminsitracion de Campanias</td>
			<td align="right">
			<div>
				<a href="./campanias/home.asp" title="Recibos"> <img src="../images/config.png" alt="Campanias"></a>
			</div>
			</td>
		</tr>	
		<%End If %>	 
		
	</table>
<%End If%>


<% if   Session("currentUser") = "admin"  or Session("currentUser") = "millanes" or     Session("currentUser") = "eillanes" or Session("currentUser") = "rrhh" Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>CHAT BOT<br/></td> 
		</tr>
		
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes"  or Session("currentUser") = "rrhh"  Then  %>
		<tr>
			<td>Mensajes</td>
			<td align="right">
			<div>
				<a href="./chatbot/mensajes_sofia.asp" title="Ver Mensajes"> <img src="../images/chatbot.png" alt="Ver mensajes"></a>
			</div>
			</td>
		</tr>

		<%End If%>	
		 
 
	</table>
<%End If%>
	
<% if   Session("currentUser") = "admin"  or Session("currentUser") = "millanes" or   Session("currentUser") = "rillanes" or Session("currentUser") = "eillanes"  or Session("currentUser") = "logistica" or Session("currentUser") = "rrhh" Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>CAJAS NAVIDEÑAS<br/></td> 
		</tr>
		
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes"  or Session("currentUser") = "rrhh"  Then  %>
		<tr>
			<td>Gestión</td>
			<td align="right">
			<div>
				<a href="./empretienda/home.asp" title="Gestionar Campañas"> <img src="../images/config.png" alt="Gestionar Campañas"></a>
			</div>
			</td>
		</tr>

		<%End If%>	
		
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "eillanes"  or Session("currentUser") = "millanes" Then  %>
		<tr>
			<td>Subir Datos Empretienda</td>
			<td align="right">
			<div>
				<a href="./subirordenes.asp" title="Subir Ordernes"> <img src="../images/upload.png" alt="Subir Ordenes"></a>
			</div>
			</td>
		</tr>
		
		<%End If%>	

		<%if Session("currentUser") = "admin"  or Session("currentUser") = "eillanes"  or Session("currentUser") = "millanes" Then  %>
		 <tr>
			<td>Procesar Ordenes</td>
			<td align="right">
			<div>
				<a href="./ordenes.asp" title="Procesar Ordenes"> <img src="../images/transferencias.png" alt="Procesar Ordenes"></a>
			</div>
			</td>
		</tr> 
		
		
		<%End If%>	

		<%if Session("currentUser") = "admin"  or Session("currentUser") = "eillanes"  or Session("currentUser") = "rillanes" Then  %>

		<tr>
			<td>Cobranza</td>
			<td align="right">
			<div>
				<a href="./imputarOrdenes.asp" title="Cobranza"> <img src="../images/cobranza2.png" alt="Subir extracto"></a>
			</div>
			</td>
		</tr>		

		<%End If%>	
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "eillanes"  or Session("currentUser") = "millanes" or Session("currentUser") = "rillanes" Then  %>
		<tr>
			<td>Listados Ordenes pendientes</td>
			<td align="right">
			<div>
				<a href="./ordenespendientes.asp" title="Ordenes Pendientes"> <img src="../images/reporte.png" alt="Ordenes Pendientes"></a>
			</div>
			</td>
		</tr>
		
		<%End If%>	
		
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "eillanes"  or Session("currentUser") = "millanes"  or Session("currentUser") = "logistica"  Then  %>
		
		
		<tr>
			<td>Logistica</td>
			<td align="right">
			<div>
				<a href="./menu.asp" title="Logisitca"> <img src="../images/control.png" alt="Logitica"></a>
			</div>
			</td>
		</tr>
		
		<%End If%>	
		  
 
	</table>
<%End If%>
	
<% if   Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio"  Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>Transportista<br/></td> 
		</tr>
		
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then  %>
		<tr>
			<td>Alta Vehiculo</td>
			<td align="right">
			<div>
				<a href="./repartovehiculos.asp" title="Vehiculos"> <img src="../images/alta.png" alt="Modificación de Vehiculos"></a>
			</div>
			</td>
		</tr>
		<%End If%>	
		
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio"  Then  %>
		<tr>
			<td>Modificación Telefono Transportista</td>
			<td align="right">
			<div>
				<a href="Transportistas_Telefonos_Modificar.asp" title="Modificación Telefono"> <img src="../images/Telefono.png" alt="Modificación de Telefonos"></a>
			</div>
			</td>
		</tr>
		<%End If%>	
 
		  
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then  %>
		<tr>
			<td>Vincular HDR con reparto</td>
			<td align="right">
			<div>
				<a href="TransportistasAsociarHDR.asp" title="Vincular HDR con Repartidores"> <img src="../images/vincular2.png" alt="Vincular HDR"></a>
			</div>
			</td>
		</tr>
		<%End If%>		

		<%if Session("currentUser") = "ad2min"  or Session("currentUser") = "lso2sa" or Session("currentUser") = "mi2llanes" or Session("currentUser") = "sbosio" Then  %>
		<tr>
			<td>Controlar HDR</td>
			<td align="right">
			<div>
				<a href="/administracion/controlhdrv2.asp" title="Control"> <img src="../images/Control.png" alt="Controlar HDR"></a>
			</div>
			</td>
		</tr>
		<%End If%>	

		<%if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
		<tr>
			<td>Nuevo Controlar HDR</td>
			<td align="right">
			<div>
				<a href="/administracion/controlhdr.asp" title="Control"> <img src="../images/Control.png" alt="Controlar HDR"></a>
			</div>
			</td>
		</tr>
		<%End If%>			
		
 
 
	</table>
<%End If%>

<% if    Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes"  or Session("currentUser") = "sbosio" Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>FACTURA ELECTRONICA<br/></td> 
		</tr>
		
		<%if  Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then  %>
		<tr>
			<td>Envio WhastApp</td>
			<td align="right">
			<div>
				<a href="./enviarwhatsapp.asp" title="Envio por whastapp"> <img src="../images/whastapplogo2.png" alt="Envio Por WhastApp"></a>
			</div>
			</td>
		</tr>
		<%End If%>	
		<%if  Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then  %>
		<tr>
			<td>Reenvio WhastApp</td>
			<td align="right">
			<div>
				<a href="./reenviowhatsapp.asp" title="Reevio por whastapp"> <img src="../images/reenviar.png" alt="Reenvio Por WhastApp"></a>
			</div>
			</td>
		</tr>
		<%End If%>			
		<%if  Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes"  or Session("currentUser") = "sbosio" Then  %>
		<tr>
			<td>Adherir Cliente</td>
			<td align="right">
			<div>
				<a href="./facturaDigital.asp" title="Adherir al cliente"> <img src="../images/alta.png" alt="Adherir cliente"></a>
			</div>
			</td>
		</tr>
		<%End If%>
		<%if  Session("currentUser") = "admin"  or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then  %>
		<tr>
			<td>Listado de mensajes con error o excluidos</td>
			<td align="right">
			<div>
				<a href="./facturadigitalpendientes.asp" title="Listado de mensajes con error o excluidos"> <img src="../images/reporte.png" alt="Listado de mensajes con error o excluidos"></a>
			</div>
			</td>
		</tr>
		<%End If%>
	</table>
<%End If%>
	
	
<% if Session("currentUser") = "illanes"  OR  Session("currentUser") = "hsarao" OR Session("currentUser") = "admin" or Session("currentUser") = "hojaderuta" or Session("currentUser") = "pxp"  or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then%>

	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>RECLAMOS<br/></td> 
		</tr>
	<% if Session("currentUser") = "illanes"  OR  Session("currentUser") = "hsarano" OR Session("currentUser") = "admin"  or Session("currentUser") = "fliporace"  or Session("currentUser") = "millanes"  Then%>				
			<tr>
				<td>
					Estadísticas
				</td>
				<td align="right">
				<div>
					<a href="./dashboardV3.asp" title="Ver estadisticas"><img src="../images/estad.png" alt="Ver estadisticas"></a>
				</div>
				</td>
			</tr>
			<tr>
				<td>Revisión de reclamos</td>
				<td align="right">
					<div>
						<a href="./reclamoscheck.asp" title="Ver Reclamo"> <img src="../images/eye.png" alt="Ver Reclamo"></a>
					</div>
				</td>		
			</tr>
			<tr>
				<td>Buscar reclamo cerrado</td>
				<td align="right">
				<div>
					<a href="./buscarreclamo.asp" title="Buscar Reclamo"> <img src="../images/view.png" alt="Buscar Reclamo"></a>
				</div>
				</td>
			</tr>
			<tr>
				<td>Descargar Reclamos</td>
				<td align="right">
				<div>
					<a href="./descargareclamos.asp" title="Descargar Reclamo"> <img src="../images/excel.png" alt="Descargar Reclamo"></a>
				</div>
				</td>
			</tr>
	<%End If%>
	<% if Session("currentUser") = "hojaderuta" OR  Session("currentUser") = "hsarano" OR Session("currentUser") = "admin" or Session("currentUser") = "lsosa"   or Session("currentUser") = "millanes" or   Session("currentUser") = "fliporace" or Session("currentUser") = "sbosio" Then%>		
			<tr>
				<td>Subir Hoja de Ruta</td>
				<td align="right">
				<div>
					<a href="./subirhojaderuta.asp" title="hoja de ruta"> <img src="../images/upload.png" alt="Subir Hoja de Ruta"></a>
				</div>
				</td>
			</tr>	
			<tr>
				<td>Ver Hojas de Ruta</td>
				<td align="right">
				<div>
					<a href="./hojaderuta.asp" title="hoja de ruta"> <img src="../images/eye.png" alt="Ver Hoja de rutas"></a>
				</div>
				</td>
			</tr>	
				
	<%End If%>
	<% if Session("currentUser") = "admin" Then%>		
			<tr>
				<td>Configuracion de Feriados</td>
				<td align="right">
				<div>
					<a href="./feriados.asp" title="Feriados"> <img src="../images/calendar2_.gif" alt="Feriados"></a>
				</div>
				</td>
			</tr>				
	<%End If%>
	
	<% if Session("currentUser") = "admin" OR  Session("currentUser") = "hsarano" or Session("currentUser") = "rrhh" or Session("currentUser") = "pxp" or Session("currentUser") = "illanes"   or Session("currentUser") = "millanes"   or Session("currentUser") = "fliporace"  Then%>		
			<tr>
				<td>Cargar Reclamo</td>
				<td align="right">
				<div>
					<a href="./pxpcargarreclamo.asp" title="Alta reclamo"> <img src="../images/reporte.png" alt="Feriados"></a>
				</div>
				</td>
			</tr>				
	<%End If%>	

		</table>
				
<%End If%>

<% if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "controlador" or Session("currentUser") = "rrhh" OR  Session("currentUser") = "hsarano" or  Session("currentUser") = "fliporace"   Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>PEDIDO POR PEDIDO<br/></td> 
		</tr>

		<tr>
			<td>Reporte de Eficiencia Tiempo Real</td>
			<td align="right">
			<div>
				<a href="../logistica/reporte_control.asp" title="Reporte RT"> <img src="../images/control.png" alt="Reporte RT"></a>
			</div>
			</td>
		</tr>	

		<%if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR  Session("currentUser") = "hsarano" or  Session("currentUser") = "fliporace"   Then%>		
		
		<tr>
			<td>Subir Hoja De Ruta para pedido por pedido</td>
			<td align="right">
			<div>
				<a href="./subirhdrpxp.asp" title="hoja de ruta"> <img src="../images/upload.png" alt="Subir Hoja de Ruta"></a>
			</div>
			</td>
		</tr>	
		<%end if%>
		
		<tr>
			<td>Pendientes de control</td>
			<td align="right">
			<div>
				<a href="./pxppendientescontrol.asp" title="Ver Pendientes"> <img src="../images/eye.png" alt="Ver Pendientes"></a>
			</div>
			</td>
		</tr>			
 
		<%if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR  Session("currentUser") = "hsarano" or  Session("currentUser") = "fliporace"   Then%>		
		
		<tr>
			<td>Reportes</td>
			<td align="right">
			<div>
				<a href="./pxpreportes.asp" title="Estados"> <img src="../images/estad.png" alt="Ver avances"></a>
			</div>
			</td>
		</tr>	
	
		<%end if%>
		<tr>
			<td>Controlar Hojas de Ruta</td>
			<td align="right">
			<div>
				<a href="./pxpcontrol.asp" title="Controlar Pedidos"> <img src="../images/reporte.png" alt="Controlar pedidos"></a>
			</div>
			</td>
		</tr>
	<%if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR  Session("currentUser") = "hsarano" or  Session("currentUser") = "fliporace"   Then%>		
				
		<tr>
			<td>Cerrar Hoja De Ruta</td>
			<td align="right">
			<div>
				<a href="./pxpcerrarhojaderuta.asp" title="Cerrar hoja de ruta"> <img src="../images/candado.png" alt="Cerrar una hoja de ruta"></a>
			</div>
			</td>
		</tr>
	
		<tr>
			<td>Agregar o quitar Armador</td>
			<td align="right">
			<div>
				<a href="./pxpabmarmador.asp" title="Modificar Armador"> <img src="../images/eye.png" alt="Modificar Armador"></a>
			</div>
			</td>
		</tr>	
			<tr>
			<td>Agregar o quitar Controlador</td>
			<td align="right">
			<div>
				<a href="./pxpabmcontrolador.asp" title="Modificar Controlador"> <img src="../images/eye.png" alt="Modificar Controlador"></a>
			</div>
			</td>
		</tr>	
		</tr>	
			<tr>
			<td>Historico Por Periodos</td>
			<td align="right">
			<div>
				<a href="./pxpdescargarperiodo.asp" title="Historico Por Periodos"> <img src="../images/excel.png" alt="Descargar Datos"></a>
			</div>
			</td>
		</tr>	
		<%end if%>
		</tr>	
			<tr>
			<td>Reclamos por Controlador</td>
			<td align="right">
			<div>
				<a href="./pxpdescargarreclamosxcontrolador.asp" title="Relcamos por controlador"> <img src="../images/excel.png" alt="Descargar Datos"></a>
			</div>
			</td>
		</tr>		
		<%if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR  Session("currentUser") = "hsarano" or  Session("currentUser") = "fliporace"   Then%>		
		<tr>
		<td>Alta Nomina</td>
			<td align="right">
			<div>
				<a href="./altanomina.asp" title="Agregar"> <img src="../images/alta.png" alt="Agregar"></a>
			</div>
			</td>
		</tr>		
		<%end if%>

	</table>
<%End If%>


<% if   Session("currentUser") = "admin" or Session("currentUser") = "1001"   Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>VIPO<br/></td> 
		</tr>
		
		<tr>
			<td>Controlar Fotos</td>
			<td align="right">
			<div>
				<a href="./vendedores.asp" title="hoja de ruta"> <img src="../images/reporte.png" alt="Controlar Fotos"></a>
			</div>
			</td>
		</tr>		
  
		</tr>	
			<tr>
			<td>Descargar por periodos</td>
			<td align="right">
			<div>
				<a href="./vendescargarestadoporperiodos.asp" title="Historico Por Periodos"> <img src="../images/excel.png" alt="Descargar Datos"></a>
			</div>
			</td>
		</tr>
		
	</table>
<%End If%>




<% if   Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion"  or Session("currentUser") = "rillanes" or Session("currentUser") = "bancos"  or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "eillanes"    or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "sbosio"  or Session("currentUser") = "ça"  or Session("currentUser") = "cromero"  or Session("currentUser") = "ahenriquez" or Session("currentUser") = "ccordero" or Session("currentUser") = "mperla" Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>Transferencias<br/></td> 
		</tr>
		
		<%if Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion"    or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "lsosa" Then %>	
		<tr>
			<td>Subir Extractos Bancarios</td>
			<td align="right">
			<div>
				<a href="./subirextracto.asp" title="Subir Extracto"> <img src="../images/upload.png" alt="Subir extracto"></a>
			</div>
			</td>
		</tr>	
		<%End If %>		
		
		<%if Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion"    or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "lsosa" Then %>	
		<tr>
			<td>Procesar extractos</td>
			<td align="right">
			<div>
				<a href="./extractos.asp" title="Procesar Extractos"> <img src="../images/eye.png" alt="Procesar Extractos"></a>
			</div>
			</td>
		</tr>
	
		<%End If %>	
		<%if Session("currentUser") = "rillanes" or Session("currentUser") = "admin" or Session("currentUser") = "bancos"  or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa"  or Session("currentUser") = "Administracion" or Session("currentUser") = "eillanes"   or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "sbosio"  or Session("currentUser") = "mperla"  or Session("currentUser") = "cromero"  or Session("currentUser") = "ahenriquez" or Session("currentUser") = "ccordero" Then %>	
		<tr>
			<td>Imputar transferencias</td>
			<td align="right">
			<div>
				<a href="./imputartransferencias.asp" title="Imputar Transferencias"> <img src="../images/transferencias.png" alt="Imputar Transferencias"></a>
			</div>
			</td>
		</tr>	
		<%End If%>	 	
		<tr>
			<td>Listados transferencias pendientes</td>
			<td align="right">
			<div>
				<a href="./transferenciaspendientes.asp" title="Transferencias Pendientes"> <img src="../images/reporte.png" alt="Transferencias Pendientes"></a>
			</div>
			</td>
		</tr>	
	 
		
	</table>
<%End If%>




<% if   Session("currentUser") = "admin" or Session("currentUser") = "logistica"  Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>Control Reparto<br/></td> 
		</tr>
		
		<%if Session("currentUser") = "admin" or Session("currentUser") = "logistica"  Then %>	
		<tr>
			<td>Pendientes</td>
			<td align="right">
			<div>
				<a href="./logistica/validarpendientes.asp" title="Validar Pendientes"> <img src="../images/control.png" alt="Validar Pendientes"></a>
			</div>
			</td>
		</tr>	
		<%End If %>		
		<%if Session("currentUser") = "admin" or Session("currentUser") = "logistica"  Then %>	
		<tr>
			<td>Check List Vehiculos</td>
			<td align="right">
			<div>
				<a href="./logistica/vehiculos_home.asp" title="Check Vehiculos"> <img src="../images/vehiculo.png" alt="Vehiculos"></a>
			</div>
			</td>
		</tr>	
		<%End If %>				
				 
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "logistica"  Then  %>
				<tr>
					<td>Links Reparto</td>
					<td align="right">
					<div>
						<a href="TransportistasHDR.asp" title="Ver accesos de reparto"> <img src="../images/vincular2.png" alt="Ver Accesos de Reparto HDR"></a>
					</div>
					</td>
				</tr>
		<%End If%>	
		<%if Session("currentUser") = "admin"  or Session("currentUser") = "logistica"  Then  %>
				<tr>
					<td>Ingreso y control de camiones</td>
					<td align="right">
					<div>
						<a href="./camiones/camiones_facturas.asp" title="Ingreso y control de mercaderia"> <img src="../images/despacho.png" alt="Ingreso y Control"></a>
					</div>
					</td>
				</tr>
		<%End If%>	
		
	</table>
<%End If%>

<% if   Session("currentUser") = "admin" or Session("currentUser") = "logistica"  Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>Sensores<br/></td> 
		</tr> 	
		<%if Session("currentUser") = "admin" or Session("currentUser") = "logistica"  Then %>	
		<tr>
			<td>Control Temperatura</td>
			<td align="right">
			<div>
				<a href="./sensores/sensores_dashboard.asp" title="Sensores"> <img src="../images/temp.png" alt="Sensores"></a>
			</div>
			</td>
		</tr>	
		<%End If %>	 
		
	</table>
<%End If%>

<% if   Session("currentUser") = "rrhh" or Session("currentUser") = "admin"  Then%>		
	<table border="0" align="center" cellpadding="0" cellspacing="0" class="contentTable table-width-lg" style="width: 100%;margin-bottom: 15px;">
		<tbody> 
		<tr>
			<td colspan="2"height="20">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tbody><tr class="tableHeader">
					<td colspan="2" align="left"><span class="formHeader">MENU</span></td>
					</tr>
				</tbody>
			</table>
			</td>
		</tr>
		</tbody>
		<tr align="center" class="columnTop">
			<td height="20" colspan="2"><b>RRHH<br/></td> 
		</tr> 	
		<%if Session("currentUser") = "admin" or Session("currentUser") = "rrhh"  Then %>	
		<tr>
			<td>Portal Empleados</td>
			<td align="right">
			<div>
				<a href="./rrhh/rrhh_recibos.asp" title="Recibos"> <img src="../images/control.png" alt="Sensores"></a>
			</div>
			</td>
		</tr>	
		<%End If %>	 
		
	</table>
<%End If%>



				
	<input type="button" class="btn1" value="Salir" onClick="fnsalir(this.form)" style="width:80px;">			
    </div>
	
	
	</form>
	 
	  

<script language="javascript">
function fnsalir(Fm){
	Fm.doWhat.value = -2; 
	Fm.submit(); 
}  

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


