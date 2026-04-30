<%@ Language=VBScript %> 

<!--#INCLUDE FILE="_upload.asp"-->
<%
'ON ERROR RESUME NEXT
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_ventas.asp" --><%
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
	Session("username") = ""
	Session("password") = ""
	Session("fromRedirect") = 0
	response.redirect "../vendedores.asp"
End If 
  
    

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
	<form name="FF" method="post" action="menusuper.asp">  
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
    <div class="wrap" align="center">


<% if Session("currentUser") = "illanes"  OR Session("currentUser") = "admin" or Session("currentUser") = "supervisor" or Session("currentUser") = "pxp" Then%>

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
	<% if Session("currentUser") = "illanes"  OR Session("currentUser") = "admin" Then%>				
			<tr>
				<td>
					Estadísticas
				</td>
				<td align="right">
				<div>
					<a href="./dashboardV2.asp" title="Ver estadisticas"><img src="../images/estad.png" alt="Ver estadisticas"></a>
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
	<% if Session("currentUser") = "hojaderuta"  OR Session("currentUser") = "admin" Then%>		
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
	
	<% if Session("currentUser") = "admin" or Session("currentUser") = "rrhh" or Session("currentUser") = "pxp" or Session("currentUser") = "illanes" Then%>		
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

<% if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "controlador" or Session("currentUser") = "rrhh"   Then%>		
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
			<td>Subir Hoja De Ruta para pedido por pedido</td>
			<td align="right">
			<div>
				<a href="./subirhdrpxp.asp" title="hoja de ruta"> <img src="../images/upload.png" alt="Subir Hoja de Ruta"></a>
			</div>
			</td>
		</tr>		

		<tr>
			<td>Reportes</td>
			<td align="right">
			<div>
				<a href="./pxpreportes.asp" title="Estados"> <img src="../images/estad.png" alt="Ver avances"></a>
			</div>
			</td>
		</tr>	
	
		<tr>
			<td>Controlar Hojas de Ruta</td>
			<td align="right">
			<div>
				<a href="./pxpcontrol.asp" title="Controlar Pedidos"> <img src="../images/reporte.png" alt="Controlar pedidos"></a>
			</div>
			</td>
		</tr>		
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
		</tr>	
			<tr>
			<td>Reclamos por Controlador</td>
			<td align="right">
			<div>
				<a href="./pxpdescargarreclamosxcontrolador.asp" title="Relcamos por controlador"> <img src="../images/excel.png" alt="Descargar Datos"></a>
			</div>
			</td>
		</tr>		
		<tr>
		<td>Alta Nomina</td>
			<td align="right">
			<div>
				<a href="./altanomina.asp" title="Agregar"> <img src="../images/alta.png" alt="Agregar"></a>
			</div>
			</td>
		</tr>			
		
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


