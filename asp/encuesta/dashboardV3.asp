<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 

Dim BlackList, ErrorPage
BlackList = Array("cursor","exec","execute",_
				  "CREATE","truncate","delete",_
                  "nchar", "varchar", "nvarchar", "iframe"_
                  )
'Note: We can include following keyword to make a stronger scan but it will also 
'protect users to input these words even those are valid input
'  "!", "char", "alter", "begin", "cast", "create",  
ErrorPage = "./error.asp?msg=" &  Server.URLEncode("Invalid Character Entered")
 

For Each s in Request.Form
	If ( CheckStringForSQL(Request.Form(s),"form") ) Then
		PrepareReport("Post Varibale")
		' Redirect to an error page
		Response.Redirect(ErrorPage)
	End If
Next
%><!--#include virtual="./includes/sql-check.asp"--><%

If Request.QueryString.Count>0 Then
	Set objRequest = Request.QueryString
Else
	Set objRequest = Request.Form
End If

' Set objRequest = Request.Form
doWhat = objRequest("doWhat")
doWhatPre = objRequest("doWhatPre")
idReclamo = replace(objRequest("idReclamo"),"'","")
txtresol  = objRequest("txtresolucion") 
subirarchivo = objRequest("subirarchivo")
fileupload = objRequest("cmd")

'submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If 
  
    
'ON ERROR RESUME NEXT
sSQL = "EXEC report.usp_Reclamos_Resolucion_Motivos_Rep" 
' Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, dbCon
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	response.redirect "./error.asp"
End If
 
FALTANTE_DE_MERCADERIA_ENVIADA = dbRS("FALTANTE_DE_MERCADERIA_ENVIADA")
FALTANTE_DE_MERCADERIA_SIN_STOCK = dbRS("FALTANTE_DE_MERCADERIA_SIN_STOCK")
ERROR_UNIDAD_DE_MEDIDA = dbRS("ERROR_UNIDAD_DE_MEDIDA")
MERCADERIA_EQUIVOCADA = dbRS("MERCADERIA_EQUIVOCADA")
MERCADERIA_DANADA = dbRS("MERCADERIA_DANADA")
DEVOLUCION_DE_MERCADERIA = dbRS("DEVOLUCION_DE_MERCADERIA")
OTRO = dbRS("OTRO")
' NO_DEFINIDO = dbRS("NO_DEFINIDO")
FECHA = dbRS("RANGO")
TotalRC = dbRS("TOTALRC")
TotalCL = dbRS("TOTALCL") 

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
<script>
window.onload = function () {

var chart = new CanvasJS.Chart("chartContainer", {
	theme: "dark2",
	exportFileName: "Doughnut Chart",
	exportEnabled: true,
	animationEnabled: true,
	title:{
		text: "Motivos Reclamos <%=FECHA%> "
	},
	legend:{
		cursor: "pointer",
		itemclick: explodePie
	},
	data: [{
		type: "doughnut",
		innerRadius: 90,
		showInLegend: true,
		toolTipContent: "<b>{name}</b>: #{y} (#percent%)",
		indexLabel: "{name} - #percent%",
		dataPoints: [
			{ y: <%=FALTANTE_DE_MERCADERIA_SIN_STOCK%>, name: "FALTANTE DE MERCADERIA SIN STOCK" },
			{ y: <%=FALTANTE_DE_MERCADERIA_ENVIADA%>, name: "FALTANTE DE MERCADERIA ENVIADA" },
			{ y: <%=ERROR_UNIDAD_DE_MEDIDA%>, name: "ERROR UNIDAD DE MEDIDA" },
			{ y: <%=MERCADERIA_EQUIVOCADA%>, name: "MERCADERIA EQUIVOCADA" },
			{ y: <%=MERCADERIA_DANADA%>, name: "MERCADERIA DAÑADA" },
			{ y: <%=DEVOLUCION_DE_MERCADERIA%>, name: "DEVOLUCION DE MERCADERIA" },
			{ y: <%=OTRO%>, name: "OTRO"},
			//{ y: <%=NO_DEFINIDO%>, name: "SIN DATOS" }
		]
	}]
});
chart.render();
 
function explodePie (e) {
	if(typeof (e.dataSeries.dataPoints[e.dataPointIndex].exploded) === "undefined" || !e.dataSeries.dataPoints[e.dataPointIndex].exploded) {
		e.dataSeries.dataPoints[e.dataPointIndex].exploded = true;
	} else {
		e.dataSeries.dataPoints[e.dataPointIndex].exploded = false;
	}
	e.chart.render();
}

}
</script>
 
  
</HEAD>
<body> 

	<form name="FF" method="post" action="menu.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
 
	
    <div class="wrap" style="width: 70%" align="center">
	<h1>Reclamos por mes</h1> 
	<!--#include file="./includes/REP_reclamos_por_mesV3.asp" -->	
	</div>
	   
	<div class="wrap" style="width: 70%" align="center">
	<h1>Reclamos Cerrados y Clasificados por Motivos: <%=Total%></h1> 
	<span  class="pagetitle"> 
	<div>Total Reclamos: <%=TotalRC%></div> 
	<div>Total Clasificados: <%=TotalCL%></div> 
	<div>% Clasificados: <%=ROUND(TotalCL*100/TotalRC,2)%></div> 
	</span>
	<div id="chartContainer" style="height: 370px; max-width: 920px; margin: 0px auto;"></div>
	<script src="/includes/canvasjs.min.js"></script>	 
	</div>	
    <div class="wrap" style="width: 70%" align="center">
	<h1>Reclamos por Motivos</h1> 
	<!--#include file="./includes/REP_reclamos_por_motivos.asp" -->	
	<input type="button" class="btn1" value="Volver" onClick="resetForm(this.form)">
	</div>	
    </form>
<script language="javascript">
  
function resetForm(Fm) { 
	Fm.submit(); 
}   
	 
 

</script>	  
	
</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


