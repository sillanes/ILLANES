<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 
 

Dim BlackList, ErrorPage
BlackList = Array("/*", "*/", "@@",_
                  "cursor", _
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

 
Set objRequest = Request.Form
doWhat = objRequest("doWhat") 
PeriodoID = CLng("0" & objRequest("PeriodoID") )
EmpleadoID = CLng("0" & objRequest("EmpleadoID") )
 
 
'submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If
 
If doWhat = "-1"  Then 
	response.redirect "../menu.asp"
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
    <style>
        li {
            cursor: pointer;
			margin: 15px 0;
        }
    </style>
 
  
</HEAD>
<body>

	
 	
	
    <div class="wrap" style="width: 70%" align="center">
	<h1>DESCARGA RECLAMOS POR CONTROLADOR</h1>
 <%
	if doWhat="" or doWhat>="0"then 
	%>
 
	<!--#include file="./includes/pxpPeriodosComboReclamos.asp" -->


	 
 <%
	End If
%> 
		 
<%	 
	if  doWhat="" or  doWhat>="0" then 
%>


	<!--#include file="./includes/pxpControladoresReclamos.asp" -->
	<!--#include file="./includes/pxpControladoresReclamosChart.asp" -->
 
<script>
window.onload = function () {

var chart = new CanvasJS.Chart("chartContainer", {
	theme: "dark2",
	exportFileName: "Doughnut Chart",
	exportEnabled: true,
	animationEnabled: true,
	title:{
		text: "Motivos Reclamos"
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
			{ y: <%=FALTANTE_DE_MERCADERIA%>, name: "FALTANTE DE MERCADERIA" },
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
 
	<br/>
	<br/>
	
	<span  class="pagetitle">
	<H1>Total Reclamos: <%=TotalRC%></h1> 
	</span>
	<div id="chartContainer" style="height: 370px; max-width: 920px; margin: 0px auto;"></div>
	<script src="/includes/canvasjs.min.js"></script>	 

	
	 
 <%
	End If	 
%>	

	<BR/>
	
	<form name="FF2" method="post" action="pxpdescargarreclamosxcontrolador.asp">

	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>"> 	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
	</form>
	 </div>	
	
<script language="javascript">

function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
}


function chkForm(Fm,prm){
	var msg  = "";
	var txt = "";
	var reg = /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/;
  
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.submit(); 
	}
}

 
</script>	  

</body>
</HTML>
<%

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

dbCon.Close
Set dbCon = Nothing%> 


