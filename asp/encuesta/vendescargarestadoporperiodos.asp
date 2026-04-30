<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_ventas.asp" --><%
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

If doWhat = "-2"  Then 
	submit_logout = "Salir"
	Session("currentUser") = "" 
End If
 
'submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../menusuper.asp"
End If
 
If doWhat = "-1"  Then 
	response.redirect "../menu.asp"
End If 
  
SupervisorID = Session("currentUser")
SupervisorNombre = Session("username")
 

Set objRequest = Nothing
%>
<HTML>
<HEAD>
<TITLE>ILLANES HNOS SRL</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="../includes/style.css">  
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
<link rel="stylesheet" type="text/css" href="../includes/css/bars.css">
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

	
 	
	
    <div style="overflow-x: auto;" align="center">
	<h1>DESCARGA ESTADO POR VENDEDOR</h1>
 <%
	if doWhat="" or doWhat="0"then 
	%> 
	<!--#include file="./includes/venPeriodosComboObjetivos.asp" -->


	 
 <%
	End If
%> 
 <%
	if  doWhat="" or doWhat>="0"  then 
%>	 	
 	<form name="FF2" method="post" action="vendescargarestadoporperiodos.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">  
	<input type="hidden" name="SupervisorID" value="<%=SupervisorID%>">

	<!--#include file="./includes/Supervisores_Vendedores_ListWithExport.asp" --> 
	<div>
		<br /> 
		<input type="button" class="btn1" value="Menu" onClick="volvermenu(this.form)" style="width:80px;">
		<input type="button" class="btn1" value="Salir" onClick="fnsalir(this.form)" style="width:80px;">
	</div>
	
	</Form>
<%
	End If
%>	 
 
	</div>
	
<script language="javascript">
function fnsalir(Fm){
	Fm.doWhat.value = -2; 
	Fm.submit(); 
}  
function ShowComments(x) {
  alert(x);
}
function volvermenu(Fm){
	Fm.doWhat.value = -1; 
	Fm.submit(); 
}

function chkForm2(Fm,prm,param1,param2){
	var msg  = ""; 
	if(prm == 1){ 
		Fm.doWhat.value=prm;
		Fm.VendedorID.value=param1;
		Fm.VendedorNombre.value=param2;
	}	
	if(prm == 2){ 
		Fm.doWhat.value=prm;
		Fm.ClienteID.value=param1;
		Fm.ClienteNombre.value=param2;
	}		
	if(msg != "") alert(msg);
	else{  
		Fm.submit(); 
	}
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


