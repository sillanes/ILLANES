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
                  "cursor","exec","execute",_
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
StartDate = Replace(objRequest("StartDate"),".","/")
 

'submit_logout = Request.Form("submit_logout")
If submit_logout = "Salir" or Session("currentUser") = "" Then
	Session("currentUser") = ""
	response.redirect "../login.asp"
End If
 
If doWhat = "-1"  Then 
	response.redirect "../menu.asp"
End If

If doWhat = "1"  Then 
	dbStartDate = year(StartDate)  & right("0"&month(StartDate),2) &  right("0"&day(StartDate),2)
	sSQL = "EXEC dbo.feriados_upsert '" & dbStartDate & "'," & doWhat
	'Response.write(sSQL)
	Set dbRS=Server.CreateObject("ADODB.Recordset")
	dbRS.open sSQL, dbCon
	
End If

If doWhat = "2"  Then 
	dbStartDate = objRequest("Fecha")
	sSQL = "EXEC dbo.feriados_upsert '" & dbStartDate & "'," & doWhat
	'Response.write(sSQL)
	Set dbRS=Server.CreateObject("ADODB.Recordset")
	dbRS.open sSQL, dbCon
	
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


	<form name="FF" method="post" action="feriados.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>">
	<input type="hidden" name="fecha" value="<%=iStartDate%>">
	
	
 
	
    <div class="wrap" style="width: 70%" align="center">
	<h1>Feriados</h1>
 <%
	if doWhat>=""then 
	%>
	
	<table width="0" border="0" align="center" cellpadding="0" cellspacing="1" class="tableForm table-width-lg">
	  <tr>		
		<td valign="top" id="SDate_show"><SPAN title="Fecha No laboral">Fecha no laboral</SPAN>:&nbsp;
		<input type="Text" name="StartDate" id="StartDate"maxlength="10" class="width-20" value="<%= StartDate%>" chktype="date|1">
		<img src="../images/calendar.gif" alt="Calendar" id="StartDate_img" >
		<input type="button" class="btn1" value="Agregar" onClick="chkForm(this.form, 1)" style="width:80px;">
		</td>
	  </tr>
	 </table>


	<!--#include file="./includes/getFeriados.asp" -->
 
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
	
	 
 <%
	End If
%>	
	

    </div>
	
	
	
	</Form>
	
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
 
function chkForm2(Fm,prm,param1){
	var msg  = "";
	if(prm == 3){ 
		Fm.fecha.value=param1;
	}		
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.fecha.value=param1;
		Fm.submit(); 
	}
}


if ((document.FF.elements['doWhat'].value) >= 0) {
	Calendar.setup({displayArea:"SDate_show", inputField:"StartDate", ifFormat:"dd/mm/y", button:"StartDate_img", singleClick:true});
}

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


