<%@ Language=VBScript %> 
 
<%
'ON ERROR RESUME NEXT
'RESPONSE.BUFFER = TRUE
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_administracion.asp" --><%
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
periodoID = objRequest("periodoID")   
sSQL = objRequest("sSQL") 
nombrearchivo = objRequest("nombrearchivo")   


'Response.Write("chksincuit: "&chksincuit& "<br>"&VbCrLf)

'response.write "doWhat>" &  doWhat
'response.write "ExtractoRow>" &  extractorow


if doWhat="" or doWhat<="3" Then
	Session("FileUploaded") = ""
End If

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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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
 

	<form  action="../includes/app_excel_adm.asp" method="post" name="FEXCEL"> 
	<input type="hidden" name="doWhat" value="<%=doWhat%>">  
	<input type="hidden" name="periodoID" value="<%=periodoid%>"> 
 
	 
	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
	
	</form>
	

	<form name="FF" method="post" action="transferenciaspendientesV2.asp">
	<input type="hidden" name="doWhat" value="<%=doWhat%>"> 
	<input type="hidden" name="periodoID" value="<%=periodoid%>">  
	 
	<input type="hidden" name="sSQL" value="<%=sSQL%>">
	<input type="hidden" name="ReportName" value="<%=nombrearchivo%>">
	
 
	
    <div class="wrap" style="width: 70%" align="center">

 <%
	if doWhat="" or doWhat="0"then 
	%>
 
 	<h1>Transferencias Pendientes</h1>
   
	<!--#include file="./includes/TransferenciasPendientes.asp" --> 

	
 	<input type="button" class="btn1" value="Volver Menu" onClick="volvermenu(this.form)" style="width:80px;">
 
 <%
	End If
%>	 

<%
	if doWhat="1"then 
	' 
	%>
	<h1>Transferencias Pendientes </h1>
	<!--#include file="./includes/TransferenciasPendientesMostrar.asp" --> 
	<input type="button" class="btn1" value="Volver" onClick="chkForm(this.form,0)" style="width:80px;">
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
if(msg != "") alert(msg);
else{
	Fm.doWhat.value = prm;
	Fm.periodoID.value=param1;
	Fm.submit(); 
}
}

function chkFormXls(Fm,prm,param1,param2,param3){
	var msg  = "";	
	if(msg != "") alert(msg);
	else{
		Fm.doWhat.value = prm;
		Fm.periodoID.value=param1;
		Fm.sSQL.value=param2;
		Fm.ReportName.value=param3;
		Fm.submit(); 
	}
}

</script>	  

</body>
</HTML>
<%dbCon.Close
Set dbCon = Nothing%> 


