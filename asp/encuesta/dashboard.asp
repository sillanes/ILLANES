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
 
	
    <div class="wrap" style="width: 70%" align="center">
	<h1>Reclamos por mes</h1> 
	<!--#include file="./includes/REP_reclamos_por_mes.asp" -->	
	</div>
    <div class="wrap" style="width: 70%" align="center">
	<h1>Reclamos por dia</h1> 
	<!--#include file="./includes/REP_reclamos_por_dia.asp" -->	
	
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


