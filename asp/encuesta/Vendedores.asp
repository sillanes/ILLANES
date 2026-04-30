<% Option Explicit %>

<%
	Dim dbCon, dbRS, cM, sSQL
	%><!--#include file="./includes/db_command_const.asp" --><%
	Server.ScriptTimeout = 300
	Set dbCon = Server.CreateObject("ADODB.Connection")
	%><!--#include file="./includes/db_con_open_ventas.asp" --><%
	dbCon.CommandTimeout = 0 

	Dim username, password, submit, nameuser,fromRedirect
	Dim message, isValid

	Dim ismobile,Regex,Match
	ismobile = Is_Mobile() 
	Session("isMobile") = ismobile
	
	fromRedirect = Session("fromRedirect")
	
	message = ""
	if fromRedirect = 1 Then
		username = Session("username")
		password = Session("password")
		submit = Session("submit")
	Else
		Session("username")=""
		Session("password")=""
		username = Request.Form("username")
		password = Request.Form("password")
		submit = Request.Form("submit")
	
	End IF
	'Response.write fromRedirect
	'Response.write "<br/>"
	'Response.write submit
	'Response.write "<br/>"
	'Response.write username
	'Response.write "<br/>"
	'Response.write password
	

	If submit = "Salir" Then
	
		Session("fromRedirect") = 0 
	End If
	
	If submit="Iniciar" or fromRedirect = 1 Then

		sSQL = "EXEC Vendedor_Login '" & username  & "','" & password & "'"
		Response.write(sSQL)
		Set dbRS=Server.CreateObject("ADODB.Recordset")
		dbRS.open sSQL, dbCon 
		IF ERR.NUMBER <> 0 or dbRS.EOF  THEN
			RESPONSE.CLEAR
			response.redirect "./error.asp"
		End If

		If  NOT dbRS.EOF  Then
			isValid = dbRS("isValid")
			nameuser = dbRS("Nombre")
		End if

	End If

	If submit = "Salir" Then
	' remove session
		Session("currentUser") = ""
		Session("isMobile") = ""
		Session("username") = "" 
		Session("password") = "" 
		Session("fromRedirect") = 0 
	ElseIf submit = "Iniciar" Then
		If username = "" Then
		message = message & "username is required <br/>"
		End If
		If password = "" Then
		message = message & "password is required <br/>"
		End If

		 
		If (isValid=0)Then
		  message = message & "username or password is wrong"
		Else
		  Session("username") = nameuser
		  Session("currentUser") = username
		  Session("isMobile") = ismobile
		End If
	  End If
	  
	 Function Is_Mobile()
	  Set Regex = New RegExp
	  With Regex
		.Pattern = "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm|ipad)"
		.IgnoreCase = True
		.Global = True
	  End With
	  Match = Regex.test(Request.ServerVariables("HTTP_USER_AGENT"))
	  If Match then
		Is_Mobile = True
	  Else
		Is_Mobile = False
	  End If
	End Function
	

  

	  
	  
%>
 
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


	<div style="overflow-x: auto;">
   
  <h1>Inicio</h1>

  <% If Session ("currentUser") = "" Then  %>
    <form method="POST" action="vendedores.asp">
      Usuario<br/>
      <input type="text" name="username" placeholder="Usuario"/> <br/>
      Contraseña<br/>
      <input type="password" name="password" placeholder="Contraseña"/> <br/>
      <input type="submit" name="submit" value="Iniciar" />
      <br/>
      <i><%= message %></i>
    </form>
  <% ElseIf Session("currentUser") <> "" Then
		if Session("currentUser") = "1001" Then
			Response.Redirect "../supervisorescontrol.asp"
		Else
			Response.Redirect "../vendedorcarga.asp"
		End IF 
	 End If 
%>
  
  </div>
</body>
</HTML>


  
