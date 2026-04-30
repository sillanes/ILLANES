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

 		<BODY BGCOLOR=”#C0C0C0″>

<div class="wrap" align="center">
<%
 If ( IsEmpty(msgError) ) Then
    msgErrorview = "Lo sentimos, ah ocurrido un error inesperado" 
  Else  
    msgErrorview = msgError 
  End If
%>

<FONT FACE=”ARIAL”><img src="./images/errors.png" width="30" height="30"> <%=msgErrorview%><BR></FONT>
</div>

</BODY>

</HTML>
 