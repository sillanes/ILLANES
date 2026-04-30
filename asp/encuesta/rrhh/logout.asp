<%@ Language="VBScript" %>
<%
' 1. Limpia todas las variables de la sesión actual
Session.Contents.RemoveAll()

' 2. Destruye la sesión actual
Session.Abandon()

Response.Redirect "/login.asp"
%>