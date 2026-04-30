<!-- conexion.asp -->
<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB;Data Source=SERVIDOR;Initial Catalog=WhatsAppAPI;User ID=USUARIO;Password=CLAVE;"
%>
