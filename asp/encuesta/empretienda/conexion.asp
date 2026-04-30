<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DRIVER=SQL Server;SERVER=192.168.200.13,3306;APP=Encuesta;DATABASE=administracion", "administracion_web", "Ill@nes.%2025+"
%>