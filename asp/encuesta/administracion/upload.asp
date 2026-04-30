<HTML>
<BODY>

<%

Set Upload = Server.CreateObject("Persits.Upload.1")
Count=  Upload.Save("c:\Reclamos\Archivos\")

%>

<% = Count %> ficheros subidos.

</BODY>
</HTML>