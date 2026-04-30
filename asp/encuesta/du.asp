<%
Set Upload = Server.CreateObject("Persits.Upload")
Upload.CodePage = 65001
Upload.OverwriteFiles = False
Temp = "C:\Reclamos\Archivos\"
Upload.Save(Temp)
response.write("llego")
%>