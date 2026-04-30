<%@ Language="VBScript" %>
<%
Dim total
total = Request.TotalBytes

If total = 0 Then
  Response.Write "No llego nada"
  Response.End
End If

Dim binData
binData = Request.BinaryRead(total)

Dim fso, stm, ruta
ruta = "C:\Administracion\Reparto\Comprobantes\test.bin"

Set stm = Server.CreateObject("ADODB.Stream")
stm.Type = 1
stm.Open
stm.Write binData
stm.SaveToFile ruta, 2
stm.Close

Response.Write "OK"
%>
