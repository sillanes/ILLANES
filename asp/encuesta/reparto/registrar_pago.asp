<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function


Dim hdrid, clienteid, totalCobrado, i, pago, cod, pagosXml, detallesXml
Dim facturaCount, itemCount, facturasXml, formaspagoXml
Dim facturaid, articulo, cantidad
Dim key, value

hdrid = Request.Form("hdrid")
clienteid = Request.Form("clienteid")
esCC = Request.Form("esCC") 

For Each key In Request.Form
  Response.Write key & " = " & Request.Form(key) & "<br/>"
Next


' Armar XML de formas de pago
formaspagoXml = "<Pagos>"
 
For i = 1 To 3
	pago =  replace( Request.Form("pago" & i),",","") 
	pago = IIf(pago = "", 0.00 , pago)
	cod = Request.Form("codigo" & i)  
	formaspagoXml = formaspagoXml & "<Pago>"
	formaspagoXml = formaspagoXml & "<FormaDePagoID>" & cod & "</FormaDePagoID>"
	formaspagoXml = formaspagoXml & "<Importe>" & pago & "</Importe>"
	formaspagoXml = formaspagoXml & "</Pago>"

Next
 
formaspagoXml = formaspagoXml & "</Pagos>"
 
 
' Armar XML de ítems de factura
detallesXml = "<Detalles>"
For Each key In Request.Form
	If Left(key, 5) = "item_" Then
		value = Request.Form(key)
		If IsNumeric(value) And value <> "" Then
			cantidad = CLng(value)
			  
			response.write "Key: " & key & "---- value: "& value 

			' Key: item_FACTURAID_ARTICULO
			Dim partes: partes = Split(key, "_")
			facturaid = partes(1)
			articulo = partes(2)

			detallesXml = detallesXml & "<Item>"
			detallesXml = detallesXml & "<FacturaID>" & facturaid & "</FacturaID>"
			detallesXml = detallesXml & "<Articulo>" & articulo & "</Articulo>"
			detallesXml = detallesXml & "<Cantidad>" & cantidad & "</Cantidad>"
			detallesXml = detallesXml & "</Item>"

		End If
	End If
Next
detallesXml = detallesXml & "</Detalles>"

response.write "detallesXml" & detallesXml

response.write "<br/>"

' Debug
'Response.Write "<hr><strong>DEBUG:</strong><br>"
'Response.Write "HDRID: " & hdrid & "<br>"
'Response.Write "ClienteID: " & clienteid & "<br>" 
 Response.Write "<pre>PagosXml: " & Server.HTMLEncode(formaspagoXml) & "</pre>"
 Response.Write "<pre>DetallesXml: " & Server.HTMLEncode(detallesXml) & "</pre><hr>"  

 

sSQL = "EXEC [cobranza].[Transportista_HojaDeRuta_RegistrarPago] "& hdrid &","& clienteid &",'"&  formaspagoXml &"','"&  detallesXml &"'" 
Response.write(sSQL)
Set dbRS=Server.CreateObject("ADODB.Recordset")
dbRS.open sSQL, conn
IsParentInactive = False
IF ERR.NUMBER <> 0 THEN
	RESPONSE.CLEAR
	Response.Redirect "error_pago.asp?mensaje="
End If

hasError= dbRS("hasError")
ErrorMessage= dbRS("ErrorMessage") 

Set dbRS = Nothing 
If hasError = 1 Then
    Dim mensajeError
    mensajeError = Server.URLEncode(ErrorMessage)
    Response.Redirect "error_pago.asp?mensaje=" & mensajeError
End If

Response.Redirect "hojaderuta.asp?hdrid=" & hdrid
%>
