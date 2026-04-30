<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"


' ──────────────────────────────────────
'  ❱  Variables básicas
' ──────────────────────────────────────
Dim hdrid        : hdrid        = Request.Form("hdrid")
Dim clienteid    : clienteid    = Request.Form("clienteid")
Dim esCC         : esCC         = CInt(Request.Form("esCC"))
Dim cond 		 : cond 		= Request.Form("cond") 
Dim dni          : dni          = Trim(Request.Form("dni"))     ' ← NUEVO
Dim totalCobrado : totalCobrado = 0
Dim observaciones : observaciones = Trim(Request.Form("observaciones"))
' 👉 NUEVO: GEOLOCALIZACIÓN
Dim latitud      : latitud      = Trim(Request.Form("latitud"))
Dim longitud     : longitud     = Trim(Request.Form("longitud"))

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function


Dim  i, pago, cod, pagosXml, detallesXml
Dim facturaCount, itemCount, facturasXml, formaspagoXml
Dim facturaid, articulo, cantidad
Dim key, value

cond = Request.Form("cond") 

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


' ──────────────────────────────────────
'  XML de FACTURAS ANULADAS
' ──────────────────────────────────────
Dim facturasAnuladasXml
facturasAnuladasXml = "<FacturasAnuladas>"

For Each key In Request.Form
    If Left(key, 8) = "anulada_" Then
        Dim anuladaFacturaID
        anuladaFacturaID = Replace(key, "anulada_", "")
        
        If IsNumeric(anuladaFacturaID) Then
            facturasAnuladasXml = facturasAnuladasXml & _
                "<FacturaID>" & anuladaFacturaID & "</FacturaID>"
        End If
    End If
Next

facturasAnuladasXml = facturasAnuladasXml & "</FacturasAnuladas>"


' Debug
'Response.Write "<hr><strong>DEBUG:</strong><br>"
Response.Write "HDRID: " & hdrid & "<br>"
Response.Write "Cond: " & cond & "<br>"
Response.Write "ClienteID: " & clienteid & "<br>" 
Response.Write "<pre>PagosXml: " & Server.HTMLEncode(formaspagoXml) & "</pre>"
Response.Write "<pre>DetallesXml: " & Server.HTMLEncode(detallesXml) & "</pre><hr>"  
Response.Write "<pre>FacturasAnuladasXml: " & Server.HTMLEncode(facturasAnuladasXml) & "</pre>"

Dim sdni
sdni= IIf(dni<>"" , "'" & dni & "'" , "NULL")
 
'response.end

' sSQL = "EXEC [cobranza].[Transportista_HojaDeRuta_RegistrarPago] "& hdrid &","& clienteid &",'"&  formaspagoXml &"','"&  detallesXml &"','" & Replace(cond,"'","''") &"','" & DNI & "'" & ", '" & Replace(observaciones,"'","''") & "'"

sSQL = "EXEC [cobranza].[Transportista_HojaDeRuta_RegistrarPago] " & _
        hdrid & "," & _
        clienteid & ",'" & _
        formaspagoXml & "','" & _
        detallesXml & "','" & _
        Replace(cond,"'","''") & "'," & _
        IIf(dni<>"","'"&dni&"'","NULL") & ",'" & _
        Replace(observaciones,"'","''") & "', '" & _
        facturasAnuladasXml & "'," & _
        IIf(latitud<>"", Replace(latitud,",","."), "NULL") & "," & _
        IIf(longitud<>"", Replace(longitud,",","."), "NULL")
		
		
Response.Write "<pre>" & Server.HTMLEncode(sSQL) & "</pre>"
Response.End

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

Response.Redirect "hojaderutav3.asp?hdrid=" & hdrid
%>