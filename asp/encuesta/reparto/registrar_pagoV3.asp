<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim hdrid        : hdrid        = Request.Form("hdrid")
Dim clienteid    : clienteid    = Request.Form("clienteid")
Dim esCC         : esCC         = 0
If Trim(Request.Form("esCC"))<>"" Then esCC = CInt(Request.Form("esCC"))

Dim cond         : cond         = Request.Form("cond")
Dim dni          : dni          = Trim(Request.Form("dni"))
Dim observaciones: observaciones = Trim(Request.Form("observaciones"))

Dim latitud      : latitud      = Trim(Request.Form("latitud"))
Dim longitud     : longitud     = Trim(Request.Form("longitud"))

Function IIf(bClause, sTrue, sFalse)
  If CBool(bClause) Then IIf = sTrue Else IIf = sFalse
End Function

Dim key, value

'==== DEBUG (activar si necesitás) ====
'For Each key In Request.Form
'  Response.Write key & " = " & Server.HTMLEncode(Request.Form(key)) & "<br/>"
'Next
'Response.End

'────────────────────────────────────────────
'  FACTURAS ABIERTAS (lazy-load)
'────────────────────────────────────────────
Dim dictAbiertas : Set dictAbiertas = CreateObject("Scripting.Dictionary")
For Each key In Request.Form
  If Left(key,16) = "factura_abierta_" Then
    Dim fidA
    fidA = Replace(key, "factura_abierta_", "")
    If fidA<>"" Then
      If Not dictAbiertas.Exists(fidA) Then dictAbiertas.Add fidA, True
    End If
  End If
Next

'────────────────────────────────────────────
'  FACTURAS ANULADAS (anula completa)
'────────────────────────────────────────────
Dim dictAnuladas : Set dictAnuladas = CreateObject("Scripting.Dictionary")
Dim facturasAnuladasXml
facturasAnuladasXml = "<FacturasAnuladas>"

For Each key In Request.Form
  If Left(key, 8) = "anulada_" Then
    Dim anuladaFacturaID
    anuladaFacturaID = Replace(key, "anulada_", "")
    If anuladaFacturaID <> "" Then
      If Not dictAnuladas.Exists(anuladaFacturaID) Then dictAnuladas.Add anuladaFacturaID, True
      facturasAnuladasXml = facturasAnuladasXml & "<FacturaID>" & anuladaFacturaID & "</FacturaID>"
    End If
  End If
Next

facturasAnuladasXml = facturasAnuladasXml & "</FacturasAnuladas>"

'────────────────────────────────────────────
'  XML de FORMAS DE PAGO
'────────────────────────────────────────────
Dim i, pago, cod, formaspagoXml
formaspagoXml = "<Pagos>"

For i = 1 To 3
  pago = Trim(Request.Form("pago" & i))
  pago = Replace(pago, ",", ".") ' ✅ no borrar decimales
  If pago = "" Then pago = "0"

  cod = Trim(Request.Form("codigo" & i))
  If cod = "" Then cod = CStr(i)

  formaspagoXml = formaspagoXml & "<Pago>"
  formaspagoXml = formaspagoXml & "<FormaDePagoID>" & cod & "</FormaDePagoID>"
  formaspagoXml = formaspagoXml & "<Importe>" & pago & "</Importe>"
  formaspagoXml = formaspagoXml & "</Pago>"
Next

formaspagoXml = formaspagoXml & "</Pagos>"

'────────────────────────────────────────────
'  XML de ÍTEMS (solo si factura fue ABIERTA y NO ANULADA)
'  clave item_... viene del ajax (lo respetamos como estaba)
'────────────────────────────────────────────
Dim detallesXml
detallesXml = "<Detalles>"

For Each key In Request.Form
  If Left(key, 5) = "item_" Then

    Dim partes, facturaid, articulo, cantidad
    partes = Split(key, "_")

    If UBound(partes) >= 2 Then
      facturaid = partes(1)
      articulo  = partes(2)

      ' ✅ SOLO SI: abierta y no anulada
      If dictAbiertas.Exists(facturaid) Then
        If Not dictAnuladas.Exists(facturaid) Then

          value = Request.Form(key)
          If IsNumeric(value) And Trim(value) <> "" Then
            cantidad = CLng(value)

            detallesXml = detallesXml & "<Item>"
            detallesXml = detallesXml & "<FacturaID>" & facturaid & "</FacturaID>"
            detallesXml = detallesXml & "<Articulo>" & articulo & "</Articulo>"
            detallesXml = detallesXml & "<Cantidad>" & cantidad & "</Cantidad>"
            detallesXml = detallesXml & "</Item>"
          End If

        End If
      End If

    End If
  End If
Next

detallesXml = detallesXml & "</Detalles>"

'────────────────────────────────────────────
'  Ejecutar SP
'────────────────────────────────────────────
Dim sSQL
sSQL = "EXEC [cobranza].[Transportista_HojaDeRuta_RegistrarPago] " & _
        hdrid & "," & _
        clienteid & ",'" & _
        formaspagoXml & "','" & _
        detallesXml & "','" & _
        Replace(cond,"'","''") & "'," & _
        IIf(dni<>"","'"&Replace(dni,"'","''")&"'","NULL") & ",'" & _
        Replace(observaciones,"'","''") & "', '" & _
        facturasAnuladasXml & "'," & _
        IIf(latitud<>"", Replace(latitud,",","."), "NULL") & "," & _
        IIf(longitud<>"", Replace(longitud,",","."), "NULL")

'Response.Write "<pre>" & Server.HTMLEncode(sSQL) & "</pre>"
'Response.End

Dim dbRS
Set dbRS = Server.CreateObject("ADODB.Recordset")
dbRS.Open sSQL, conn

If Err.Number <> 0 Then
  Response.Clear
  Response.Redirect "error_pago.asp?mensaje="
End If

Dim hasError, ErrorMessage
hasError = dbRS("hasError")
ErrorMessage = dbRS("ErrorMessage")

dbRS.Close
Set dbRS = Nothing

If hasError = 1 Then
  Response.Redirect "error_pago.asp?mensaje=" & Server.URLEncode(ErrorMessage)
End If

'========================================================
' VALIDACIÓN: solo pedir comprobante si Cheque o Transferencia > 0
'========================================================
Function ParseMoneyAR(ByVal v)
  Dim s
  s = Trim(CStr(v))
  If s = "" Then ParseMoneyAR = 0 : Exit Function

  s = Replace(s, "$", "")
  s = Replace(s, " ", "")

  ' AR: miles "." decimales ","
  s = Replace(s, ".", "")
  s = Replace(s, ",", ".")

  If IsNumeric(s) Then
    ParseMoneyAR = CDbl(s)
  Else
    ParseMoneyAR = 0
  End If
End Function

Dim redir, returnUrl
redir = Trim(Request.Form("redir_to_comprobantes"))
returnUrl = Trim(Request.Form("return_after_comprobante"))

Dim impCheque, impTransferencia
impCheque = ParseMoneyAR(Request.Form("pago2"))          ' Cheque/Echeq
impTransferencia = ParseMoneyAR(Request.Form("pago3"))   ' Transferencia

If redir = "1" Then

  If (impCheque > 0) Or (impTransferencia > 0) Then
    Response.Redirect "hojaderuta_comprobantes.asp?hdrid=" & Server.URLEncode(Request.Form("hdrid")) & _
                      "&clienteid=" & Server.URLEncode(Request.Form("clienteid")) & _
                      "&from=entregar" & _
                      "&return=" & Server.URLEncode(returnUrl)
    Response.End
  Else
    ' No corresponde pedir comprobante
    If returnUrl <> "" Then
      Response.Redirect returnUrl
    Else
      Response.Redirect "hojaderutav3.asp?hdrid=" & Server.URLEncode(Request.Form("hdrid"))
    End If
    Response.End
  End If

End If

Response.Redirect "hojaderutav3.asp?hdrid=" & hdrid
%>
