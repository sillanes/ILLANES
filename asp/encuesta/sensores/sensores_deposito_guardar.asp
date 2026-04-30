<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sensores_common.asp" -->
<%
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
  Response.Redirect "sensores_deposito_nuevo.asp"
End If

' =========================
' SafeLong: convierte a Long sin romper
' - soporta "1", "1 ", "1.0", "1,0"
' - si no puede, devuelve 0
' =========================
Function SafeLong(v)
  On Error Resume Next

  If IsNull(v) Or IsEmpty(v) Then
    SafeLong = 0
    Exit Function
  End If

  Dim s : s = Trim(CStr(v))
  If s = "" Then
    SafeLong = 0
    Exit Function
  End If

  ' normalizar decimales si vino "1,0" o "1.0"
  s = Replace(s, ",", ".")
  ' si trae letras u otros símbolos, no es válido
  If Not IsNumeric(s) Then
    SafeLong = 0
    Exit Function
  End If

  ' pasar por CDbl y luego CLng para evitar "no coinciden los tipos"
  Dim n : n = CDbl(s)
  If Err.Number <> 0 Then
    Err.Clear
    SafeLong = 0
    Exit Function
  End If

  SafeLong = CLng(Fix(n))
  If Err.Number <> 0 Then
    Err.Clear
    SafeLong = 0
    Exit Function
  End If

  On Error GoTo 0
End Function

' =========================
' Parse decimal robusto (acepta 7.8 / 7,8 / 1.234,56 / 1,234.56)
' Devuelve Variant Double o Null si no se puede parsear
' =========================
Function ParseDecimalAny(s)
  On Error Resume Next
  Dim v, hasDot, hasComma

  v = Trim(CStr(s))
  If v = "" Then
    ParseDecimalAny = Null
    Exit Function
  End If

  ' dejar solo dígitos, signo y separadores
  Dim i, ch, out
  out = ""
  For i = 1 To Len(v)
    ch = Mid(v,i,1)
    If (ch >= "0" And ch <= "9") Or ch="-" Or ch="." Or ch="," Then
      out = out & ch
    End If
  Next
  v = out

  hasDot = (InStr(v,".") > 0)
  hasComma = (InStr(v,",") > 0)

  ' "." y "," => el último separador es decimal
  If hasDot And hasComma Then
    If InStrRev(v,",") > InStrRev(v,".") Then
      ' 1.234,56  => miles "." / decimal ","
      v = Replace(v, ".", "")
      ' queda 1234,56 (ok para CDbl ES)
    Else
      ' 1,234.56  => miles "," / decimal "."
      v = Replace(v, ",", "")
      v = Replace(v, ".", ",")
    End If
  ElseIf hasDot And Not hasComma Then
    ' 7.8 => pasar a 7,8 para CDbl ES
    v = Replace(v, ".", ",")
  Else
    ' solo coma => ok
  End If

  Dim n : n = CDbl(v)
  If Err.Number <> 0 Then
    Err.Clear
    ParseDecimalAny = Null
    Exit Function
  End If
  On Error GoTo 0

  ParseDecimalAny = n
End Function

' =========================
' Inputs
' =========================
Dim depositoid, posicion, temperaturaStr, observacion, usuario, responsableid
depositoid      = Trim(Request.Form("depositoid"))
posicion        = UCase(Trim(Request.Form("posicion")))
temperaturaStr  = Trim(Request.Form("temperatura"))
observacion     = Trim(Request.Form("observacion"))
responsableid   = Trim(Request.Form("responsableid"))
usuario         = Nz(Session("NombreTransportista"), Nz(Session("currentUser"), ""))

Dim errMsg
errMsg = ""

' Validaciones mínimas
Dim depIdLng : depIdLng = SafeLong(depositoid)
Dim respIdLng : respIdLng = SafeLong(responsableid)

If depIdLng <= 0 Then errMsg = "Seleccioná un depósito."
If errMsg = "" And posicion = "" Then errMsg = "Seleccioná una posición."
If errMsg = "" And temperaturaStr = "" Then errMsg = "Ingresá la temperatura."
If errMsg = "" And respIdLng <= 0 Then errMsg = "Seleccioná un responsable."

Dim temperaturaVal
temperaturaVal = Null
If errMsg = "" Then
  temperaturaVal = ParseDecimalAny(temperaturaStr)
  If IsNull(temperaturaVal) Then
    errMsg = "Temperatura inválida. Usá por ejemplo -18.5 o -18,5"
  End If
End If

If errMsg <> "" Then
  Response.Redirect "sensores_deposito_nuevo.asp?depositoid=" & Server.URLEncode(Nz(depositoid,"0")) & "&msg=" & Server.URLEncode(errMsg)
End If

' =========================
' Guardar
' =========================
On Error Resume Next

Dim cmdIns, rsOut
Set cmdIns = Server.CreateObject("ADODB.Command")
Set cmdIns.ActiveConnection = conn
cmdIns.CommandType = 4 ' adCmdStoredProc
cmdIns.CommandText = "dbo.usp_Sensores_Deposito_Ins"

cmdIns.Parameters.Append cmdIns.CreateParameter("@DepositoID", 3, 1, , depIdLng) ' adInteger
cmdIns.Parameters.Append cmdIns.CreateParameter("@Posicion",   200, 1, 20, posicion) ' adVarChar

' DECIMAL(6,2) => adNumeric = 131
Dim pTemp
Set pTemp = cmdIns.CreateParameter("@Temperatura", 131, 1) ' adNumeric input
pTemp.Precision = 6
pTemp.NumericScale = 2
pTemp.Value = CDbl(temperaturaVal)
cmdIns.Parameters.Append pTemp
Set pTemp = Nothing

cmdIns.Parameters.Append cmdIns.CreateParameter("@Usuario",     200, 1, 100, usuario) ' adVarChar
cmdIns.Parameters.Append cmdIns.CreateParameter("@Observacion", 200, 1, 300, observacion) ' adVarChar

' NUEVO: ResponsableID
cmdIns.Parameters.Append cmdIns.CreateParameter("@ResponsableID", 3, 1, , respIdLng) ' adInteger

Set rsOut = cmdIns.Execute

If Err.Number <> 0 Then
  errMsg = "Error guardando lectura: " & Err.Description & " | " & GetAdoErrors()
  Err.Clear
End If

On Error GoTo 0

Set rsOut = Nothing
Set cmdIns = Nothing

If errMsg <> "" Then
  Response.Redirect "sensores_deposito_nuevo.asp?depositoid=" & Server.URLEncode(CStr(depIdLng)) & "&msg=" & Server.URLEncode(errMsg)
Else
  Response.Redirect "sensores_deposito_nuevo.asp?depositoid=" & Server.URLEncode(CStr(depIdLng)) & "&msg=" & Server.URLEncode("Lectura guardada OK.")
End If
%>