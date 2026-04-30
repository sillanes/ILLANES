<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<!--#include file="sensores_common.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("currentUser") = "" Then
  Response.Redirect "login.asp"
End If

' ===== Helpers locales =====
Function SqlStr(s)
  If IsNull(s) Then s = ""
  SqlStr = "'" & Replace(CStr(s),"'","''") & "'"
End Function

Function SqlNullIfEmpty(s)
  If Trim(CStr(s)) = "" Then
    SqlNullIfEmpty = "NULL"
  Else
    SqlNullIfEmpty = SqlStr(s)
  End If
End Function

Function SqlLongOrNull(v)
  If Trim(CStr(v)) = "" Then
    SqlLongOrNull = "NULL"
  ElseIf IsNumeric(v) Then
    SqlLongOrNull = CLng(v)
  Else
    SqlLongOrNull = "NULL"
  End If
End Function

Function SqlNum(v)
  ' número con punto decimal (soporta coma)
  Dim s
  s = Trim(CStr(v))
  s = Replace(s, ",", ".")
  If s = "" Then
    SqlNum = "NULL"
  Else
    SqlNum = s
  End If
End Function

Function DoExec(sqlText, ByRef outErr)
  outErr = ""
  On Error Resume Next
  conn.Execute sqlText
  If Err.Number <> 0 Then
    outErr = Err.Description & " | " & GetAdoErrors()
    Err.Clear
  End If
  On Error GoTo 0
End Function

' ===== Inputs =====
Dim patente, tempAbasto, tempCarga, observacion, repartidorid, ayudanteid
patente     = UCase(Trim(Request.Form("patente")))
tempAbasto  = Trim(Request.Form("temp_abasto"))
tempCarga   = Trim(Request.Form("temp_carga"))
observacion = Trim(Request.Form("observacion"))
repartidorid = Trim(Request.Form("repartidorid"))
ayudanteid  = Trim(Request.Form("ayudanteid"))

If patente = "" Then
  Response.Redirect "sensores_vehiculo_nuevo.asp?msg=" & Server.URLEncode("Falta patente.")
End If

If Trim(repartidorid) = "" Or Not IsNumeric(repartidorid) Then
  Response.Redirect "sensores_vehiculo_nuevo.asp?msg=" & Server.URLEncode("Seleccioná repartidor.")
End If

If tempAbasto = "" Or tempCarga = "" Then
  Response.Redirect "sensores_vehiculo_nuevo.asp?msg=" & Server.URLEncode("Faltan temperaturas.")
End If

Dim usuario
usuario = Nz(Session("currentUser"), Nz(Session("NombreTransportista"), ""))
If Trim(usuario) = "" Then usuario = "web"

' ===== Ejecutar 2 inserts (ABASTO y CARGA) =====
Dim err1, err2, sql1, sql2, ridSql, aidSql, obsBase

ridSql = SqlLongOrNull(repartidorid)
aidSql = SqlLongOrNull(ayudanteid)
obsBase = observacion

' --- 1) ABASTO ---
' Intento firma extendida: (Patente, Temperatura, Usuario, Observacion, TipoTemp, RepartidorID, AyudanteID)
sql1 = "EXEC dbo.usp_Sensores_Vehiculo_Ins " & _
       SqlStr(patente) & ", " & _
       SqlNum(tempAbasto) & ", " & _
       SqlStr(usuario) & ", " & _
       SqlNullIfEmpty(obsBase) & ", " & _
       SqlStr("ABASTO") & ", " & _
       ridSql & ", " & _
       aidSql

DoExec sql1, err1

' Fallback firma vieja: (Patente, Temperatura, Usuario, Observacion)
If err1 <> "" Then
  err1 = "" ' limpio y reintento
  sql1 = "EXEC dbo.usp_Sensores_Vehiculo_Ins " & _
         SqlStr(patente) & ", " & _
         SqlNum(tempAbasto) & ", " & _
         SqlStr(usuario) & ", " & _
         SqlNullIfEmpty(Trim(obsBase & " [ABASTO] RepartidorID=" & repartidorid & IIf(Trim(ayudanteid)<>""," AyudanteID=" & ayudanteid,"")))
  DoExec sql1, err1
End If

' --- 2) CARGA ---
sql2 = "EXEC dbo.usp_Sensores_Vehiculo_Ins " & _
       SqlStr(patente) & ", " & _
       SqlNum(tempCarga) & ", " & _
       SqlStr(usuario) & ", " & _
       SqlNullIfEmpty(obsBase) & ", " & _
       SqlStr("CARGA") & ", " & _
       ridSql & ", " & _
       aidSql

DoExec sql2, err2

If err2 <> "" Then
  err2 = ""
  sql2 = "EXEC dbo.usp_Sensores_Vehiculo_Ins " & _
         SqlStr(patente) & ", " & _
         SqlNum(tempCarga) & ", " & _
         SqlStr(usuario) & ", " & _
         SqlNullIfEmpty(Trim(obsBase & " [CARGA] RepartidorID=" & repartidorid & IIf(Trim(ayudanteid)<>""," AyudanteID=" & ayudanteid,"")))
  DoExec sql2, err2
End If

' ===== Resultado =====
If err1 <> "" Or err2 <> "" Then
  Dim msgErr
  msgErr = "Error guardando lecturas. " & _
           IIf(err1<>"","ABASTO: " & err1 & " ","") & _
           IIf(err2<>"","CARGA: " & err2,"")
  Response.Write "<html><head><meta charset='UTF-8'></head><body style='font-family:Arial;'>"
  Response.Write "<h3>Error</h3><div style='color:#b00020;'>" & Server.HTMLEncode(msgErr) & "</div>"
  Response.Write "<br><a href='sensores_vehiculo_nuevo.asp'>Volver</a>"
  Response.Write "</body></html>"
  Response.End
End If

Response.Redirect "sensores_vehiculo_nuevo.asp?msg=" & Server.URLEncode("Lecturas guardadas OK (Abasto y Carga).")
%>