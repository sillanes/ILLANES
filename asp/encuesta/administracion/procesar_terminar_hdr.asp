<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "login.asp"
End If

Dim hojaRutaID
hojaRutaID = Request.Form("hdrid")

If hojaRutaID = "" Or Not IsNumeric(hojaRutaID) Then
    Response.Write "<p style='color:red;'>ID de hoja de ruta inválido.</p>"
    Response.End
End If

' Captura de campos
Dim TotalHDR, Gastos, Efectivo, Cheques, Transferencia, DepositoBancario, Diferencia, Pendientes, Errores, Anuladas, PendientesJustificadas, Observaciones

TotalHDR = Replace(Request.Form("TotalHDR"), ",", "")
Gastos = Replace(Request.Form("Gastos"), ",", "")
Efectivo = Replace(Request.Form("Efectivo"), ",", "")
Cheques = Replace(Request.Form("Cheques"), ",", "")
Transferencia = Replace(Request.Form("Transferencia"), ",", "")
DepositoBancario = Replace(Request.Form("DepositoBancario"), ",", "")
Diferencia = Replace(Request.Form("Diferencia"), ",", "")
Pendientes = Request.Form("Pendientes")
Errores = Request.Form("Errores")
Anuladas = Request.Form("Anuladas")
PendientesJustificadas = Request.Form("PendientesJustificadas")
Observaciones = Replace(Request.Form("Observaciones"), "'", "''") ' Escapar comillas simples

' Validaciones básicas (puedes hacer más)
If Not IsNumeric(TotalHDR) Then TotalHDR = 0
If Not IsNumeric(Gastos) Then Gastos = 0
If Not IsNumeric(Efectivo) Then Efectivo = 0
If Not IsNumeric(Cheques) Then Cheques = 0
If Not IsNumeric(Transferencia) Then Transferencia = 0
If Not IsNumeric(DepositoBancario) Then DepositoBancario = 0
If Not IsNumeric(Diferencia) Then Diferencia = 0
If Not IsNumeric(Pendientes) Then Pendientes = 0
If Not IsNumeric(Errores) Then Errores = 0
If Not IsNumeric(Anuladas) Then Anuladas = 0
If Not IsNumeric(PendientesJustificadas) Then PendientesJustificadas = 0

' Ejecutar Stored Procedure
Dim sSQL, cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 1 ' adCmdText
sSQL = "EXEC usp_Transportista_HojaDeRuta_Terminar " & hojaRutaID & "," & TotalHDR & "," & Gastos & "," & Efectivo & "," & Cheques & "," & Transferencia & "," & DepositoBancario & "," & Diferencia & "," & Pendientes & "," & Errores & "," & Anuladas & "," & PendientesJustificadas & ",'" & Observaciones & "'"
cmd.CommandText = sSQL

On Error Resume Next
cmd.Execute
If Err.Number <> 0 Then
    Response.Write "<p style='color:red;'>Error al ejecutar stored procedure: " & Err.Description & "</p>"
    Response.End
End If
On Error GoTo 0

Set cmd = Nothing
conn.Close
Set conn = Nothing

Response.Redirect "controlhdr.asp?hdrid=" & hojaRutaID & "&msg=ok"
%>
