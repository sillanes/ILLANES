<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else
        IIf = sFalse
    End If
End Function

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Function SafeSql(v)
    SafeSql = Replace(Trim("" & v), "'", "''")
End Function

Dim cabeceraID, camionID, itemID, cantidadControlada, observacion, completo
Dim usuarioActual, sql

cabeceraID         = CLng(0 & Request.Form("CabeceraID"))
camionID           = CLng(0 & Request.Form("CamionID"))
itemID             = CLng(0 & Request.Form("ItemID"))
cantidadControlada = Trim("" & Request.Form("CantidadControlada"))
observacion        = Trim("" & Request.Form("Observacion"))
completo           = 0

If Trim("" & Request.Form("Completo")) = "1" Then
    completo = 1
End If

usuarioActual = Trim("" & Session("currentUser"))
If usuarioActual = "" Then usuarioActual = "sistema"

If cabeceraID <= 0 Or itemID <= 0 Then
    If camionID > 0 Then
        Response.Redirect "camiones_facturas.asp?CamionID=" & camionID & "&err=" & Server.URLEncode("Parametros invalidos.")
    Else
        Response.Redirect "camiones_facturas.asp?err=" & Server.URLEncode("Parametros invalidos.")
    End If
End If

If cantidadControlada = "" And completo = 0 Then
    Response.Redirect "camiones_factura_control.asp?CabeceraID=" & cabeceraID & "&CamionID=" & camionID & "&err=" & Server.URLEncode("Debe ingresar una cantidad o marcar completo.")
End If

sql = "EXEC dbo.usp_Camiones_Cabecera_Item_Control_Ins " & _
      "@ItemID=" & itemID & ", " & _
      "@CantidadControlada=" & IIf(cantidadControlada = "", "NULL", "'" & SafeSql(cantidadControlada) & "'") & ", " & _
      "@Completo=" & completo & ", " & _
      "@Observacion=" & IIf(observacion = "", "NULL", "'" & SafeSql(observacion) & "'") & ", " & _
      "@Usuario='" & SafeSql(usuarioActual) & "'"

On Error Resume Next
conn.Execute sql

If Err.Number <> 0 Then
    Response.Redirect "camiones_factura_control.asp?CabeceraID=" & cabeceraID & "&CamionID=" & camionID & "&err=" & Server.URLEncode("Error al guardar el item: " & Err.Description)
Else
    Response.Redirect "camiones_factura_control.asp?CabeceraID=" & cabeceraID & "&CamionID=" & camionID & "&msg=" & Server.URLEncode("Item guardado correctamente.")
End If
%>