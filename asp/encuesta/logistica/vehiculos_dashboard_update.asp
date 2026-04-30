<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

Dim checklistID, item
checklistID = Request.QueryString("ChecklistID")
item = Request.QueryString("Item")

If checklistID = "" Or item = "" Then
  Response.Write "Faltan parámetros"
  Response.End
End If

Set cmd = Server.CreateObject("ADODB.Command")
With cmd
  .ActiveConnection = conn
  .CommandType = 4
  .CommandText = "usp_Vehiculos_Checklist_CorregirItem"
  .Parameters.Append .CreateParameter("@ChecklistID", 3, 1, , CLng(checklistID))
  .Parameters.Append .CreateParameter("@Item", 202, 1, 200, item)
  .Execute
End With

Response.Write "Item marcado como corregido correctamente."
%>
