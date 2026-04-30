<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then Response.Redirect "../login.asp"

Dim CampaniaID
CampaniaID = Request.QueryString("CampaniaID")

If IsNumeric(CampaniaID) And CampaniaID <> "" Then
    On Error Resume Next
 
    conn.Execute "EXEC dbo.usp_Campania_Pendientes_del " & CampaniaID
    If Err.Number <> 0 Then
        Response.Redirect "campania_modificar.asp?error=1"
    Else
        Response.Redirect "campania_modificar.asp?success=1"
    End If
    On Error GoTo 0
Else
    Response.Redirect "campania_modificar.asp?error=2"
End If
%>
