<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "application/json"
Response.Buffer = True
Response.Clear



Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

' =========================
' Helpers JSON (ANTI-NULL)
' =========================
Function JsonNum(v)
    If IsNull(v) Then
        JsonNum = "null"
    Else
        JsonNum = Replace(CStr(v), ",", ".")
    End If
End Function

Function JsonStr(v)
    If IsNull(v) Then
        JsonStr = ""
        Exit Function
    End If

    v = CStr(v)
    v = Replace(v, "\", "\\")
    v = Replace(v, """", "\""")
    JsonStr = v
End Function

' =========================
' Parámetro
' =========================
Dim hojaderutaid
hojaderutaid = Trim(Request("hojaderutaid"))

If hojaderutaid = "" Then
    Response.Write "[]"
    Response.End
End If

' =========================
' Stored Procedure
' =========================
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "tracking.usp_GPS_Clientes_HojaRuta"

cmd.Parameters.Append cmd.CreateParameter("@HojaDeRutaID", 3, 1, , CLng(hojaderutaid))

Set rs = cmd.Execute

Dim first
first = True

Response.Write "["

Do While Not rs.EOF

    If Not first Then Response.Write ","
    first = False

    Response.Write "{"
    Response.Write """clienteid"":" & rs("ClienteID") & ","
    Response.Write """nombre"":""" & JsonStr(rs("Nombre")) & ""","
    Response.Write """lat"":" & JsonNum(rs("Latitud")) & ","
    Response.Write """lon"":" & JsonNum(rs("Longitud")) & ","
    Response.Write """direccion"":""" & JsonStr(rs("Direccion")) & ""","
    Response.Write """orden"":" & IIf(IsNull(rs("Orden")), "null", rs("Orden"))
    Response.Write "}"

    rs.MoveNext
Loop

Response.Write "]"

rs.Close
Set rs = Nothing
Set cmd = Nothing

Response.End
%>
