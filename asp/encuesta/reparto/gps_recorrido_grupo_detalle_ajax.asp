<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.ContentType = "application/json"
Response.CharSet = "UTF-8"
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


Dim grupo, hojaderutaid
grupo = CLng(Request("grupo"))
hojaderutaid = CLng(Request("hojaderutaid"))

If  hojaderutaid = 0 Then
    Response.Write "[]"
    Response.End
End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4
cmd.CommandText = "tracking.usp_GPS_Recorrido_Grupo_Detalle" 
cmd.Parameters.Append cmd.CreateParameter("@HojaDeRutaID", 3, 1, , hojaderutaid)
cmd.Parameters.Append cmd.CreateParameter("@GrupoID", 3, 1, , grupo)

Set rs = cmd.Execute

Dim first : first = True
Response.Write "["

Do Until rs.EOF

    If Not first Then Response.Write ","
    first = False

    Response.Write "{"
    Response.Write """lat"":" & Replace(rs("Latitud"), ",", ".") & ","
    Response.Write """lon"":" & Replace(rs("Longitud"), ",", ".") & ","
    Response.Write """fecha"":""" & rs("FechaGPS") & ""","
    Response.Write """speed"":" & JsonNum(rs("Velocidad")) & ","
    Response.Write """accuracy"":" & Replace(rs("Accuracy"), ",", ".")  
    Response.Write "}"

    rs.MoveNext
Loop

Response.Write "]"

rs.Close
Set rs = Nothing
Set cmd = Nothing
Response.End
%>
