<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "application/json"
Response.Buffer = True
Response.Clear

' 🔑 CLAVE: fuerza punto decimal en ADO/VBScript
Session.LCID = 1033   ' English - United States

On Error Resume Next

' =========================
' FUNCIONES JSON SAFE
' =========================
Function JsonNum(v)
    If IsNull(v) Or Trim(CStr(v)) = "" Then
        JsonNum = "0"
    Else
        JsonNum = Replace(CStr(v), ",", ".")
    End If
End Function

Function JsonStr(s)
    If IsNull(s) Then s = ""
    s = CStr(s)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonStr = s
End Function

' =========================
' PARÁMETROS
' =========================
Dim transportistaid, fecha
transportistaid = Trim(Request("transportistaid"))
fecha = Trim(Request("fecha"))   ' yyyy-mm-dd

If transportistaid = "" Or fecha = "" Then
    Response.Write "{""error"":""Faltan parámetros""}"
    Response.End
End If

' =========================
' EJECUTAR STORED PROCEDURE
' =========================
Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandType = 4   ' adCmdStoredProc
cmd.CommandText = "tracking.usp_GPS_Puntos_Dia"

cmd.Parameters.Append cmd.CreateParameter("@TransportistaID", 3, 1, , CLng(transportistaid))
cmd.Parameters.Append cmd.CreateParameter("@Fecha", 200, 1, 10, fecha)


Set rs = cmd.Execute

If Err.Number <> 0 Then
    Response.Write "{""error"":""" & JsonStr(Err.Description) & """}"
    Response.End
End If

' =========================
' ARMAR JSON
' =========================
Dim first
first = True

Response.Write "["

Do While Not rs.EOF

    If Not first Then Response.Write ","
    first = False

    Response.Write "{"
    Response.Write """lat"":" & JsonNum(rs("latitud")) & ","
    Response.Write """lon"":" & JsonNum(rs("longitud")) & ","
    Response.Write """fecha"":""" & JsonStr(rs("fecha")) & ""","
    Response.Write """accuracy"":" & JsonNum(rs("accuracy")) & ","
    Response.Write """speed"":" & JsonNum(rs("speed")) & ","
    Response.Write """provider"":""" & JsonStr(rs("provider")) & ""","
    Response.Write """origen"":""" & JsonStr(rs("origen")) & """"
    Response.Write "}"

    rs.MoveNext
Loop

Response.Write "]"

rs.Close
Set rs = Nothing
Set cmd = Nothing

Response.End
%>
