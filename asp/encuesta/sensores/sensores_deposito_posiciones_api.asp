<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

' Si querés obligar sesión también en API:
If Session("currentUser") = "" Then
  Response.Status = "401 Unauthorized"
  Response.ContentType = "application/json"
  Response.Write "{""ok"":false,""error"":""No autorizado""}"
  Response.End
End If

Response.ContentType = "application/json"

Function JsonEscape(s)
  If IsNull(s) Then s = ""
  s = CStr(s)
  s = Replace(s, "\", "\\")
  s = Replace(s, """", "\""")
  s = Replace(s, vbCrLf, "\n")
  s = Replace(s, vbCr, "\n")
  s = Replace(s, vbLf, "\n")
  JsonEscape = s
End Function

Dim depositoid
depositoid = Trim(Request.QueryString("depositoid"))
If depositoid = "" Then depositoid = "0"

If Not IsNumeric(depositoid) Or CLng(depositoid) <= 0 Then
  Response.Write "{""ok"":false,""error"":""depositoid inválido""}"
  Response.End
End If

On Error Resume Next

Dim cmd, rs, json, first
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = 4 'adCmdStoredProc
cmd.CommandText = "dbo.usp_Sensores_Deposito_Posiciones_Sel"
cmd.Parameters.Append cmd.CreateParameter("@DepositoID", 3, 1, , CLng(depositoid)) ' adInteger input

Set rs = cmd.Execute

If Err.Number <> 0 Then
  json = "{""ok"":false,""error"":""" & JsonEscape(Err.Description) & """}"
  Err.Clear
  Response.Write json
  Response.End
End If

json = "{""ok"":true,""items"":["
first = True

Do While Not rs.EOF
  Dim p
  p = UCase(Trim(CStr(rs("Posicion"))))
  If first Then
    first = False
  Else
    json = json & ","
  End If
  json = json & "{""posicion"":""" & JsonEscape(p) & """}"
  rs.MoveNext
Loop

json = json & "]}"

Response.Write json

On Error Resume Next
If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
Set rs = Nothing
Set cmd = Nothing
On Error GoTo 0
%>
