<%
' ===== Sensores Common =====

Response.CodePage = 65001
Response.CharSet  = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

' =========================
' Helpers Null-Safe
' =========================
Function Nz(v, d)
  If IsNull(v) Or IsEmpty(v) Then
    Nz = d
    Exit Function
  End If

  Dim s
  On Error Resume Next
  s = CStr(v)
  If Err.Number <> 0 Then
    Err.Clear
    Nz = d
    Exit Function
  End If
  On Error GoTo 0

  If Trim(s) = "" Then
    Nz = d
  Else
    Nz = v
  End If
End Function

Function HtmlEncode(s)
  If IsNull(s) Or IsEmpty(s) Then
    HtmlEncode = ""
  Else
    HtmlEncode = Server.HTMLEncode(CStr(s))
  End If
End Function

Function FormatoTemp(v)
  If IsNull(v) Or IsEmpty(v) Then
    FormatoTemp = "-"
  Else
    On Error Resume Next
    Dim n : n = CDbl(v)
    If Err.Number <> 0 Then
      Err.Clear
      FormatoTemp = "-"
      Exit Function
    End If
    On Error GoTo 0
    FormatoTemp = Replace(FormatNumber(n, 2), ",", ".") & " &deg;C"
  End If
End Function

Function FormatoFechaHora(v)
  If IsNull(v) Or IsEmpty(v) Then
    FormatoFechaHora = "-"
  Else
    On Error Resume Next
    Dim dt : dt = CDate(v)
    If Err.Number <> 0 Then
      Err.Clear
      FormatoFechaHora = "-"
      Exit Function
    End If
    On Error GoTo 0
    FormatoFechaHora = Right("0"&Day(dt),2) & "/" & Right("0"&Month(dt),2) & "/" & Year(dt) & " " & _
                       Right("0"&Hour(dt),2) & ":" & Right("0"&Minute(dt),2)
  End If
End Function

Function PosBadgeClass(pos)
  Dim p
  If IsNull(pos) Or IsEmpty(pos) Then pos = ""
  p = UCase(Trim(CStr(pos)))
  Select Case p
    Case "ALTA":  PosBadgeClass = "bg-danger-subtle text-danger border border-danger-subtle"
    Case "MEDIA": PosBadgeClass = "bg-warning-subtle text-warning border border-warning-subtle"
    Case "BAJA":  PosBadgeClass = "bg-info-subtle text-info border border-info-subtle"
    Case Else:    PosBadgeClass = "bg-secondary-subtle text-secondary border border-secondary-subtle"
  End Select
End Function

Function GetAdoErrors()
  On Error Resume Next
  Dim s, i
  s = ""
  If Not conn Is Nothing Then
    If conn.Errors.Count > 0 Then
      For i = 0 To conn.Errors.Count - 1
        s = s & "ADO[" & i & "]: " & conn.Errors(i).Number & " - " & conn.Errors(i).Description & " | "
      Next
    End If
    conn.Errors.Clear
  End If
  GetAdoErrors = s
  On Error GoTo 0
End Function

Function SensoresExecRS(spName)
  ' Ejecuta EXEC dbo.<spName> y devuelve Recordset (o Nothing si falla)
  On Error Resume Next
  Dim rs, sql
  Set rs = Nothing
  sql = "EXEC " & spName
  Set rs = conn.Execute(sql)
  If Err.Number <> 0 Then
    Set rs = Nothing
    Err.Clear
  End If
  Set SensoresExecRS = rs
  On Error GoTo 0
End Function

Function SensoresExecRS_Params(cmdText, cmdObj)
  ' Placeholder por si más adelante querés usar ADODB.Command parametrizado.
End Function

Function IIf(bClause, sTrue, sFalse)
  If CBool(bClause) Then
    IIf = sTrue
  Else
    IIf = sFalse
  End If
End Function

Dim nombreHeader
nombreHeader = Nz(Session("NombreTransportista"), Nz(Session("currentUser"), ""))
%>