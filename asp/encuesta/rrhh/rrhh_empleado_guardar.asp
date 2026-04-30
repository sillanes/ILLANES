<%@LANGUAGE="VBScript" CODEPAGE="65001"%>
<!--#include file="conexion.asp" -->
<!--#include file="sidebar.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Session.LCID = 11274

If Trim(Session("currentUser") & "") = "" Then
    Response.Redirect "login.asp"
End If

Function Nz(v, defaultValue)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function 
On Error Resume Next

Dim EmpleadoID, Legajo, CUIL, DNI, Apellido, Nombre, TelefonoWhatsApp, Email, Activo, ActivoNum
Dim cmd, rs
Dim redirectUrl, msg

EmpleadoID = 0
If IsNumeric(Request.Form("EmpleadoID")) Then
    EmpleadoID = CLng(Request.Form("EmpleadoID"))
End If

Legajo = Trim(Request.Form("Legajo") & "")
CUIL = Trim(Request.Form("CUIL") & "")
DNI = Trim(Request.Form("DNI") & "")
Apellido = Trim(Request.Form("Apellido") & "")
Nombre = Trim(Request.Form("Nombre") & "")
TelefonoWhatsApp = Trim(Request.Form("TelefonoWhatsApp") & "")
Email = Trim(Request.Form("Email") & "")
Activo = Trim(Request.Form("Activo") & "")

If EmpleadoID <= 0 Then
    Response.Redirect "rrhh_empleados.asp?ok=0&msg=" & Server.URLEncode("EmpleadoID inválido.")
End If

If Activo = "1" Then
    ActivoNum = 1
Else
    ActivoNum = 0
End If

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = Conn
cmd.CommandType = 4
cmd.CommandText = "rrhh.usp_Empleado_Upd"

cmd.Parameters.Append cmd.CreateParameter("@EmpleadoID", 3, 1, , EmpleadoID)
cmd.Parameters.Append cmd.CreateParameter("@Legajo", 200, 1, 50, Legajo)
cmd.Parameters.Append cmd.CreateParameter("@CUIL", 200, 1, 20, CUIL)
cmd.Parameters.Append cmd.CreateParameter("@DNI", 200, 1, 20, DNI)
cmd.Parameters.Append cmd.CreateParameter("@Apellido", 200, 1, 100, Apellido)
cmd.Parameters.Append cmd.CreateParameter("@Nombre", 200, 1, 100, Nombre)
cmd.Parameters.Append cmd.CreateParameter("@TelefonoWhatsApp", 200, 1, 50, TelefonoWhatsApp)
cmd.Parameters.Append cmd.CreateParameter("@Email", 200, 1, 150, Email)
cmd.Parameters.Append cmd.CreateParameter("@Activo", 11, 1, , ActivoNum)

Set rs = cmd.Execute

If Err.Number <> 0 Then
    msg = "Error al guardar: " & Err.Description
    Err.Clear

    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing

    Conn.Close
    Set Conn = Nothing

    Response.Redirect "rrhh_empleado_editar.asp?EmpleadoID=" & EmpleadoID & "&ok=0&msg=" & Server.URLEncode(msg)
End If

msg = "Empleado actualizado correctamente."

If Not rs Is Nothing Then
    If Not rs.EOF Then
        If Not IsNull(rs("Mensaje")) Then
            msg = CStr(rs("Mensaje"))
        End If
    End If
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

Set cmd = Nothing
Conn.Close
Set Conn = Nothing

Response.Redirect "rrhh_empleado_editar.asp?EmpleadoID=" & EmpleadoID & "&ok=1&msg=" & Server.URLEncode(msg)
%>