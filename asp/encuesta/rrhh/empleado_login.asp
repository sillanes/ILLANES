<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

'========================================================
' Helpers
'========================================================
Function Nz(v, alt)
    If IsNull(v) Or IsEmpty(v) Or Trim(CStr(v & "")) = "" Then
        Nz = alt
    Else
        Nz = v
    End If
End Function

Function TieneMayuscula(txt)
    Dim i, ch
    TieneMayuscula = False
    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        If ch >= "A" And ch <= "Z" Then
            TieneMayuscula = True
            Exit Function
        End If
    Next
End Function

Function TieneMinuscula(txt)
    Dim i, ch
    TieneMinuscula = False
    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        If ch >= "a" And ch <= "z" Then
            TieneMinuscula = True
            Exit Function
        End If
    Next
End Function

Function TieneNumero(txt)
    Dim i, ch
    TieneNumero = False
    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        If ch >= "0" And ch <= "9" Then
            TieneNumero = True
            Exit Function
        End If
    Next
End Function

Function ValidarPasswordFuerte(pwd)
    ValidarPasswordFuerte = ""

    If Len(pwd) < 8 Then
        ValidarPasswordFuerte = "La nueva contraseña debe tener al menos 8 caracteres."
        Exit Function
    End If

    If Not TieneMayuscula(pwd) Then
        ValidarPasswordFuerte = "La nueva contraseña debe tener al menos una letra mayúscula."
        Exit Function
    End If

    If Not TieneMinuscula(pwd) Then
        ValidarPasswordFuerte = "La nueva contraseña debe tener al menos una letra minúscula."
        Exit Function
    End If

    If Not TieneNumero(pwd) Then
        ValidarPasswordFuerte = "La nueva contraseña debe tener al menos un número."
        Exit Function
    End If
End Function

Sub LimpiarLoginTemporal()
    Session.Contents.Remove("Tmp_UsuarioID")
    Session.Contents.Remove("Tmp_EmpleadoID")
    Session.Contents.Remove("Tmp_Usuario")
    Session.Contents.Remove("Tmp_RolID")
    Session.Contents.Remove("Tmp_Legajo")
    Session.Contents.Remove("Tmp_Nombre")
    Session.Contents.Remove("Tmp_DebeCambiarClave")
End Sub

Sub CrearSesionDefinitiva()
    Session("Empleado_UsuarioID")        = Session("Tmp_UsuarioID")
    Session("Empleado_EmpleadoID")       = Session("Tmp_EmpleadoID")
    Session("Empleado_Usuario")          = Session("Tmp_Usuario")
    Session("Empleado_RolID")            = Session("Tmp_RolID")
    Session("Empleado_Legajo")           = Session("Tmp_Legajo")
    Session("Empleado_Nombre")           = Session("Tmp_Nombre")
    Session("Empleado_DebeCambiarClave") = False
End Sub

Sub EjecutarUltimoLogin(ByVal pUsuarioID)
    On Error Resume Next

    Dim cmdUpd
    Set cmdUpd = Nothing
    Set cmdUpd = Server.CreateObject("ADODB.Command")
    Set cmdUpd.ActiveConnection = conn
    cmdUpd.CommandType = 4
    cmdUpd.CommandText = "empleado.usp_Usuario_UltimoLogin_Upd"
    cmdUpd.Parameters.Append cmdUpd.CreateParameter("@UsuarioID", 3, 1, , CLng(pUsuarioID))
    cmdUpd.Execute , , 128

    Set cmdUpd = Nothing
    On Error GoTo 0
End Sub

'========================================================
' Variables
'========================================================
Dim msgError, msgInfo, usuario, clave, accion
Dim nuevaClave, repetirClave
Dim mostrarCambioClave

msgError = ""
msgInfo = ""
usuario = Trim(Request.Form("usuario"))
clave   = Trim(Request.Form("clave"))
accion  = LCase(Trim(Request.Form("accion")))

nuevaClave   = Trim(Request.Form("nueva_clave"))
repetirClave = Trim(Request.Form("repetir_clave"))

mostrarCambioClave = False

If CBool(Nz(Session("Tmp_DebeCambiarClave"), False)) Then
    mostrarCambioClave = True
End If

'========================================================
' LOGIN
'========================================================
If accion = "login" Then

    LimpiarLoginTemporal()

    If usuario = "" Or clave = "" Then
        msgError = "Ingresá usuario y contraseña."
    Else
        On Error Resume Next

        Dim cmdLogin, rsLogin, loginError
        Dim debeCambiarClave

        loginError = ""
        debeCambiarClave = False

        Set cmdLogin = Nothing
        Set rsLogin  = Nothing

        Set cmdLogin = Server.CreateObject("ADODB.Command")
        Set cmdLogin.ActiveConnection = conn
        cmdLogin.CommandType = 4
        cmdLogin.CommandText = "empleado.usp_Empleado_Login"

        cmdLogin.Parameters.Append cmdLogin.CreateParameter("@Usuario", 200, 1, 100, usuario)
        cmdLogin.Parameters.Append cmdLogin.CreateParameter("@PasswordPlano", 200, 1, 200, clave)

        Set rsLogin = cmdLogin.Execute

        If Err.Number <> 0 Then
            loginError = "Error al validar el acceso: " & Err.Description
            Err.Clear
        End If

        On Error GoTo 0

        If loginError <> "" Then
            msgError = loginError

        ElseIf Not IsObject(rsLogin) Then
            msgError = "No fue posible validar el usuario."

        ElseIf rsLogin.EOF Then
            msgError = "Usuario o contraseña incorrectos."

        Else
            If Not IsNull(rsLogin("DebeCambiarClave")) Then
                debeCambiarClave = CBool(rsLogin("DebeCambiarClave"))
            End If

            Session("Tmp_UsuarioID")        = CLng(rsLogin("UsuarioID"))
            Session("Tmp_EmpleadoID")       = CLng(rsLogin("EmpleadoID"))
            Session("Tmp_Usuario")          = CStr(Nz(rsLogin("Usuario"), ""))
            Session("Tmp_RolID")            = CLng(Nz(rsLogin("RolID"), 0))
            Session("Tmp_Legajo")           = CStr(Nz(rsLogin("Legajo"), ""))
            Session("Tmp_Nombre")           = CStr(Nz(rsLogin("ApellidoNombre"), ""))
            Session("Tmp_DebeCambiarClave") = debeCambiarClave

            If debeCambiarClave Then
                mostrarCambioClave = True
                msgInfo = "Por seguridad, debés cambiar tu contraseña antes de continuar."
            Else
                CrearSesionDefinitiva()
                EjecutarUltimoLogin Session("Tmp_UsuarioID")
                LimpiarLoginTemporal()
                Response.Redirect "empleado_recibos.asp"
            End If
        End If

        If IsObject(rsLogin) Then
            If rsLogin.State = 1 Then rsLogin.Close
            Set rsLogin = Nothing
        End If

        If IsObject(cmdLogin) Then
            Set cmdLogin = Nothing
        End If
    End If
End If

'========================================================
' CAMBIO DE CLAVE OBLIGATORIO
'========================================================
If accion = "cambiarclave" Then
    mostrarCambioClave = True

    If Nz(Session("Tmp_UsuarioID"), "") = "" Then
        msgError = "La sesión para cambio de clave expiró. Ingresá nuevamente."
        mostrarCambioClave = False
        LimpiarLoginTemporal()

    ElseIf nuevaClave = "" Or repetirClave = "" Then
        msgError = "Completá la nueva contraseña y su confirmación."

    ElseIf nuevaClave <> repetirClave Then
        msgError = "La confirmación de la contraseña no coincide."

    Else
        Dim errPwd
        errPwd = ValidarPasswordFuerte(nuevaClave)

        If errPwd <> "" Then
            msgError = errPwd
        Else
            On Error Resume Next

            Dim cmdCambio, rsCambio, cambioError, cambioOk, cambioMsg
            cambioError = ""
            cambioOk = False
            cambioMsg = ""

            Set cmdCambio = Nothing
            Set rsCambio  = Nothing

            Set cmdCambio = Server.CreateObject("ADODB.Command")
            Set cmdCambio.ActiveConnection = conn
            cmdCambio.CommandType = 4
            cmdCambio.CommandText = "empleado.usp_Usuario_Clave_Cambiar"

            cmdCambio.Parameters.Append cmdCambio.CreateParameter("@UsuarioID", 3, 1, , CLng(Session("Tmp_UsuarioID")))
            cmdCambio.Parameters.Append cmdCambio.CreateParameter("@PasswordNuevaPlano", 200, 1, 200, nuevaClave)

            Set rsCambio = cmdCambio.Execute

            If Err.Number <> 0 Then
                cambioError = "Error al cambiar la contraseña: " & Err.Description
                Err.Clear
            End If

            On Error GoTo 0

            If cambioError <> "" Then
                msgError = cambioError

            ElseIf IsObject(rsCambio) Then
                If Not rsCambio.EOF Then
                    cambioOk  = CBool(Nz(rsCambio("Ok"), False))
                    cambioMsg = CStr(Nz(rsCambio("Mensaje"), ""))
                End If

                If rsCambio.State = 1 Then rsCambio.Close
                Set rsCambio = Nothing
            Else
                cambioOk = True
            End If

            If cambioOk Then
                CrearSesionDefinitiva()
                EjecutarUltimoLogin Session("Tmp_UsuarioID")
                LimpiarLoginTemporal()
                Response.Redirect "empleado_recibos.asp"
            Else
                If cambioMsg = "" Then cambioMsg = "No fue posible actualizar la contraseña."
                msgError = cambioMsg
            End If

            If IsObject(cmdCambio) Then
                Set cmdCambio = Nothing
            End If
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Portal del Empleado - Login</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="estilos.css">
<style>
body{
    margin:0;
    font-family:Arial, Helvetica, sans-serif;
    background:#f4f6f9;
}
.login-wrap{
    min-height:100vh;
    display:flex;
    align-items:center;
    justify-content:center;
    padding:20px;
    box-sizing:border-box;
}
.login-card{
    width:100%;
    max-width:420px;
    background:#fff;
    border:1px solid #ddd;
    border-radius:12px;
    padding:24px;
    box-sizing:border-box;
    box-shadow:0 2px 10px rgba(0,0,0,.06);
}
.login-card h1{
    margin:0 0 8px 0;
    font-size:28px;
}
.login-card p{
    margin:0 0 20px 0;
    color:#666;
}
.campo{
    margin-bottom:14px;
}
.campo label{
    display:block;
    margin-bottom:6px;
    font-size:13px;
    color:#555;
}
.campo input{
    width:100%;
    padding:12px;
    border:1px solid #ccc;
    border-radius:8px;
    box-sizing:border-box;
    font-size:15px;
}
.campo input[readonly]{
    background:#f7f7f7;
}
.btn{
    width:100%;
    background:#2b7cff;
    color:#fff;
    border:none;
    padding:12px;
    border-radius:8px;
    cursor:pointer;
    font-size:15px;
}
.btn:hover{
    background:#1f68d8;
}
.alert-error{
    background:#f8d7da;
    color:#721c24;
    border-radius:8px;
    padding:12px 14px;
    margin-bottom:14px;
}
.alert-info{
    background:#e7f1ff;
    color:#0c5460;
    border-radius:8px;
    padding:12px 14px;
    margin-bottom:14px;
}
.small{
    font-size:12px;
    color:#666;
    margin-top:14px;
    line-height:1.5;
}
.requisitos{
    margin-top:12px;
    background:#f8f9fa;
    border:1px solid #e3e6ea;
    border-radius:8px;
    padding:10px 12px;
    font-size:12px;
    color:#555;
    line-height:1.5;
}
</style>
</head>
<body>
<div class="login-wrap">
    <div class="login-card">
        <h1>Portal del Empleado</h1>

        <% If mostrarCambioClave Then %>
            <p>Actualizá tu contraseña para poder ingresar.</p>
        <% Else %>
            <p>Ingresá con tu usuario y contraseña.</p>
        <% End If %>

        <% If msgError <> "" Then %>
            <div class="alert-error"><%=Server.HTMLEncode(msgError)%></div>
        <% End If %>

        <% If msgInfo <> "" Then %>
            <div class="alert-info"><%=Server.HTMLEncode(msgInfo)%></div>
        <% End If %>

        <% If Not mostrarCambioClave Then %>

            <form method="post" action="empleado_login.asp" autocomplete="off">
                <input type="hidden" name="accion" value="login">

                <div class="campo">
                    <label for="usuario">Usuario</label>
                    <input type="text" name="usuario" id="usuario" value="<%=Server.HTMLEncode(usuario)%>">
                </div>

                <div class="campo">
                    <label for="clave">Contraseña</label>
                    <input type="password" name="clave" id="clave">
                </div>

                <button type="submit" class="btn">Ingresar</button>
            </form>

            <div class="small">
                Si es tu primer ingreso, usá la contraseña inicial definida por RRHH.
            </div>

        <% Else %>

            <form method="post" action="empleado_login.asp" autocomplete="off">
                <input type="hidden" name="accion" value="cambiarclave">

                <div class="campo">
                    <label>Usuario</label>
                    <input type="text" value="<%=Server.HTMLEncode(CStr(Nz(Session("Tmp_Usuario"), "")))%>" readonly>
                </div>

                <div class="campo">
                    <label>Empleado</label>
                    <input type="text" value="<%=Server.HTMLEncode(CStr(Nz(Session("Tmp_Nombre"), "")))%>" readonly>
                </div>

                <div class="campo">
                    <label for="nueva_clave">Nueva contraseña</label>
                    <input type="password" name="nueva_clave" id="nueva_clave">
                </div>

                <div class="campo">
                    <label for="repetir_clave">Confirmar nueva contraseña</label>
                    <input type="password" name="repetir_clave" id="repetir_clave">
                </div>

                <button type="submit" class="btn">Guardar nueva contraseña</button>
            </form>

            <div class="requisitos">
                La contraseña debe tener al menos 8 caracteres, una letra mayúscula, una minúscula y un número.
            </div>

        <% End If %>
    </div>
</div>
</body>
</html>
<%
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
%>