<%@ Language="VBScript" %>

<% Option Explicit %>
<!--#include file="conexion.asp" -->
<%
Session.Timeout = 20
 

%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Login Transportistas</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container mt-5">
    <div class="row justify-content-center">
        <div class="col-md-4">
            <div class="card shadow-lg p-4">
                <h4 class="text-center mb-4">Ingreso de Transportistas</h4>
                <form method="post" action="login.asp">
                    <div class="mb-3">
                        <label class="form-label">Patente</label>
                        <input type="text" name="patente" class="form-control" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Contraseña</label>
                        <input type="password" name="contrasena" class="form-control" required>
                    </div>
                    <%   
                    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
                        Dim patente, contrasena, cmd, rs, IPCliente
                        patente = Trim(Request.Form("patente"))
                        contrasena = Trim(Request.Form("contrasena"))
						IPCliente = Request.ServerVariables("REMOTE_ADDR")

                        Set cmd = Server.CreateObject("ADODB.Command")
                        Set cmd.ActiveConnection = conn
                        cmd.CommandText = "Reparto.dbo.usp_Login_ValidarTransportista"
                        cmd.CommandType = 4

                        cmd.Parameters.Append cmd.CreateParameter("@Patente", 200, 1, 50, patente)
                        cmd.Parameters.Append cmd.CreateParameter("@Contrasena", 200, 1, 50, contrasena)
						cmd.Parameters.Append cmd.CreateParameter("@IPCliente", 200, 1, 50, IPCliente)     ' adVarChar, adParamInput

                        Set rs = cmd.Execute
                        If Not rs.EOF Then
						
							IF rs("ID") >0 THEN 
								
								Session("TransportistaID") = rs("ID")
								Session("patente") = rs("Patente")
								Session("NombreTransportista") = rs("NombreTransportista")
								Session("bloqueado") = rs("Bloqueado")
								Response.Redirect "home.asp"
								
							ElseIF rs("ID") = -1 Then	
								Response.Write "<div class='alert alert-danger'>Usuario bloqueado por intentos fallidos</div>"
							Else
								Response.Write "<div class='alert alert-danger'>Patente o contraseña incorrecta.</div>"
							End If
                        Else
                            Response.Write "<div class='alert alert-danger'>Patente o contraseña incorrecta.</div>"
                        End If

                        rs.Close
                        Set rs = Nothing
                        Set cmd = Nothing
                    End If
                    %>
                    <div class="d-grid mt-3">
                        <button type="submit" class="btn btn-primary">Ingresar</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>
</body>
</html>