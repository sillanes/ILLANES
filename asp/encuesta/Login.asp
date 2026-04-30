<% Option Explicit %>

<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

Session("fromRedirect") = 0

Dim username, password, submit
Dim message

message  = ""
username = Trim(Request.Form("username"))
password = Trim(Request.Form("password"))
submit   = Request.Form("submit")

If submit = "Salir" Then

    Session("currentUser") = ""

ElseIf submit = "Iniciar" Then

    If username = "" Then
        message = message & "El usuario es requerido.<br/>"
    End If

    If password = "" Then
        message = message & "La contraseña es requerida.<br/>"
    End If

    If message = "" Then

        If          (username <> "hsarano" or password <> "hsarano2025") _
            AND (username <> "admin" or password <> "ill4nes%2023") _
            AND (username <> "logistica" or password <> "illanes$2025") _
            AND (username <> "millanes" or password <> "illanes%2023") _
            AND (username <> "hojaderuta" or password <> "illanes%2023") _
            AND (username <> "pxp" or password <> "illanes%2024") _
            AND (username <> "controlador" or password <> "illanes$2024") _
            AND (username <> "rrhh" or password <> "RRHH*2024") _
            AND (username <> "1001" or password <> "Super*2024") _
            AND (username <> "bancos" or password <> "Illanes$2024") _
            AND (username <> "vsaffe" or password <> "Illanes$2024") _
            AND (username <> "AdministracionChau" or password <> "Illanes.2024") _
            AND (username <> "rillanes" or password <> "Rillanes.2024") _
            AND (username <> "lsosa" or password <> "Lilo.2024") _
            AND (username <> "fliporace" or password <> "Illanes$2025") _
            AND (username <> "sbosio" or password <> "Sbo.753") _
            AND (username <> "mperla" or password <> "Mpe.242") _
            AND (username <> "cromero" or password <> "Cro.548") _
            AND (username <> "ahenriquez" or password <> "Ahe.321") _
            AND (username <> "ccordero" or password <> "Cco.408") _
            AND (username <> "eillanes" or password <> "Nolasabes.2024") _
        Then
            message = "Usuario o contraseña incorrectos."
        Else
            Session("currentUser") = username
        End If

    End If

End If
%>

<!DOCTYPE html>
<html lang="es">
<head>

	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <title>ILLANES HNOS SRL</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />

    <link rel="stylesheet" type="text/css" href="../includes/style.css">
    <link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
    <link rel="stylesheet" type="text/css" href="../includes/calendar_cool.css" media="all" />

    <script type="text/javascript" src="../includes/calendar_cool.js"></script>
    <script type="text/javascript" src="../includes/copy.js"></script>

    <style>
        * {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            padding: 0;
            min-height: 100%;
            font-family: "Segoe UI", Arial, sans-serif;
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 45%, #2563eb 100%);
        }

        body {
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }

        .login-page {
            width: 100%;
            max-width: 420px;
        }

        .login-card {
            background: rgba(255, 255, 255, 0.96);
            border-radius: 22px;
            padding: 34px 32px;
            box-shadow: 0 24px 60px rgba(0, 0, 0, 0.28);
            border: 1px solid rgba(255, 255, 255, 0.35);
        }

        .login-logo {
            width: 74px;
            height: 74px;
            margin: 0 auto 18px auto;
            border-radius: 22px;
            background: linear-gradient(135deg, #2563eb, #0f172a);
            display: flex;
            align-items: center;
            justify-content: center;
            color: #ffffff;
            font-size: 32px;
            font-weight: 800;
            letter-spacing: -1px;
            box-shadow: 0 12px 28px rgba(37, 99, 235, 0.35);
        }

        .login-title {
            margin: 0;
            text-align: center;
            font-size: 26px;
            color: #111827;
            font-weight: 800;
        }

        .login-subtitle {
            text-align: center;
            margin: 8px 0 28px 0;
            color: #6b7280;
            font-size: 14px;
        }

        .form-group {
            margin-bottom: 18px;
        }

        .form-group label {
            display: block;
            margin-bottom: 7px;
            font-size: 14px;
            color: #374151;
            font-weight: 700;
        }

        .form-control {
            width: 100%;
            height: 48px;
            border: 1px solid #d1d5db;
            border-radius: 13px;
            padding: 0 14px;
            font-size: 15px;
            outline: none;
            color: #111827;
            background: #f9fafb;
            transition: all 0.2s ease;
        }

        .form-control:focus {
            border-color: #2563eb;
            background: #ffffff;
            box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.14);
        }

        .btn-login {
            width: 100%;
            height: 50px;
            border: none;
            border-radius: 14px;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: #ffffff;
            font-size: 16px;
            font-weight: 800;
            cursor: pointer;
            transition: all 0.2s ease;
            margin-top: 6px;
        }

        .btn-login:hover {
            transform: translateY(-1px);
            box-shadow: 0 12px 24px rgba(37, 99, 235, 0.34);
        }

        .btn-login:active {
            transform: translateY(0);
        }

        .login-message {
            margin-top: 18px;
            padding: 12px 14px;
            border-radius: 12px;
            background: #fee2e2;
            color: #991b1b;
            border: 1px solid #fecaca;
            font-size: 14px;
            line-height: 1.4;
            text-align: center;
            font-weight: 600;
        }

        .login-footer {
            margin-top: 24px;
            text-align: center;
            color: #e5e7eb;
            font-size: 13px;
        }

        @media (max-width: 480px) {
            body {
                padding: 16px;
            }

            .login-card {
                padding: 28px 22px;
                border-radius: 18px;
            }

            .login-title {
                font-size: 23px;
            }
        }
    </style>
</head>

<body>

<%
If Session("currentUser") = "" Then
%>

    <div class="login-page">
        <div class="login-card">

            <div class="login-logo">SG</div>

            <h1 class="login-title">Iniciar sesión</h1>
            <div class="login-subtitle">Acceso al sistema de Illanes Hnos SRL</div>

            <form method="POST" action="login.asp" autocomplete="off">

                <div class="form-group">
                    <label for="username">Usuario</label>
                    <input 
                        type="text" 
                        id="username"
                        name="username" 
                        class="form-control"
                        placeholder="Ingrese su usuario"
                        value="<%=Server.HTMLEncode(username)%>"
                    />
                </div>

                <div class="form-group">
                    <label for="password">Contraseña</label>
                    <input 
                        type="password" 
                        id="password"
                        name="password" 
                        class="form-control"
                        placeholder="Ingrese su contraseña"
                    />
                </div>

                <input type="submit" name="submit" value="Iniciar" class="btn-login" />

                <% If message <> "" Then %>
                    <div class="login-message">
                        <%= message %>
                    </div>
                <% End If %>

            </form>
        </div>

        <div class="login-footer">
            © <%=Year(Date())%> Illanes Hnos SRL
        </div>
    </div>

<%
ElseIf Session("currentUser") <> "" Then

    If Session("currentUser") = "controlador" Then

        Response.Redirect "../pxpcontrol.asp"

    ElseIf Session("currentUser") = "1001" Then

        Session("fromRedirect") = 1
        Session("username") = username
        Session("password") = password
        Session("submit") = submit

        Response.Redirect "../menusuper.asp"

    Else

        Response.Redirect "../menu.asp"

    End If

End If
%>

</body>
</html>