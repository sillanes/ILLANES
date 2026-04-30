<%@ Language=VBScript %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Session.CodePage  = 65001

If Session("currentUser") = "" Then
    Response.Redirect "./vendedores.asp"
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <title>ILLANES HNOS SRL - Alta Clientes</title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link rel="stylesheet" type="text/css" href="../includes/style.css">
    <link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
    <style>
        * {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            padding: 0;
            min-height: 100%;
            font-family: "Segoe UI", Arial, sans-serif;
            background: #eef3f8;
            color: #172033;
        }

        .page-shell {
            min-height: 100vh;
            padding: 28px 18px;
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 48%, #2563eb 100%);
        }

        .page-container {
            width: 100%;
            max-width: 760px;
            margin: 0 auto;
        }

        .page-header {
            color: #ffffff;
            margin-bottom: 20px;
        }

        .page-title {
            margin: 0;
            font-size: 28px;
            font-weight: 900;
        }

        .page-subtitle {
            margin-top: 6px;
            color: #dbeafe;
            font-size: 14px;
            font-weight: 600;
        }

        .page-card {
            background: rgba(255, 255, 255, 0.96);
            border-radius: 22px;
            padding: 26px;
            box-shadow: 0 24px 58px rgba(15, 23, 42, 0.28);
            border: 1px solid rgba(255, 255, 255, 0.48);
        }

        .empty-title {
            margin: 0 0 8px 0;
            font-size: 20px;
            color: #0f172a;
            font-weight: 900;
        }

        .empty-text {
            margin: 0 0 22px 0;
            color: #5b6472;
            line-height: 1.5;
            font-size: 15px;
            font-weight: 600;
        }

        .btn-row {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }

        .btn-primary {
            display: inline-block;
            border: none;
            border-radius: 14px;
            padding: 12px 18px;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: #ffffff;
            font-size: 15px;
            font-weight: 900;
            text-decoration: none;
            cursor: pointer;
        }

        .btn-primary:hover {
            color: #ffffff;
            text-decoration: none;
            box-shadow: 0 12px 24px rgba(37, 99, 235, 0.28);
        }

        @media (max-width: 560px) {
            .page-shell {
                padding: 20px 12px;
            }

            .page-card {
                padding: 20px;
                border-radius: 18px;
            }

            .page-title {
                font-size: 23px;
            }
        }
    </style>
</head>
<body>
    <div class="page-shell">
        <div class="page-container">
            <div class="page-header">
                <h1 class="page-title">Alta Clientes</h1>
                <div class="page-subtitle">Vendedor: <%=Server.HTMLEncode(Session("currentUser"))%></div>
            </div>

            <div class="page-card">
                <h2 class="empty-title">Pantalla en preparacion</h2>
                <p class="empty-text">
                    Esta seccion queda creada para armar el alta de clientes en el siguiente paso.
                </p>

                <div class="btn-row">
                    <a class="btn-primary" href="./menuvendedores.asp">Volver al menu</a>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
