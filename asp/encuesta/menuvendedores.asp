<%@ Language=VBScript %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Session.CodePage  = 65001

Dim doWhat
doWhat = Trim(Request("doWhat"))

If doWhat = "-2" Then
    Session("currentUser") = ""
    Session("username") = ""
    Session("isMobile") = ""
    Session("fromRedirect") = 0
    Session("password") = ""
    Response.Redirect "./vendedores.asp"
End If

If Session("currentUser") = "" Then
    Response.Redirect "./vendedores.asp"
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <title>ILLANES HNOS SRL - Vendedores</title>
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
            color: #172033;
            background: #eef3f8;
        }

        body {
            min-height: 100vh;
        }

        .seller-shell {
            min-height: 100vh;
            padding: 28px 18px;
            background:
                linear-gradient(135deg, rgba(15, 23, 42, 0.92), rgba(37, 99, 235, 0.82)),
                url("../images/Logo Grupo Illanes.jpg") center center / cover no-repeat;
        }

        .seller-container {
            width: 100%;
            max-width: 980px;
            margin: 0 auto;
        }

        .seller-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 18px;
            margin-bottom: 22px;
            color: #ffffff;
        }

        .seller-brand {
            display: flex;
            align-items: center;
            gap: 14px;
        }

        .seller-logo {
            width: 64px;
            height: 64px;
            border-radius: 20px;
            background: rgba(255, 255, 255, 0.96);
            color: #1d4ed8;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 27px;
            font-weight: 900;
            box-shadow: 0 16px 34px rgba(0, 0, 0, 0.24);
        }

        .seller-title {
            margin: 0;
            font-size: 28px;
            line-height: 1.15;
            font-weight: 900;
        }

        .seller-subtitle {
            margin-top: 5px;
            font-size: 14px;
            color: #dbeafe;
            font-weight: 600;
        }

        .seller-user {
            background: rgba(255, 255, 255, 0.14);
            border: 1px solid rgba(255, 255, 255, 0.26);
            border-radius: 999px;
            padding: 9px 14px;
            color: #ffffff;
            font-weight: 800;
            white-space: nowrap;
        }

        .seller-panel {
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid rgba(255, 255, 255, 0.48);
            border-radius: 24px;
            padding: 24px;
            box-shadow: 0 24px 58px rgba(15, 23, 42, 0.28);
        }

        .seller-section-title {
            margin: 0 0 15px 0;
            color: #0f172a;
            font-size: 18px;
            font-weight: 900;
        }

        .seller-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 16px;
        }

        .seller-card {
            min-height: 150px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            gap: 18px;
            padding: 20px;
            border-radius: 18px;
            text-decoration: none;
            color: #101827;
            background: #f8fafc;
            border: 1px solid #e5e7eb;
            border-left: 7px solid #2563eb;
            transition: transform 0.18s ease, box-shadow 0.18s ease, background 0.18s ease;
        }

        .seller-card:hover {
            transform: translateY(-3px);
            background: #ffffff;
            color: #101827;
            text-decoration: none;
            box-shadow: 0 16px 32px rgba(15, 23, 42, 0.15);
        }

        .seller-card.vipo {
            border-left-color: #14b8a6;
        }

        .seller-card.alta {
            border-left-color: #f59e0b;
        }

        .seller-card.export {
            border-left-color: #16a34a;
        }

        .seller-card-top {
            display: flex;
            align-items: center;
            gap: 14px;
        }

        .seller-icon {
            width: 54px;
            height: 54px;
            min-width: 54px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 16px;
            background: #ffffff;
            box-shadow: inset 0 0 0 1px #e5e7eb;
        }

        .seller-icon img {
            max-width: 34px;
            max-height: 34px;
            object-fit: contain;
        }

        .seller-card-title {
            font-size: 20px;
            font-weight: 900;
            line-height: 1.2;
        }

        .seller-card-text {
            color: #5b6472;
            font-size: 14px;
            line-height: 1.45;
            font-weight: 600;
        }

        .seller-card-action {
            color: #1d4ed8;
            font-size: 14px;
            font-weight: 900;
        }

        .seller-actions {
            display: flex;
            justify-content: center;
            margin-top: 22px;
        }

        .seller-logout {
            min-width: 150px;
            border: none;
            border-radius: 14px;
            padding: 12px 18px;
            background: #dc2626;
            color: #ffffff;
            font-size: 15px;
            font-weight: 900;
            cursor: pointer;
            box-shadow: 0 10px 20px rgba(220, 38, 38, 0.22);
        }

        .seller-logout:hover {
            background: #b91c1c;
        }

        @media (max-width: 720px) {
            .seller-shell {
                padding: 20px 12px;
            }

            .seller-header {
                align-items: flex-start;
                flex-direction: column;
            }

            .seller-title {
                font-size: 23px;
            }

            .seller-user {
                white-space: normal;
            }

            .seller-panel {
                padding: 16px;
                border-radius: 18px;
            }

            .seller-grid {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
<form name="FF" method="post" action="menuvendedores.asp">
    <input type="hidden" name="doWhat" value="">

    <div class="seller-shell">
        <div class="seller-container">

            <div class="seller-header">
                <div class="seller-brand">
                    <div class="seller-logo">VE</div>
                    <div>
                        <h1 class="seller-title">Menu de vendedores</h1>
                        <div class="seller-subtitle">Acceso rapido a las tareas comerciales</div>
                    </div>
                </div>

                <div class="seller-user">
                    <%=Server.HTMLEncode(Session("currentUser"))%>
                    <% If Trim(Session("username")) <> "" Then %>
                        - <%=Server.HTMLEncode(Session("username"))%>
                    <% End If %>
                </div>
            </div>

            <div class="seller-panel">
                <% If Session("currentUser") = "1001" Or Session("currentUser") = "admin" Then %>
                    <h2 class="seller-section-title">VIPO</h2>

                    <div class="seller-grid">
                        <a class="seller-card vipo" href="./supervisorescontrol.asp" title="Controlar Fotos">
                            <div class="seller-card-top">
                                <div class="seller-icon">
                                    <img src="../images/reporte.png" alt="Controlar Fotos">
                                </div>
                                <div class="seller-card-title">Controlar Fotos</div>
                            </div>
                            <div class="seller-card-text">
                                Revision y control de fotos cargadas por vendedores.
                            </div>
                            <div class="seller-card-action">Ingresar</div>
                        </a>

                        <a class="seller-card export" href="./vendescargarestadoporperiodos.asp" title="Descargar por periodos">
                            <div class="seller-card-top">
                                <div class="seller-icon">
                                    <img src="../images/excel.png" alt="Descargar por periodos">
                                </div>
                                <div class="seller-card-title">Descargar por periodos</div>
                            </div>
                            <div class="seller-card-text">
                                Exportacion historica de datos por periodo.
                            </div>
                            <div class="seller-card-action">Ingresar</div>
                        </a>
                    </div>
                <% Else %>
                    <h2 class="seller-section-title">Operaciones</h2>

                    <div class="seller-grid">
                        <a class="seller-card vipo" href="./vendedorcarga.asp" title="VIPO">
                            <div class="seller-card-top">
                                <div class="seller-icon">
                                    <img src="../images/reporte.png" alt="VIPO">
                                </div>
                                <div class="seller-card-title">VIPO</div>
                            </div>
                            <div class="seller-card-text">
                                Carga de fotos y comprobantes para objetivos del vendedor.
                            </div>
                            <div class="seller-card-action">Ingresar</div>
                        </a>

                        <a class="seller-card alta" href="./altacliente.asp" title="Alta Clientes">
                            <div class="seller-card-top">
                                <div class="seller-icon">
                                    <img src="../images/alta.png" alt="Alta Clientes">
                                </div>
                                <div class="seller-card-title">Alta Clientes</div>
                            </div>
                            <div class="seller-card-text">
                                Registro de nuevos clientes para el circuito comercial.
                            </div>
                            <div class="seller-card-action">Ingresar</div>
                        </a>
                    </div>
                <% End If %>

                <div class="seller-actions">
                    <input type="button" class="seller-logout" value="Salir" onclick="fnsalir(this.form)">
                </div>
            </div>

        </div>
    </div>
</form>

<script type="text/javascript">
function fnsalir(Fm){
    Fm.doWhat.value = "-2";
    Fm.submit();
}
</script>
</body>
</html>
