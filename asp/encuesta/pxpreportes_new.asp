<%@ Language=VBScript %>
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Session.CodePage = 65001
%>

<%
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%

Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0 

Dim BlackList, ErrorPage
BlackList = Array("cursor","exec","execute","CREATE","truncate","delete","nchar","varchar","nvarchar","iframe")

ErrorPage = "./error.asp?msg=" & Server.URLEncode("Invalid Character Entered")

For Each s in Request.Form
    If (CheckStringForSQL(Request.Form(s),"form")) Then
        PrepareReport("Post Varibale")
        Response.Redirect(ErrorPage)
    End If
Next
%><!--#include virtual="./includes/sql-check.asp"--><%

If Request.QueryString.Count > 0 Then
    Set objRequest = Request.QueryString
Else
    Set objRequest = Request.Form
End If

doWhat = objRequest("doWhat")
doWhatPre = objRequest("doWhatPre")
idReclamo = Replace(objRequest("idReclamo"),"'","")
txtresol = objRequest("txtresolucion")
subirarchivo = objRequest("subirarchivo")
fileupload = objRequest("cmd")

If submit_logout = "Salir" Or Session("currentUser") = "" Then
    Session("currentUser") = ""
    Response.Redirect "../login.asp"
End If

Set objRequest = Nothing
%>

<!DOCTYPE html>
<html>
<head>
    <title>ILLANES HNOS SRL</title>

    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />

    <link rel="stylesheet" type="text/css" href="../includes/style.css">
    <link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
    <link rel="stylesheet" type="text/css" href="../includes/calendar_cool.css" media="all" />

    <script type="text/javascript" src="../includes/calendar_cool.js"></script>
    <script type="text/javascript" src="../includes/copy.js"></script>

    <style>
        body {
            background: #f4f6f9;
            margin: 0;
            padding: 0;
            font-family: Arial, Helvetica, sans-serif;
        }

        .page-container {
            width: 100%;
            max-width: 1200px;
            margin: 0 auto;
            padding: 18px;
            box-sizing: border-box;
        }

        .page-header {
            background: #ffffff;
            border-radius: 16px;
            padding: 18px 20px;
            margin-bottom: 18px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.08);
            text-align: center;
        }

        .page-header h1 {
            margin: 0;
            font-size: 24px;
            color: #102a53;
        }

        .reports-grid {
            display: grid;
            grid-template-columns: 1fr;
            gap: 18px;
        }

        .report-card {
            background: #ffffff;
            border-radius: 16px;
            padding: 18px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.08);
            box-sizing: border-box;
            overflow: hidden;
        }

        .report-title {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 14px;
            padding-bottom: 10px;
            border-bottom: 1px solid #e5e7eb;
        }

        .report-title-icon {
            width: 38px;
            height: 38px;
            border-radius: 12px;
            background: #eef4ff;
            color: #0d6efd;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            font-weight: bold;
        }

        .report-title h2 {
            margin: 0;
            font-size: 18px;
            color: #102a53;
        }

        .report-content {
            width: 100%;
            overflow-x: auto;
        }

        .report-content table {
            width: 100% !important;
            max-width: 100%;
            border-collapse: collapse;
        }

        .report-content th,
        .report-content td {
            padding: 8px;
            font-size: 13px;
            white-space: nowrap;
        }

        .actions {
            margin-top: 20px;
            text-align: center;
        }

        .btn-volver {
            border: none;
            border-radius: 12px;
            padding: 12px 22px;
            background: #0d6efd;
            color: #ffffff;
            font-weight: bold;
            cursor: pointer;
            min-width: 160px;
            transition: all .2s ease;
        }

        .btn-volver:hover {
            background: #0b5ed7;
            transform: translateY(-1px);
        }

        @media (min-width: 900px) {
            .reports-grid {
                grid-template-columns: 1fr 1fr;
            }
        }

        @media (max-width: 600px) {
            .page-container {
                padding: 10px;
            }

            .page-header {
                padding: 14px;
                border-radius: 12px;
            }

            .page-header h1 {
                font-size: 19px;
            }

            .report-card {
                padding: 12px;
                border-radius: 12px;
            }

            .report-title h2 {
                font-size: 16px;
            }

            .report-title-icon {
                width: 34px;
                height: 34px;
                font-size: 17px;
            }

            .report-content th,
            .report-content td {
                font-size: 12px;
                padding: 6px;
            }

            .btn-volver {
                width: 100%;
            }
        }
.rep-list {
    display: flex;
    flex-direction: column;
    gap: 10px;
}

.rep-row-card {
    background: #f9fafb;
    border: 1px solid #e5e7eb;
    border-left: 5px solid #0d6efd;
    border-radius: 14px;
    padding: 12px;
}

.rep-row-main {
    display: flex;
    gap: 12px;
    margin-bottom: 10px;
}

.rep-periodo {
    flex: 1;
}

.rep-metrics {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
    gap: 10px;
}

.rep-metric {
    background: #ffffff;
    border-radius: 12px;
    padding: 10px;
    border: 1px solid #e5e7eb;
}

.rep-label {
    display: block;
    font-size: 12px;
    color: #6b7280;
    margin-bottom: 4px;
}

.rep-metric strong,
.rep-periodo strong {
    font-size: 15px;
    color: #102a53;
}

.rep-metric.money {
    background: #f8fafc;
}

.rep-metric.total {
    border-color: #0d6efd;
}

.rep-empty {
    background: #fff3cd;
    color: #856404;
    padding: 14px;
    border-radius: 12px;
    text-align: center;
    font-weight: bold;
}

@media (max-width: 600px) {
    .rep-row-main {
        flex-direction: column;
        gap: 6px;
    }

    .rep-metrics {
        grid-template-columns: 1fr 1fr;
    }
}		
    </style>
</head>

<body>

<form name="FF" method="post" action="menu.asp">
    <input type="hidden" name="doWhat" value="<%=doWhat%>">

    <div class="page-container">

        <div class="page-header">
            <h1>Reportes Pedido por Pedido</h1>
        </div>

        <div class="reports-grid">

            <section class="report-card">
                <div class="report-title">
                    <div class="report-title-icon">A</div>
                    <h2>Reporte Total Armados por mes</h2>
                </div>
                <div class="report-content">
                    <!--#include file="./includes/REP_Armados_Por_Mes_New.asp" -->
                </div>
            </section>

            <section class="report-card">
                <div class="report-title">
                    <div class="report-title-icon">C</div>
                    <h2>Reporte Total Controlados por mes</h2>
                </div>
                <div class="report-content">
                    <!--#include file="./includes/REP_Controlados_Por_Mes.asp" -->
                </div>
            </section>

            <section class="report-card">
                <div class="report-title">
                    <div class="report-title-icon">AM</div>
                    <h2>Reporte Armadores por mes</h2>
                </div>
                <div class="report-content">
                    <!--#include file="./includes/REP_Armadores_Por_Mes.asp" -->
                </div>
            </section>

            <section class="report-card">
                <div class="report-title">
                    <div class="report-title-icon">CM</div>
                    <h2>Reporte Controladores por mes</h2>
                </div>
                <div class="report-content">
                    <!--#include file="./includes/REP_Controladores_Por_Mes.asp" -->
                </div>
            </section>

        </div>

        <div class="actions">
            <input type="button" class="btn-volver" value="Volver Menú" onclick="resetForm(this.form)">
        </div>

    </div>
</form>

<script language="javascript">
function resetForm(Fm) {
    Fm.submit();
}
</script>

</body>
</html>

<%
dbCon.Close
Set dbCon = Nothing
%>