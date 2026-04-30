<%@ Language=VBScript %>
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Session.CodePage  = 65001
%>

<!--#include file="includes/punto_venta.asp" -->
<!--#include file="includes/menu_v2_helpers.asp" -->
<!--#INCLUDE FILE="_upload.asp"-->
<%
'ON ERROR RESUME NEXT
Dim dbCon, dbRS, cM, sSQL, comboList, aPreguntas, aRespuestas
%><!--#include file="./includes/db_command_const.asp" --><%
Server.ScriptTimeout = 300
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_con_open_reclamos.asp" --><%
dbCon.CommandTimeout = 0

If Request.QueryString.Count>0 Then
    Set objRequest = Request.QueryString
Else
    Set objRequest = Request.Form
End If

doWhat = objRequest("doWhat")

If doWhat = "-2" Then
    submit_logout = "Salir"
	Session("currentUser") = ""
	' 1. Limpia todas las variables de la sesión actual
	Session.Contents.RemoveAll()

	' 2. Destruye la sesión actual
	Session.Abandon()
End If

If submit_logout = "Salir" or Session("currentUser") = "" Then
    Session("currentUser") = ""
	' 1. Limpia todas las variables de la sesión actual
	Session.Contents.RemoveAll()

	' 2. Destruye la sesión actual
	Session.Abandon()
    Response.Redirect "../login.asp"
	
End If

Dim NombreMenu
If Session("currentUser") = "illanes" OR Session("currentUser") = "admin" or Session("currentUser") = "hojaderuta" or Session("currentUser") = "pxp" or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" Then
    NombreMenu = "SISTEMA DE GESTIÓN DE LOGÍSTICA"
ElseIf Session("currentUser") = "admin" or Session("currentUser") = "1001" Then
    NombreMenu = "SISTEMA DE GESTIÓN DE VENTAS"
ElseIF Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion" or Session("currentUser") = "rillanes" or Session("currentUser") = "bancos" or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "sbosio" or Session("currentUser") = "mperla" or Session("currentUser") = "cromero" or Session("currentUser") = "ahenriquez" or Session("currentUser") = "ccordero" Then
    NombreMenu = "SISTEMA DE GESTIÓN ADMINISTRATIVO"
Else
    NombreMenu = "SISTEMA DE GESTIÓN DE LOGÍSTICA"
End IF

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
    <link rel="stylesheet" type="text/css" href="../includes/css/menu_v2.css">
    <link rel="stylesheet" type="text/css" href="../includes/calendar_cool.css" media="all" />
    <script type="text/javascript" src="../includes/calendar_cool.js"></script>
    <script type="text/javascript" src="../includes/copy.js"></script>
</head>
<body>
<form name="FF" method="post" action="menu.asp">
<input type="hidden" name="doWhat" value="<%=doWhat%>">

<div class="menu-container">

    <div class="menu-header">
        <h1><%=NombreMenu%></h1>
        <div class="menu-user">Bienvenidx: <span><%=UCASE(Session("currentUser"))%></span></div>
    </div>

    <div class="menu-toolbar">
        <input type="text" id="menuSearch" class="menu-search" placeholder="Buscar opción del menú..." onkeyup="filtrarMenu()">
    </div>

<% if Session("currentUser") = "admin" Then %>
    <% Call RenderSectionStart("Campanias") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "rrhh" Then %>
            <% Call RenderItem("Adminsitracion de Campanias", "./campanias/home.asp", "../images/config.png", "admin") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "millanes" or Session("currentUser") = "eillanes" or Session("currentUser") = "rrhh" Then %>
    <% Call RenderSectionStart("CHAT BOT") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" or Session("currentUser") = "rrhh" Then %>
            <% Call RenderItem("Mensajes", "./chatbot/mensajes_sofia.asp", "../images/chatbot.png", "chat") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "millanes" or Session("currentUser") = "rillanes" or Session("currentUser") = "eillanes" or Session("currentUser") = "logistica" or Session("currentUser") = "rrhh" Then %>
    <% Call RenderSectionStart("CAJAS NAVIDEÑAS") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" or Session("currentUser") = "rrhh" Then %>
            <% Call RenderItem("Gestión", "./empretienda/home.asp", "../images/config.png", "caja") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" Then %>
            <% Call RenderItem("Subir Datos Empretienda", "./subirordenes.asp", "../images/upload.png", "caja") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" Then %>
            <% Call RenderItem("Procesar Ordenes", "./ordenes.asp", "../images/transferencias.png", "caja") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "eillanes" or Session("currentUser") = "rillanes" Then %>
            <% Call RenderItem("Cobranza", "./imputarOrdenes.asp", "../images/cobranza2.png", "caja") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" or Session("currentUser") = "rillanes" Then %>
            <% Call RenderItem("Listados Ordenes pendientes", "./ordenespendientes.asp", "../images/reporte.png", "caja") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" or Session("currentUser") = "logistica" Then %>
            <% Call RenderItem("Logistica", "./menu.asp", "../images/control.png", "caja") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
    <% Call RenderSectionStart("Transportista") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Alta Vehiculo", "./repartovehiculos.asp", "../images/alta.png", "transporte") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Modificación Telefono Transportista", "Transportistas_Telefonos_Modificar.asp", "../images/Telefono.png", "transporte") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Vincular HDR con reparto", "TransportistasAsociarHDR.asp", "../images/vincular2.png", "transporte") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Nuevo Controlar HDR", "/administracion/controlhdr.asp", "../images/Control.png", "transporte") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
    <% Call RenderSectionStart("FACTURA ELECTRONICA") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Envio WhastApp", "./enviarwhatsapp.asp", "../images/whastapplogo2.png", "factura") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Reenvio WhastApp", "./reenviowhatsapp.asp", "../images/reenviar.png", "factura") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Adherir Cliente", "./facturaDigital.asp", "../images/alta.png", "factura") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Listado de mensajes con error o excluidos", "./facturadigitalpendientes.asp", "../images/reporte.png", "factura") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "illanes" OR Session("currentUser") = "hsarao" OR Session("currentUser") = "admin" or Session("currentUser") = "hojaderuta" or Session("currentUser") = "pxp" or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "sbosio" Then %>
    <% Call RenderSectionStart("RECLAMOS") %>
        <% if Session("currentUser") = "illanes" OR Session("currentUser") = "hsarano" OR Session("currentUser") = "admin" or Session("currentUser") = "fliporace" or Session("currentUser") = "millanes" Then %>
            <% Call RenderItem("Estadísticas", "./dashboardV3.asp", "../images/estad.png", "reclamos") %>
            <% Call RenderItem("Revisión de reclamos", "./reclamoscheck.asp", "../images/eye.png", "reclamos") %>
            <% Call RenderItem("Buscar reclamo cerrado", "./buscarreclamo.asp", "../images/view.png", "reclamos") %>
            <% Call RenderItem("Descargar Reclamos", "./descargareclamos.asp", "../images/excel.png", "reclamos") %>
        <% End If %>

        <% if Session("currentUser") = "hojaderuta" OR Session("currentUser") = "hsarano" OR Session("currentUser") = "admin" or Session("currentUser") = "lsosa" or Session("currentUser") = "millanes" or Session("currentUser") = "fliporace" or Session("currentUser") = "sbosio" Then %>
            <% Call RenderItem("Subir Hoja de Ruta", "./subirhojaderuta.asp", "../images/upload.png", "reclamos") %>
            <% Call RenderItem("Ver Hojas de Ruta", "./hojaderuta.asp", "../images/eye.png", "reclamos") %>
        <% End If %>

        <% if Session("currentUser") = "admin" Then %>
            <% Call RenderItem("Configuracion de Feriados", "./feriados.asp", "../images/calendar2_.gif", "reclamos") %>
        <% End If %>

        <% if Session("currentUser") = "admin" OR Session("currentUser") = "hsarano" or Session("currentUser") = "rrhh" or Session("currentUser") = "pxp" or Session("currentUser") = "illanes" or Session("currentUser") = "millanes" or Session("currentUser") = "fliporace" Then %>
            <% Call RenderItem("Cargar Reclamo", "./pxpcargarreclamo.asp", "../images/reporte.png", "reclamos") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "controlador" or Session("currentUser") = "rrhh" OR Session("currentUser") = "hsarano" or Session("currentUser") = "fliporace" Then %>
    <% Call RenderSectionStart("PEDIDO POR PEDIDO") %>
        <% Call RenderItem("Reporte de Eficiencia Tiempo Real", "../logistica/reporte_control.asp", "../images/control.png", "pxp") %>

        <% if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR Session("currentUser") = "hsarano" or Session("currentUser") = "fliporace" Then %>
            <% Call RenderItem("Subir Hoja De Ruta para pedido por pedido", "./subirhdrpxp.asp", "../images/upload.png", "pxp") %>
        <% End If %>

        <% Call RenderItem("Pendientes de control", "./pxppendientescontrol.asp", "../images/eye.png", "pxp") %>

        <% if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR Session("currentUser") = "hsarano" or Session("currentUser") = "fliporace" Then %>
            <% Call RenderItem("Reportes", "./pxpreportes.asp", "../images/estad.png", "pxp") %>
        <% End If %>

        <% Call RenderItem("Controlar Hojas de Ruta", "./pxpcontrol.asp", "../images/reporte.png", "pxp") %>

        <% if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR Session("currentUser") = "hsarano" or Session("currentUser") = "fliporace" Then %>
            <% Call RenderItem("Cerrar Hoja De Ruta", "./pxpcerrarhojaderuta.asp", "../images/candado.png", "pxp") %>
            <% Call RenderItem("Agregar o quitar Armador", "./pxpabmarmador.asp", "../images/eye.png", "pxp") %>
            <% Call RenderItem("Agregar o quitar Controlador", "./pxpabmcontrolador.asp", "../images/eye.png", "pxp") %>
            <% Call RenderItem("Historico Por Periodos", "./pxpdescargarperiodo.asp", "../images/excel.png", "pxp") %>
        <% End If %>

        <% Call RenderItem("Reclamos por Controlador", "./pxpdescargarreclamosxcontrolador.asp", "../images/excel.png", "pxp") %>

        <% if Session("currentUser") = "pxp" OR Session("currentUser") = "admin" or Session("currentUser") = "rrhh" OR Session("currentUser") = "hsarano" or Session("currentUser") = "fliporace" Then %>
            <% Call RenderItem("Alta Nomina", "./altanomina.asp", "../images/alta.png", "pxp") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "1001" Then %>
    <% Call RenderSectionStart("VIPO") %>
        <% Call RenderItem("Controlar Fotos", "./vendedores.asp", "../images/reporte.png", "vipo") %>
        <% Call RenderItem("Descargar por periodos", "./vendescargarestadoporperiodos.asp", "../images/excel.png", "vipo") %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion" or Session("currentUser") = "rillanes" or Session("currentUser") = "bancos" or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "sbosio" or Session("currentUser") = "ça" or Session("currentUser") = "cromero" or Session("currentUser") = "ahenriquez" or Session("currentUser") = "ccordero" or Session("currentUser") = "mperla" Then %>
    <% Call RenderSectionStart("Transferencias") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion" or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "lsosa" Then %>
            <% Call RenderItem("Subir Extractos Bancarios", "./subirextracto.asp", "../images/upload.png", "transferencias") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "Administracion" or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "lsosa" Then %>
            <% Call RenderItem("Procesar extractos", "./extractos.asp", "../images/eye.png", "transferencias") %>
        <% End If %>

        <% if Session("currentUser") = "rillanes" or Session("currentUser") = "admin" or Session("currentUser") = "bancos" or Session("currentUser") = "fliporace" or Session("currentUser") = "lsosa" or Session("currentUser") = "Administracion" or Session("currentUser") = "eillanes" or Session("currentUser") = "millanes" or Session("currentUser") = "vsaffe" or Session("currentUser") = "sbosio" or Session("currentUser") = "mperla" or Session("currentUser") = "cromero" or Session("currentUser") = "ahenriquez" or Session("currentUser") = "ccordero" Then %>
            <% Call RenderItem("Imputar transferencias", "./imputartransferencias.asp", "../images/transferencias.png", "transferencias") %>
        <% End If %>

        <% Call RenderItem("Listados transferencias pendientes", "./transferenciaspendientes.asp", "../images/reporte.png", "transferencias") %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "logistica" Then %>
    <% Call RenderSectionStart("Control Reparto") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "logistica" Then %>
            <% Call RenderItem("Pendientes", "./logistica/validarpendientes.asp", "../images/control.png", "logistica") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "logistica" Then %>
            <% Call RenderItem("Check List Vehiculos", "./logistica/vehiculos_home.asp", "../images/vehiculo.png", "logistica") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "logistica" Then %>
            <% Call RenderItem("Links Reparto", "TransportistasHDR.asp", "../images/vincular2.png", "logistica") %>
        <% End If %>

        <% if Session("currentUser") = "admin" or Session("currentUser") = "logistica" Then %>
            <% Call RenderItem("Ingreso y control de camiones", "./camiones/camiones_facturas.asp", "../images/despacho.png", "logistica") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "admin" or Session("currentUser") = "logistica" Then %>
    <% Call RenderSectionStart("Sensores") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "logistica" Then %>
            <% Call RenderItem("Control Temperatura", "./sensores/sensores_dashboard.asp", "../images/temp.png", "sensores") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

<% if Session("currentUser") = "rrhh" or Session("currentUser") = "admin" Then %>
    <% Call RenderSectionStart("RRHH") %>
        <% if Session("currentUser") = "admin" or Session("currentUser") = "rrhh" Then %>
            <% Call RenderItem("Portal Empleados", "./rrhh/rrhh_recibos.asp", "../images/control.png", "rrhh") %>
        <% End If %>
    <% Call RenderSectionEnd() %>
<% End If %>

    <div class="menu-actions">
        <input type="button" class="btn-salir-v2" value="Salir" onClick="fnsalir(this.form)">
    </div>

</div>
</form>

<script language="javascript">
function fnsalir(Fm){
    Fm.doWhat.value = -2;
    Fm.submit();
}

function filtrarMenu(){
    var input = document.getElementById('menuSearch');
    var filtro = (input.value || '').toLowerCase();
    var secciones = document.getElementsByClassName('menu-section');

    for (var s = 0; s < secciones.length; s++) {
        var section = secciones[s];
        var cards = section.getElementsByClassName('menu-item-card');
        var visible = 0;

        for (var i = 0; i < cards.length; i++) {
            var txt = cards[i].innerText || cards[i].textContent;
            if (txt.toLowerCase().indexOf(filtro) > -1) {
                cards[i].style.display = 'flex';
                visible++;
            } else {
                cards[i].style.display = 'none';
            }
        }

        section.style.display = (visible > 0 ? 'block' : 'none');
    }
}
</script>

</body>
</html>
<%
dbCon.Close
Set dbCon = Nothing
%>
