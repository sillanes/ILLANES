<%@ Language=VBScript %>
<%
' ==========================================
'  vendedorcarga.asp  — flujo completo
' ==========================================
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Server.ScriptTimeout = 600

' --- Seguridad básica de sesión ---
If Session("currentUser") = "" Then
    Response.Redirect "../vendedores.asp"
End If

Dim dbCon
Set dbCon = Server.CreateObject("ADODB.Connection")
%><!--#include file="./includes/db_command_const.asp" --><%
' Abre tu conexión (usa tu include existente)
%><!--#include file="./includes/db_con_open_ventas.asp" --><%

 
' -----------------------------
'  Lee parámetros / sesión
' -----------------------------
Dim doWhat, PeriodoID, ClienteID, clientenombre, ObjetivoID, objetivonombre
Dim submit_logout, volver

doWhat         = Trim(Request("doWhat"))
PeriodoID      = CLng("0" & Trim(Request("PeriodoID")))
ClienteID      = CLng("0" & Trim(Request("ClienteID")))
clientenombre  = Trim(Request("clientenombre"))
ObjetivoID     = CLng("0" & Trim(Request("ObjetivoID")))
objetivonombre = Trim(Request("objetivonombre"))

submit_logout  = Trim(Request("submit_logout"))
volver         = Trim(Request("volver"))

If (doWhat = "") Then doWhat = "0"

If doWhat = "-1" Then Response.Redirect "../vendedores.asp"

If doWhat = "-2" Or submit_logout = "Salir" Then
    Session("currentUser") = ""
    Session("username") = ""
    Session("fromRedirect") = 0
    Session("password") = ""
    Response.Redirect "../vendedores.asp"
End If

If volver = "Menu" Then Response.Redirect "../vendedores.asp"

' --- Usuario actual ---
Dim VendedorID, VendedorNombre
VendedorID     = Trim(Session("currentUser"))
VendedorNombre = Trim(Session("username"))
If VendedorID = "" Then VendedorID = "000"
Session("VendedorID") = VendedorID

' --- Reset de sesión de archivo cuando volvés a 0/1/2 ---
If (doWhat = "0" Or doWhat = "1" Or doWhat = "2") Then
    Session("FileUploaded") = ""
    Session("FileName")     = ""
End If

' --- Persistencia de contexto al pasar a paso 2 ---
If doWhat = "2" Then
    If ClienteID > 0 Then Session("ClienteID") = ClienteID
    If ObjetivoID > 0 Then Session("ObjetivoID") = ObjetivoID
    If PeriodoID  > 0 Then Session("PeriodoID")  = PeriodoID
    If clientenombre <> "" Then Session("clientenombre") = clientenombre
    If objetivonombre <> "" Then Session("objetivonombre") = objetivonombre
End If

' --- Fallback desde sesión (por si no vinieron en el request) ---
If (ClienteID = 0 And CLng("0" & Session("ClienteID")) > 0) Then ClienteID = CLng(Session("ClienteID"))
If (ObjetivoID = 0 And CLng("0" & Session("ObjetivoID")) > 0) Then ObjetivoID = CLng(Session("ObjetivoID"))
If (PeriodoID  = 0 And CLng("0" & Session("PeriodoID"))  > 0) Then PeriodoID  = CLng(Session("PeriodoID"))
If (clientenombre = "" And Session("clientenombre") <> "") Then clientenombre = Session("clientenombre")
If (objetivonombre = "" And Session("objetivonombre") <> "") Then objetivonombre = Session("objetivonombre")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>ILLANES HNOS SRL - Carga Vendedores</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="../includes/style.css">  
<link rel="stylesheet" type="text/css" href="../includes/css/new-style.css">
<link rel="stylesheet" type="text/css" href="../includes/css/bars.css">
<link rel="stylesheet" type="text/css" href="../includes/calendar_cool.css" media="all" />
<script type="text/javascript" src="../includes/calendar_cool.js"></script>
<script type="text/javascript" src="../includes/copy.js"></script>
<style>
    .upload-box{display:inline-block;padding:15px;border:1px solid #ccc;border-radius:10px;background:#fff;width:95%;max-width:560px;margin:10px 0;}
    .btn1{padding:8px 15px;border:none;border-radius:5px;background:#007bff;color:#fff;cursor:pointer}
    .btn1:hover{background:#0056b3}
    .progress{display:none;width:100%;background:#f1f1f1;border-radius:5px;margin-top:10px;overflow:hidden;}
    #progressBar{width:0%;height:25px;background:#28a745;color:#fff;text-align:center;line-height:25px;}
    .textSuccess{color:green;font-weight:bold;}
    .labelError{color:#c00;font-weight:bold;}
    .pagetitle{font-weight:bold;font-size:16px}
</style>
</head>
<body>
<div style="overflow-x:auto" align="center">
<br/>
<span class="pagetitle">Bienvenido: <%=Server.HTMLEncode(VendedorID)%> - <%=Server.HTMLEncode(VendedorNombre)%></span>
<br/><br/>

<%
' =====================================================
' Paso 0: Selección de período + lista de clientes
' =====================================================
If (doWhat = "0") Then
%>
<form name="FF" method="post" action="vendedorcarga.asp">
    <input type="hidden" name="doWhat" value="0">
    <!-- Periodo actual / selección -->
    <!-- Tu include original -->
    <!--#include file="./includes/Ventas_Periodos_Actual.asp" -->
    <input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
</form>

<form name="FF2" method="post" action="vendedorcarga.asp">
    <input type="hidden" name="doWhat" value="0">
    <input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
    <input type="hidden" name="ClienteID" value="">
    <input type="hidden" name="clientenombre" value="">
    <!-- Lista de clientes del vendedor -->
    <!--#include file="./includes/Ventas_Vendedor_Cliente_List.asp" -->
    <div><br/>
        <input type="button" class="btn1" value="Salir" onclick="fnsalir(this.form)" style="width:90px;">
    </div>
</form>
<%
End If

' =====================================================
' Paso 1: Seleccionado cliente -> objetivos
' =====================================================
If (doWhat = "1") Then
%>
<form name="FF3" method="post" action="vendedorcarga.asp">
    <input type="hidden" name="doWhat" value="1">
    <input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
    <input type="hidden" name="ClienteID" value="<%=ClienteID%>">
    <input type="hidden" name="clientenombre" value="<%=Server.HTMLEncode(clientenombre)%>">
    <input type="hidden" name="ObjetivoID" value="">
    <input type="hidden" name="objetivonombre" value="">
    <span class="pagetitle">Cliente: <%=ClienteID%> - <%=Server.HTMLEncode(clientenombre)%></span>
    <br/><br/>
    <!-- Objetivos del cliente -->
    <!--#include file="./includes/Ventas_Vendedor_Cliente_Objetivos.asp" -->
    <div><br/>
        <input type="button" class="btn1" value="Volver" onclick="chkForm(this.form,0)" style="width:90px;">
    </div>
</form>
<%
End If

' =====================================================
' Paso 2: Carga de archivo (foto/comprobante) para Objetivo
' =====================================================
If (doWhat = "2") Then
%>
<form name="FF4" method="post" action="vendedorcarga.asp">
    <input type="hidden" name="doWhat" value="2">
    <input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
    <input type="hidden" name="ClienteID" value="<%=ClienteID%>">
    <input type="hidden" name="clientenombre" value="<%=Server.HTMLEncode(clientenombre)%>">
    <input type="hidden" name="ObjetivoID" value="<%=ObjetivoID%>">
    <input type="hidden" name="objetivonombre" value="<%=Server.HTMLEncode(objetivonombre)%>">
</form>

<span class="pagetitle">Cliente: <%=ClienteID%> - <%=Server.HTMLEncode(clientenombre)%></span>
<br/><br/>
<span class="pagetitle">Regla: <%=ObjetivoID%> - <%=Server.HTMLEncode(objetivonombre)%></span>
<br/><br/>

<div class="upload-box">
    <h3>📤 Subir foto o comprobante</h3>
    <input type="file" id="fileInput" accept="image/*,application/pdf"><br/><br/>
    <button type="button" class="btn1" id="btnUpload">Subir Archivo</button>
    <div class="progress"><div id="progressBar">0%</div></div>
    <div id="message" style="margin-top:10px;font-weight:bold;"></div>
</div>

<!-- Mini-galería / comunes del vendedor -->
<!--#INCLUDE FILE="_commonfotosvendedor.asp"-->

<form name="FF4b" method="post" action="vendedorcarga.asp">
    <input type="hidden" name="doWhat" value="1">
    <input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
    <input type="hidden" name="ClienteID" value="<%=ClienteID%>">
    <input type="hidden" name="clientenombre" value="<%=Server.HTMLEncode(clientenombre)%>">
    <div><br/>
        <input type="button" class="btn1" value="Volver" onclick="chkForm(this.form,1)" style="width:90px;">
    </div>
</form>
<%
End If

' =====================================================
' Paso 3: Confirmación + guardado en BD (SP)
' =====================================================
If (doWhat = "3") Then
    Dim okFile, okName
    okFile = (Session("FileUploaded") & "") <> ""
    okName = (Session("FileName") & "") <> ""
%>
    <%
    If okFile And okName Then
        ' Asegura valores desde sesión (por si no vinieron)
        If ClienteID = 0 Then ClienteID = CLng("0" & Session("ClienteID"))
        If ObjetivoID = 0 Then ObjetivoID = CLng("0" & Session("ObjetivoID"))
        If PeriodoID  = 0 Then PeriodoID  = CLng("0" & Session("PeriodoID"))
        If clientenombre = "" Then clientenombre = Session("clientenombre")
        If objetivonombre = "" Then objetivonombre = Session("objetivonombre")
    %>
        <!-- Llama a tu SP para persistir -->
        <!--#include file="./includes/Ventas_Vendedor_Cliente_Objetivos_Guardar.asp" -->
        <div class="textSuccess">
            ✅ Archivo subido correctamente<br/> 
        </div>
    <%
    Else
    %>
        <div class="labelError">❌ No se encontró archivo cargado. Volvé a subir el archivo.</div>
    <%
    End If
    %>

<form name="FF5" method="post" action="vendedorcarga.asp">
    <input type="hidden" name="doWhat" value="1">
    <input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">
    <input type="hidden" name="ClienteID" value="<%=ClienteID%>">
    <input type="hidden" name="clientenombre" value="<%=Server.HTMLEncode(clientenombre)%>">
    <input type="hidden" name="ObjetivoID" value="<%=ObjetivoID%>">
    <input type="hidden" name="objetivonombre" value="<%=Server.HTMLEncode(objetivonombre)%>">
    <div><br/>
        <input type="button" class="btn1" value="Volver" onclick="chkForm(this.form,1)" style="width:90px;">
    </div>
</form>
<%
End If
%>

</div>

<script type="text/javascript">
function fnsalir(Fm){
    Fm.doWhat.value = -2;
    Fm.submit();
}
function chkForm(Fm,prm){
    Fm.doWhat.value = prm;
    Fm.submit();
}
function chkForm2(Fm,prm,param1,param2){
    if(prm==1){
        // Asegurate que existan los inputs en el form
        if(!Fm.ClienteID){ var i=document.createElement('input'); i.type='hidden'; i.name='ClienteID'; Fm.appendChild(i); }
        if(!Fm.clientenombre){ var j=document.createElement('input'); j.type='hidden'; j.name='clientenombre'; Fm.appendChild(j); }
        Fm.doWhat.value = prm;
        Fm.ClienteID.value = param1;
        Fm.clientenombre.value = param2;
        Fm.submit();
    }
}
function chkForm3(Fm,prm,param1,param2,param3,param4){
    if(prm==2){
        if(!Fm.ClienteID){ var a=document.createElement('input'); a.type='hidden'; a.name='ClienteID'; Fm.appendChild(a); }
        if(!Fm.clientenombre){ var b=document.createElement('input'); b.type='hidden'; b.name='clientenombre'; Fm.appendChild(b); }
        if(!Fm.ObjetivoID){ var c=document.createElement('input'); c.type='hidden'; c.name='ObjetivoID'; Fm.appendChild(c); }
        if(!Fm.objetivonombre){ var d=document.createElement('input'); d.type='hidden'; d.name='objetivonombre'; Fm.appendChild(d); }
        Fm.doWhat.value    = prm;
        Fm.ClienteID.value = param1;
        Fm.clientenombre.value = param2;
        Fm.ObjetivoID.value = param3;
        Fm.objetivonombre.value = param4;
        Fm.submit();
    }
}

function ShowComments(texto) {
  if (!texto) texto = "Sin comentarios";

  // Sanitizar para evitar romper el HTML
  texto = String(texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/\r?\n/g, "<br>");

  // Crear overlay (fondo oscuro)
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.style.position = "fixed";
  overlay.style.left = "0";
  overlay.style.top = "0";
  overlay.style.width = "100%";
  overlay.style.height = "100%";
  overlay.style.background = "rgba(0,0,0,0.4)";
  overlay.style.display = "flex";
  overlay.style.alignItems = "center";
  overlay.style.justifyContent = "center";
  overlay.style.zIndex = "9999";

  // Crear contenido del modal
  const modal = document.createElement("div");
  modal.style.background = "#fff";
  modal.style.padding = "20px";
  modal.style.maxWidth = "450px";
  modal.style.borderRadius = "10px";
  modal.style.boxShadow = "0 2px 8px rgba(0,0,0,0.3)";
  modal.style.textAlign = "left";
  modal.style.lineHeight = "1.4";
  modal.innerHTML = `
    <h3 style="margin-top:0;font-size:18px;">🗒️ Observaciones</h3>
    <div style="max-height:300px;overflow:auto;font-size:15px;">${texto}</div>
    <div style="text-align:right;margin-top:15px;">
      <button id="closeModalBtn" 
              style="padding:6px 12px;background:#007bff;color:white;
                     border:none;border-radius:4px;cursor:pointer;">
        Cerrar
      </button>
    </div>
  `;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  // ✅ Cerrar correctamente todo el modal (overlay completo)
  document.getElementById("closeModalBtn").addEventListener("click", function () {
    document.body.removeChild(overlay);
  });

  // También cerrar al hacer clic fuera del modal
  overlay.addEventListener("click", function (e) {
    if (e.target === overlay) {
      document.body.removeChild(overlay);
    }
  });
}

/* Uploader (solo en doWhat=2) */
document.addEventListener("DOMContentLoaded",function(){
    var upBtn = document.getElementById("btnUpload");
    if(!upBtn) return;

    upBtn.addEventListener("click", function(){
        var fi = document.getElementById("fileInput"),
            msg = document.getElementById("message"),
            bar = document.getElementById("progressBar");
        if(!fi || !fi.files || !fi.files.length){
            alert("Seleccione un archivo.");
            return;
        }
        var file = fi.files[0], reader = new FileReader();
        reader.onload = function(){
            var base64 = reader.result.split(',')[1] || "";
            var xhr = new XMLHttpRequest();
            xhr.open("POST","_upload_https.asp",true);
            xhr.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
            xhr.upload.onprogress = function(e){
                if(e.lengthComputable){
                    var p = Math.round((e.loaded/e.total)*100);
                    document.querySelector(".progress").style.display="block";
                    bar.style.width = p+"%";
                    bar.innerText   = p+"%";
                }
            };
            xhr.onload = function(){
                if(xhr.status===200){
                    msg.innerHTML = "<span class='textSuccess'>"+xhr.responseText+"</span>";
                    // Al confirmar subida, ir a doWhat=3 (guardado en BD)
                    setTimeout(function(){
                        var f = document.createElement("form");
                        f.method = "post"; f.action = "vendedorcarga.asp";
                        f.innerHTML = ''+
                            '<input type="hidden" name="doWhat" value="3">'+
                            '<input type="hidden" name="PeriodoID" value="<%=PeriodoID%>">'+
                            '<input type="hidden" name="ClienteID" value="<%=ClienteID%>">'+
                            '<input type="hidden" name="clientenombre" value="<%=Replace(Server.HTMLEncode(clientenombre),"""","&quot;")%>">'+
                            '<input type="hidden" name="ObjetivoID" value="<%=ObjetivoID%>">'+
                            '<input type="hidden" name="objetivonombre" value="<%=Replace(Server.HTMLEncode(objetivonombre),"""","&quot;")%>">';
                        document.body.appendChild(f);
                        f.submit();
                    }, 800);
                } else {
                    msg.innerHTML = "<span class='labelError'>❌ Error en la carga ("+xhr.status+").</span>";
                }
            };
            var params = "filename="+encodeURIComponent(file.name)+
                         "&data="+encodeURIComponent(base64)+
                         "&ClienteID="+encodeURIComponent("<%=ClienteID%>")+
                         "&ObjetivoID="+encodeURIComponent("<%=ObjetivoID%>")+
                         "&PeriodoID="+encodeURIComponent("<%=PeriodoID%>");
            xhr.send(params);
        };
        reader.readAsDataURL(file);
    });
});

/* Delegación opcional para íconos "ojo" si vienen de tus includes */
document.addEventListener("click",function(e){
    var t=e.target;
    if(t && t.classList && t.classList.contains("ver-objetivos")){
        var cid=t.getAttribute("data-clienteid")||"0";
        var cn =t.getAttribute("data-clientenombre")||"";
        var f=document.forms["FF3"]||document.forms[0];
        if(!f){ return; }
        chkForm2(f,1,cid,cn);
    }
});
</script>
</body>
</html>
<%
On Error Resume Next
dbCon.Close
Set dbCon = Nothing
%>
