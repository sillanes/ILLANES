<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Server.ScriptTimeout = 300

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Dim chatID
chatID = Trim(Request("numero"))

' ===============================
' Marcar mensajes recibidos como leídos
' ===============================
If chatID <> "" Then
    conn.Execute "UPDATE [WhatsAppAPI].sofia.WhatsApp_API_Historial " & _
                 "SET Estado='read' WHERE Numero='" & Replace(chatID,"'","") & "' " & _
                 "AND Tipo='message' AND Estado<>'read'"
End If

' ===============================
' Datos del contacto
' ===============================
Dim sqlInfo, rsInfo, nombreContacto, esChatHumano
nombreContacto = ""
esChatHumano = False

sqlInfo = "SELECT TOP 1 c.Nombre FROM [WhatsAppAPI].sofia.WhatsAPP_API_Contacto c WHERE c.WhatsAppNumero='" & Replace(chatID,"'","") & "'"
Set rsInfo = conn.Execute(sqlInfo)
If Not rsInfo.EOF Then nombreContacto = rsInfo("Nombre")
rsInfo.Close : Set rsInfo = Nothing

' Detectar si es chat humano
Dim rsChatFlag
Set rsChatFlag = conn.Execute("SELECT TOP 1 1 FROM [WhatsAppAPI].sofia.WhatsAPP_API_Historial WHERE Numero='" & Replace(chatID,"'","") & "' AND Contenido LIKE '🗣️%'")
If Not rsChatFlag.EOF Then esChatHumano = True
rsChatFlag.Close : Set rsChatFlag = Nothing
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Chat Sofía 💬</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
/* ====== ESTILOS BASE ====== */
body {
    font-family:'Segoe UI',sans-serif;
    background:#e5ddd5;
    margin:0;
    padding:0;
}
.container {
    display:flex;
    height:100vh;
    overflow:hidden;
}
.sidebar-chats {
    width:30%;
    background:#fff;
    border-right:1px solid #ddd;
    display:flex;
    flex-direction:column;
    transition:all 0.3s ease;
}
.search-box {
    padding:10px;
    background:#f6f6f6;
    border-bottom:1px solid #ddd;
}
.search-box input {
    width:100%;
    padding:8px 12px;
    border-radius:20px;
    border:1px solid #ccc;
    outline:none;
}
.chat-list {
    flex:1;
    overflow-y:auto;
}
.chat-window {
    flex:1;
    display:flex;
    flex-direction:column;
    background:#efeae2;
}
.chat-header {
    background:#075E54;
    color:white;
    padding:15px;
    font-size:16px;
    display:flex;
    align-items:center;
    justify-content:space-between;
}
.chat-messages {
    flex:1;
    padding:15px;
    overflow-y:auto;
    display:flex;
    flex-direction:column;
    scroll-behavior:smooth;
}
.message {
    max-width:75%;
    margin-bottom:10px;
    padding:10px 15px;
    border-radius:15px;
    word-wrap:break-word;
    font-size:14px;
}
.incoming { background:#fff; align-self:flex-start; border:1px solid #ddd; }
.outgoing { background:#dcf8c6; align-self:flex-end; }
.interno { background:#d1ecf1; align-self:flex-end; border:1px solid #bee5eb; }
.time { font-size:11px; color:#888; text-align:right; margin-top:3px; }
.chat-asesor {
    background:#fff3cd;
    padding:10px;
    text-align:center;
    font-size:13px;
    color:#856404;
    border-radius:10px;
    margin:10px auto;
    width:80%;
}
.empty-state {
    text-align:center;
    color:#666;
    padding:50px;
    font-size:16px;
}
.chat-input {
    display:flex;
    padding:10px;
    background:#f0f0f0;
    border-top:1px solid #ccc;
}
.chat-input input {
    flex:1;
    padding:8px;
    border-radius:20px;
    border:1px solid #ccc;
    outline:none;
}
.chat-input button {
    margin-left:8px;
    background:#25D366;
    border:none;
    color:white;
    padding:8px 16px;
    border-radius:20px;
    cursor:pointer;
}
.chat-input button:hover {
    background:#0d6e61;
}

/* ====== RESPONSIVE ====== */
@media (max-width: 768px) {
    .container {
        flex-direction:column;
        height:auto;
        margin-left:0;
    }
    .sidebar-chats {
        width:100%;
        height:auto;
        border-right:none;
    }
    iframe {
        height:400px;
    }
    .chat-window {
        width:100%;
        height:auto;
        min-height:70vh;
    }
    .chat-header {
        font-size:15px;
        padding:12px;
    }
    .message {
        font-size:13px;
        max-width:90%;
    }
    .chat-input {
        flex-direction:column;
        align-items:stretch;
    }
    .chat-input input {
        width:100%;
        margin-bottom:8px;
    }
    .chat-input button {
        width:100%;
        padding:10px;
    }
}
</style>
</head>
<body>

<div class="container">
    <!-- Panel Izquierdo -->
    <div class="sidebar-chats">
        <div class="search-box">
            <input type="text" id="searchInput" placeholder="🔍 Buscar número o mensaje...">
        </div>
        <iframe src="mensajes_sofia_list.asp" style="border:none;width:100%;height:100%;"></iframe>
    </div>

    <!-- Panel Derecho -->
    <div class="chat-window">
        <% If chatID = "" Then %>
            <div class="empty-state"><p>👈 Seleccioná un contacto para ver la conversación</p></div>
        <% Else %>
            <div class="chat-header">
                <div><i class="fa fa-user"></i> <%=nombreContacto%> (<%=chatID%>)</div>
            </div>

            <% If esChatHumano Then %>
                <div class="chat-asesor">🗣️ Chat con asesor activo</div>
            <% End If %>

            <div class="chat-messages" id="chatContainer"></div>

            <div class="chat-input">
                <input type="text" id="txtMsg" placeholder="Escribí un mensaje...">
                <button onclick="enviarMensaje()">
                    <i class="fa fa-paper-plane"></i> Enviar
                </button>
            </div>
        <% End If %>
    </div>
</div>

<script>
let numeroActual = "<%=chatID%>";

async function cargarMensajes() {
    if (!numeroActual) return;
    const contenedor = document.getElementById("chatContainer");
    contenedor.innerHTML = "<p style='color:#666;'>⏳ Cargando conversación...</p>";

    try {
        const resp = await fetch("mensajes_sofia_load.asp?numero=" + encodeURIComponent(numeroActual));
        const texto = await resp.text();
        const limpio = texto.trim().replace(/^[^\[]*/, "");
        const datos = JSON.parse(limpio);
        contenedor.innerHTML = "";

        if (!Array.isArray(datos) || datos.length === 0) {
            contenedor.innerHTML = "<p style='color:#666;'>No hay mensajes para este contacto.</p>";
            return;
        }

        let ultimoDia = "";
        datos.forEach(m => {
            const textoMsg = (m.Contenido || "").replace(/\n/g, "<br>");
            let clase = "incoming";
            switch ((m.Origen || "").toLowerCase()) {
                case "bot": clase = "outgoing"; break;
                case "interno": clase = "interno"; break;
                default: clase = "incoming";
            }
            let icon = (m.Estado === "sent") ? "✔" : (m.Estado === "read" ? "✔✔" : "🕓");
            let hh="--", mm="--", fechaISO="";
            if (m.Fecha) {
                fechaISO = m.Fecha;
                const partes = m.Fecha.split("T");
                if (partes.length === 2) {
                    const horaPartes = partes[1].split(":");
                    if (horaPartes.length >= 2) { hh = horaPartes[0]; mm = horaPartes[1]; }
                }
            }
            const div = document.createElement("div");
            div.className = "message " + clase;
            div.innerHTML = `<div>${textoMsg}</div><div class="time">${hh}:${mm} ${icon}</div>`;
            contenedor.appendChild(div);
        });
        contenedor.scrollTop = contenedor.scrollHeight;
    } catch (err) {
        console.error("⚠ Error al cargar mensajes:", err);
        contenedor.innerHTML = "<p style='color:red;'>⚠ Error al cargar mensajes</p>";
    }
}

async function enviarMensaje() {
    const txt = document.getElementById("txtMsg");
    const msg = txt.value.trim();
    if (!msg) return;
    try {
        const params = "numero=" + encodeURIComponent(numeroActual) + "&mensaje=" + encodeURIComponent(msg);
        const resp = await fetch("mensajes_sofia_send.asp", {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: params
        });
        const text = await resp.text();
        console.log("📤 Envío:", text);
        txt.value = "";
        cargarMensajes();
    } catch (err) {
        alert("❌ Error al enviar mensaje: " + err);
    }
}

setInterval(cargarMensajes, 60000);
window.onload = cargarMensajes;
</script>

</body>
</html>
