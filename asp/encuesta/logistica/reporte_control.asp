<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("currentUser") = "" Then
    Response.Redirect "login.asp"
End If
%>

<style>
/* CONTENEDOR LADO A LADO */
.grids-wrapper {
    display: flex;
    gap: 20px;
    margin-top: 20px;
}

/* TARJETAS */
.grid-box {
    background: #ffffff;
    padding: 25px;
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.12);
    flex: 1;
    min-width: 420px;
    transition: transform .2s ease-in-out;
}
.grid-box:hover {
    transform: translateY(-3px);
}

h3 {
    margin-top: 0;
    font-weight: bold;
}

/* TABLAS */
table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 15px;
}
table th {
    background: #0059B3;
    color: white;
    padding: 12px;
    font-size: 14px;
    cursor: pointer;
}
table td {
    padding: 8px 10px;
    font-size: 14px;
    border-bottom: 1px solid #ddd;
}

/* Detalle multilinea */
.detalle-cell {
    white-space: pre-line;
    font-family: monospace;
}

/* RESPONSIVE */
@media (max-width: 992px) {
    .grids-wrapper {
        flex-direction: column;
    }
}
</style>

<div class="container-fluid" >

    <h2 style="margin-top:10px;">Reporte de Control</h2>

    <div id="timestamp" style="font-size:12px;color:#666;margin-bottom:10px;">
        Cargando datos...
    </div>

    <div class="grids-wrapper">

        <div id="eficiencia_div" class="grid-box">
            <h3 style="color:#0059B3;">Eficiencia de Armadores</h3>
            Cargando...
        </div>

        <div id="errores_div" class="grid-box">
            <h3 style="color:#B30000;">Errores por Armador</h3>
            Cargando...
        </div>

    </div>

</div>

<script>
// ================================
// AUTO REFRESH CADA 5 MINUTOS
// ================================
function refrescarReportes() {
    cargar("reporte_control_ajax.asp?tipo=eficiencia", "eficiencia_div");
    cargar("reporte_control_ajax.asp?tipo=errores", "errores_div");
}

function cargar(url, divID) {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", url, true);
    xhr.onreadystatechange = function() {
        if (xhr.readyState === 4) {
            if (xhr.status === 200) {
                document.getElementById(divID).innerHTML = xhr.responseText;

                // Re-conectar el ordenamiento después de recargar AJAX
                inicializarOrdenamiento();

                document.getElementById("timestamp").innerHTML =
                    "Última actualización: " + new Date().toLocaleTimeString();
            } else {
                document.getElementById(divID).innerHTML =
                    "<p style='color:red;'>Error al cargar (" + xhr.status + ")</p>";
            }
        }
    };
    xhr.send();
}

window.onload = function () {
    refrescarReportes();
    setInterval(refrescarReportes, 300000);
};

// ======================================================
// ORDENAR TABLAS POR COLUMNA — JS PURO
// ======================================================
// ======================================================
// ORDENAR TABLAS POR COLUMNA — JS PURO + FLECHAS
// ======================================================
function inicializarOrdenamiento() {
    document.querySelectorAll("th.sortable-th").forEach(th => {
        th.onclick = function () {
            var table = th.closest("table");
            var tbody = table.querySelector("tbody");
            var index = Array.prototype.indexOf.call(th.parentNode.children, th);

            // Quitar flechas de todos los th de esta tabla
            th.parentNode.querySelectorAll("th").forEach(x => {
                x.classList.remove("asc");
                x.classList.remove("desc");
                x.innerHTML = x.innerHTML.replace(" ▲", "").replace(" ▼", "");
            });

            // Alternar estado
            var asc = th.toggleAttribute("data-asc");

            // Agregar flecha visual
            if (asc) {
                th.classList.add("asc");
                th.innerHTML += " ▲";
            } else {
                th.classList.add("desc");
                th.innerHTML += " ▼";
            }

            var rows = Array.from(tbody.querySelectorAll("tr"));

            rows.sort(function(a, b) {
                var v1 = a.children[index].innerText.trim();
                var v2 = b.children[index].innerText.trim();

                var n1 = parseFloat(v1.replace(",", "."));
                var n2 = parseFloat(v2.replace(",", "."));

                if (!isNaN(n1) && !isNaN(n2)) {
                    return asc ? (n1 - n2) : (n2 - n1);
                }

                return asc ? v1.localeCompare(v2) : v2.localeCompare(v1);
            });

            rows.forEach(r => tbody.appendChild(r));
        };
    });
}

// Reconectar el sorting cuando la tabla se recarga por AJAX
document.addEventListener("DOMContentLoaded", inicializarOrdenamiento);
</script>
