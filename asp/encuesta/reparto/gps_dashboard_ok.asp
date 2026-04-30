<%@ Language="VBScript" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>

<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<title>Tracking GPS</title>

<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />

<style>
body { font-family: Arial, Helvetica, sans-serif; margin:0; }
#map { height: 72vh; width: 100%; }

.controls {
    padding: 10px;
    background: #f5f5f5;
    display: flex;
    gap: 12px;
    align-items: center;
    flex-wrap: wrap;
}

.alert {
    padding: 6px 10px;
    font-weight: bold;
    border-radius: 4px;
}
.alert.ok { background: #d4edda; color: #155724; }
.alert.warn { background: #fff3cd; color: #856404; }
.alert.bad { background: #f8d7da; color: #721c24; }

.stats {
    padding: 8px 10px;
    background: #222;
    color: #fff;
    font-size: 14px;
}
.stats span { margin-right: 15px; }

.num-marker {
    background: #007bff;
    color: #fff;
    border-radius: 50%;
    width: 26px;
    height: 26px;
    text-align: center;
    line-height: 26px;
    font-size: 12px;
    font-weight: bold;
    border: 2px solid #fff;
}
</style>
</head>

<body>

<div class="controls">
    Transportista:
    <select id="transportista">
        <option value="">Seleccione</option>
        <%
        Dim rs
        Set rs = conn.Execute("EXEC dbo.usp_Transportista_sel 0,0")
        Do While Not rs.EOF
            Response.Write "<option value='" & rs("TransportistaID") & "'>" & rs("Nombre") & "</option>"
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
    </select>

    Fecha:
    <input type="date" id="fecha"
        value="<%=Year(Date()) & "-" & Right("0"&Month(Date()),2) & "-" & Right("0"&Day(Date()),2)%>">

    <label>
        <input type="checkbox" id="chkLinea" checked> Mostrar línea
    </label>

    <label>
        <input type="checkbox" id="chkOrden"> Mostrar orden
    </label>

    <button onclick="cargar()">Cargar</button>

    <span id="estado" class="alert ok">Sin datos</span>
</div>

<div class="stats">
    <span id="km">Km: 0</span>
    <span id="puntos">Puntos: 0</span>
</div>

<div id="map"></div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>

<script>
let map = L.map('map').setView([-38.95, -68.05], 12);

L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© OpenStreetMap'
}).addTo(map);

/* =========================
   CAPAS
========================= */
let cluster = L.markerClusterGroup();
let capaOrden = L.layerGroup();
let capaClientes = L.layerGroup();
let polyline = null;
let ultimoMarker = null;
let HojaDeRutaID = null;

/* =========================
   UTILIDADES
========================= */
function estado(txt, tipo) {
    let e = document.getElementById("estado");
    e.className = "alert " + tipo;
    e.innerText = txt;
}

function iconNumero(n) {
    return L.divIcon({
        html: `<div class="num-marker">${n}</div>`,
        iconSize: [26,26],
        iconAnchor: [13,13]
    });
}

function iconUltimo() {
    return L.divIcon({
        html: "<div style='background:red;width:16px;height:16px;border-radius:50%;border:2px solid white'></div>",
        iconSize: [16,16],
        iconAnchor: [8,8]
    });
}

function distanciaMetros(lat1, lon1, lat2, lon2) {
    const R = 6371000;
    const dLat = (lat2-lat1)*Math.PI/180;
    const dLon = (lon2-lon1)*Math.PI/180;
    const a =
        Math.sin(dLat/2)*Math.sin(dLat/2) +
        Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180) *
        Math.sin(dLon/2)*Math.sin(dLon/2);
    return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

/* =========================
   CARGA GENERAL
========================= */
function cargar() {
    resolverHojaRuta();
}

/* =========================
   HOJA DE RUTA
========================= */
function resolverHojaRuta() {

    let t = transportista.value;
    let f = fecha.value;
    if (!t || !f) return alert("Seleccione transportista y fecha");

    fetch(`gps_hojaruta_ajax.asp?transportistaid=${t}&fecha=${f}`)
        .then(r => r.json())
        .then(resp => {

            if (!resp || !resp.hojaderutaid) {
                estado("Sin hoja de ruta", "bad");
                return;
            }

            HojaDeRutaID = resp.hojaderutaid;

            cargarRecorrido();
            cargarClientes(HojaDeRutaID);
        })
        .catch(() => estado("Error hoja de ruta", "bad"));
}

/* =========================
   RECORRIDO GPS
========================= */
function cargarRecorrido() {

    let t = transportista.value;
    let f = fecha.value;
    let mostrarOrden = document.getElementById("chkOrden").checked;
    let mostrarLinea = document.getElementById("chkLinea").checked;

    // limpiar
    if (map.hasLayer(cluster)) map.removeLayer(cluster);
    if (map.hasLayer(capaOrden)) map.removeLayer(capaOrden);
    capaOrden.clearLayers();
    cluster.clearLayers();

    if (polyline) map.removeLayer(polyline);
    if (ultimoMarker) map.removeLayer(ultimoMarker);

    fetch(`gps_puntos_ajax.asp?transportistaid=${t}&fecha=${f}`)
        .then(r => r.json())
        .then(data => {

            if (!data || data.length === 0) {
                estado("Sin puntos", "bad");
                return;
            }

            data.sort((a,b) =>
                new Date(a.fecha.replace(" ","T")) -
                new Date(b.fecha.replace(" ","T"))
            );

            let ruta = [];
            let ant = null;

            data.forEach(p => {
                if (!ant) { ruta.push(p); ant = p; return; }
                let d = distanciaMetros(ant.lat, ant.lon, p.lat, p.lon);
                if (d < 10 || d > 1500) return;
                ruta.push(p);
                ant = p;
            });

            let km = 0;
            let coords = [];
            let bounds = [];

            ruta.forEach((p,i) => {

                coords.push([p.lat, p.lon]);
                bounds.push([p.lat, p.lon]);

                if (i > 0) {
                    let d = distanciaMetros(
                        ruta[i-1].lat, ruta[i-1].lon,
                        p.lat, p.lon
                    );
                    km += d / 1000;
                }

                let marker;

                if (mostrarOrden) {
                    marker = L.marker([p.lat, p.lon], {
                        icon: iconNumero(i+1)
                    });
                    capaOrden.addLayer(marker);
                } else {
                    marker = L.circleMarker([p.lat, p.lon], {
                        radius: 4,
                        color: "#007bff"
                    });
                    cluster.addLayer(marker);
                }

                marker.bindPopup(`
                    <b>Punto ${i+1}</b><br>
                    Fecha: ${p.fecha}<br>
                    Lat: ${p.lat}<br>
                    Lon: ${p.lon}<br>
                    Velocidad: ${p.speed ?? "-"} km/h<br>
                    Precisión: ${p.accuracy ?? "-"} m
                `);
            });

            if (mostrarOrden) map.addLayer(capaOrden);
            else map.addLayer(cluster);

            if (mostrarLinea) {
                polyline = L.polyline(coords, {
                    color: "#007bff",
                    weight: 4
                }).addTo(map);
            }

            let last = ruta[ruta.length-1];
            ultimoMarker = L.marker([last.lat, last.lon], {icon: iconUltimo()})
                .addTo(map)
                .bindPopup("<b>ÚLTIMA POSICIÓN</b><br>"+last.fecha);

            map.fitBounds(bounds);

            estado("OK – "+ruta.length+" puntos", "ok");
            document.getElementById("km").innerText = "Km: " + km.toFixed(2);
            document.getElementById("puntos").innerText = "Puntos: " + ruta.length;
        });
}

/* =========================
   CLIENTES
========================= */
function cargarClientes(hojaRutaId) {

    capaClientes.clearLayers();
    map.addLayer(capaClientes);

    fetch(`gps_clientes_ajax.asp?hojaderutaid=${hojaRutaId}`)
        .then(r => r.json())
        .then(data => {

            if (!data) return;

            data.forEach(c => {

                if (!c.lat || !c.lon) return;

                let marker = L.marker([c.lat, c.lon], {
                    icon: L.divIcon({
                        html: `
                          <div style="
                            width:0;height:0;
                            border-left:8px solid transparent;
                            border-right:8px solid transparent;
                            border-bottom:16px solid red;
                          "></div>`,
                        iconSize:[16,16],
                        iconAnchor:[8,16]
                    })
                });

                marker.bindPopup(
                    `<b>${c.nombre}</b><br>${c.direccion}`
                );

                capaClientes.addLayer(marker);
            });
        });
}
</script>

</body>
</html>
