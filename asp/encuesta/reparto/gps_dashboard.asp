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
<title>GPS Dashboard</title>

<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />

<style>
body { font-family: Arial, Helvetica, sans-serif; margin:0; }
#map { height: 72vh; width: 100%; }

.controls{
  padding:10px;
  background:#f5f5f5;
  display:flex;
  gap:10px;
  align-items:center;
  flex-wrap:wrap;
}
.alert{
  padding:6px 10px;
  font-weight:bold;
  border-radius:4px;
}
.alert.ok{ background:#d4edda; color:#155724; }
.alert.warn{ background:#fff3cd; color:#856404; }
.alert.bad{ background:#f8d7da; color:#721c24; }
.small { font-size:12px; opacity:.85; }
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
        Response.Write "<option value='" & rs("TransportistaID") & "'>" & rs("Apellido") & " " & rs("Nombre") & "</option>"
        rs.MoveNext
      Loop
      rs.Close
    %>
  </select>

  Fecha:
  <input type="date" id="fecha"
    value="<%=Year(Date()) & "-" & Right("0"&Month(Date()),2) & "-" & Right("0"&Day(Date()),2)%>">

  <label>
    <input type="checkbox" id="chkDetalle" checked>
    Mostrar detalle (puntos / trazado)
  </label>

  <button onclick="cargar()">Cargar</button>

  <span id="estado" class="alert ok">Sin datos</span>
  <span class="small" id="debug"></span>
</div>

<div id="map"></div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>

<script>
/* =========================
   MAPA / CAPAS
========================= */
let map = L.map('map').setView([-38.95, -68.05], 12);

L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
  attribution: '© OpenStreetMap'
}).addTo(map);

let capaGrupos  = L.layerGroup().addTo(map);
let capaDetalle = L.layerGroup();
let capaClientes = L.layerGroup();

let polyline = null;
let HojaDeRutaID = null;
let markerActivo = null;

/* =========================
   HELPERS UI
========================= */
function estado(txt, tipo){
  const e = document.getElementById("estado");
  e.className = "alert " + tipo;
  e.innerText = txt;
}

function debug(txt){
  document.getElementById("debug").innerText = txt || "";
}

/* =========================
   ICONOS
========================= */
function iconGrupo(n, activo){
  const color = activo ? "#dc3545" : "#007bff";
  return L.divIcon({
    html:`<div style="background:${color};color:#fff;width:26px;height:26px;border-radius:50%;text-align:center;line-height:26px;font-weight:bold;border:2px solid white">${n}</div>`,
    iconSize:[26,26],
    iconAnchor:[13,13]
  });
}

function iconInicio(){
  return L.divIcon({
    html:`<div style="background:green;color:white;width:30px;height:30px;border-radius:50%;text-align:center;line-height:30px;font-weight:bold;border:3px solid white">I</div>`,
    iconSize:[30,30],
    iconAnchor:[15,15]
  });
}

function iconFin(){
  return L.divIcon({
    html:`<div style="background:red;color:white;width:32px;height:32px;border-radius:50%;text-align:center;line-height:32px;font-weight:bold;border:3px solid white">F</div>`,
    iconSize:[32,32],
    iconAnchor:[16,16]
  });
}

/* =========================
   CHECK DETALLE
========================= */
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("chkDetalle").addEventListener("change", aplicarVisibilidadDetalle);
});

function aplicarVisibilidadDetalle(){
  if (!document.getElementById("chkDetalle").checked){
    if (map.hasLayer(capaDetalle)) map.removeLayer(capaDetalle);
    if (polyline) { map.removeLayer(polyline); polyline = null; }
  } else {
    if (!map.hasLayer(capaDetalle)) map.addLayer(capaDetalle);
  }
}

/* =========================
   FLUJO
========================= */
function cargar(){
  resolverHojaRuta();
}

function resolverHojaRuta(){
  const t = transportista.value;
  const f = fecha.value;
  if (!t || !f) return alert("Seleccione transportista y fecha");

  estado("Buscando hoja de ruta...", "warn");

  fetch(`gps_hojaruta_ajax.asp?transportistaid=${t}&fecha=${f}`)
    .then(r=>r.json())
    .then(resp=>{
      if (!resp.hojaderutaid){
        estado("Sin hoja de ruta", "bad");
        return;
      }
      HojaDeRutaID = resp.hojaderutaid;
      cargarClientes(HojaDeRutaID);
      cargarGrupos();
    });
}

function limpiarDetalle(){
  capaDetalle.clearLayers();
  if (polyline){ map.removeLayer(polyline); polyline=null; }
}

/* =========================
   GRUPOS (INICIO / FIN)
========================= */
function cargarGrupos(){

  capaGrupos.clearLayers();
  limpiarDetalle();
  markerActivo = null;

  estado("Cargando recorrido...", "warn");

  fetch(`gps_recorrido_grupo_ajax.asp?hojaderutaid=${HojaDeRutaID}`)
    .then(r=>r.json())
    .then(data=>{

      if (!data || data.length===0){
        estado("Sin recorrido", "bad");
        return;
      }

      const bounds=[];

      data.forEach((g, idx)=>{
        const orden = idx+1;
        const lat = g.lat ?? g.Latitud;
        const lon = g.lon ?? g.Longitud;
        const grupoId = g.GrupoID ?? g.grupo;
        const minutos =
		  g.minutos ??
		  g.Minutos ??
		  g.tiempo ??
		  g.Tiempo ??
		  g.permanencia ??
		  g.Permanencia ??
		  g.tiempo_min ??
		  g.TiempoMin ??
		  "";


        if (lat==null || lon==null) return;

        let icono = iconGrupo(orden,false);
        let extra = "";

        if (idx===0){
          icono = iconInicio();
          extra = "<br><b>🟢 INICIO DEL RECORRIDO</b>";
        }
        if (idx===data.length-1){
          icono = iconFin();
          extra = "<br><b>🔴 ÚLTIMA POSICIÓN</b>";
        }



        const m = L.marker([lat,lon],{icon:icono});
        m._grupoId = grupoId;
        m._orden = orden;

		let textoPermanencia = minutos !== ""
		  ? `Permanencia: ${minutos} min<br>`
		  : "";

		m.bindPopup(`
		  <b>Grupo ${orden}</b><br>
		  ${extra}
		  ${textoPermanencia}
		  <b>Clic para ver detalle</b>
		`);


        m.on("click",()=>{
          if (markerActivo)
            markerActivo.setIcon(iconGrupo(markerActivo._orden,false));

          markerActivo=m;
          m.setIcon(iconGrupo(orden,true));

          if (!chkDetalle.checked) return;
          cargarDetalleGrupo(grupoId);
        });

        capaGrupos.addLayer(m);
        bounds.push([lat,lon]);
      });

      map.fitBounds(bounds);
      estado(`OK - ${data.length} grupos`, "ok");
      aplicarVisibilidadDetalle();
    });
}

/* =========================
   DETALLE
========================= */
function cargarDetalleGrupo(grupoId){

  limpiarDetalle();
  map.addLayer(capaDetalle);

  fetch(`gps_recorrido_grupo_detalle_ajax.asp?hojaderutaid=${HojaDeRutaID}&grupo=${grupoId}`)
    .then(r=>r.json())
    .then(data=>{
      if (!data || data.length===0) return;

      const coords=[];
      const bounds=[];

      data.forEach(p=>{
        const lat=p.lat??p.Latitud;
        const lon=p.lon??p.Longitud;
        if (lat==null||lon==null) return;

        coords.push([lat,lon]);
        bounds.push([lat,lon]);
		const fecha =
			  p.fecha ??
			  p.Fecha ??
			  p.FechaHoraGPS ??
			  p.datetime ??
			  p.DateTime ??
			  "";

			const velocidad =
			  p.velocidad ??
			  p.Velocidad ??
			  p.speed ??
			  p.Speed ??
			  "";

			const precision =
			  p.accuracy ??
			  p.Accuracy ??
			  p.precision ??
			  p.Precision ??
			  "";
        L.circleMarker([lat,lon],{radius:4,color:"#dc3545"})
          

			.bindPopup(`
			  <b>${fecha}</b><br>
			  Velocidad: ${velocidad !== "" ? velocidad + " km/h" : "N/D"}<br>
			  Precisión: ${precision !== "" ? precision + " m" : "N/D"}
			`)

          .addTo(capaDetalle);
      });

      if (chkDetalle.checked){
        polyline=L.polyline(coords,{color:"#dc3545",weight:4}).addTo(map);
      }

      map.fitBounds(bounds);
    });
}

/* =========================
   CLIENTES
========================= */
function cargarClientes(hojaRutaId){
  capaClientes.clearLayers();
  map.addLayer(capaClientes);

  fetch(`gps_clientes_ajax.asp?hojaderutaid=${hojaRutaId}`)
    .then(r=>r.json())
    .then(data=>{
      data.forEach(c=>{
        if (!c.lat||!c.lon) return;

        L.marker([c.lat,c.lon],{
          icon:L.divIcon({
            html:`<div style="width:0;height:0;border-left:8px solid transparent;border-right:8px solid transparent;border-bottom:26px solid red"></div>`,
            iconSize:[16,16],
            iconAnchor:[8,16]
          })
        })
        .bindPopup(`<b>${c.nombre}</b><br>${c.direccion}`)
        .addTo(capaClientes);
      });
    });
}
</script>

</body>
</html>
