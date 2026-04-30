<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="conexion.asp" -->
<%
Response.CodePage = 65001
Response.CharSet  = "UTF-8"
Server.ScriptTimeout = 120

If Session("currentUser") = "" Then
    Response.Redirect "../login.asp"
End If

Function Qs(name)
  Dim v: v = Request.QueryString(name)
  If IsNull(v) Or v = "" Then Qs = "" Else Qs = v
End Function

Dim tipo, fecha, km, vehiculoid
tipo = Qs("tipo")
fecha = Qs("fecha")
km = Qs("km")
vehiculoid = Qs("vehiculoid")

' Fechas precargadas desde vehiculos_home.asp
Dim RTO, BROMA, SEGURO, SERVICIO, KMSERV
RTO      = Qs("rto")
BROMA    = Qs("broma")
SEGURO   = Qs("seguro")
SERVICIO = Qs("servicio")
KMSERV   = Qs("kmserv")  ' lo tenemos, pero en el flujo nuevo sólo lo usamos si corresponde

If tipo = "" Or fecha = "" Or km = "" Or vehiculoid = "" Then
    Response.Write "Faltan parámetros. Volvé atrás."
    Response.End
End If
%>

<!--#include file="sidebar.asp" -->

<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0"> 
<link rel="stylesheet" href="estilos.css">
<link rel="stylesheet" href="ventas.css">
<title>Vehículos – Checklist <%=Server.HTMLEncode(tipo)%></title>
<style>
  body{background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;color:#222;margin:0}
  .container{max-width:900px;margin:20px auto;padding:0 15px}
  .card{background:#fff;border-radius:16px;box-shadow:0 10px 25px rgba(0,0,0,.08);padding:22px;margin-bottom:16px}
  .step-title{font-size:19px;font-weight:800;margin:10px 0 6px}
  .step-sub{color:#666;margin:0 0 14px}
  .categoria{margin-top:10px;padding:8px 10px;background:#eef2ff;border-radius:10px;font-weight:700;color:#1e40af}
  .item{display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-bottom:1px dashed #e5e7eb}
  .item:last-child{border-bottom:0}
  .opt{display:flex;gap:8px}
  .btn{border:0;border-radius:12px;padding:8px 14px;cursor:pointer;font-weight:700;transition:all .15s ease}
  .btn-primary{background:#2563eb;color:#fff}
  .btn-soft{background:#eef2ff;color:#1e40af}
  .btn-ghost{background:#fff;border:1px solid #d8e0ea;color:#1f2937}
  .btn.active-si{background:#16a34a;color:#fff}
  .btn.active-no{background:#dc2626;color:#fff}
  .footer-actions{display:flex;justify-content:space-between;gap:10px;margin-top:16px}
  .hidden{display:none}
  input[type="date"],input[type="text"],input[type="number"],textarea{
    width:100%;padding:10px 12px;border:1px solid #dce3ea;border-radius:10px;font-size:15px;background:#fff
  }
  input.error, textarea.error{border-color:#dc2626;background:#fee2e2}

  /* Barra de progreso */
  .progress-container{margin-bottom:12px}
  .progress-text{font-weight:600;color:#1e3a8a;margin-bottom:6px}
  .progress-bar{width:100%;background:#e5e7eb;border-radius:10px;height:10px;overflow:hidden}
  .progress-bar-inner{height:10px;background:#2563eb;width:0%;border-radius:10px;transition:width .3s ease}
</style>
</head>
<body>
<!--#include file="header.asp" -->

<div class="container">
  <div class="card">
    <div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap">
      <div><strong>Tipo:</strong> <%=Server.HTMLEncode(tipo)%></div>
      <div><strong>Fecha:</strong> <%=Server.HTMLEncode(fecha)%></div>
      <div><strong>KM:</strong> <%=Server.HTMLEncode(km)%></div>
      <div><strong>VehículoID:</strong> <%=Server.HTMLEncode(vehiculoid)%></div>
    </div>
  </div>

  <div id="wizard"></div>

  <div class="card" id="footerNav">
    <div class="footer-actions">
      <button class="btn btn-ghost" id="btnPrev" onclick="prevStep()">◀ Anterior</button>
      <button class="btn btn-primary" id="btnNext" onclick="nextStep()">Siguiente ▶</button>
    </div>
  </div>
</div>

<form id="frmGuardar" class="hidden" method="post" action="vehiculos_guardar.asp">
  <input type="hidden" name="tipo" value="<%=Server.HTMLEncode(tipo)%>" />
  <input type="hidden" name="fecha" value="<%=Server.HTMLEncode(fecha)%>" />
  <input type="hidden" name="km" value="<%=Server.HTMLEncode(km)%>" />
  <input type="hidden" name="vehiculoid" value="<%=Server.HTMLEncode(vehiculoid)%>" />
  <input type="hidden" id="observaciones" name="observaciones" value="" />
  <input type="hidden" id="dni" name="dni" value="" />
  <textarea name="detalles_xml" id="detalles_xml" class="hidden"></textarea>
</form>

<script>
function toggleSidebar(){
  document.querySelector('.sidebar').classList.toggle('open');
}

const Tipo   = "<%=Replace(tipo, """", "\" )%>";
const RTO    = "<%=RTO%>";
const BROMA  = "<%=BROMA%>";
const SEGURO = "<%=SEGURO%>";
const SERV   = "<%=SERVICIO%>";
// KMSERV existe pero NO lo precargamos en el flujo: se pedirá sólo si cambia Ultimo Servicio

// =============================================
// Helper para definir ítems
// =============================================
function p(categoria, sub, item, type='bool', valorInicial=''){
  return { categoria:categoria, sub:sub, item:item, type:type, valorInicial:valorInicial };
}

// =============================================
// FLUJO DIARIO (todos los SI/NO)
// =============================================
const diario = [
  {
    titulo:"Documentación vehículo",
    items:[
      p("Documentación vehiculo","", "Documentación para circular")
    ]
  },
  {
    titulo:"Condiciones Generales",
    items:[
      // Elementos de seguridad obligatorios (SubCategoria)
      p("Condiciones Generales","Elementos de seguridad obligatorios","Cinturones de seguridad"),
      p("Condiciones Generales","Elementos de seguridad obligatorios","Matafuego"),
      p("Condiciones Generales","Elementos de seguridad obligatorios","Botiquin"),

      // Habitáculo (SubCategoria)
      p("Condiciones Generales","Habitaculo","Bocina"),
      p("Condiciones Generales","Habitaculo","Limpieza general"),

      // Neumáticos (SubCategoria)
      p("Condiciones Generales","Neumáticos","Presión"),

      // Parabrisas (SubCategoria)
      p("Condiciones Generales","Parabrisas","Parabrisas"),
      p("Condiciones Generales","Parabrisas","Limpia parabrisas"),
      p("Condiciones Generales","Parabrisas","Lava parabrisas"),

      // Luces (SubCategoria)
      p("Condiciones Generales","Luces","Posición"),
      p("Condiciones Generales","Luces","Bajas"),
      p("Condiciones Generales","Luces","Altas"),
      p("Condiciones Generales","Luces","Freno"),
      p("Condiciones Generales","Luces","Giro"),
      p("Condiciones Generales","Luces","Balizas"),
      p("Condiciones Generales","Luces","Retroceso"),
      p("Condiciones Generales","Luces","Interior"),

      // Mecánica (SubCategoria)
      p("Condiciones Generales","Mecanica","Aceite motor"),
      p("Condiciones Generales","Mecanica","Líquido refrigerante"),
      p("Condiciones Generales","Mecanica","Freno de mano"),
      p("Condiciones Generales","Mecanica","Nivel de agua"),
      p("Condiciones Generales","Mecanica","Nivel de aceite"),

      // Caja de Carga (SubCategoria)
      p("Condiciones Generales","Caja de Carga","Puerta de carga"),
      p("Condiciones Generales","Caja de Carga","Puerta lateral"),
      p("Condiciones Generales","Caja de Carga","Equipo de frio"),

      // Exterior (SubCategoria)
      p("Condiciones Generales","Exterior","Cerraduras"),
      p("Condiciones Generales","Exterior","Visagras"),
      p("Condiciones Generales","Exterior","Levantavidrio"),
      p("Condiciones Generales","Exterior","Espejos retrovisores"),
      p("Condiciones Generales","Exterior","Limpieza general")
    ]
  },
  {
    titulo:"Observaciones y cierre",
    items:[]
  }
];

// =============================================
// FLUJO MENSUAL (fechas + SI/NO + servicio)
// =============================================
const mensual = [
  {
    titulo:"Documentación vehiculo (fechas)",
    items:[
      p("Documentación vehiculo","", "RTO Vencimiento","date", RTO),
      p("Documentación vehiculo","", "BROMATOLOGIA Vencimiento","date", BROMA),
      p("Documentación vehiculo","", "SEGURO Vencimiento","date", SEGURO),      
	  // Servicio (especial: fecha + Km)
      p("Servicio","","Ultimo Servicio","date", SERV),
      p("Servicio","","Km Servicio","number","")
    ]
  },
  {
    titulo:"Condiciones Generales",
    items:[
      // Elementos de seguridad obligatorios
      p("Condiciones Generales","","Cinturones de seguridad"),
      p("Condiciones Generales","","Apoyacabezas"),
      p("Condiciones Generales","","Balizas triangulo"),
      p("Condiciones Generales","","Extintor 5 kg"),
      p("Condiciones Generales","","Botiquin"),

      // Habitáculo
      p("Condiciones Generales","","Bocina"),
      p("Condiciones Generales","","Velocimetro"),
      p("Condiciones Generales","","Calefaccion"),
      p("Condiciones Generales","","Aire acondicionado"),
      p("Condiciones Generales","","Limpieza general"),
      p("Condiciones Generales","","Tapizado"),

      // Neumáticos
      p("Condiciones Generales","","Delantero derecho"),
      p("Condiciones Generales","","Delantero izquierdo"),
      p("Condiciones Generales","","Trasero derecho"),
      p("Condiciones Generales","","Trasero izquierdo"),
      p("Condiciones Generales","","Ajuste de tuercas"),
      p("Condiciones Generales","","Auxilio"),

      // Parabrisas
      p("Condiciones Generales","","Parabrisas"),
      p("Condiciones Generales","","Limpia parabrisas"),
      p("Condiciones Generales","","Lava parabrisas"),
      p("Condiciones Generales","","Liquido lavaparabrisas"),

      // Luces
      p("Condiciones Generales","","Posición"),
      p("Condiciones Generales","","Bajas"),
      p("Condiciones Generales","","Altas"),
      p("Condiciones Generales","","Freno"),
      p("Condiciones Generales","","Giro"),
      p("Condiciones Generales","","Balizas"),
      p("Condiciones Generales","","Retroceso"),
      p("Condiciones Generales","","luces habitaculo"),

      // Mecánica
      p("Condiciones Generales","","Aceite motor"),
      p("Condiciones Generales","","Aceite Caja"),
      p("Condiciones Generales","","Líquido refrigerante"),
      p("Condiciones Generales","","Líquido de frenos"),
      p("Condiciones Generales","","Freno de mano"),
      p("Condiciones Generales","","Líquido direcc. Asistida"),
      p("Condiciones Generales","","nivel agua"),
      p("Condiciones Generales","","nivel aceite"),
      p("Condiciones Generales","","Caño de escape"),
      p("Condiciones Generales","","soporte cardan"),


    ]
  },
  {
    titulo:"Caja de Carga",
    items:[
      p("Caja de Carga","","Puerta de carga"),
      p("Caja de Carga","","Piso de carga"),
      p("Caja de Carga","","Cubre caja"),
      p("Caja de Carga","","banda reflectaria"),
      p("Caja de Carga","","Equipo de frio"),
      p("Caja de Carga","","Estanterias"),
      p("Caja de Carga","","Pisaderas"),
      p("Caja de Carga","","agarradera"),
      p("Caja de Carga","","luces")
    ]
  },
  {
    titulo:"Exterior",
    items:[
      p("Exterior","","Paragolpes"),
      p("Exterior","","Puertas habitaculo"),
      p("Exterior","","Cristales no parabrisas"),
      p("Exterior","","Espejos retrovisores"),
      p("Exterior","","Limpieza general"),
      p("Exterior","","ganchos puertas")
    ]
  },
  {
    titulo:"Observaciones y cierre",
    items:[]
  }
];

// =============================================
// Selección del flujo
// =============================================
let flow = (Tipo === "Mensual") ? mensual : diario;
let currentStep = 0;
let respuestas = [];

// =============================================
// Render principal
// =============================================
function render(){
  const w = document.getElementById("wizard");
  w.innerHTML = "";

  const step = flow[currentStep];
  const card = document.createElement("div");
  card.className = "card";

  // ---- Progreso ----
  const prog = document.createElement("div");
  prog.className = "progress-container";
  const txt = document.createElement("div");
  txt.className = "progress-text";
  txt.textContent = "Paso " + (currentStep+1) + " de " + flow.length;
  const bar = document.createElement("div");
  bar.className = "progress-bar";
  const inner = document.createElement("div");
  inner.className = "progress-bar-inner";
  inner.style.width = Math.round(((currentStep+1)/flow.length)*100) + "%";
  bar.appendChild(inner);
  prog.appendChild(txt);
  prog.appendChild(bar);
  card.appendChild(prog);

  // ---- Título y subtítulo ----
  const h = document.createElement("h3");
  h.className = "step-title";
  h.textContent = step.titulo;
  card.appendChild(h);

  const sub = document.createElement("div");
  sub.className = "step-sub";
  sub.textContent = (step.items.length > 0)
    ? "Completá los ítems:"
    : "Completá Observaciones y DNI para cerrar.";
  card.appendChild(sub);

  // Para luego manejar el caso especial de Servicio
  let ultimoServicioInput = null;
  let rowKmServicio = null;

  if(step.items.length > 0){
    // Agrupar por Categoria
    const grupos = {};
    step.items.forEach(it=>{
      if(!grupos[it.categoria]) grupos[it.categoria] = [];
      grupos[it.categoria].push(it);
    });

    Object.keys(grupos).forEach(cat=>{
      const catDiv = document.createElement("div");
      catDiv.className = "categoria";
      catDiv.textContent = cat;
      card.appendChild(catDiv);

      grupos[cat].forEach(it=>{
        const row = document.createElement("div");
        row.className = "item";

        const lbl = document.createElement("div");
        lbl.textContent = (it.sub ? it.sub + " – " : "") + it.item;
        row.appendChild(lbl);

        const opts = document.createElement("div");
        opts.className = "opt";

        if(it.type === "date"){
          const input = document.createElement("input");
          input.type = "date";
          input.setAttribute("data-item", it.item);

          if(it.valorInicial){
            input.value = it.valorInicial;
            setValue(it, it.valorInicial);   // precarga en respuestas
          }

          input.onchange = e => {
            setValue(it, e.target.value || "");
          };

          // Guardamos referencia especial si es Ultimo Servicio
          if(it.item === "Ultimo Servicio"){
            ultimoServicioInput = input;
          }

          opts.appendChild(input);

        } else if(it.type === "number"){
          const input = document.createElement("input");
          input.type = "number";
          input.min  = "0";
          input.step = "1";
          input.setAttribute("data-item", it.item);

          if(it.valorInicial){
            input.value = it.valorInicial;
            setValue(it, it.valorInicial);
          }

          input.onchange = e => setValue(it, e.target.value || "");

          opts.appendChild(input);

          if(it.item === "Km Servicio"){
            row.id = "row-km-servicio";
            rowKmServicio = row;
          }

        } else {
          // SI / NO
          const si = document.createElement("button");
          si.type="button"; si.className="btn btn-soft"; si.textContent="Sí";
          const no = document.createElement("button");
          no.type="button"; no.className="btn btn-ghost"; no.textContent="No";

          // por defecto NO
          setValue(it, "NO");
          no.classList.add("active-no");

          si.onclick = ()=>{
            setValue(it, "SI");
            si.classList.add("active-si");
            no.classList.remove("active-no");
          };
          no.onclick = ()=>{
            setValue(it, "NO");
            no.classList.add("active-no");
            si.classList.remove("active-si");
          };

          opts.appendChild(si);
          opts.appendChild(no);
        }

        row.appendChild(opts);
        card.appendChild(row);
      });
    });

    // Lógica especial: Km Servicio visible solo si Ultimo Servicio cambia
    if(ultimoServicioInput && rowKmServicio){
      function actualizarKmServicio(){
        const valor = ultimoServicioInput.value || "";
        // Si no había fecha previa (SERV vacío) => mostrar KM si cargan alguna fecha
        if(!SERV){
          if(valor !== ""){
            rowKmServicio.classList.remove("hidden");
          } else {
            rowKmServicio.classList.add("hidden");
            clearKmServicioRespuesta();
          }
        } else {
          // Si había fecha previa: sólo mostrar si la cambian
          if(valor !== "" && valor !== SERV){
            rowKmServicio.classList.remove("hidden");
          } else {
            rowKmServicio.classList.add("hidden");
            clearKmServicioRespuesta();
          }
        }
      }

      function clearKmServicioRespuesta(){
        const idx = respuestas.findIndex(r => r.Item === "Km Servicio");
        if(idx >= 0){
          respuestas[idx].Valor = "";
        }
      }

      // Estado inicial
      actualizarKmServicio();
      // Evento cambio
      ultimoServicioInput.addEventListener("change", actualizarKmServicio);
    }

  } else {
    // Último paso: observaciones + DNI
    const wrapObs = document.createElement("div");
    wrapObs.style.marginTop = "8px";
    wrapObs.innerHTML = "<label>Observaciones</label><textarea id='txtObs' rows='4' placeholder='Escribí observaciones…'></textarea>";
    card.appendChild(wrapObs);

    const wrapDni = document.createElement("div");
    wrapDni.style.marginTop = "12px";
    wrapDni.innerHTML = "<label>DNI (quien realizó el control)</label><input type='text' id='txtDni' placeholder='Ej: 30111222'>";
    card.appendChild(wrapDni);

    const save = document.createElement("div");
    save.style.marginTop = "16px";
    const btn = document.createElement("button");
    btn.type="button"; btn.className="btn btn-primary"; btn.textContent="Finalizar checklist ✓";
    btn.onclick = enviar;
    save.appendChild(btn);
    card.appendChild(save);
  }

  w.appendChild(card);

  // Navegación
  const btnPrev = document.getElementById("btnPrev");
  const btnNext = document.getElementById("btnNext");

  btnPrev.style.display = (currentStep === 0) ? "none" : "inline-block";
  btnNext.style.display = (currentStep === flow.length - 1) ? "none" : "inline-block";
}

// =============================================
// Manejo de respuestas
// =============================================
function setValue(it,val){
  const idx = respuestas.findIndex(r =>
    r.Categoria === it.categoria &&
    r.SubCategoria === it.sub &&
    r.Item === it.item
  );
  if(idx >= 0){
    respuestas[idx].Valor = val;
  } else {
    respuestas.push({
      Categoria:   it.categoria,
      SubCategoria:it.sub,
      Item:        it.item,
      Valor:       val
    });
  }
}

// =============================================
// Navegación entre pasos
// =============================================
function nextStep(){
  if(currentStep < flow.length - 1){
    if(!validarStep()) return;
    currentStep++;
    render();
  }
}

function prevStep(){
  if(currentStep > 0){
    currentStep--;
    render();
  }
}

// =============================================
// Validación del paso actual
// =============================================
function validarStep(){
  const step = flow[currentStep];
  if(step.items.length === 0){
    return true;
  }

  for(let it of step.items){
    // Caso especial: Km Servicio oculto => no exigir valor
    if(it.item === "Km Servicio"){
      const rowKm = document.getElementById("row-km-servicio");
      if(rowKm && rowKm.classList.contains("hidden")){
        continue;
      }
    }

    const idx = respuestas.findIndex(r =>
      r.Categoria === it.categoria &&
      r.SubCategoria === it.sub &&
      r.Item === it.item
    );
    if(idx < 0 || respuestas[idx].Valor === "" || respuestas[idx].Valor === null){
      alert("Debés completar todos los ítems antes de continuar.");
      return false;
    }
  }
  return true;
}

// =============================================
// XML + envío
// =============================================
function xmlEscape(s){
  return String(s||"")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&apos;");
}

function buildXML(){
  let x = "<Detalles>";
  respuestas.forEach(r=>{
    x += "<D Categoria='" + xmlEscape(r.Categoria) +
         "' SubCategoria='" + xmlEscape(r.SubCategoria) +
         "' Item='" + xmlEscape(r.Item) +
         "' Valor='" + xmlEscape(r.Valor) + "'/>";
  });
  x += "</Detalles>";
  return x;
}

function enviar(){
  const dniInput = document.getElementById("txtDni");
  const dni = (dniInput.value || "").trim();
  if(!dni){
    dniInput.classList.add("error");
    dniInput.focus();
    alert("Ingresá el DNI para cerrar el checklist.");
    return;
  }
  dniInput.classList.remove("error");

  document.getElementById("observaciones").value = document.getElementById("txtObs").value || "";
  document.getElementById("dni").value           = dni;
  document.getElementById("detalles_xml").value  = buildXML();
  document.getElementById("frmGuardar").submit();
}

// Inicializar
render();
</script>
</body>
</html>
