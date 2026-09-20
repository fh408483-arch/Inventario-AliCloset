/* ==========================================================
   Ali Closet · Sistema de inventario — app.js
   -----------------------------------------------------------
   ANTES DE USAR:
   1) Pegá tus credenciales de Supabase abajo (URL y anon key).
   2) Corré el esquema SQL en el SQL Editor.
      OJO: para las medidas en cm, corré una sola vez:
        alter table variantes add column if not exists medidas jsonb;
   3) Creá tu usuario en Authentication → Users → Add user
      (marcá "Auto Confirm User"). Ya no hay auto-registro en la web.
   4) Para el Asistente IA, desplegá la Edge Function "asistente-ia"
      (ver asistente-ia.ts) con tu GEMINI_API_KEY.
   ========================================================== */

const SUPABASE_URL = "https://urqhwneninvrayjcvzmk.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVycWh3bmVuaW52cmF5amN2em1rIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODkzMjU0NDYsImV4cCI6MjEwNDkwMTQ0Nn0._j1BssjEh7AX9KqKJ8_uq7DRKKMxK6TudbE0wqsg-Mc";
const BUCKET = "productos";

const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

/* ================= Utilidades ================= */
const $ = (id) => document.getElementById(id);
const val = (id) => ($(id) ? $(id).value.trim() : "");
const num = (id) => { const v = parseFloat($(id) ? $(id).value : ""); return isNaN(v) ? null : v; };
const money = (v) => new Intl.NumberFormat("es-SV", { style: "currency", currency: "USD" }).format(Number(v) || 0);
const esc = (s) => (s == null ? "" : String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])));
const fecha = (d) => new Date(d).toLocaleDateString("es-SV");
const cap = (s) => (s ? s[0].toUpperCase() + s.slice(1) : s);

function toast(msg, tipo) {
  const t = $("toast");
  t.textContent = msg;
  t.className = "toast" + (tipo ? " " + tipo : "");
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (t.hidden = true), 3500);
}

/* ---- Fechas: formato local para <input type="date"> ---- */
function fmtInput(d) {
  const off = d.getTimezoneOffset() * 60000;
  return new Date(d.getTime() - off).toISOString().slice(0, 10);
}
function rangoISO(desdeId, hastaId) {
  const d = val(desdeId), h = val(hastaId);
  const desde = d ? new Date(d + "T00:00:00") : null;
  const hasta = h ? new Date(h + "T23:59:59.999") : null;
  return { desde, hasta, dTxt: d, hTxt: h };
}
function ponerMesActual(desdeId, hastaId) {
  const now = new Date();
  const primero = new Date(now.getFullYear(), now.getMonth(), 1);
  if (!val(desdeId)) $(desdeId).value = fmtInput(primero);
  if (!val(hastaId)) $(hastaId).value = fmtInput(now);
}

/* ---- Excel (SheetJS) ---- */
function exportarExcel(nombreArchivo, hojas) {
  if (!window.XLSX) { toast("No cargó la librería de Excel. Revisá tu conexión.", "err"); return; }
  const wb = XLSX.utils.book_new();
  hojas.forEach((h) => {
    const filas = (h.filas && h.filas.length) ? h.filas : [{ "Sin datos": "" }];
    const ws = XLSX.utils.json_to_sheet(filas);
    XLSX.utils.book_append_sheet(wb, ws, (h.nombre || "Hoja").slice(0, 31));
  });
  XLSX.writeFile(wb, nombreArchivo);
  toast("Excel descargado ✓", "ok");
}
const hoy = () => fmtInput(new Date());

/* ---- Códigos por tipo de prenda (prefijos) ---- */
const TIPOS = [
  { codigo: "V",   nombre: "Vestidos" },
  { codigo: "C",   nombre: "Camisas mujer" },
  { codigo: "CH",  nombre: "Camisas hombre" },
  { codigo: "MQ",  nombre: "Maquillaje" },
  { codigo: "LEN", nombre: "Lencería" },
];
async function traerSkus() {
  const { data } = await sb.from("variantes").select("sku");
  return (data || []).map((v) => (v.sku || "").toUpperCase());
}
function siguienteCodigo(prefijo, skus) {
  const re = new RegExp("^" + prefijo + "-(\\d+)$");
  let max = 0, ancho = 4;
  skus.forEach((s) => {
    const m = re.exec(s);
    if (m) { const n = parseInt(m[1], 10); if (n >= max) { max = n; ancho = m[1].length; } }
  });
  return prefijo + "-" + String(max + 1).padStart(ancho, "0");
}

/* ---- Paleta para gráficos (tokens del logo) ---- */
const PAL = ["#b08a4f", "#262932", "#3c6f52", "#8a6a37", "#918b7c", "#b23a48", "#c9a86a", "#5a5d67"];
const GRID = "#ece5d8";
Chart.defaults.font.family = "Inter, system-ui, sans-serif";
Chart.defaults.color = "#918b7c";
const charts = {};
const holderDe = (id) => document.querySelector(`.chart-holder[data-for="${id}"]`);
function pintar(id, config) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
  const holder = holderDe(id);
  holder.innerHTML = `<canvas id="${id}"></canvas>`;
  charts[id] = new Chart($(id).getContext("2d"), config);
}
function vaciarChart(id, texto) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
  holderDe(id).innerHTML = `<div class="chart-vacio">${esc(texto)}</div>`;
}

/* ================= AUTH ================= */
const loginForm = $("login-form");

loginForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  $("login-error").textContent = "";
  const email = val("email");
  const password = val("password");
  if (!email || !password) { $("login-error").textContent = "Escribí correo y contraseña."; return; }
  $("btn-login").disabled = true; $("btn-login").textContent = "Entrando…";
  const { error } = await sb.auth.signInWithPassword({ email, password });
  $("btn-login").disabled = false; $("btn-login").textContent = "Iniciar sesión";
  if (error) { $("login-error").textContent = "Correo o contraseña incorrectos."; return; }
  entrarApp();
});

$("btn-logout").addEventListener("click", async () => {
  await sb.auth.signOut();
  $("app").classList.add("is-hidden");
  $("login-screen").classList.remove("is-hidden");
});

async function entrarApp() {
  $("login-screen").classList.add("is-hidden");
  $("app").classList.remove("is-hidden");
  irA("inicio");
}

/* ================= NAVEGACIÓN ================= */
const app = $("app");

document.querySelectorAll(".nav-item[data-view]").forEach((btn) => {
  btn.addEventListener("click", () => irA(btn.dataset.view));
});
$("hamburger").addEventListener("click", () => app.classList.toggle("menu-open"));
$("backdrop").addEventListener("click", () => app.classList.remove("menu-open"));

function irA(vista) {
  document.querySelectorAll(".nav-item[data-view]").forEach((b) => b.classList.toggle("is-active", b.dataset.view === vista));
  document.querySelectorAll(".view").forEach((v) => v.classList.toggle("is-active", v.id === "view-" + vista));
  app.classList.remove("menu-open");
  if (vista === "inicio") { ponerMesActual("dash-desde", "dash-hasta"); cargarInicio(); }
  if (vista === "productos") { ponerMesActual("pr-desde", "pr-hasta"); cargarProductos(); }
  if (vista === "historial") cargarHistorialCodigos();
  if (vista === "kardex") cargarKardexSelect();
  if (vista === "ventas") { ponerMesActual("vt-desde", "vt-hasta"); cargarVentas(); }
  if (vista === "reportes") { ponerMesActual("rep-desde", "rep-hasta"); cargarReportes(); }
}

/* ================= INICIO (dashboard) ================= */
$("dash-aplicar").addEventListener("click", cargarInicio);
$("dash-excel").addEventListener("click", exportarInicio);

let _dashData = null;

async function cargarInicio() {
  const { desde, hasta } = rangoISO("dash-desde", "dash-hasta");

  // Snapshot de inventario (no depende de fechas)
  const { data: stock, error } = await sb.from("stock_actual").select("*");
  if (error) { toast("No se pudo cargar el inicio.", "err"); return; }

  const valor = (stock || []).reduce((a, r) => a + Math.max(0, r.existencias) * (r.costo_promedio || 0), 0);
  const productos = new Set((stock || []).map((r) => r.producto_id)).size;
  const bajos = (stock || []).filter((r) => r.existencias <= r.stock_minimo);

  // Ventas del rango
  let qv = sb.from("ventas").select("*");
  if (desde) qv = qv.gte("fecha", desde.toISOString());
  if (hasta) qv = qv.lte("fecha", hasta.toISOString());
  const { data: ventas } = await qv.order("fecha", { ascending: true });
  const totalRango = (ventas || []).reduce((a, v) => a + Number(v.total), 0);

  // Items vendidos del rango (para unidades, top y categoría)
  let qi = sb.from("venta_items").select("cantidad, precio_unitario, variantes(sku, talla, color, tono, productos(nombre, categoria)), ventas!inner(fecha)");
  if (desde) qi = qi.gte("ventas.fecha", desde.toISOString());
  if (hasta) qi = qi.lte("ventas.fecha", hasta.toISOString());
  const { data: items } = await qi;
  const unidades = (items || []).reduce((a, it) => a + it.cantidad, 0);

  $("dash-metrics").innerHTML = `
    ${metric("Ventas del período", money(totalRango))}
    ${metric("Unidades vendidas", unidades)}
    ${metric("Valor inventario", money(valor))}
    ${metric("Bajo mínimo", bajos.length, bajos.length > 0)}`;

  $("dash-lowstock").innerHTML = bajos.length
    ? bajos.slice(0, 8).map((r) => filaStock(r)).join("")
    : `<p class="vacio">Todo bien: nada por agotarse. ✨</p>`;

  dibujarCharts(ventas || [], items || []);

  _dashData = { ventas: ventas || [], items: items || [], bajos, valor, totalRango, unidades };
}

function dibujarCharts(ventas, items) {
  // 1) Ventas por día
  const porDia = {};
  ventas.forEach((v) => { const k = fmtInput(new Date(v.fecha)); porDia[k] = (porDia[k] || 0) + Number(v.total); });
  const dias = Object.keys(porDia).sort();
  if (dias.length) {
    pintar("chart-ventas-dia", {
      type: "line",
      data: { labels: dias.map((d) => d.slice(5)), datasets: [{ data: dias.map((d) => porDia[d]), borderColor: PAL[0], backgroundColor: "rgba(176,138,79,.12)", fill: true, tension: .3, pointRadius: 3, pointBackgroundColor: PAL[0] }] },
      options: baseOpts(true),
    });
  } else vaciarChart("chart-ventas-dia", "Sin ventas en este período.");

  // 2) Ventas por canal
  const porCanal = {};
  ventas.forEach((v) => { const c = v.canal || "Otro"; porCanal[c] = (porCanal[c] || 0) + Number(v.total); });
  const canales = Object.keys(porCanal);
  if (canales.length) {
    pintar("chart-canal", {
      type: "doughnut",
      data: { labels: canales, datasets: [{ data: canales.map((c) => porCanal[c]), backgroundColor: PAL, borderColor: "#fff", borderWidth: 2 }] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } },
    });
  } else vaciarChart("chart-canal", "Sin ventas en este período.");

  // 3) Top productos vendidos (unidades)
  const porProd = {};
  items.forEach((it) => { const n = it.variantes?.productos?.nombre || it.variantes?.sku || "—"; porProd[n] = (porProd[n] || 0) + it.cantidad; });
  const top = Object.entries(porProd).sort((a, b) => b[1] - a[1]).slice(0, 6);
  if (top.length) {
    pintar("chart-top", {
      type: "bar",
      data: { labels: top.map((t) => t[0]), datasets: [{ data: top.map((t) => t[1]), backgroundColor: PAL[0], borderRadius: 6 }] },
      options: { indexAxis: "y", ...baseOpts(false) },
    });
  } else vaciarChart("chart-top", "Sin ventas en este período.");

  // 4) Ventas por categoría (unidades)
  const porCat = {};
  items.forEach((it) => { const c = it.variantes?.productos?.categoria || "otro"; porCat[c] = (porCat[c] || 0) + it.cantidad; });
  const cats = Object.keys(porCat);
  if (cats.length) {
    pintar("chart-categoria", {
      type: "doughnut",
      data: { labels: cats.map((c) => c[0].toUpperCase() + c.slice(1)), datasets: [{ data: cats.map((c) => porCat[c]), backgroundColor: [PAL[0], PAL[2], PAL[3], PAL[4]], borderColor: "#fff", borderWidth: 2 }] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } },
    });
  } else vaciarChart("chart-categoria", "Sin ventas en este período.");
}
function baseOpts(money_y) {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      x: { grid: { display: false } },
      y: { grid: { color: GRID }, ticks: money_y ? { callback: (v) => "$" + v } : {} },
    },
  };
}

async function exportarInicio() {
  if (!_dashData) { await cargarInicio(); }
  const { ventas, items, bajos } = _dashData;
  const { dTxt, hTxt } = rangoISO("dash-desde", "dash-hasta");
  const resumen = [
    { Métrica: "Período", Valor: `${dTxt || "inicio"} a ${hTxt || "hoy"}` },
    { Métrica: "Ventas del período", Valor: _dashData.totalRango },
    { Métrica: "Unidades vendidas", Valor: _dashData.unidades },
    { Métrica: "Valor inventario", Valor: _dashData.valor },
    { Métrica: "Productos bajo mínimo", Valor: bajos.length },
  ];
  const ventasR = ventas.map((v) => ({ Fecha: fecha(v.fecha), Canal: v.canal || "", Cliente: v.cliente || "", Total: Number(v.total) }));
  const topR = Object.entries(items.reduce((a, it) => { const n = it.variantes?.productos?.nombre || "—"; a[n] = (a[n] || 0) + it.cantidad; return a; }, {}))
    .sort((a, b) => b[1] - a[1]).map(([Producto, Unidades]) => ({ Producto, Unidades }));
  const bajosR = bajos.map((r) => ({ Producto: r.nombre, SKU: r.sku, Existencias: r.existencias, "Stock mínimo": r.stock_minimo }));
  exportarExcel(`inicio_${hoy()}.xlsx`, [
    { nombre: "Resumen", filas: resumen },
    { nombre: "Ventas", filas: ventasR },
    { nombre: "Top productos", filas: topR },
    { nombre: "Bajo mínimo", filas: bajosR },
  ]);
}

const metric = (label, value, alerta) =>
  `<div class="metric"><p class="label">${label}</p><p class="value${alerta ? " alerta" : ""}">${value}</p></div>`;

function filaStock(r) {
  const attrs = [r.talla, r.color, r.tono].filter(Boolean).join(" · ");
  const badge = r.existencias <= 0 ? `<span class="badge badge-danger">Agotado</span>`
    : r.existencias <= r.stock_minimo ? `<span class="badge badge-danger">${r.existencias} uds</span>`
    : `<span class="badge badge-ok">${r.existencias} uds</span>`;
  return `<div class="fila">
    <img class="fila-thumb" src="${(r.imagenes && r.imagenes[0]) || ""}" alt="" onerror="this.style.visibility='hidden'"/>
    <div class="fila-main"><div class="t">${esc(r.nombre)}</div><div class="s">${esc(attrs || r.sku)}</div></div>
    <div class="fila-right">${badge}</div></div>`;
}

/* ================= PRODUCTOS (catálogo) ================= */
let _stockCache = [];
let _archivados = new Set();

async function idsArchivados() {
  const { data } = await sb.from("productos").select("id").eq("activo", false);
  return new Set((data || []).map((p) => p.id));
}

async function cargarProductos() {
  const { data, error } = await sb.from("stock_actual").select("*").order("nombre");
  if (error) { toast("No se pudieron cargar los productos.", "err"); return; }
  _stockCache = data;
  _archivados = await idsArchivados();
  renderProductos();
}
$("buscar-producto").addEventListener("input", renderProductos);
$("filtro-categoria").addEventListener("change", renderProductos);
$("ver-archivados").addEventListener("change", renderProductos);

function renderProductos() {
  const q = val("buscar-producto").toLowerCase();
  const cat = val("filtro-categoria");
  const verArch = $("ver-archivados").checked;
  const rows = _stockCache.filter((r) =>
    (!cat || r.categoria === cat) &&
    (!q || (r.nombre || "").toLowerCase().includes(q) || (r.sku || "").toLowerCase().includes(q)) &&
    (verArch || !_archivados.has(r.producto_id))
  );
  $("productos-lista").innerHTML = rows.length
    ? `<div class="panel">${rows.map((r) => productoFila(r, _archivados.has(r.producto_id))).join("")}</div>`
    : `<div class="panel"><p class="vacio">No hay productos para mostrar.</p></div>`;

  $("productos-lista").querySelectorAll("[data-archivar]").forEach((b) =>
    b.addEventListener("click", () => archivarProducto(b.dataset.archivar, true)));
  $("productos-lista").querySelectorAll("[data-restaurar]").forEach((b) =>
    b.addEventListener("click", () => archivarProducto(b.dataset.restaurar, false)));
}

function productoFila(r, archivado) {
  const attrs = [r.talla, r.color, r.tono].filter(Boolean).join(" · ");
  const accion = archivado
    ? `<button class="btn-row" data-restaurar="${r.producto_id}">Restaurar</button>`
    : `<button class="btn-row" data-archivar="${r.producto_id}">Archivar</button>`;
  return `<div class="fila${archivado ? " archivada" : ""}" data-vid="${r.variante_id}">
    <img class="fila-thumb" src="${(r.imagenes && r.imagenes[0]) || ""}" alt="" onerror="this.style.visibility='hidden'"/>
    <div class="fila-main">
      <div class="t">${esc(r.nombre)} ${attrs ? `<span class="s">· ${esc(attrs)}</span>` : ""}</div>
      <div class="s">${esc(r.sku)} · ${money(r.precio_venta)}</div>
    </div>
    <div class="fila-right prod-acc">
      <span class="badge ${r.existencias <= r.stock_minimo ? "badge-danger" : "badge-ok"}">${r.existencias} uds</span>
      ${archivado ? `<span class="badge badge-warn">Archivado</span>` : ""}
      ${accion}
    </div></div>`;
}

async function archivarProducto(id, archivar) {
  const { error } = await sb.from("productos").update({ activo: !archivar }).eq("id", id);
  if (error) {
    toast(/activo/.test(error.message || "") ? "Falta la columna 'activo'. Corré el SQL indicado." : msgError(error), "err");
    return;
  }
  toast(archivar ? "Producto archivado ✓" : "Producto restaurado ✓", "ok");
  cargarProductos();
}

/* ---- Ver producto: doble clic en la fila o clic en la foto ---- */
function abrirVerDesdeFila(fila) {
  const r = _stockCache.find((x) => String(x.variante_id) === fila.dataset.vid);
  if (r) verProducto(r);
}
$("productos-lista").addEventListener("dblclick", (e) => {
  if (e.target.closest("button")) return;
  const fila = e.target.closest(".fila[data-vid]");
  if (fila) abrirVerDesdeFila(fila);
});
$("productos-lista").addEventListener("click", (e) => {
  if (!e.target.classList.contains("fila-thumb")) return;
  const fila = e.target.closest(".fila[data-vid]");
  if (fila) abrirVerDesdeFila(fila);
});

/* ---- Descargar una imagen (fallback: abrir en pestaña) ---- */
async function descargarImagen(url, nombre) {
  try {
    const res = await fetch(url);
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nombre || "foto.jpg";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1500);
  } catch (e) { window.open(url, "_blank"); }
}

/* ---- Ver producto: muestra TODO lo que se llenó ---- */
async function verProducto(r) {
  document.body.classList.add("no-scroll");
  $("modal-ver").classList.remove("is-hidden");
  $("ver-titulo").textContent = r.nombre || "Producto";
  $("ver-body").innerHTML = `<p class="vacio">Cargando…</p>`;

  const { data } = await sb.from("variantes")
    .select("*, productos(nombre, marca, descripcion, imagenes, categoria)")
    .eq("id", r.variante_id).single();
  const full = data || {};
  const p = full.productos || {};
  const imgs = (p.imagenes || r.imagenes || []).filter(Boolean);
  const med = full.medidas || null;

  const filas = [];
  const add = (l, v) => { if (v != null && String(v).trim() !== "") filas.push([l, v]); };
  add("Código", full.sku || r.sku);
  add("Categoría", cap(p.categoria || r.categoria));
  add("Marca", p.marca || r.marca);
  add("Talla / medida", full.talla);
  add("Color", full.color);
  add("Material", full.material);
  add("Tono", full.tono);
  add("Contenido", full.contenido);
  add("Precio de venta", money(full.precio_venta ?? r.precio_venta));
  add("Existencias", `${r.existencias} uds`);
  add("Stock mínimo", full.stock_minimo);

  let medHtml = "";
  if (med && typeof med === "object") {
    const et = { busto: "Busto", cintura: "Cintura", cadera: "Cadera", largo: "Largo", hombro: "Hombros", manga: "Manga" };
    const chips = Object.entries(et)
      .filter(([k]) => med[k] != null && String(med[k]).trim() !== "")
      .map(([k, lbl]) => `<div class="medida-chip"><span>${lbl}</span><strong>${esc(med[k])} cm</strong></div>`).join("");
    const otras = med.otras ? `<p class="medida-otras">${esc(med.otras)}</p>` : "";
    if (chips || otras) medHtml = `<div class="ver-seccion"><h4>Medidas</h4><div class="medidas-chips">${chips}</div>${otras}</div>`;
  }

  const galHtml = imgs.length
    ? `<div class="ver-seccion">
         <div class="ver-seccion-head"><h4>Fotos</h4>
           <button class="btn-mini" id="ver-descargar-todas"><i class="ti ti-download"></i> Descargar todas</button></div>
         <div class="galeria">${imgs.map((u, i) => `
           <figure class="galeria-item">
             <img src="${esc(u)}" alt="" onerror="this.closest('.galeria-item').style.display='none'"/>
             <button class="galeria-dl" data-url="${esc(u)}" data-i="${i + 1}" aria-label="Descargar"><i class="ti ti-download"></i></button>
           </figure>`).join("")}</div></div>`
    : `<p class="vacio">Este producto no tiene fotos subidas.</p>`;

  const descHtml = p.descripcion
    ? `<div class="ver-seccion"><h4>Descripción</h4><p class="ver-desc">${esc(p.descripcion)}</p></div>` : "";

  $("ver-body").innerHTML = `
    ${galHtml}
    <div class="ver-seccion"><h4>Detalle</h4>
      <div class="ver-meta">${filas.map(([l, v]) => `<div><span>${esc(l)}</span><strong>${esc(v)}</strong></div>`).join("")}</div>
    </div>
    ${medHtml}
    ${descHtml}`;

  const base = (r.sku || "producto").replace(/\s+/g, "-");
  $("ver-body").querySelectorAll(".galeria-dl").forEach((b) =>
    b.addEventListener("click", () => descargarImagen(b.dataset.url, `${base}-${b.dataset.i}.jpg`)));
  const todas = $("ver-descargar-todas");
  if (todas) todas.addEventListener("click", () => imgs.forEach((u, i) => setTimeout(() => descargarImagen(u, `${base}-${i + 1}.jpg`), i * 400)));
}
$("ver-close").addEventListener("click", () => { $("modal-ver").classList.add("is-hidden"); document.body.classList.remove("no-scroll"); });
$("modal-ver").addEventListener("click", (e) => { if (e.target.id === "modal-ver") { $("modal-ver").classList.add("is-hidden"); document.body.classList.remove("no-scroll"); } });

/* ---- Productos: reporte por fechas (ingresos / ventas) ---- */
let _prRows = [];
$("pr-filtrar").addEventListener("click", cargarProductosReporte);
$("pr-excel").addEventListener("click", () => {
  if (!_prRows.length) { toast("Primero filtrá para tener datos.", "err"); return; }
  exportarExcel(`productos_movimientos_${hoy()}.xlsx`, [{ nombre: "Movimientos", filas: _prRows }]);
});

async function cargarProductosReporte() {
  const { desde, hasta } = rangoISO("pr-desde", "pr-hasta");
  const tipo = val("pr-tipo");
  let q = sb.from("movimientos").select("fecha, tipo, cantidad, costo_unitario, referencia, variantes(sku, talla, color, tono, precio_venta, productos(nombre, categoria))");
  if (desde) q = q.gte("fecha", desde.toISOString());
  if (hasta) q = q.lte("fecha", hasta.toISOString());
  if (tipo !== "todos") q = q.eq("tipo", tipo);
  const { data, error } = await q.order("fecha", { ascending: false });
  if (error) { toast("No se pudo cargar el reporte.", "err"); return; }

  _prRows = (data || []).map((m) => {
    const v = m.variantes || {};
    const attrs = [v.talla, v.color, v.tono].filter(Boolean).join(" ");
    return {
      Fecha: fecha(m.fecha),
      Producto: (v.productos?.nombre || "—") + (attrs ? " " + attrs : ""),
      SKU: v.sku || "",
      Categoría: v.productos?.categoria || "",
      Movimiento: etiquetaTipo(m.tipo),
      Cantidad: m.cantidad,
      "Costo unit.": m.costo_unitario != null ? m.costo_unitario : "",
      Referencia: m.referencia || "",
    };
  });

  renderPrTabla();
}

$("pr-buscar").addEventListener("input", renderPrTabla);
function renderPrTabla() {
  const q = val("pr-buscar").toLowerCase();
  const rows = _prRows.filter((r) => !q || (r.Producto || "").toLowerCase().includes(q) || (r.SKU || "").toLowerCase().includes(q));
  if (!_prRows.length) { $("pr-tabla").innerHTML = `<p class="vacio">No hubo movimientos en ese rango.</p>`; return; }
  if (!rows.length) { $("pr-tabla").innerHTML = `<p class="vacio">Nada coincide con la búsqueda.</p>`; return; }
  $("pr-tabla").innerHTML = `<table class="data">
    <thead><tr><th>Fecha</th><th>Producto</th><th>SKU</th><th>Movimiento</th><th class="num">Cant.</th><th class="num">Costo</th><th>Ref.</th></tr></thead>
    <tbody>${rows.map((r) => `<tr>
      <td>${esc(r.Fecha)}</td><td>${esc(r.Producto)}</td><td><span class="s">${esc(r.SKU)}</span></td>
      <td>${esc(r.Movimiento)}</td><td class="num">${r.Cantidad}</td>
      <td class="num">${r["Costo unit."] !== "" ? money(r["Costo unit."]) : "—"}</td><td>${esc(r.Referencia)}</td></tr>`).join("")}</tbody></table>`;
}

/* ================= MODAL nuevo producto ================= */
let categoria = "ropa";
let fotos = [];
const modal = $("modal-producto");
const form = $("form-producto");

$("btn-nuevo-producto").addEventListener("click", () => { modal.classList.remove("is-hidden"); document.body.classList.add("no-scroll"); });
$("modal-close").addEventListener("click", cerrarModal);
$("np-cancelar").addEventListener("click", cerrarModal);
function cerrarModal() { modal.classList.add("is-hidden"); document.body.classList.remove("no-scroll"); }
modal.addEventListener("click", (e) => { if (e.target.id === "modal-producto") cerrarModal(); });

/* Sugerir código al elegir el tipo de prenda */
$("np-tipo-codigo").addEventListener("change", async () => {
  const pref = val("np-tipo-codigo");
  if (!pref) { $("np-sku-hint").textContent = ""; return; }
  $("np-sku-hint").textContent = "Buscando el siguiente código…";
  const skus = await traerSkus();
  const code = siguienteCodigo(pref, skus);
  $("np-sku").value = code;
  $("np-sku-hint").textContent = "Sugerido: " + code + " · podés cambiarlo a mano";
  const cat = pref === "MQ" ? "maquillaje" : "ropa";
  const btn = document.querySelector(`#form-producto .seg-btn[data-cat="${cat}"]`);
  if (btn && !btn.classList.contains("is-active")) btn.click();
});

document.querySelectorAll("#form-producto .seg-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    categoria = btn.dataset.cat;
    document.querySelectorAll("#form-producto .seg-btn").forEach((b) => b.classList.toggle("is-active", b === btn));
    $("np-campos-ropa").classList.toggle("is-hidden", categoria !== "ropa");
    $("np-campos-maquillaje").classList.toggle("is-hidden", categoria !== "maquillaje");
  });
});

$("np-fotos").addEventListener("change", () => {
  for (const file of $("np-fotos").files) { if (fotos.length >= 5) break; if (file.type.startsWith("image/")) fotos.push(file); }
  $("np-fotos").value = ""; renderPreviews();
});
function renderPreviews() {
  $("np-previews").innerHTML = "";
  fotos.forEach((file, i) => {
    const url = URL.createObjectURL(file);
    const div = document.createElement("div");
    div.className = "thumb";
    div.innerHTML = `<img src="${url}" alt=""/><button type="button" aria-label="Quitar">&times;</button>`;
    div.querySelector("button").addEventListener("click", () => { fotos.splice(i, 1); renderPreviews(); });
    $("np-previews").appendChild(div);
  });
}
function generarSku(nombre) {
  const base = (nombre || "PROD").toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]+/g, "-").slice(0, 10).replace(/-$/, "");
  return base + "-" + Math.floor(1000 + Math.random() * 9000);
}

/* ---- Recoge las medidas en cm del formulario (o null si no hay ninguna) ---- */
function recogerMedidas() {
  const campos = { busto: "med-busto", cintura: "med-cintura", cadera: "med-cadera", largo: "med-largo", hombro: "med-hombro", manga: "med-manga" };
  const m = {}; let hay = false;
  for (const [k, id] of Object.entries(campos)) { const v = num(id); if (v != null) { m[k] = v; hay = true; } }
  const otras = val("med-otras"); if (otras) { m.otras = otras; hay = true; }
  return hay ? m : null;
}

async function subirFotos() {
  const urls = [];
  for (const file of fotos) {
    const ext = file.name.split(".").pop();
    const path = crypto.randomUUID() + "." + ext;
    const { error } = await sb.storage.from(BUCKET).upload(path, file);
    if (error) throw error;
    urls.push(sb.storage.from(BUCKET).getPublicUrl(path).data.publicUrl);
  }
  return urls;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  document.querySelectorAll("#form-producto .error").forEach((el) => (el.textContent = ""));
  const nombre = val("np-nombre"); const precio = num("np-precio");
  let ok = true;
  if (!nombre) { setErr("np-nombre", "Escribí un nombre."); ok = false; }
  if (precio === null || precio <= 0) { setErr("np-precio", "Poné un precio válido."); ok = false; }
  if (!ok) return;

  $("np-guardar").disabled = true; $("np-guardar").textContent = "Guardando…";
  try {
    const imagenes = await subirFotos();
    const { data: prod, error: e1 } = await sb.from("productos").insert({
      nombre, categoria, marca: val("np-marca") || null, descripcion: val("np-descripcion") || null, imagenes,
    }).select("id").single();
    if (e1) throw e1;

    const variante = { producto_id: prod.id, sku: val("np-sku") || generarSku(nombre), precio_venta: precio, stock_minimo: num("np-minimo") || 0 };
    if (categoria === "ropa") {
      variante.talla = val("np-talla") || null;
      variante.color = val("np-color") || null;
      variante.material = val("np-material") || null;
      const medidas = recogerMedidas();
      if (medidas) variante.medidas = medidas;
    }
    else { variante.tono = val("np-tono") || null; variante.contenido = val("np-contenido") || null; }
    const { data: varRow, error: e2 } = await sb.from("variantes").insert(variante).select("id").single();
    if (e2) throw e2;

    const stock = num("np-stock");
    if (stock && stock > 0) {
      const mov = { variante_id: varRow.id, tipo: "entrada", cantidad: stock, costo_unitario: num("np-costo"), referencia: "Alta de producto" };
      if (categoria === "maquillaje") { mov.lote = val("np-lote") || null; mov.fecha_caducidad = val("np-caducidad") || null; }
      const { error: e3 } = await sb.from("movimientos").insert(mov);
      if (e3) throw e3;
    }

    toast("Producto guardado ✓", "ok");
    form.reset(); fotos = []; renderPreviews(); cerrarModal();
    $("np-sku-hint").textContent = "";
    cargarProductos();
  } catch (err) {
    console.error(err);
    toast(msgError(err), "err");
  } finally {
    $("np-guardar").disabled = false; $("np-guardar").textContent = "Guardar producto";
  }
});
function setErr(campo, msg) { const el = document.querySelector(`.error[data-for="${campo}"]`); if (el) el.textContent = msg; }
function msgError(err) {
  return err && err.message && /row-level security|jwt|auth/i.test(err.message)
    ? "Sesión vencida. Volvé a iniciar sesión." : "No se pudo guardar. Revisá la conexión.";
}

/* ================= KARDEX ================= */
let _kardexRows = [];
let _kardexNombre = "";

async function cargarKardexSelect() {
  const { data } = await sb.from("stock_actual").select("*").order("nombre");
  const sel = $("kx-variante");
  sel.innerHTML = (data || []).map((r) => {
    const attrs = [r.talla, r.color, r.tono].filter(Boolean).join(" ");
    return `<option value="${r.variante_id}" data-cat="${r.categoria}" data-nombre="${esc(r.nombre)} ${esc(attrs)}">${esc(r.nombre)} ${esc(attrs)} · ${esc(r.sku)}</option>`;
  }).join("");
  sel.onchange = () => { ajustarKardexForm(); cargarKardexTabla(); };
  ajustarKardexForm();
  cargarKardexTabla();
}
function ajustarKardexForm() {
  const opt = $("kx-variante").selectedOptions[0];
  const esMaq = opt && opt.dataset.cat === "maquillaje";
  document.querySelector(".acordeon").classList.toggle("show-maq", !!esMaq);
  const tipo = val("kx-tipo");
  $("kx-costo-wrap").style.display = tipo === "salida" ? "none" : "block";
}
$("kx-tipo").addEventListener("change", ajustarKardexForm);

async function cargarKardexTabla() {
  const opt = $("kx-variante").selectedOptions[0];
  const vid = $("kx-variante").value;
  _kardexNombre = opt ? opt.dataset.nombre : "";
  if (!vid) { $("kardex-tabla").innerHTML = `<p class="vacio">No hay variantes.</p>`; _kardexRows = []; return; }
  const { data } = await sb.from("movimientos").select("*").eq("variante_id", vid).order("fecha", { ascending: true });
  if (!data || !data.length) { $("kardex-tabla").innerHTML = `<p class="vacio">Sin movimientos todavía.</p>`; _kardexRows = []; return; }

  let qty = 0, avg = 0, filas = "";
  _kardexRows = [];
  data.forEach((m) => {
    if (m.tipo === "entrada") { const nuevo = qty + m.cantidad; avg = nuevo ? ((qty * avg) + (m.cantidad * (m.costo_unitario || 0))) / nuevo : 0; qty = nuevo; }
    else if (m.tipo === "salida") { qty -= m.cantidad; }
    else { qty += m.cantidad; }
    const signo = m.tipo === "salida" ? -m.cantidad : m.cantidad;
    _kardexRows.unshift({
      Fecha: fecha(m.fecha), Tipo: etiquetaTipo(m.tipo), Cantidad: signo,
      "Costo unit.": m.costo_unitario != null ? m.costo_unitario : "", Saldo: qty,
      "Costo prom.": Number(avg.toFixed(4)), Referencia: m.referencia || "",
    });
    filas = `<tr>
      <td>${fecha(m.fecha)}</td>
      <td>${etiquetaTipo(m.tipo)}</td>
      <td class="num">${signo > 0 ? "+" : ""}${signo}</td>
      <td class="num">${m.costo_unitario != null ? money(m.costo_unitario) : "—"}</td>
      <td class="num">${qty}</td>
      <td class="num">${money(avg)}</td>
      <td>${esc(m.referencia || "")}</td>
    </tr>` + filas;
  });
  $("kardex-tabla").innerHTML = `<table class="data">
    <thead><tr><th>Fecha</th><th>Tipo</th><th class="num">Cant.</th><th class="num">Costo</th><th class="num">Saldo</th><th class="num">Cº prom.</th><th>Ref.</th></tr></thead>
    <tbody>${filas}</tbody></table>`;
}
const etiquetaTipo = (t) => ({ entrada: "Entrada", salida: "Salida", ajuste: "Ajuste" }[t] || t);

$("kx-excel").addEventListener("click", () => {
  if (!_kardexRows.length) { toast("Elegí una variante con movimientos.", "err"); return; }
  exportarExcel(`kardex_${(_kardexNombre || "producto").trim().replace(/\s+/g, "-")}_${hoy()}.xlsx`, [{ nombre: "Kardex", filas: _kardexRows }]);
});

$("btn-kx-guardar").addEventListener("click", async () => {
  const vid = $("kx-variante").value;
  const tipo = val("kx-tipo");
  const cantidad = num("kx-cantidad");
  if (!vid || !cantidad || cantidad === 0) { toast("Poné una cantidad válida.", "err"); return; }
  const mov = { variante_id: vid, tipo, cantidad: Math.abs(cantidad), referencia: val("kx-referencia") || null };
  if (tipo === "ajuste") mov.cantidad = cantidad;
  if (tipo !== "salida") mov.costo_unitario = num("kx-costo");
  const opt = $("kx-variante").selectedOptions[0];
  if (opt && opt.dataset.cat === "maquillaje") { mov.lote = val("kx-lote") || null; mov.fecha_caducidad = val("kx-caducidad") || null; }
  const { error } = await sb.from("movimientos").insert(mov);
  if (error) { toast(msgError(error), "err"); return; }
  toast("Movimiento registrado ✓", "ok");
  ["kx-cantidad", "kx-costo", "kx-referencia", "kx-lote"].forEach((id) => ($(id).value = ""));
  cargarKardexTabla();
});

/* ================= VENTAS ================= */
let carrito = [];

async function cargarVentas() {
  const [{ data }, arch] = await Promise.all([
    sb.from("stock_actual").select("*").order("nombre"),
    idsArchivados(),
  ]);
  const sel = $("vt-variante");
  sel.innerHTML = (data || []).filter((r) => !arch.has(r.producto_id)).map((r) => {
    const attrs = [r.talla, r.color, r.tono].filter(Boolean).join(" ");
    return `<option value="${r.variante_id}" data-precio="${r.precio_venta}" data-stock="${r.existencias}" data-nombre="${esc(r.nombre)} ${esc(attrs)}">${esc(r.nombre)} ${esc(attrs)} · ${esc(r.sku)}</option>`;
  }).join("");
  sel.onchange = prefijarPrecio; prefijarPrecio();
}
function prefijarPrecio() {
  const opt = $("vt-variante").selectedOptions[0];
  if (opt) $("vt-precio").value = opt.dataset.precio;
}
$("btn-vt-add").addEventListener("click", () => {
  const opt = $("vt-variante").selectedOptions[0];
  if (!opt) return;
  const cantidad = num("vt-cantidad"); const precio = num("vt-precio");
  if (!cantidad || cantidad <= 0 || precio == null) { toast("Revisá cantidad y precio.", "err"); return; }
  const stock = parseFloat(opt.dataset.stock);
  const yaEnCarrito = carrito.filter((i) => i.variante_id === opt.value).reduce((a, i) => a + i.cantidad, 0);
  if (!isNaN(stock) && cantidad + yaEnCarrito > stock) {
    toast(`Ojo: solo hay ${stock} en stock de ese producto.`, "err"); return;
  }
  carrito.push({ variante_id: opt.value, nombre: opt.dataset.nombre, cantidad, precio });
  renderCarrito();
});
function renderCarrito() {
  if (!carrito.length) { $("vt-cart").innerHTML = ""; return; }
  const total = carrito.reduce((a, i) => a + i.cantidad * i.precio, 0);
  $("vt-cart").innerHTML = carrito.map((i, idx) => `
    <div class="fila">
      <div class="fila-main"><div class="t">${esc(i.nombre)}</div><div class="s">${i.cantidad} × ${money(i.precio)}</div></div>
      <div class="fila-right">${money(i.cantidad * i.precio)}
        <button class="modal-close" data-idx="${idx}" style="margin-left:8px;"><i class="ti ti-x"></i></button></div>
    </div>`).join("") + `<div class="fila"><div class="fila-main"><strong>Total</strong></div><div class="fila-right"><strong>${money(total)}</strong></div></div>`;
  $("vt-cart").querySelectorAll("button[data-idx]").forEach((b) => b.addEventListener("click", () => { carrito.splice(+b.dataset.idx, 1); renderCarrito(); }));
}
$("btn-vt-guardar").addEventListener("click", async () => {
  if (!carrito.length) { toast("Agregá al menos un producto.", "err"); return; }
  const total = carrito.reduce((a, i) => a + i.cantidad * i.precio, 0);
  try {
    const { data: venta, error: e1 } = await sb.from("ventas").insert({ canal: val("vt-canal"), cliente: val("vt-cliente") || null, total }).select("id").single();
    if (e1) throw e1;
    const items = carrito.map((i) => ({ venta_id: venta.id, variante_id: i.variante_id, cantidad: i.cantidad, precio_unitario: i.precio }));
    const { error: e2 } = await sb.from("venta_items").insert(items);
    if (e2) throw e2;
    const movs = carrito.map((i) => ({ variante_id: i.variante_id, tipo: "salida", cantidad: i.cantidad, referencia: "Venta " + venta.id.slice(0, 8) }));
    const { error: e3 } = await sb.from("movimientos").insert(movs);
    if (e3) throw e3;
    toast("Venta registrada ✓", "ok");
    carrito = []; renderCarrito(); $("vt-cliente").value = "";
    cargarVentas();
  } catch (err) { console.error(err); toast(msgError(err), "err"); }
});

/* ---- Ventas: filtro por fechas + Excel ---- */
let _ventasRows = [];
let _ventasItems = [];
$("vt-filtrar").addEventListener("click", filtrarVentas);
$("vt-excel").addEventListener("click", () => {
  if (!_ventasRows.length) { toast("Primero filtrá para tener datos.", "err"); return; }
  exportarExcel(`ventas_${hoy()}.xlsx`, [
    { nombre: "Ventas", filas: _ventasRows },
    { nombre: "Detalle", filas: _ventasItems },
  ]);
});

async function filtrarVentas() {
  const { desde, hasta } = rangoISO("vt-desde", "vt-hasta");
  let q = sb.from("ventas").select("*");
  if (desde) q = q.gte("fecha", desde.toISOString());
  if (hasta) q = q.lte("fecha", hasta.toISOString());
  const { data, error } = await q.order("fecha", { ascending: false });
  if (error) { toast("No se pudo filtrar.", "err"); return; }

  // Detalle (items) del rango
  let qi = sb.from("venta_items").select("cantidad, precio_unitario, variantes(sku, productos(nombre)), ventas!inner(fecha, canal, cliente)");
  if (desde) qi = qi.gte("ventas.fecha", desde.toISOString());
  if (hasta) qi = qi.lte("ventas.fecha", hasta.toISOString());
  const { data: items } = await qi;

  _ventasRows = (data || []).map((v) => ({ Fecha: fecha(v.fecha), Canal: v.canal || "", Cliente: v.cliente || "", Total: Number(v.total) }));
  _ventasItems = (items || []).map((it) => ({
    Fecha: fecha(it.ventas?.fecha), Canal: it.ventas?.canal || "", Cliente: it.ventas?.cliente || "",
    Producto: it.variantes?.productos?.nombre || it.variantes?.sku || "—",
    Cantidad: it.cantidad, "Precio unit.": Number(it.precio_unitario), Subtotal: it.cantidad * Number(it.precio_unitario),
  }));

  renderVentasLista();
}

$("vt-buscar").addEventListener("input", renderVentasLista);
function renderVentasLista() {
  const q = val("vt-buscar").toLowerCase();
  const rows = _ventasRows.filter((v) => !q || (v.Canal || "").toLowerCase().includes(q) || (v.Cliente || "").toLowerCase().includes(q));
  if (!_ventasRows.length) { $("ventas-historial").innerHTML = `<p class="vacio">No hubo ventas en ese rango.</p>`; return; }
  if (!rows.length) { $("ventas-historial").innerHTML = `<p class="vacio">Nada coincide con la búsqueda.</p>`; return; }
  const totalRango = rows.reduce((a, v) => a + v.Total, 0);
  $("ventas-historial").innerHTML =
    `<div class="fila"><div class="fila-main"><strong>Total mostrado</strong></div><div class="fila-right"><strong>${money(totalRango)}</strong></div></div>` +
    rows.map((v) => `<div class="fila">
      <div class="fila-main"><div class="t">${money(v.Total)}</div><div class="s">${esc(v.Canal || "")}${v.Cliente ? " · " + esc(v.Cliente) : ""}</div></div>
      <div class="fila-right"><span class="s">${esc(v.Fecha)}</span></div></div>`).join("");
}

/* ================= REPORTES ================= */
$("rep-mes").addEventListener("change", () => {
  const m = val("rep-mes"); // "YYYY-MM"
  if (!m) return;
  const [y, mm] = m.split("-").map(Number);
  const primero = new Date(y, mm - 1, 1);
  const ultimo = new Date(y, mm, 0);
  $("rep-desde").value = fmtInput(primero);
  $("rep-hasta").value = fmtInput(ultimo);
});
$("rep-filtrar").addEventListener("click", cargarReportes);
$("rep-excel").addEventListener("click", exportarReportes);

let _repVendidos = [], _repMargen = [], _repBajos = [];

async function cargarReportes() {
  const { desde, hasta } = rangoISO("rep-desde", "rep-hasta");
  const { data: stock } = await sb.from("stock_actual").select("*");

  // Más vendidos (respeta el rango de fechas)
  let qi = sb.from("venta_items").select("cantidad, variantes(sku, productos(nombre)), ventas!inner(fecha)");
  if (desde) qi = qi.gte("ventas.fecha", desde.toISOString());
  if (hasta) qi = qi.lte("ventas.fecha", hasta.toISOString());
  const { data: items } = await qi;
  const acum = {};
  (items || []).forEach((it) => {
    const nombre = it.variantes?.productos?.nombre || it.variantes?.sku || "—";
    acum[nombre] = (acum[nombre] || 0) + it.cantidad;
  });
  _repVendidos = Object.entries(acum).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([Producto, Unidades]) => ({ Producto, Unidades }));
  $("rep-vendidos").innerHTML = _repVendidos.length
    ? _repVendidos.map((r) => `<div class="fila"><div class="fila-main"><div class="t">${esc(r.Producto)}</div></div><div class="fila-right"><span class="badge badge-warn">${r.Unidades} vendidas</span></div></div>`).join("")
    : `<p class="vacio">Sin ventas en el período elegido.</p>`;

  // Margen por producto (situación actual)
  _repMargen = (stock || []).map((r) => ({ Producto: r.nombre, SKU: r.sku, Venta: r.precio_venta, Costo: r.costo_promedio || 0, Margen: r.precio_venta - (r.costo_promedio || 0) }))
    .sort((a, b) => b.Margen - a.Margen);
  $("rep-margen").innerHTML = _repMargen.length ? `<table class="data">
    <thead><tr><th>Producto</th><th class="num">Venta</th><th class="num">Costo</th><th class="num">Margen</th></tr></thead>
    <tbody>${_repMargen.map((r) => `<tr><td>${esc(r.Producto)} <span class="s">${esc(r.SKU)}</span></td><td class="num">${money(r.Venta)}</td><td class="num">${money(r.Costo)}</td><td class="num">${money(r.Margen)}</td></tr>`).join("")}</tbody></table>`
    : `<p class="vacio">Sin datos.</p>`;

  // Bajo mínimo (situación actual)
  const bajos = (stock || []).filter((r) => r.existencias <= r.stock_minimo);
  _repBajos = bajos.map((r) => ({ Producto: r.nombre, SKU: r.sku, Existencias: r.existencias, "Stock mínimo": r.stock_minimo }));
  $("rep-bajomin").innerHTML = bajos.length ? bajos.map(filaStock).join("") : `<p class="vacio">Nada por agotarse. ✨</p>`;
}

function exportarReportes() {
  const { dTxt, hTxt } = rangoISO("rep-desde", "rep-hasta");
  exportarExcel(`reportes_${dTxt || "inicio"}_a_${hTxt || "hoy"}.xlsx`, [
    { nombre: "Mas vendidos", filas: _repVendidos },
    { nombre: "Margen", filas: _repMargen },
    { nombre: "Bajo minimo", filas: _repBajos },
  ]);
}

/* ================= ASISTENTE IA (opcional) =================
   Se activa solo si existen los elementos del chat en el HTML.
   Así el resto de la app nunca se rompe si no está el chat. */
function addMsg(texto, quien) {
  const div = document.createElement("div");
  div.className = "msg msg-" + quien;
  div.textContent = texto;
  $("chat-mensajes").appendChild(div);
  $("chat-mensajes").scrollTop = $("chat-mensajes").scrollHeight;
  return div;
}
async function enviarChat() {
  const pregunta = val("chat-texto");
  if (!pregunta) return;
  addMsg(pregunta, "user");
  $("chat-texto").value = "";
  const pensando = addMsg("Pensando…", "bot");
  try {
    const { data: stock } = await sb.from("stock_actual").select("nombre, sku, categoria, existencias, stock_minimo, costo_promedio, precio_venta");
    const inicioMes = new Date(); inicioMes.setDate(1); inicioMes.setHours(0, 0, 0, 0);
    const { data: ventas } = await sb.from("ventas").select("total, fecha, canal").gte("fecha", inicioMes.toISOString());
    const { data: { session } } = await sb.auth.getSession();

    const res = await fetch(`${SUPABASE_URL}/functions/v1/asistente-ia`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON_KEY, "Authorization": "Bearer " + (session?.access_token || SUPABASE_ANON_KEY) },
      body: JSON.stringify({ pregunta, datos: { inventario: stock, ventas_del_mes: ventas } }),
    });
    const j = await res.json().catch(() => ({}));
    if (!res.ok) {
      const detalle = j.error ? ` (${j.error})` : ` (código ${res.status})`;
      pensando.textContent = "El asistente respondió con un error" + detalle + ". Revisá que la Edge Function esté desplegada y que GEMINI_API_KEY sea válida.";
      return;
    }
    pensando.textContent = j.respuesta || "No pude generar una respuesta.";
  } catch (err) {
    console.error(err);
    pensando.textContent = "No pude conectar con el asistente. Verificá que la Edge Function \"asistente-ia\" esté desplegada.";
  }
}
if ($("btn-chat")) {
  $("btn-chat").addEventListener("click", enviarChat);
  $("chat-texto").addEventListener("keydown", (e) => { if (e.key === "Enter") enviarChat(); });
  document.querySelectorAll(".chip[data-q]").forEach((c) => c.addEventListener("click", () => { $("chat-texto").value = c.dataset.q; enviarChat(); }));
}

/* ================= HISTORIAL DE CÓDIGOS ================= */
let _histRows = [];       // todos los códigos (con _fecha para filtrar)
let _histFiltrado = [];   // tras fecha + búsqueda

$("hs-filtrar").addEventListener("click", renderHistorial);
$("hs-buscar").addEventListener("input", renderHistorial);
$("hs-excel").addEventListener("click", () => {
  if (!_histFiltrado.length) { toast("No hay códigos para exportar.", "err"); return; }
  const filas = _histFiltrado.map(({ _fecha, ...r }) => r);
  exportarExcel(`codigos_${hoy()}.xlsx`, [{ nombre: "Códigos", filas }]);
});

async function cargarHistorialCodigos() {
  // 1) Próximo código sugerido por tipo
  const skus = await traerSkus();
  $("proximos-codigos").innerHTML = TIPOS.map((t) =>
    `<div class="codigo-card"><div class="ct">${esc(t.nombre)}</div><div class="cc">${siguienteCodigo(t.codigo, skus)}</div></div>`
  ).join("");

  // 2) Lista de códigos creados
  const { data: vars, error } = await sb.from("variantes").select("*, productos(nombre, categoria)");
  if (error) { toast("No se pudo cargar el historial.", "err"); $("hs-tabla").innerHTML = `<p class="vacio">No se pudo cargar.</p>`; return; }

  // Fecha de creación: created_at si existe; si no, el primer movimiento
  const { data: movs } = await sb.from("movimientos").select("variante_id, fecha").order("fecha", { ascending: true });
  const primerMov = {};
  (movs || []).forEach((m) => { if (!primerMov[m.variante_id]) primerMov[m.variante_id] = m.fecha; });

  _histRows = (vars || []).map((v) => {
    const f = v.created_at || v.inserted_at || primerMov[v.id] || null;
    const attrs = [v.talla, v.color, v.tono].filter(Boolean).join(" ");
    return {
      _fecha: f,
      Código: v.sku || "",
      Producto: (v.productos?.nombre || "—") + (attrs ? " " + attrs : ""),
      Categoría: v.productos?.categoria || "",
      "Precio venta": v.precio_venta != null ? v.precio_venta : "",
      Creado: f ? fecha(f) : "—",
    };
  }).sort((a, b) => String(b._fecha || "").localeCompare(String(a._fecha || "")));

  renderHistorial();
}

function renderHistorial() {
  const { desde, hasta } = rangoISO("hs-desde", "hs-hasta");
  const q = val("hs-buscar").toLowerCase();
  _histFiltrado = _histRows.filter((r) => {
    const f = r._fecha ? new Date(r._fecha) : null;
    if (desde && (!f || f < desde)) return false;
    if (hasta && (!f || f > hasta)) return false;
    if (q && !((r.Código || "").toLowerCase().includes(q) || (r.Producto || "").toLowerCase().includes(q))) return false;
    return true;
  });
  if (!_histRows.length) { $("hs-tabla").innerHTML = `<p class="vacio">Todavía no hay códigos creados.</p>`; return; }
  if (!_histFiltrado.length) { $("hs-tabla").innerHTML = `<p class="vacio">Nada coincide con el filtro.</p>`; return; }
  $("hs-tabla").innerHTML = `<table class="data">
    <thead><tr><th>Código</th><th>Producto</th><th>Categoría</th><th class="num">Precio</th><th>Creado</th></tr></thead>
    <tbody>${_histFiltrado.map((r) => `<tr>
      <td><strong>${esc(r.Código)}</strong></td>
      <td>${esc(r.Producto)}</td>
      <td>${esc(r.Categoría)}</td>
      <td class="num">${r["Precio venta"] !== "" ? money(r["Precio venta"]) : "—"}</td>
      <td>${esc(r.Creado)}</td></tr>`).join("")}</tbody></table>`;
}
/* ================= KEEP-ALIVE (mantener despierta la BD) ================= */
function pingBD() {
  sb.from("productos").select("id").limit(1).then(() => {}, () => {});
}
pingBD();                          // al abrir la app
setInterval(pingBD, 4 * 60 * 1000); // y cada 4 minutos mientras esté abierta+

/* ================= ARRANQUE ================= */
(async function init() {
  const { data: { session } } = await sb.auth.getSession();
  if (session) entrarApp();
})();
