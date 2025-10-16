// js/formulario.js
import { getToken, login } from "./msal-app.js";

// === CONFIG: link del CSV (compartido) ===
const DOTACION_CSV_LINK = "https://peruntp-my.sharepoint.com/:x:/g/personal/ntp_pyc_peruntp_onmicrosoft_com1/Eb7YRaHsqu1As2lCtAdQ504BljYOH0Tllibd2Rev0RvGlQ?e=DQM391";

// --- utils para /shares/{shareId} ---
function base64UrlEncode(input) {
  return btoa(input).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}
function buildShareId(url) {
  return "u!" + base64UrlEncode(url);
}

// Resuelve el vínculo compartido a driveId/itemId
async function resolveDriveItemFromShare(token, sharedUrl) {
  const shareId = buildShareId(sharedUrl);
  const resp = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!resp.ok) {
    const t = await resp.text();
    throw new Error(`No se pudo resolver el vínculo del CSV: ${resp.status} ${resp.statusText}\n${t}`);
  }
  const item = await resp.json();
  return { driveId: item.parentReference?.driveId, itemId: item.id, name: item.name };
}

// Descarga el CSV como texto y lo parsea con PapaParse
async function fetchCsvRowsFromShare(sharedUrl) {
  console.log("[INIT] Intentando obtener token para leer CSV…");
  const token = await getToken();
  if (!token) {
    // getToken() ya habrá lanzado loginRedirect si no había sesión.
    console.warn("[AUTH] No hay token todavía (seguramente va a redirigir).");
    return [];
  }

  const { driveId, itemId, name } = await resolveDriveItemFromShare(token, sharedUrl);
  console.log("[GRAPH] CSV resuelto:", { driveId, itemId, name });

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`No se pudo descargar el CSV: ${res.status} ${res.statusText}\n${t}`);
  }
  const csvText = await res.text();
  console.log("[CSV] bytes:", csvText.length);

  const parsed = Papa.parse(csvText, { header: true, skipEmptyLines: true });
  if (parsed.errors?.length) console.warn("PapaParse errors:", parsed.errors);
  const rows = parsed.data.map(row => {
    const out = {};
    for (const k of Object.keys(row)) out[k?.toLowerCase().trim()] = String(row[k] ?? "").trim();
    return out;
  });
  console.log("[CSV] filas parseadas:", rows.length);
  return rows;
}

// Llena un <select> con valores únicos de una columna
function fillSelectUniqueFromColumn(selectId, rows, columnKey, { sort = true, withEmpty = true, label = null } = {}) {
  const sel = document.getElementById(selectId);
  if (!sel) return;
  const set = new Set();
  rows.forEach(r => {
    const v = (r[columnKey] ?? "").trim();
    if (v) set.add(v);
  });
  let values = Array.from(set);
  if (sort) values.sort((a, b) => a.localeCompare(b, "es"));

  sel.innerHTML = "";
  if (withEmpty) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "-- Seleccione --";
    sel.appendChild(opt);
  }
  values.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = (typeof label === "function") ? label(v) : v;
    sel.appendChild(opt);
  });
  console.log(`[UI] ${selectId} cargado con ${values.length} opciones`);
}

// Carga el CSV y llena tus combos
async function cargarDotacionYLlenarSelects() {
  try {
    const rows = await fetchCsvRowsFromShare(DOTACION_CSV_LINK);
    if (!rows.length) {
      console.warn("[UI] No se cargaron filas (posible login pendiente).");
      return;
    }
    // columnas: dni | usuario_icc | asesor | supervisor | … | campania | …
    fillSelectUniqueFromColumn("campania",      rows, "campania");
    fillSelectUniqueFromColumn("dni_asesor",    rows, "dni");
    fillSelectUniqueFromColumn("usuario_icc",   rows, "usuario_icc");
    fillSelectUniqueFromColumn("nombre_asesor", rows, "asesor");
    fillSelectUniqueFromColumn("supervisor",    rows, "supervisor");
    // Si no existe "monitor" en ese CSV, no intentes llenarlo desde aquí.
    console.log("✔ Selects cargados");
  } catch (err) {
    console.error(err);
    alert("No se pudieron cargar las listas desde el CSV: " + err.message);
  }
}

// Helper para obtener valor (soporta inputs/select/textarea)
function val(selector) {
  const el = document.querySelector(selector);
  return el ? (el.value ?? "").toString().trim() : "";
}

// Construye el array con el orden de columnas de tu tabla de Excel
function construirFilaDesdeFormulario() {
  const fecha_llamada   = val('input[name="fecha_llamada"]');
  const campania        = val("#campania");
  const tipo_focalizado = "NTP";
  const fecha_revision  = val('input[name="fecha_revision"]');
  const telefono_any    = document.querySelectorAll('input')[4]?.value || "";
  const dni_cliente     = document.querySelectorAll('input')[5]?.value || "";
  const dni_asesor      = val("#dni_asesor");
  const usuario_icc     = val("#usuario_icc");
  const nombre_asesor   = val("#nombre_asesor");
  const supervisor      = val("#supervisor");
  const monitor         = val("#monitor"); // si no existe, quita esta línea y su columna
  const semana          = document.querySelectorAll('input')[12]?.value || "";
  const cuartil         = document.querySelectorAll('input')[13]?.value || "";
  const quien_comunica  = val("#quien_comunica");
  const motivo_contacto = val("#motivo_contacto");
  const tipo_plan       = val("#tipo_plan");
  const amerita_ofrec   = val("#amerita_ofrecimiento");
  const fecha_envio_iso = new Date().toISOString();

  return [
    fecha_llamada,
    campania,
    tipo_focalizado,
    fecha_revision,
    telefono_any,
    dni_cliente,
    dni_asesor,
    usuario_icc,
    nombre_asesor,
    supervisor,
    monitor,
    semana,
    cuartil,
    quien_comunica,
    motivo_contacto,
    tipo_plan,
    amerita_ofrec,
    fecha_envio_iso
  ];
}

// === Escritura en Excel (otra hoja/tabla) ===
const EXCEL_SHARED_LINK = "https://peruntp-my.sharepoint.com/:x:/g/personal/ntp_pyc_peruntp_onmicrosoft_com1/Efsjwdnnjl5Cl8I3eJqQehgBc2lN3DJo7fo0MBwQfFJKPw?e=RcSq2f";
const TABLE_NAME  = "fc_formulario_cross";

async function addRowToExcel(values) {
  const token = await getToken();
  if (!token) return;

  const { driveId, itemId } = await resolveDriveItemFromShare(token, EXCEL_SHARED_LINK);
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows/add`;

  const r = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ values: [values] })
  });

  if (!r.ok) {
    const err = await r.text();
    throw new Error(`Error al insertar fila: ${r.status} ${r.statusText}\n${err}`);
  }
  return r.json();
}

// === Inicio de la UI ===
document.addEventListener("DOMContentLoaded", () => {
  // Botón de login
  const btn = document.getElementById("btnLogin");
  if (btn) {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      console.log("[UI] Click en Iniciar sesión");
      login(); // lanza loginRedirect
    });
  }

  // Intento de carga (si no hay sesión, getToken dispara login y al volver se ejecuta de nuevo)
  cargarDotacionYLlenarSelects();

  // Manejo del submit
  const form = document.getElementById("formCroselling");
  const logEl = document.createElement("pre");
  form.parentElement.appendChild(logEl);
  const log = (m) => { logEl.textContent += m + "\n"; };

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    try {
      log("Enviando...");
      const fila = construirFilaDesdeFormulario();
      await addRowToExcel(fila);
      log("✅ Fila agregada en Excel.");
      form.reset();
    } catch (err) {
      console.error(err);
      log("❌ " + err.message);
    }
  });
});

import { login, debugToken } from "./msal-app.js";

document.addEventListener("DOMContentLoaded", async () => {
  // 1) prueba: ¿existe MSAL?
  console.log("[DEBUG] MSAL redirectUri:", window.__msal?.msalConfig?.auth?.redirectUri);

  // 2) intenta obtener token; si no hay sesión, msal hará redirect
  const tok = await debugToken();
  if (!tok) {
    console.log("[DEBUG] No token todavía. Pulsa el botón 'Iniciar sesión' si no te redirige solo.");
  }

  // 3) engancha el botón (por si quieres forzar login)
  document.getElementById("btnLogin")?.addEventListener("click", (e) => {
    e.preventDefault();
    console.log("[UI] Click en Iniciar sesión");
    login();
  });
});
