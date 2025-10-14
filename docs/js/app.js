/***************
 * CONFIG MSAL *
 ***************/
const TENANT_ID = "f88cba71-d226-4d73-be87-c972ecafc1f5"; // <- tu tenant
const CLIENT_ID = "a6e88ae9-4e68-494c-8207-c0b81401ba46";  // <- tu app (cliente)

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    // Debe estar registrado EXACTAMENTE en Azure como Redirect URI (SPA)
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: { cacheLocation: "localStorage" }
};

const SCOPES_READ  = ["Files.Read", "offline_access"];
const SCOPES_WRITE = ["Files.ReadWrite", "offline_access"]; // para guardar en Excel

const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;

async function initAuth() {
  try { await msalInstance.handleRedirectPromise(); } catch {}
  const accs = msalInstance.getAllAccounts();
  if (accs.length) account = accs[0];
}
async function login(scopes) {
  if (!account) {
    await msalInstance.loginPopup({ scopes });
    await initAuth();
  }
}
async function getToken(scopes) {
  if (!account) throw new Error("Inicia sesión primero.");
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes, account });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ scopes });
    return r.accessToken;
  }
}

/*************************
 * RUTAS (formato Graph) *
 *************************/
// Tu CSV (para desplegables) – ruta relativa en OneDrive del usuario que inicia sesión
const CSV_PATH = "/Documents/FORMULARIOS/data/dotacion_calidad.csv";

// Tu Excel de destino (para guardar respuestas)
const XLSX_PATH = "/Documents/FORMULARIOS/data/calidad_focalizado_cross.xlsx";
const TABLE_NAME = "Tabla1"; // <- cambia si tu tabla tiene otro nombre

/***********************
 * HELPERS de Microsoft Graph
 ***********************/
async function downloadCsvByPath(pathRel, token) {
  // Lee un archivo (texto) del OneDrive del usuario actual
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(pathRel)}:/content`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`Graph ${r.status}: ${await r.text()}`);
  return await r.text();
}

async function getItemIdsByPath(pathRel, token) {
  // Devuelve { driveId, itemId } para un item por ruta
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:${encodeURI(pathRel)}`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`No se encontró: ${pathRel}\n${await r.text()}`);
  const j = await r.json();
  return { driveId: j.parentReference.driveId, itemId: j.id };
}

async function addRowToExcelTable(driveId, itemId, tableName, token, valuesArray) {
  // Inserta una fila en una Tabla de Excel
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables/${encodeURIComponent(tableName)}/rows/add`;
  const body = JSON.stringify({ values: [ valuesArray ] });
  const r = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body
  });
  if (!r.ok) throw new Error(`Error al insertar fila: ${await r.text()}`);
  return r.json();
}

/*****************
 * UI utilities  *
 *****************/
function fillSelect(el, values) {
  el.innerHTML = "";
  const uniq = [...new Set(values.filter(v => v != null && v !== ""))].sort();
  uniq.unshift(""); // opción vacía
  for (const v of uniq) {
    const opt = document.createElement("option");
    opt.value = v; opt.textContent = v;
    el.appendChild(opt);
  }
}

/**********************
 * Cargar desplegables *
 **********************/
async function cargarDesplegables() {
  await login(SCOPES_READ);
  const token = await getToken(SCOPES_READ);

  const csvText = await downloadCsvByPath(CSV_PATH, token);
  const parsed = Papa.parse(csvText, { header: true, skipEmptyLines: true });
  const rows = parsed.data; // [{...},{...}]

  // Mapea aquí los nombres de columnas del CSV (ajústalos a tu archivo real)
  // Ejemplo de columnas: Campaña, DNI_Asesor, Usuario_ICC, Nombre_Asesor, Supervisor, Monitor, Quien_Comunica, Motivo_Contacto, Tipo_Plan_SVA, Amerita_Ofrecimiento
  const get = (col) => rows.map(r => r[col]);

  fillSelect(document.getElementById("campania"),        get("Campaña"));
  fillSelect(document.getElementById("dni_asesor"),      get("DNI_Asesor"));
  fillSelect(document.getElementById("usuario_icc"),     get("Usuario_ICC"));
  fillSelect(document.getElementById("nombre_asesor"),   get("Nombre_Asesor"));
  fillSelect(document.getElementById("supervisor"),      get("Supervisor"));
  fillSelect(document.getElementById("monitor"),         get("Monitor"));

  fillSelect(document.getElementById("quien_comunica"),  get("Quien_Comunica"));
  fillSelect(document.getElementById("motivo_contacto"), get("Motivo_Contacto"));
  fillSelect(document.getElementById("tipo_plan"),       get("Tipo_Plan_SVA"));
  fillSelect(document.getElementById("amerita_ofrecimiento"), get("Amerita_Ofrecimiento"));
}

/*****************
 * Guardar form  *
 *****************/
async function guardarFormulario(e) {
  e.preventDefault();

  // 1) Capturar datos
  const fd = new FormData(e.target);
  // Si pasas fechas por querystring, precárgalas así:
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.has("fecha_llamada"))  fd.set("fecha_llamada",  urlParams.get("fecha_llamada"));
  if (urlParams.has("fecha_revision")) fd.set("fecha_revision", urlParams.get("fecha_revision"));

  // 2) (Opcional) valida y ordena columnas según el orden de la Tabla en Excel
  // EJEMPLO de orden (ajústalo al orden real de columnas en tu Tabla1):
  const values = [
    fd.get("fecha_llamada") || "",
    fd.get("campania") || "",
    "NTP", // Tipo Focalizado (readonly en tu HTML)
    fd.get("fecha_revision") || "",
    fd.get("dni_asesor") || "",
    fd.get("usuario_icc") || "",
    fd.get("nombre_asesor") || "",
    fd.get("supervisor") || "",
    fd.get("monitor") || "",
    fd.get("quien_comunica") || "",
    fd.get("motivo_contacto") || "",
    fd.get("tipo_plan") || "",
    fd.get("amerita_ofrecimiento") || "",
    // ... añade el resto de campos que quieras guardar, en el MISMO orden de columnas del Excel
  ];

  // 3) Token con permisos de escritura
  await login(SCOPES_WRITE);
  const token = await getToken(SCOPES_WRITE);

  // 4) Resolver drive/item del Excel, e insertar fila
  const { driveId, itemId } = await getItemIdsByPath(XLSX_PATH, token);
  await addRowToExcelTable(driveId, itemId, TABLE_NAME, token, values);

  alert("✅ Respuesta guardada en el Excel.");
}

/****************
 * Arranque     *
 ****************/
document.addEventListener("DOMContentLoaded", async () => {
  await initAuth();

  // Pre-cargar fechas desde querystring (si vienen en la URL)
  const p = new URLSearchParams(window.location.search);
  if (p.has("fecha_llamada"))  document.querySelector('[name="fecha_llamada"]').value  = p.get("fecha_llamada");
  if (p.has("fecha_revision")) document.querySelector('[name="fecha_revision"]').value = p.get("fecha_revision");

  // Cargar desplegables desde el CSV
  try {
    await cargarDesplegables();
  } catch (err) {
    console.error(err);
    alert(err.message);
  }

  // Hook al submit del formulario
  const form = document.getElementById("formCroselling");
  form.addEventListener("submit", guardarFormulario);
});
