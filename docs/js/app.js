// js/app.js
async function cargarCSV() {
  const response = await fetch('data/dotacion_calidad.csv');
  const data = await response.text();

  // Convertir CSV a objetos
  const filas = data.trim().split('\n').map(linea => linea.split(','));
  const headers = filas[0];
  const registros = filas.slice(1).map(f => {
    const obj = {};
    headers.forEach((h, i) => (obj[h.trim()] = f[i]?.trim()));
    return obj;
  });

  return registros;
}

function llenarSelect(id, opciones, valorColumna) {
  const select = document.getElementById(id);
  if (!select) return;

  const valoresUnicos = [...new Set(opciones.map(o => o[valorColumna]).filter(Boolean))];

  select.innerHTML = '<option value="">-- Seleccionar --</option>';
  valoresUnicos.forEach(valor => {
    const option = document.createElement('option');
    option.value = valor;
    option.textContent = valor;
    select.appendChild(option);
  });
}

document.addEventListener('DOMContentLoaded', async () => {
  const asesores = await cargarCSV();

  llenarSelect('dni_asesor', asesores, 'DNI');
  llenarSelect('usuario_icc', asesores, 'Usuario ICC');
  llenarSelect('nombre_asesor', asesores, 'Asesor');
  llenarSelect('supervisor', asesores, 'Supervisor');
  llenarSelect('campania', asesores, 'CAMPAÑA');

  // También podrías llenar automáticamente el campo "monitor" si lo tienes en el CSV.
});
