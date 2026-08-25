let globalData = [];
let activeChannel = 'todos';

document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
});

function setupEventListeners() {
  const fileInput = document.getElementById('excel-file');
  if (fileInput) {
    fileInput.addEventListener('change', handleFileUpload);
  }

  const channelButtons = document.querySelectorAll('.btn-channel');
  channelButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      channelButtons.forEach(b => b.classList.remove('active'));
      e.target.classList.add('active');
      activeChannel = e.target.getAttribute('data-channel');
      renderKanban();
    });
  });

  const searchInput = document.getElementById('search-input');
  if (searchInput) {
    searchInput.addEventListener('input', renderKanban);
  }
}

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheet];
    
    // Convertir la hoja directamente a un array de objetos utilizando los encabezados nativos
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

    // Filtrar filas completamente vacías o encabezados repetidos
    globalData = jsonData.filter(row => {
      const rowString = JSON.stringify(row).toLowerCase();
      return !rowString.includes('nota de venta') && Object.values(row).some(val => val !== '');
    });

    const syncInfo = document.getElementById('sync-info');
    if (syncInfo) {
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
      syncInfo.textContent = `Actualizado: ${timeStr}`;
    }

    renderKanban();
  };
  reader.readAsArrayBuffer(file);
}

// Búsqueda insensible a mayúsculas/minúsculas y espacios en las llaves del objeto
function getFieldValue(item, posiblesNombres) {
  const itemKeys = Object.keys(item);
  for (const nombre of posiblesNombres) {
    const targetKey = nombre.toLowerCase().trim();
    const foundKey = itemKeys.find(k => k.toLowerCase().trim() === targetKey);
    if (foundKey && item[foundKey] !== undefined && item[foundKey] !== null && item[foundKey] !== '') {
      return item[foundKey].toString().trim();
    }
  }
  return '';
}

function filterData() {
  const searchTerm = (document.getElementById('search-input')?.value || '').toLowerCase();

  return globalData.filter(item => {
    // Busca en la columna C1 o variaciones de Canal
    const canalVal = getFieldValue(item, ['c1', 'canal', 'canal venta', 'canal de venta']).toLowerCase();

    // Filtro por Canal de Venta
    let matchesChannel = false;
    if (activeChannel === 'todos') {
      matchesChannel = true;
    } else if (activeChannel === 'retail') {
      matchesChannel = canalVal.includes('retail');
    } else if (activeChannel === 'a despachar') {
      matchesChannel = canalVal.includes('despachar') || canalVal.includes('despacho');
    } else if (activeChannel === 'a retirar por cliente') {
      matchesChannel = canalVal.includes('retirar') || canalVal.includes('retiro');
    } else if (activeChannel === 'ecommerce') {
      matchesChannel = canalVal.includes('ecommerce') || canalVal.includes('e-commerce') || canalVal.includes('web');
    }

    // Filtro por Buscador General
    const nv = getFieldValue(item, ['nv', 'n° nv', 'nota venta', 'nota de venta', 'nro nv']).toLowerCase();
    const cliente = getFieldValue(item, ['cliente', 'nombre cliente', 'razon social', 'razón social']).toLowerCase();
    const vendedor = getFieldValue(item, ['vendedor', 'nombre vendedor']).toLowerCase();

    const matchesSearch = !searchTerm || nv.includes(searchTerm) || cliente.includes(searchTerm) || vendedor.includes(searchTerm);

    return matchesChannel && matchesSearch;
  });
}

function renderKanban() {
  const filtered = filterData();

  const cols = {
    entregado: [],
    programado: [],
    porProgramar: [],
    despacho: []
  };

  filtered.forEach(item => {
    const estado = getFieldValue(item, ['estado', 'estado nv', 'estado logistico', 'situacion', 'est.']).toLowerCase();

    if (estado.includes('entregado') || estado.includes('concluido') || estado.includes('finalizado')) {
      cols.entregado.push(item);
    } else if (estado.includes('programado')) {
      cols.programado.push(item);
    } else if (estado.includes('despacho') || estado.includes('transito') || estado.includes('tránsito') || estado.includes('en ruta')) {
      cols.despacho.push(item);
    } else {
      cols.porProgramar.push(item);
    }
  });

  updateColumnUI('cards-entregado', 'count-entregado', cols.entregado);
  updateColumnUI('cards-programado', 'count-programado', cols.programado);
  updateColumnUI('cards-por-programar', 'count-por-programar', cols.porProgramar);
  updateColumnUI('cards-despacho', 'count-despacho', cols.despacho);
}

function updateColumnUI(containerId, countId, items) {
  const container = document.getElementById(containerId);
  const countBadge = document.getElementById(countId);

  if (countBadge) countBadge.textContent = items.length;
  if (!container) return;

  container.innerHTML = '';

  items.forEach(item => {
    const card = document.createElement('div');
    card.className = 'card';

    // Extracción limpia de campos probando variantes comunes
    const nv = getFieldValue(item, ['nv', 'n° nv', 'nota venta', 'nota de venta', 'nro nv']) || 'N/A';
    const cliente = getFieldValue(item, ['cliente', 'nombre cliente', 'razon social', 'razón social']) || 'Cliente no especificado';
    const vendedor = getFieldValue(item, ['vendedor', 'nombre vendedor']) || 'Sin asignar';
    const fecha = getFieldValue(item, ['fecha nv', 'fecha', 'fecha de emisión', 'f. nv']) || 'N/A';
    const compromiso = getFieldValue(item, ['compromiso', 'fecha compromiso', 'f. compromiso']) || 'N/A';

    card.innerHTML = `
      <span class="card-nv">NV: #${nv}</span>
      <div class="card-client">${cliente}</div>
      <div class="card-field">Vendedor: <strong>${vendedor}</strong></div>
      <div class="card-field">Fecha NV: <strong>${fecha}</strong></div>
      <div class="card-field">Compromiso: <strong>${compromiso}</strong></div>
      <span class="badge-ontime">Entregado a Tiempo</span>
    `;

    container.appendChild(card);
  });
}
