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
    globalData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    const syncInfo = document.getElementById('sync-info');
    if (syncInfo) {
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
      syncInfo.textContent = `Actualizado: ${timeStr}`;
    }

    renderKanban();
  };
  reader.readAsArrayBuffer(file);
}

function filterData() {
  const searchTerm = (document.getElementById('search-input')?.value || '').toLowerCase();

  return globalData.filter(item => {
    // Lectura flexible de la columna C1
    const canalVal = (item['C1'] || item['c1'] || item['CANAL'] || item['Canal'] || '').toString().toLowerCase();

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

    // Filtro por Búsqueda General
    const nv = (item['NV'] || item['N° NV'] || item['Nota Venta'] || '').toString().toLowerCase();
    const cliente = (item['Cliente'] || item['CLIENTE'] || '').toString().toLowerCase();
    const vendedor = (item['Vendedor'] || item['VENDEDOR'] || '').toString().toLowerCase();

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
    const estado = (item['Estado'] || item['ESTADO'] || '').toString().toLowerCase();

    if (estado.includes('entregado') || estado.includes('concluido')) {
      cols.entregado.push(item);
    } else if (estado.includes('programado')) {
      cols.programado.push(item);
    } else if (estado.includes('despacho') || estado.includes('transito') || estado.includes('tránsito')) {
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

    const nv = item['NV'] || item['N° NV'] || item['Nota Venta'] || 'N/A';
    const cliente = item['Cliente'] || item['CLIENTE'] || 'Cliente no especificado';
    const vendedor = item['Vendedor'] || item['VENDEDOR'] || 'Sin asignar';
    const fecha = item['Fecha NV'] || item['FECHA'] || 'N/A';
    const compromiso = item['Compromiso'] || item['FECHA COMPROMISO'] || 'N/A';

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
