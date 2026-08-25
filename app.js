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
    
    // Convierte la hoja tomando los nombres de la primera fila
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

    // Normaliza los nombres de las columnas (elimina espacios y convierte a minúsculas)
    globalData = rawData.map(row => {
      const cleanRow = {};
      Object.keys(row).forEach(key => {
        cleanRow[key.trim().toLowerCase()] = row[key] ? row[key].toString().trim() : '';
      });
      return cleanRow;
    });

    const syncInfo = document.getElementById('sync-info');
    if (syncInfo) {
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      syncInfo.textContent = `Actualizado: ${timeStr}`;
    }

    renderKanban();
  };
  reader.readAsArrayBuffer(file);
}

// Función auxiliar para leer campos tolerando nombres variados
function getValue(item, keys) {
  for (const k of keys) {
    if (item[k] !== undefined && item[k] !== '') {
      return item[k];
    }
  }
  return '';
}

function filterData() {
  const searchTerm = (document.getElementById('search-input')?.value || '').toLowerCase();

  return globalData.filter(item => {
    // 1. Canal de Venta (Columna C: Canal)
    const canalVal = getValue(item, ['canal', 'c1']).toLowerCase();

    // 2. Filtros por botones laterales
    let matchesChannel = false;
    if (activeChannel === 'todos') {
      matchesChannel = true;
    } else if (activeChannel === 'retail') {
      matchesChannel = canalVal.includes('retail');
    } else if (activeChannel === 'a despachar') {
      matchesChannel = canalVal.includes('despachar') || canalVal.includes('despacho');
    } else if (activeChannel === 'a retirar cliente' || activeChannel === 'a retirar por cliente') {
      matchesChannel = canalVal.includes('retirar') || canalVal.includes('retiro');
    } else if (activeChannel === 'ecommerce') {
      matchesChannel = canalVal.includes('ecommerce') || canalVal.includes('e-commerce') || canalVal.includes('web');
    }

    // 3. Filtro de búsqueda por N.Venta o Nombre Cliente
    const nv = getValue(item, ['n.venta', 'n° nv', 'nv']).toLowerCase();
    const cliente = getValue(item, ['nombre cliente', 'cliente']).toLowerCase();

    const matchesSearch = !searchTerm || nv.includes(searchTerm) || cliente.includes(searchTerm);

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
    const estadoGeneral = getValue(item, ['estado', 'estado nv']).toLowerCase();
    const estadoLogistico = getValue(item, ['estado logistico', 'despacho', 'situacion']).toLowerCase();
    
    const textoCompleto = `${estadoGeneral} ${estadoLogistico}`;

    const esAprobado = estadoGeneral.includes('aprobad');
    const esProgramado = textoCompleto.includes('programad');
    const esDespacho = textoCompleto.includes('despacho') || textoCompleto.includes('transito') || textoCompleto.includes('tránsito');

    // 1. Entregado / Concluido
    if (estadoGeneral.includes('concluid') || estadoGeneral.includes('entregad') || estadoGeneral.includes('finalizad')) {
      cols.entregado.push(item);
    } 
    // 2. En Despacho / Tránsito (Requiere estar Aprobado Y Programado Y en Despacho)
    else if (esAprobado && esProgramado && esDespacho) {
      cols.despacho.push(item);
    } 
    // 3. Programado (Requiere estar Aprobado Y Programado)
    else if (esAprobado && esProgramado) {
      cols.programado.push(item);
    } 
    // 4. Por Programar (Resto de casos, ej. Aprobadas sin programar)
    else {
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

    const nv = getValue(item, ['n.venta', 'nv']) || 'N/A';
    const cliente = getValue(item, ['nombre cliente']) || 'Cliente no especificado';
    const fecha = getValue(item, ['fecha de nv', 'fecha nv']) || 'N/A';
    const estado = getValue(item, ['estado']) || 'Sin Estado';

    card.innerHTML = `
      <span class="card-nv">NV: #${nv}</span>
      <div class="card-client">${cliente}</div>
      <div class="card-field">Fecha NV: <strong>${fecha}</strong></div>
      <div class="card-field">Estado: <strong>${estado}</strong></div>
      <span class="badge-ontime">Procesado</span>
    `;

    container.appendChild(card);
  });
}
