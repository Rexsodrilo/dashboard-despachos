let globalDataRaw = [];
let groupedNvData = [];
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
    
    // Forzar la lectura de la hoja "NNVO_Corte"
    let targetSheetName = 'NNVO_Corte';
    if (!workbook.SheetNames.includes(targetSheetName)) {
      targetSheetName = workbook.SheetNames[0]; // fallback
    }
    const worksheet = workbook.Sheets[targetSheetName];
    
    // Convertir hoja a objeto JSON
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

    // Estandarizar claves (minúsculas y sin espacios)
    globalDataRaw = rawData.map(row => {
      const cleanRow = {};
      Object.keys(row).forEach(key => {
        cleanRow[key.trim().toLowerCase()] = row[key] ? row[key].toString().trim() : '';
      });
      return cleanRow;
    });

    // Agrupar filas por Nota de Venta (NVNumero)
    groupDataByNV();

    const syncInfo = document.getElementById('sync-info');
    if (syncInfo) {
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      syncInfo.textContent = `Actualizado: ${timeStr}`;
    }

    renderKanban();
  };
  reader.readAsArrayBuffer(file);
}

// Función auxiliar para leer propiedades de forma tolerante a nombres de columna
function getValue(item, keys) {
  for (const k of keys) {
    if (item[k] !== undefined && item[k] !== null && item[k] !== '') {
      return item[k];
    }
  }
  return '';
}

// Agrupa las líneas de productos bajo una misma tarjeta NV
function groupDataByNV() {
  const groups = {};

  globalDataRaw.forEach(item => {
    const nvEstado = getValue(item, ['nvestado', 'estado']).toUpperCase();

    // EXCLUIR estado 'N' (Nulo)
    if (nvEstado === 'N') return;

    const nvNumero = getValue(item, ['nvnumero', 'n.venta', 'nv']);
    if (!nvNumero) return;

    if (!groups[nvNumero]) {
      groups[nvNumero] = {
        nvNumero: nvNumero,
        nvEstado: nvEstado,
        tipoCliente: getValue(item, ['tipo de cliente', 'canal']),
        fechaCreacion: getValue(item, ['fechacreacion', 'fecha nv']),
        nomAux: getValue(item, ['nomaux', 'nombre cliente']),
        venDes: getValue(item, ['vendes', 'vendedor']),
        enPicking: getValue(item, ['en picking si/no']).toUpperCase() === 'SI',
        fechaCoordinada: getValue(item, ['fecha_coordinada', 'fechacoordinada']),
        horaCoordinada: getValue(item, ['hora_coordinada', 'horacoordinada']),
        cargado: getValue(item, ['cargado si/no']).toUpperCase() === 'SI',
        recibido: getValue(item, ['recibido si/no']).toUpperCase() === 'SI',
        items: []
      };
    }

    // Agregar detalle de productos
    groups[nvNumero].items.push({
      codProd: getValue(item, ['codprod', 'código']),
      detProd: getValue(item, ['detprod', 'producto']),
      cant: getValue(item, ['nvcant', 'cantidad']),
      stock: getValue(item, ['stockdisponible', 'stock'])
    });
  });

  groupedNvData = Object.values(groups);
}

function filterData() {
  const searchTerm = (document.getElementById('search-input')?.value || '').toLowerCase();

  return groupedNvData.filter(item => {
    // 1. Filtrar por Canal / Tipo de cliente (Columna D)
    const canalVal = item.tipoCliente.toLowerCase();
    let matchesChannel = false;
    
    if (activeChannel === 'todos') {
      matchesChannel = true;
    } else if (activeChannel === 'retail') {
      matchesChannel = canalVal.includes('retail');
    } else if (activeChannel === 'a despachar' || activeChannel === 'despacho') {
      matchesChannel = canalVal.includes('despach');
    } else if (activeChannel === 'a retirar cliente' || activeChannel === 'retiro') {
      matchesChannel = canalVal.includes('retir');
    } else if (activeChannel === 'ecommerce') {
      matchesChannel = canalVal.includes('ecom') || canalVal.includes('web');
    } else if (activeChannel === 'pendiente') {
      // Filtro especial para pendientes ('P')
      matchesChannel = item.nvEstado === 'P';
    }

    // 2. Búsqueda por NV, Cliente o Vendedor
    const nv = item.nvNumero.toLowerCase();
    const cliente = item.nomAux.toLowerCase();
    const vendedor = item.venDes.toLowerCase();

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
    despacho: [],
    pendiente: []
  };

  filtered.forEach(item => {
    const estado = item.nvEstado;
    const tieneFechaCoord = item.fechaCoordinada !== '' && item.fechaCoordinada !== '0';

    // 1. PENDIENTE (Estado 'P' - En espera de aprobación o modificación)
    if (estado === 'P') {
      cols.pendiente.push(item);
    }
    // 2. ENTREGADO / CONCLUIDO (Estado 'C')
    else if (estado === 'C') {
      cols.entregado.push(item);
    } 
    // 3. EN DESPACHO / TRÁNSITO (Estado 'A' + Picking: SI + Fecha Coordinada + Cargado: SI)
    else if (estado === 'A' && item.enPicking && tieneFechaCoord && item.cargado) {
      cols.despacho.push(item);
    } 
    // 4. PROGRAMADO (Estado 'A' + Picking: SI + Fecha Coordinada)
    else if (estado === 'A' && item.enPicking && tieneFechaCoord) {
      cols.programado.push(item);
    } 
    // 5. POR PROGRAMAR (Estado 'A' sin cumplir condiciones anteriores)
    else {
      cols.porProgramar.push(item);
    }
  });

  updateColumnUI('cards-entregado', 'count-entregado', cols.entregado);
  updateColumnUI('cards-programado', 'count-programado', cols.programado);
  updateColumnUI('cards-por-programar', 'count-por-programar', cols.porProgramar);
  updateColumnUI('cards-despacho', 'count-despacho', cols.despacho);
  updateColumnUI('cards-pendiente', 'count-pendiente', cols.pendiente);
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

    const horarioText = (item.horaCoordinada && item.horaCoordinada !== '0') 
      ? item.horaCoordinada 
      : 'Horario abierto';

    // Construcción del HTML de productos para el acordeón desplegable
    let itemsTable = `
      <div class="card-items-detail" style="display:none; margin-top:10px; font-size:12px; border-top:1px solid #ddd; padding-top:5px;">
        <table style="width:100%; border-collapse:collapse;">
          <thead>
            <tr style="text-align:left; background:#f5f5f5;">
              <th>Cód</th>
              <th>Producto</th>
              <th>Cant</th>
              <th>Stock</th>
            </tr>
          </thead>
          <tbody>
    `;

    item.items.forEach(prod => {
      itemsTable += `
        <tr>
          <td>${prod.codProd}</td>
          <td>${prod.detProd}</td>
          <td>${prod.cant}</td>
          <td>${prod.stock}</td>
        </tr>
      `;
    });

    itemsTable += `</tbody></table></div>`;

    card.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center;">
        <span class="card-nv">NV: #${item.nvNumero}</span>
        <span class="badge-ontime" style="cursor:pointer;" onclick="toggleDetails(this)">📦 ${item.items.length} Prod. ▾</span>
      </div>
      <div class="card-client">${item.nomAux}</div>
      <div class="card-field">Vendedor: <strong>${item.venDes}</strong></div>
      <div class="card-field">Fecha NV: <strong>${item.fechaCreacion}</strong></div>
      <div class="card-field">Horario: <strong>${horarioText}</strong></div>
      <div class="card-field">Estado NV: <strong>${item.nvEstado}</strong></div>
      ${itemsTable}
    `;

    container.appendChild(card);
  });
}

// Función global para expandir o colapsar la lista de productos
window.toggleDetails = function(btnElement) {
  const card = btnElement.closest('.card');
  const details = card.querySelector('.card-items-detail');
  if (details.style.display === 'none') {
    details.style.display = 'block';
    btnElement.textContent = btnElement.textContent.replace('▾', '▴');
  } else {
    details.style.display = 'none';
    btnElement.textContent = btnElement.textContent.replace('▴', '▾');
  }
};
