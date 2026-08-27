let globalDataRaw = [];
let groupedNvData = [];
let activeChannel = 'todos';
let maxFileDate = new Date();

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

function isFlagYes(val) {
  if (!val) return false;
  const str = val.toString().trim().toUpperCase();
  return str === 'SI' || str === '1' || str === 'S' || str === 'TRUE';
}

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    
    let targetSheetName = 'NNVO_Corte';
    if (!workbook.SheetNames.includes(targetSheetName)) {
      targetSheetName = workbook.SheetNames[0];
    }
    const worksheet = workbook.Sheets[targetSheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

    globalDataRaw = rawData.map(row => {
      const cleanRow = {};
      Object.keys(row).forEach(key => {
        cleanRow[key.trim().toLowerCase()] = row[key] ? row[key].toString().trim() : '';
      });
      return cleanRow;
    });

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

function getValue(item, keys) {
  for (const k of keys) {
    if (item[k] !== undefined && item[k] !== null && item[k] !== '') {
      return item[k];
    }
  }
  return '';
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
}

function groupDataByNV() {
  const groups = {};
  let latestDate = new Date(0);

  globalDataRaw.forEach(item => {
    const nvEstado = getValue(item, ['nvestado', 'estado']).toUpperCase();

    // Excluir 'N' (Nulo)
    if (nvEstado === 'N') return;

    const nvNumero = getValue(item, ['nvnumero', 'n.venta', 'nv']);
    if (!nvNumero) return;

    const strFecha = getValue(item, ['fechacreacion', 'fecha nv']);
    const dateObj = parseDate(strFecha);
    if (dateObj && dateObj > latestDate) {
      latestDate = dateObj;
    }

    const montoRaw = getValue(item, ['total linea', 'nvtotlinea']).replace(/[^0-9.-]+/g, "");
    const montoVal = parseFloat(montoRaw) || 0;

    const pickingVal = getValue(item, ['en picking si/no']);
    const cargadoVal = getValue(item, ['cargado si/no']);
    const recibidoVal = getValue(item, ['recibido si/no']);
    const fechaCoordVal = getValue(item, ['fecha_coordinada', 'fechacoordinada']);
    
    const tipoCliente = getValue(item, ['tipo de cliente', 'canal']);
    const valorDespacho = getValue(item, ['valor a depacho', 'valor despacho']);

    if (!groups[nvNumero]) {
      groups[nvNumero] = {
        nvNumero: nvNumero,
        nvEstado: nvEstado,
        tipoCliente: tipoCliente,
        valorDespacho: valorDespacho,
        fechaCreacion: strFecha,
        fechaObj: dateObj,
        nomAux: getValue(item, ['nomaux', 'nombre cliente']),
        venDes: getValue(item, ['vendes', 'vendedor']),
        ordenCompra: getValue(item, ['orden de compra', 'oc']),
        enPicking: isFlagYes(pickingVal),
        pickingRaw: pickingVal.toUpperCase(),
        tieneFechaCoord: isFlagYes(fechaCoordVal) || (fechaCoordVal !== '' && fechaCoordVal !== '0'),
        horaCoordinada: getValue(item, ['hora_coordinada', 'horacoordinada']),
        cargado: isFlagYes(cargadoVal),
        recibido: isFlagYes(recibidoVal),
        motivoP: getValue(item, ['motivop', 'motivo']),
        montoTotal: 0,
        items: []
      };
    }

    groups[nvNumero].montoTotal += montoVal;

    groups[nvNumero].items.push({
      codProd: getValue(item, ['codprod', 'código']),
      detProd: getValue(item, ['detprod', 'producto']),
      cant: getValue(item, ['nvcant', 'cantidad']),
      stock: getValue(item, ['stockdisponible', 'stock'])
    });
  });

  maxFileDate = latestDate.getTime() > 0 ? latestDate : new Date();
  
  // Evaluar regla especial de "Bajo valor de despacho" por grupo
  groupedNvData = Object.values(groups).map(item => {
    const esHabitual = item.tipoCliente.toUpperCase().includes('HABITUAL');
    const esBajoMonto = item.valorDespacho.toLowerCase().includes('bajo');
    const esEstadoA o P = item.nvEstado === 'A' || item.nvEstado === 'P';

    if (esEstadoA o P && esHabitual && esBajoMonto) {
      item.esBajoMontoPendiente = true;
      item.motivoP = 'Bajo valor del despacho';
    } else {
      item.esBajoMontoPendiente = false;
    }
    return item;
  });
}

function filterData() {
  const searchTerm = (document.getElementById('search-input')?.value || '').toLowerCase();

  return groupedNvData.filter(item => {
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
      matchesChannel = item.nvEstado === 'P' || item.pickingRaw === 'NO' || item.esBajoMontoPendiente;
    }

    const nv = item.nvNumero.toLowerCase();
    const cliente = item.nomAux.toLowerCase();
    const vendedor = item.venDes.toLowerCase();
    const oc = item.ordenCompra.toLowerCase();

    const matchesSearch = !searchTerm || nv.includes(searchTerm) || cliente.includes(searchTerm) || vendedor.includes(searchTerm) || oc.includes(searchTerm);

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

  let totalMontoSemana = 0;
  let totalMontoDiaAnt = 0;

  const msPerDay = 24 * 60 * 60 * 1000;

  filtered.forEach(item => {
    const estado = item.nvEstado;
    
    let diffDays = 0;
    if (item.fechaObj) {
      diffDays = Math.floor((maxFileDate - item.fechaObj) / msPerDay);
    }

    if (diffDays <= 7) totalMontoSemana += item.montoTotal;
    if (diffDays === 1) totalMontoDiaAnt += item.montoTotal;

    // Jerarquía de Clasificación en Columnas:

    // 1. ENTREGADO / CONCLUIDO (Prioridad Absoluta): Estado 'C' O Recibido = SI
    if (estado === 'C' || item.recibido) {
      if (diffDays <= 7) cols.entregado.push(item);
    } 
    // 2. PENDIENTE: Estado 'P', Picking 'NO', o Regla 'Bajo valor de despacho'
    else if (estado === 'P' || item.pickingRaw === 'NO' || item.esBajoMontoPendiente) {
      cols.pendiente.push(item);
    } 
    // 3. EN DESPACHO / TRÁNSITO: Estado 'A' + Picking SI + Fecha Coord + Cargado SI
    else if (estado === 'A' && item.enPicking && item.tieneFechaCoord && item.cargado) {
      if (diffDays <= 1) cols.despacho.push(item);
    } 
    // 4. PROGRAMADO: Estado 'A' + Picking SI + Fecha Coord
    else if (estado === 'A' && item.enPicking && item.tieneFechaCoord) {
      if (diffDays <= 1) cols.programado.push(item);
    } 
    // 5. POR PROGRAMAR: Restantes en Estado 'A'
    else {
      if (diffDays <= 1) cols.porProgramar.push(item);
    }
  });

  updateColumnUI('cards-entregado', 'count-entregado', cols.entregado);
  updateColumnUI('cards-programado', 'count-programado', cols.programado);
  updateColumnUI('cards-por-programar', 'count-por-programar', cols.porProgramar);
  updateColumnUI('cards-despacho', 'count-despacho', cols.despacho);
  updateColumnUI('cards-pendiente', 'count-pendiente', cols.pendiente);

  updateTotalsUI(totalMontoSemana, totalMontoDiaAnt);
}

function updateTotalsUI(semana, diaAnt) {
  const elSemana = document.getElementById('total-semana');
  const elDiaAnt = document.getElementById('total-dia-anterior');

  const fmt = (val) => '$' + Math.round(val).toLocaleString('es-CL');

  if (elSemana) elSemana.textContent = fmt(semana);
  if (elDiaAnt) elDiaAnt.textContent = fmt(diaAnt);
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

    const ocHtml = (item.ordenCompra && item.ordenCompra !== '0') 
      ? `<div class="card-field" style="color:#d9534f; font-weight:bold;">O.C.: ${item.ordenCompra}</div>` 
      : '';

    const esPendiente = item.nvEstado === 'P' || item.pickingRaw === 'NO' || item.esBajoMontoPendiente;
    const motivoHtml = (esPendiente && item.motivoP && item.motivoP !== '0')
      ? `<div class="card-field" style="background:#fff3cd; color:#856404; padding:2px 5px; border-radius:3px; margin-top:4px;">Motivo: <strong>${item.motivoP}</strong></div>`
      : '';

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
      <div class="card-client">${item.nomAux || 'Cliente no especificado'}</div>
      ${ocHtml}
      <div class="card-field">Vendedor: <strong>${item.venDes || 'Sin Vendedor'}</strong></div>
      <div class="card-field">Fecha NV: <strong>${item.fechaCreacion || 'N/A'}</strong></div>
      <div class="card-field">Horario: <strong>${horarioText}</strong></div>
      <div class="card-field">Monto NV: <strong>$${Math.round(item.montoTotal).toLocaleString('es-CL')}</strong></div>
      <div class="card-field">Estado: <strong>${item.nvEstado}</strong></div>
      ${motivoHtml}
      ${itemsTable}
    `;

    container.appendChild(card);
  });
}

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
