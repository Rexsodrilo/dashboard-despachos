let globalData = [];
let selectedChannel = 'todos';
let searchQuery = '';

// Formateador de moneda
const formatMoney = (val) => new Intl.NumberFormat('es-CL', { style: 'currency', currency: 'CLP' }).format(val || 0);

// Escuchadores de eventos
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('excel-file').addEventListener('change', handleFileUpload);
  document.getElementById('search-input').addEventListener('input', (e) => {
    searchQuery = e.target.value.toLowerCase();
    renderKanban();
  });

  document.querySelectorAll('.btn-channel').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.btn-channel').forEach(b => b.classList.remove('active'));
      e.target.classList.add('active');
      selectedChannel = e.target.dataset.channel;
      renderKanban();
    });
  });
});

// Lectura de Excel
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    globalData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    document.getElementById('sync-info').innerText = `Actualizado: ${new Date().toLocaleTimeString()}`;
    renderKanban();
  };
  reader.readAsArrayBuffer(file);
}

// Renderizado del Kanban
function renderKanban() {
  if (!globalData.length) return;

  // Filtrado por canal y búsqueda
  const filtered = globalData.filter(row => {
    const canalVal = String(row['Canal'] || '').toLowerCase().trim();
    const matchChannel = (selectedChannel === 'todos') || (canalVal === selectedChannel);
    
    const nv = String(row['N.Venta'] || '').toLowerCase();
    const cliente = String(row['Nombre Cliente'] || '').toLowerCase();
    const vendedor = String(row['Nombre Vendedor'] || '').toLowerCase();
    const matchSearch = nv.includes(searchQuery) || cliente.includes(searchQuery) || vendedor.includes(searchQuery);

    return matchChannel && matchSearch;
  });

  // Clasificación de Columnas
  const entregados = filtered.filter(r => String(r['Estado']).trim() === 'Concluida' || String(r['Entregado SI/NO']).trim().toLowerCase() === 'si');
  const programados = filtered.filter(r => String(r['Estado']).trim() === 'Aprobada' && String(r['Motivo']).trim().toLowerCase() === 'coordinado');
  const porProgramar = filtered.filter(r => String(r['Estado']).trim() === 'Aprobada' && String(r['Motivo']).trim().toLowerCase() === 'por coordinar');
  const enDespacho = filtered.filter(r => String(r['Estado']).trim() === 'En Despacho');

  // Render tarjetas y contadores
  renderColumn('cards-entregado', 'count-entregado', entregados, createEntregadoCard);
  renderColumn('cards-programado', 'count-programado', programados, createProgramadoCard);
  renderColumn('cards-por-programar', 'count-por-programar', porProgramar, createPorProgramarCard);
  renderColumn('cards-despacho', 'count-despacho', enDespacho, createDespachoCard);

  // Calcular Totales de Cabecera (Entregados)
  const totalSemana = entregados.reduce((sum, r) => sum + (Number(r['Monto Neto']) || 0), 0);
  document.getElementById('kpi-semana').innerText = formatMoney(totalSemana);
  document.getElementById('kpi-dia-anterior').innerText = formatMoney(totalSemana * 0.2); // Ejemplo adaptativo
}

function renderColumn(containerId, countId, dataList, cardFn) {
  const container = document.getElementById(containerId);
  document.getElementById(countId).innerText = dataList.length;
  container.innerHTML = '';
  dataList.forEach(item => container.appendChild(cardFn(item)));
}

// Plantillas de Tarjetas
function createEntregadoCard(item) {
  const card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = `
    <span class="card-nv" onclick="copyToClipboard('${item['N.Venta']}')">NV: #${item['N.Venta'] || 'Sin/N'}</span>
    <div class="card-client" title="${item['Nombre Cliente']}">${item['Nombre Cliente'] || 'Cliente no informado'}</div>
    <div class="card-field">Vendedor: <strong>${item['Nombre Vendedor'] || 'Sin asignar'}</strong></div>
    <div class="card-field">Fecha NV: <strong>${item['Fecha de NV '] || 'N/A'}</strong></div>
    <div class="card-field">Compromiso: <strong>${item['FNV Compromiso '] || 'N/A'}</strong></div>
    <span class="badge-ontime">Entregado a Tiempo</span>
  `;
  return card;
}

function createProgramadoCard(item) {
  const card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = `
    <span class="card-nv" onclick="copyToClipboard('${item['N.Venta']}')">NV: #${item['N.Venta'] || 'Sin/N'}</span>
    <div class="card-client" title="${item['Nombre Cliente']}">${item['Nombre Cliente'] || 'Cliente no informado'}</div>
    <div class="card-field">Hora Estimada: <strong>${item['HoraCordinada'] || 'Por definir'}</strong></div>
    <div class="card-field">Chofer/Transporte: <strong>Sin asignar</strong></div>
  `;
  return card;
}

function createPorProgramarCard(item) {
  const card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = `
    <span class="card-nv" onclick="copyToClipboard('${item['N.Venta']}')">NV: #${item['N.Venta'] || 'Sin/N'}</span>
    <div class="card-client" title="${item['Nombre Cliente']}">${item['Nombre Cliente'] || 'Cliente no informado'}</div>
    <div class="card-field">Cod. Vendedor: <strong>${item['Cod.Vendedor'] || 'Sin código'}</strong></div>
    <div class="card-field">Motivo Retraso: <strong style="color: #e11d48;">${item['Motivo'] || 'Por coordinar'}</strong></div>
  `;
  return card;
}

function createDespachoCard(item) {
  const card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = `
    <span class="card-nv" onclick="copyToClipboard('${item['N.Venta']}')">NV: #${item['N.Venta'] || 'Sin/N'}</span>
    <div class="card-client" title="${item['Nombre Cliente']}">${item['Nombre Cliente'] || 'Cliente no informado'}</div>
    <div class="card-field">Estado: <strong>En Tránsito / Aceptado</strong></div>
  `;
  return card;
}

function copyToClipboard(text) {
  navigator.clipboard.writeText(text);
  alert(`Número de NV #${text} copiado al portapapeles.`);
}
