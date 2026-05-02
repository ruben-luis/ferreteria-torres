// ===== DATOS =====
// ===== CARGA DIFERIDA DE LIBRERÍAS =====
const _scripts = {};
function cargarScript(url) {
  if (_scripts[url]) return _scripts[url];
  _scripts[url] = new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[src="${url}"]`);
    if (existing) { resolve(); return; }
    const s = document.createElement('script');
    s.src = url;
    s.onload  = resolve;
    s.onerror = () => reject(new Error('No se pudo cargar: ' + url));
    document.head.appendChild(s);
  });
  return _scripts[url];
}

let _pdfListo   = false;
let _excelListo = false;

async function prepararPDF() {
  if (_pdfListo) return;
  await cargarScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  await cargarScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js');
  await cargarScript('logo-data.js');
  _pdfListo = true;
}

async function prepararExcel() {
  if (_excelListo) return;
  await cargarScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
  _excelListo = true;
}

// ===== DATOS =====
let productos        = JSON.parse(localStorage.getItem('ferreteriaProductos'))        || [];
let ventas           = JSON.parse(localStorage.getItem('ferreteriaVentas'))           || [];
let cortes           = JSON.parse(localStorage.getItem('ferreteriaCortes'))           || [];
let ordenes          = JSON.parse(localStorage.getItem('ferreteriaOrdenes'))          || [];
let pagosProveedor   = JSON.parse(localStorage.getItem('ferreteriaPagosProveedor'))   || [];
let folioActual      = parseInt(localStorage.getItem('ferreteriaFolio')      || '0');
let ordenFolioActual = parseInt(localStorage.getItem('ferreteriaOrdenFolio') || '0');

function siguienteFolio() {
  folioActual++;
  localStorage.setItem('ferreteriaFolio', String(folioActual));
  return String(folioActual).padStart(3, '0');
}

function siguienteOrdenFolio() {
  ordenFolioActual++;
  localStorage.setItem('ferreteriaOrdenFolio', String(ordenFolioActual));
  return 'OC-' + String(ordenFolioActual).padStart(4, '0');
}
let carrito    = [];
let idEliminar = null;
let catActivaPos = '';
let pagInv     = 1;
let pagAlertas  = 1;
let pagReabasto = 1;
let pagPos      = 1;
let _alertasTodas     = [];
let _reabastoFiltrados = [];
let _reabastoState     = {}; // { id: { checked, cantidad } }
const POR_PAG_INV      = 50;
const POR_PAG_ALERTAS  = 25;
const POR_PAG_REABASTO = 25;
const POR_PAG_POS      = 15;

// ===== INICIO =====
document.addEventListener('DOMContentLoaded', () => {
  mostrarFecha();
  renderTodo();
  iniciarEventos();
});

let _guardarTimer;
function guardar() {
  clearTimeout(_guardarTimer);
  _guardarTimer = setTimeout(() => {
    localStorage.setItem('ferreteriaProductos',      JSON.stringify(productos));
    localStorage.setItem('ferreteriaVentas',         JSON.stringify(ventas));
    localStorage.setItem('ferreteriaCortes',         JSON.stringify(cortes));
    localStorage.setItem('ferreteriaOrdenes',        JSON.stringify(ordenes));
    localStorage.setItem('ferreteriaPagosProveedor', JSON.stringify(pagosProveedor));
  }, 200);
}

function toast(msg, tipo = 'ok', ms = 3500) {
  const el = document.createElement('div');
  el.className = `toast toast-${tipo}`;
  el.textContent = msg;
  document.body.appendChild(el);
  requestAnimationFrame(() => requestAnimationFrame(() => el.classList.add('visible')));
  setTimeout(() => { el.classList.remove('visible'); setTimeout(() => el.remove(), 300); }, ms);
}

// ===== RESPALDO =====
const _CLAVES_RESPALDO = [
  'ferreteriaProductos',
  'ferreteriaVentas',
  'ferreteriaCortes',
  'ferreteriaOrdenes',
  'ferreteriaOrdenFolio',
  'ferreteriaPagosProveedor',
  'ferreteriaFolio',
];

function exportarRespaldo() {
  const datos = {};
  _CLAVES_RESPALDO.forEach(k => { datos[k] = localStorage.getItem(k); });
  const vacio = !datos.ferreteriaProductos || datos.ferreteriaProductos === '[]';
  if (vacio) { toast('No hay datos para exportar', 'error'); return; }
  const blob = new Blob([JSON.stringify(datos, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `respaldo-ferreteria-${hoy()}.json`;
  a.click();
  URL.revokeObjectURL(url);
  toast('Respaldo completo exportado correctamente', 'ok');
}

function restaurarRespaldo(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const datos = JSON.parse(e.target.result);
      if (typeof datos !== 'object' || Array.isArray(datos)) throw new Error();
      _CLAVES_RESPALDO.forEach(k => localStorage.removeItem(k));
      _CLAVES_RESPALDO.forEach(k => { if (datos[k] != null) localStorage.setItem(k, datos[k]); });
      toast('Respaldo restaurado. Recargando...', 'ok', 2000);
      setTimeout(() => location.reload(), 2000);
    } catch { toast('Archivo inválido', 'error'); }
  };
  reader.readAsText(file);
}

// Fecha/hora local (sin conversión UTC) para almacenamiento coherente
function fechaISOLocal() {
  const d   = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function mostrarFecha() {
  const opciones = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  document.getElementById('fecha-hoy').textContent =
    new Date().toLocaleDateString('es-MX', opciones);
}

// ===== NAVEGACIÓN =====
function iniciarEventos() {
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', e => {
      e.preventDefault();
      const vista = item.dataset.vista;
      navegarA(vista);
    });
  });

  document.getElementById('btn-abrir-modal').addEventListener('click', () => abrirModal());
  document.getElementById('btn-cerrar-modal').addEventListener('click', cerrarModal);
  document.getElementById('btn-cancelar').addEventListener('click', cerrarModal);
  document.getElementById('modal-overlay').addEventListener('click', e => {
    if (e.target === e.currentTarget) cerrarModal();
  });

  document.getElementById('form-producto').addEventListener('submit', guardarProducto);

  let _busTimer;
  document.getElementById('buscador').addEventListener('input', () => {
    clearTimeout(_busTimer);
    _busTimer = setTimeout(renderTabla, 200);
  });
  document.getElementById('filtro-categoria').addEventListener('change', renderTabla);
  document.getElementById('filtro-stock').addEventListener('change', renderTabla);

  document.getElementById('btn-exportar').addEventListener('click', exportarCSV);
  let _posBusTimer;
  document.getElementById('pos-buscador').addEventListener('input', () => {
    clearTimeout(_posBusTimer);
    pagPos = 1;
    _posBusTimer = setTimeout(renderPosGrid, 200);
  });
  document.getElementById('btn-cobrar').addEventListener('click', procesarVenta);
  document.getElementById('btn-limpiar').addEventListener('click', limpiarCarrito);

  document.getElementById('btn-cerrar-caja').addEventListener('click', cerrarCaja);
  document.getElementById('btn-add-pago-prov').addEventListener('click', agregarPagoProveedor);
  document.getElementById('btn-pdf-semanal').addEventListener('click', generarReporteSemanal);
  document.getElementById('btn-excel-semanal').addEventListener('click', () => exportarExcelVentas('semanal'));
  document.getElementById('btn-excel-mensual').addEventListener('click', () => exportarExcelVentas('mensual'));
  document.getElementById('btn-nueva-orden').addEventListener('click', generarOrden);
  document.getElementById('btn-buscar-orden').addEventListener('click', buscarOrdenRecibir);
  document.getElementById('input-recibir-orden').addEventListener('keydown', e => {
    if (e.key === 'Enter') buscarOrdenRecibir();
  });

  document.getElementById('btn-cerrar-eliminar').addEventListener('click', () => cerrarModalEliminar());
  document.getElementById('btn-cancelar-eliminar').addEventListener('click', () => cerrarModalEliminar());
  document.getElementById('btn-confirmar-eliminar').addEventListener('click', confirmarEliminar);
  document.getElementById('modal-eliminar').addEventListener('click', e => {
    if (e.target === e.currentTarget) cerrarModalEliminar();
  });
}

function navegarA(vista) {
  document.querySelectorAll('.vista').forEach(v => v.classList.add('oculto'));
  document.getElementById(`vista-${vista}`).classList.remove('oculto');
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('activo'));
  document.querySelector(`[data-vista="${vista}"]`).classList.add('activo');

  const titulos = {
    dashboard: 'Dashboard',
    inventario: 'Inventario',
    pos: 'Punto de Venta',
    ventas: 'Historial de Ventas',
    caja: 'Corte de Caja',
    compras: 'Órdenes de Compra',
    alertas: 'Alertas de Stock'
  };
  document.getElementById('titulo-vista').textContent = titulos[vista] || vista;

  if (vista === 'inventario') renderTabla();
  if (vista === 'pos') renderPos();
  if (vista === 'ventas') renderHistorial();
  if (vista === 'caja') renderCaja();
  if (vista === 'compras') renderCompras();
  if (vista === 'alertas') renderAlertas();
}

// ===== RENDER GENERAL =====
function renderTodo() {
  renderMetricas();
  const vistaId = document.querySelector('.vista:not(.oculto)')?.id;
  if (vistaId === 'vista-dashboard')  renderDashboard();
  else if (vistaId === 'vista-inventario') renderTabla();
  else if (vistaId === 'vista-alertas')    renderAlertas();
  else if (vistaId === 'vista-ventas')     renderHistorial();
  else if (vistaId === 'vista-caja')       renderCaja();
  else if (vistaId === 'vista-compras')    renderCompras();
}

// ===== MÉTRICAS + BADGE =====
function renderMetricas() {
  const total = productos.length;
  let valor = 0, alertas = 0;
  const cats = new Set();
  for (const p of productos) {
    valor += p.cantidad * p.precio;
    cats.add(p.categoria);
    if (p.cantidad <= p.stockMin) alertas++;
  }
  document.getElementById('total-productos').textContent = total;
  document.getElementById('valor-inventario').textContent = '$' + valor.toLocaleString('es-MX', { minimumFractionDigits: 2 });
  document.getElementById('total-alertas').textContent = alertas;
  document.getElementById('total-categorias').textContent = cats.size;
  const badge = document.getElementById('badge-alertas');
  badge.textContent = alertas;
  alertas > 0 ? badge.classList.add('visible') : badge.classList.remove('visible');
}

function actualizarBadge() { renderMetricas(); }

// ===== ESTADO DE STOCK =====
function estadoStock(p) {
  if (p.cantidad === 0) return 'critico';
  if (p.cantidad <= p.stockMin) return p.cantidad <= p.stockMin * 0.5 ? 'critico' : 'bajo';
  return 'ok';
}

function etiquetaEstado(estado) {
  const map = { ok: ['stock-ok', 'Normal'], bajo: ['stock-bajo', 'Stock bajo'], critico: ['stock-critico', 'Crítico'] };
  const [cls, txt] = map[estado];
  return `<span class="estado-badge ${cls}">${txt}</span>`;
}

// ===== DASHBOARD =====
function renderDashboard() {
  const criticos = productos.filter(p => estadoStock(p) !== 'ok')
    .sort((a, b) => a.cantidad - b.cantidad).slice(0, 5);

  const recientes = [...productos].reverse().slice(0, 5);

  const listaCrit = document.getElementById('lista-criticos');
  listaCrit.innerHTML = criticos.length ? criticos.map(p => `
    <div class="dash-item">
      <div>
        <p class="dash-item-nombre">${p.nombre}</p>
        <p class="dash-item-cat">${p.categoria}</p>
      </div>
      <span class="dash-item-stock ${estadoStock(p) === 'critico' ? 'stock-critico' : 'stock-bajo'}">
        ${p.cantidad} ${p.unidad}
      </span>
    </div>`) .join('') : '<p class="dash-vacío">✅ Sin productos críticos</p>';

  const listaRec = document.getElementById('lista-recientes');
  listaRec.innerHTML = recientes.length ? recientes.map(p => `
    <div class="dash-item">
      <div>
        <p class="dash-item-nombre">${p.nombre}</p>
        <p class="dash-item-cat">${p.categoria}</p>
      </div>
      <span class="dash-item-stock stock-ok">$${Number(p.precio).toFixed(2)}</span>
    </div>`).join('') : '<p class="dash-vacío">No hay productos aún</p>';
}

// ===== TABLA =====
let _tablaFiltrados = [];

function renderTabla() {
  const busqueda  = document.getElementById('buscador').value.toLowerCase();
  const catFiltro = document.getElementById('filtro-categoria').value;
  const stockFiltro = document.getElementById('filtro-stock').value;

  _tablaFiltrados = productos.filter(p => {
    const coincide = p.nombre.toLowerCase().includes(busqueda) ||
                     (p.codigo   || '').toLowerCase().includes(busqueda) ||
                     (p.clave    || '').toLowerCase().includes(busqueda) ||
                     (p.proveedor|| '').toLowerCase().includes(busqueda);
    const cat   = !catFiltro   || p.categoria === catFiltro;
    const stock = !stockFiltro || estadoStock(p) === stockFiltro;
    return coincide && cat && stock;
  });

  pagInv = 1;
  renderTablaPage();
}

function renderTablaPage() {
  const tbody = document.getElementById('cuerpo-tabla');
  const vacio = document.getElementById('tabla-vacia');
  const pagEl = document.getElementById('pag-inv');

  if (_tablaFiltrados.length === 0) {
    tbody.innerHTML = '';
    vacio.classList.remove('oculto');
    pagEl.innerHTML = '';
    return;
  }

  vacio.classList.add('oculto');

  const ini    = (pagInv - 1) * POR_PAG_INV;
  const pagina = _tablaFiltrados.slice(ini, ini + POR_PAG_INV);

  tbody.innerHTML = pagina.map(p => `
    <tr>
      <td style="font-family:monospace;color:var(--naranja);font-size:11px;white-space:nowrap">${p.codigo || '—'}</td>
      <td class="td-nombre-cell" title="${p.nombre}"><strong>${p.nombre}</strong></td>
      <td><span class="cat-badge">${p.categoria}</span></td>
      <td><strong>${p.cantidad}</strong></td>
      <td>${p.unidad}</td>
      <td>$${Number(p.precio).toFixed(2)}</td>
      <td>$${(p.cantidad * p.precio).toLocaleString('es-MX', { minimumFractionDigits: 2 })}</td>
      <td style="max-width:100px;overflow:hidden;text-overflow:ellipsis" title="${p.proveedor || ''}">${p.proveedor || '—'}</td>
      <td>${p.stockMin}</td>
      <td>${etiquetaEstado(estadoStock(p))}</td>
      <td>
        <div class="acciones">
          <button class="btn-editar"   onclick="abrirEditar('${p.id}')">Editar</button>
          <button class="btn-eliminar" onclick="abrirEliminar('${p.id}')">Eliminar</button>
        </div>
      </td>
    </tr>`).join('');

  // Paginación
  const total = Math.ceil(_tablaFiltrados.length / POR_PAG_INV);
  if (total <= 1) { pagEl.innerHTML = ''; return; }

  let h = `<button class="btn-pag-i" onclick="irPagInv(${pagInv-1})" ${pagInv===1?'disabled':''}>‹</button>`;
  for (let i = 1; i <= total; i++) {
    if (total > 8 && Math.abs(i - pagInv) > 2 && i !== 1 && i !== total) {
      if (i === pagInv - 3 || i === pagInv + 3) h += `<span style="color:var(--texto-gris)">…</span>`;
      continue;
    }
    h += `<button class="btn-pag-i ${i===pagInv?'activo':''}" onclick="irPagInv(${i})">${i}</button>`;
  }
  h += `<button class="btn-pag-i" onclick="irPagInv(${pagInv+1})" ${pagInv===total?'disabled':''}>›</button>`;
  h += `<span class="pag-info-i">${_tablaFiltrados.length} productos · pág ${pagInv}/${total}</span>`;
  pagEl.innerHTML = h;
}

function irPagInv(n) {
  const total = Math.ceil(_tablaFiltrados.length / POR_PAG_INV);
  if (n < 1 || n > total) return;
  pagInv = n;
  renderTablaPage();
  document.getElementById('tabla-inventario').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== ALERTAS =====
function renderAlertas() {
  _alertasTodas = productos.filter(p => estadoStock(p) !== 'ok')
    .sort((a, b) => a.cantidad - b.cantidad);

  const sinAlertas = document.getElementById('sin-alertas');
  if (_alertasTodas.length === 0) {
    document.getElementById('lista-alertas').innerHTML = '';
    document.getElementById('pag-alertas').innerHTML  = '';
    sinAlertas.classList.remove('oculto');
    return;
  }
  sinAlertas.classList.add('oculto');
  pagAlertas = 1;
  renderAlertasPage();
}

function renderAlertasPage() {
  const grid  = document.getElementById('lista-alertas');
  const pagEl = document.getElementById('pag-alertas');
  const ini   = (pagAlertas - 1) * POR_PAG_ALERTAS;
  const pagina = _alertasTodas.slice(ini, ini + POR_PAG_ALERTAS);

  grid.innerHTML = pagina.map(p => {
    const estado = estadoStock(p);
    return `
      <div class="alerta-card ${estado}">
        <div class="alerta-icono">${estado === 'critico' ? '🔴' : '🟡'}</div>
        <div>
          <p class="alerta-nombre">${p.nombre}</p>
          <p class="alerta-detalle">
            Categoría: ${p.categoria}<br>
            Existencia: <strong>${p.cantidad} ${p.unidad}</strong> · Mínimo: ${p.stockMin}<br>
            ${p.proveedor ? 'Proveedor: ' + p.proveedor : ''}
          </p>
        </div>
      </div>`;
  }).join('');

  const total = Math.ceil(_alertasTodas.length / POR_PAG_ALERTAS);
  if (total <= 1) { pagEl.innerHTML = ''; return; }

  let h = `<button class="btn-pag-i" onclick="irPagAlertas(${pagAlertas-1})" ${pagAlertas===1?'disabled':''}>‹</button>`;
  for (let i = 1; i <= total; i++) {
    if (total > 8 && Math.abs(i - pagAlertas) > 2 && i !== 1 && i !== total) {
      if (i === pagAlertas - 3 || i === pagAlertas + 3) h += `<span style="color:var(--texto-gris)">…</span>`;
      continue;
    }
    h += `<button class="btn-pag-i ${i===pagAlertas?'activo':''}" onclick="irPagAlertas(${i})">${i}</button>`;
  }
  h += `<button class="btn-pag-i" onclick="irPagAlertas(${pagAlertas+1})" ${pagAlertas===total?'disabled':''}>›</button>`;
  h += `<span class="pag-info-i">${_alertasTodas.length} alertas · pág ${pagAlertas}/${total}</span>`;
  pagEl.innerHTML = h;
}

function irPagAlertas(n) {
  const total = Math.ceil(_alertasTodas.length / POR_PAG_ALERTAS);
  if (n < 1 || n > total) return;
  pagAlertas = n;
  renderAlertasPage();
  document.getElementById('vista-alertas').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== MODAL AGREGAR/EDITAR =====
function abrirModal(id = null) {
  document.getElementById('form-producto').reset();
  document.getElementById('producto-id').value = '';

  if (id) {
    const p = productos.find(x => x.id === id);
    if (!p) return;
    document.getElementById('modal-titulo').textContent = 'Editar Producto';
    document.getElementById('producto-id').value = p.id;
    document.getElementById('campo-nombre').value = p.nombre;
    document.getElementById('campo-categoria').value = p.categoria;
    document.getElementById('campo-cantidad').value = p.cantidad;
    document.getElementById('campo-unidad').value = p.unidad;
    document.getElementById('campo-precio').value = p.precio;
    document.getElementById('campo-stock-min').value = p.stockMin;
    document.getElementById('campo-proveedor').value = p.proveedor || '';
    document.getElementById('campo-codigo').value    = p.codigo    || '';
  } else {
    document.getElementById('modal-titulo').textContent = 'Agregar Producto';
  }

  document.getElementById('modal-overlay').classList.remove('oculto');
}

function cerrarModal() {
  document.getElementById('modal-overlay').classList.add('oculto');
}

function abrirEditar(id) { abrirModal(id); }

function guardarProducto(e) {
  e.preventDefault();
  const id = document.getElementById('producto-id').value;

  const datos = {
    nombre:    document.getElementById('campo-nombre').value.trim(),
    categoria: document.getElementById('campo-categoria').value,
    cantidad:  Number(document.getElementById('campo-cantidad').value),
    unidad:    document.getElementById('campo-unidad').value,
    precio:    Number(document.getElementById('campo-precio').value),
    stockMin:  Number(document.getElementById('campo-stock-min').value),
    proveedor: document.getElementById('campo-proveedor').value.trim(),
    codigo:    document.getElementById('campo-codigo').value.trim(),
  };

  if (id) {
    const idx = productos.findIndex(p => p.id === id);
    productos[idx] = { ...productos[idx], ...datos };
  } else {
    datos.id = Date.now().toString();
    datos.fechaAlta = fechaISOLocal();
    productos.push(datos);
  }

  guardar();
  cerrarModal();
  renderTodo();
}

// ===== ELIMINAR =====
function abrirEliminar(id) {
  idEliminar = id;
  document.getElementById('modal-eliminar').classList.remove('oculto');
}

function cerrarModalEliminar() {
  idEliminar = null;
  document.getElementById('modal-eliminar').classList.add('oculto');
}

function confirmarEliminar() {
  if (!idEliminar) return;
  productos = productos.filter(p => p.id !== idEliminar);
  guardar();
  cerrarModalEliminar();
  renderTodo();
}

// ===== PUNTO DE VENTA =====
function renderPos() {
  renderPosCategorias();
  renderPosGrid();
  renderCarrito();
}

function renderPosCategorias() {
  const cats = ['', ...new Set(productos.map(p => p.categoria))];
  const wrap = document.getElementById('pos-categorias');
  wrap.innerHTML = cats.map(c => `
    <button class="cat-btn ${catActivaPos === c ? 'activo' : ''}"
      data-cat="${c}"
      onclick="filtrarCatPos('${c}')">
      ${c === '' ? 'Todos' : c}
    </button>`).join('');
}

function filtrarCatPos(cat) {
  catActivaPos = cat;
  pagPos = 1;
  document.querySelectorAll('#pos-categorias .cat-btn').forEach(btn => {
    btn.classList.toggle('activo', btn.dataset.cat === cat);
  });
  renderPosGrid();
}

function renderPosGrid() {
  const busq = document.getElementById('pos-buscador').value.toLowerCase();
  const filtrados = productos.filter(p => {
    const coincide = p.nombre.toLowerCase().includes(busq) || (p.codigo || '').toLowerCase().includes(busq);
    const cat = !catActivaPos || p.categoria === catActivaPos;
    return coincide && cat;
  });

  const total  = Math.ceil(filtrados.length / POR_PAG_POS) || 1;
  if (pagPos > total) pagPos = total;
  const ini    = (pagPos - 1) * POR_PAG_POS;
  const pagina = filtrados.slice(ini, ini + POR_PAG_POS);

  const grid = document.getElementById('pos-grid');
  grid.innerHTML = pagina.length ? pagina.map(p => `
    <div class="pos-producto-card ${p.cantidad <= 0 ? 'sin-stock' : ''}"
         onclick="agregarAlCarrito('${p.id}')">
      <p class="pos-prod-nombre">${p.nombre}</p>
      <p class="pos-prod-cat">${p.categoria}</p>
      <p class="pos-prod-precio">$${Number(p.precio).toFixed(2)}</p>
      <p class="pos-prod-stock">${p.cantidad <= 0 ? '❌ Sin stock' : `Stock: ${p.cantidad} ${p.unidad}`}</p>
    </div>`).join('')
  : '<p style="color:var(--texto-gris);font-size:12px;padding:20px">No hay productos</p>';

  const pag = document.getElementById('pag-pos');
  if (total <= 1) { pag.innerHTML = ''; return; }
  let h = `<button class="btn-pag-i" onclick="irPagPos(${pagPos-1})" ${pagPos===1?'disabled':''}>‹</button>`;
  for (let i = 1; i <= total; i++) {
    if (total > 8 && Math.abs(i - pagPos) > 2 && i !== 1 && i !== total) { h += '…'; continue; }
    h += `<button class="btn-pag-i ${i===pagPos?'activo':''}" onclick="irPagPos(${i})">${i}</button>`;
  }
  h += `<button class="btn-pag-i" onclick="irPagPos(${pagPos+1})" ${pagPos===total?'disabled':''}>›</button>`;
  h += `<span class="pag-info-i">${filtrados.length} productos · pág ${pagPos}/${total}</span>`;
  pag.innerHTML = h;
}

function irPagPos(n) {
  const busq = document.getElementById('pos-buscador').value.toLowerCase();
  const filtrados = productos.filter(p => {
    const coincide = p.nombre.toLowerCase().includes(busq) || (p.codigo || '').toLowerCase().includes(busq);
    const cat = !catActivaPos || p.categoria === catActivaPos;
    return coincide && cat;
  });
  const total = Math.ceil(filtrados.length / POR_PAG_POS) || 1;
  pagPos = Math.max(1, Math.min(n, total));
  renderPosGrid();
}

function agregarAlCarrito(id) {
  const prod = productos.find(p => p.id === id);
  if (!prod || prod.cantidad <= 0) return;
  const enCarrito = carrito.find(c => c.id === id);
  if (enCarrito) {
    if (enCarrito.qty >= prod.cantidad) return;
    enCarrito.qty++;
  } else {
    carrito.push({ id: prod.id, nombre: prod.nombre, precio: prod.precio, precioVenta: prod.precio, qty: 1, unidad: prod.unidad });
  }
  renderCarrito();
}

function cambiarCantidad(id, delta) {
  const item = carrito.find(c => c.id === id);
  const prod = productos.find(p => p.id === id);
  if (!item) return;
  item.qty += delta;
  if (item.qty <= 0) carrito = carrito.filter(c => c.id !== id);
  if (prod && item.qty > prod.cantidad) item.qty = prod.cantidad;
  renderCarrito();
}

function quitarDelCarrito(id) {
  carrito = carrito.filter(c => c.id !== id);
  renderCarrito();
}

function actualizarPrecioVenta(id, valor) {
  const item = carrito.find(c => c.id === id);
  if (!item) return;
  const v = parseFloat(valor);
  item.precioVenta = isNaN(v) || v < 0 ? 0 : v;
  const total = carrito.reduce((s, i) => s + i.precioVenta * i.qty, 0);
  document.getElementById('carrito-subtotal').textContent = '$' + total.toFixed(2);
  document.getElementById('carrito-total').textContent = '$' + total.toFixed(2);
}

function limpiarCarrito() {
  carrito = [];
  renderCarrito();
}

function renderCarrito() {
  const wrap = document.getElementById('carrito-items');
  const btnCobrar = document.getElementById('btn-cobrar');

  if (carrito.length === 0) {
    wrap.innerHTML = '<p class="carrito-vacio">Selecciona productos del panel izquierdo</p>';
    btnCobrar.disabled = true;
  } else {
    wrap.innerHTML = carrito.map(item => `
      <div class="carrito-item">
        <p class="carrito-item-nombre">${item.nombre}</p>
        <div class="carrito-precios">
          <span class="precio-ref">Ref: $${Number(item.precio).toFixed(2)}</span>
          <label class="precio-venta-label">Precio venta $
            <input type="number" class="precio-venta-input" min="0" step="0.01"
              value="${Number(item.precioVenta).toFixed(2)}"
              onchange="actualizarPrecioVenta('${item.id}', this.value)"
              oninput="actualizarPrecioVenta('${item.id}', this.value)">
          </label>
        </div>
        <div class="carrito-item-controles">
          <div class="ctrl-cantidad">
            <button class="btn-ctrl" onclick="cambiarCantidad('${item.id}', -1)">−</button>
            <span class="ctrl-num">${item.qty}</span>
            <button class="btn-ctrl" onclick="cambiarCantidad('${item.id}', 1)">+</button>
          </div>
          <span class="carrito-item-precio">$${(item.precioVenta * item.qty).toFixed(2)}</span>
          <button class="btn-quitar" onclick="quitarDelCarrito('${item.id}')">🗑</button>
        </div>
      </div>`).join('');
    btnCobrar.disabled = false;
  }

  const total = carrito.reduce((s, i) => s + i.precioVenta * i.qty, 0);
  document.getElementById('carrito-subtotal').textContent = '$' + total.toFixed(2);
  document.getElementById('carrito-total').textContent = '$' + total.toFixed(2);
}

function procesarVenta() {
  if (carrito.length === 0) return;

  const metodoPago = document.querySelector('input[name="metodo-pago"]:checked')?.value || 'efectivo';

  carrito.forEach(item => {
    const prod = productos.find(p => p.id === item.id);
    if (prod) prod.cantidad -= item.qty;
  });

  const venta = {
    id: Date.now().toString(),
    folio: siguienteFolio(),
    fecha: fechaISOLocal(),
    items: carrito.map(i => ({ ...i })),
    total: carrito.reduce((s, i) => s + i.precioVenta * i.qty, 0),
    metodoPago
  };

  ventas.unshift(venta);
  guardar();
  carrito = [];
  renderCarrito();
  renderTodo();
  renderPosGrid();
  toast(`✅ Venta #${venta.folio} registrada — $${venta.total.toFixed(2)} (${venta.metodoPago})`, 'ok', 4000);
}

// ===== HISTORIAL DE VENTAS =====
function renderHistorial() {
  const totalVentas = ventas.length;
  const totalIngresos = ventas.reduce((s, v) => s + v.total, 0);
  const totalArticulos = ventas.reduce((s, v) => s + v.items.reduce((a, i) => a + i.qty, 0), 0);

  document.getElementById('ventas-resumen').innerHTML = `
    <div class="metrica-card">
      <div class="metrica-icono naranja">🧾</div>
      <div>
        <p class="metrica-valor">${totalVentas}</p>
        <p class="metrica-label">Ventas realizadas</p>
      </div>
    </div>
    <div class="metrica-card">
      <div class="metrica-icono verde">💵</div>
      <div>
        <p class="metrica-valor">$${totalIngresos.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</p>
        <p class="metrica-label">Ingresos totales</p>
      </div>
    </div>
    <div class="metrica-card">
      <div class="metrica-icono azul">📦</div>
      <div>
        <p class="metrica-valor">${totalArticulos}</p>
        <p class="metrica-label">Artículos vendidos</p>
      </div>
    </div>`;

  const tbody = document.getElementById('historial-body');
  const vacio = document.getElementById('historial-vacio');

  if (ventas.length === 0) {
    tbody.innerHTML = '';
    vacio.classList.remove('oculto');
    return;
  }

  vacio.classList.add('oculto');
  const iconoPago = { efectivo: '💵', tarjeta: '💳', transferencia: '📲' };
  tbody.innerHTML = ventas.map(v => {
    const fecha = new Date(v.fecha).toLocaleString('es-MX', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit'
    });
    const prods = v.items.map(i => `${i.nombre} (x${i.qty})`).join(', ');
    const totalArts = v.items.reduce((s, i) => s + i.qty, 0);
    const mp = v.metodoPago || 'efectivo';
    const mpLabel = { efectivo: 'Efectivo', tarjeta: 'Tarjeta', transferencia: 'Transferencia' }[mp] || mp;
    return `<tr>
      <td style="font-family:monospace;color:var(--naranja);font-weight:700">#${v.folio || '—'}</td>
      <td>${fecha}</td>
      <td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${prods}">${prods}</td>
      <td>${totalArts}</td>
      <td><span class="badge-pago badge-pago-${mp}">${iconoPago[mp] || ''} ${mpLabel}</span></td>
      <td><strong style="color:var(--verde)">$${v.total.toFixed(2)}</strong></td>
      <td style="white-space:nowrap">
        <button class="btn-nota" onclick="generarNotaVenta('${v.id}')">⬇ PDF</button>
        <button class="btn-nota" style="margin-left:4px" onclick="imprimirTicket('${v.id}')">🖨️ Ticket</button>
      </td>
    </tr>`;
  }).join('');
}

// ===== EXPORTAR CSV =====
function exportarCSV() {
  if (productos.length === 0) { toast('No hay productos para exportar.', 'error'); return; }

  const encabezado = ['Código', 'Nombre', 'Categoría', 'Cantidad', 'Unidad', 'Precio Unitario', 'Valor Total', 'Proveedor', 'Stock Mínimo', 'Estado'];
  const filas = productos.map(p => [
    p.codigo || '', p.nombre, p.categoria, p.cantidad, p.unidad,
    p.precio.toFixed(2), (p.cantidad * p.precio).toFixed(2),
    p.proveedor || '', p.stockMin, estadoStock(p)
  ]);

  const esc = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csv = [encabezado, ...filas].map(f => f.map(esc).join(',')).join('\n');
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `inventario-ferreteria-torres-${hoy()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ===== NOTA DE VENTA PDF =====
async function generarNotaVenta(ventaId) {
  const venta = ventas.find(v => v.id === ventaId);
  if (!venta) return;
  try { await prepararPDF(); } catch { toast('No se pudo cargar la librería PDF. Verifica tu conexión.', 'error'); return; }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  const folio    = venta.folio || 'S/N';
  const fecha    = new Date(venta.fecha);
  const fechaStr = fecha.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const horaStr  = fecha.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
  const mpLabels = { efectivo: 'Efectivo', tarjeta: 'Tarjeta', transferencia: 'Transferencia' };
  const mpLabel  = mpLabels[venta.metodoPago] || 'Efectivo';

  const subtotalSinIva = venta.total / 1.16;
  const iva            = venta.total - subtotalSinIva;

  // ── ENCABEZADO ──
  doc.setFillColor(249, 115, 22);
  doc.rect(0, 0, 210, 40, 'F');
  if (typeof LOGO_B64 !== 'undefined') doc.addImage(LOGO_B64, 'JPEG', 174, 4, 28, 28);
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(20);
  doc.setFont('helvetica', 'bold');
  doc.text('Ferretería Las Torres', 14, 13);
  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.text('Tlapalería y Ferretería', 14, 20);
  doc.setFontSize(14);
  doc.setFont('helvetica', 'bold');
  doc.text('NOTA DE VENTA', 14, 31);
  doc.setFontSize(11);
  doc.text(`Folio #${folio}`, 166, 31, { align: 'right' });

  // ── INFO ──
  doc.setTextColor(60, 60, 60);
  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.text(`Fecha: ${fechaStr}`, 14, 50);
  doc.text(`Hora: ${horaStr}`, 14, 56);
  doc.text(`Método de pago: ${mpLabel}`, 14, 62);
  doc.setDrawColor(230, 230, 230);
  doc.line(14, 66, 196, 66);

  // ── TABLA DE ARTÍCULOS ──
  doc.autoTable({
    startY: 70,
    head: [['Producto', 'Cant.', 'P. Unitario', 'Subtotal']],
    body: venta.items.map(item => {
      const pu = Number(item.precioVenta ?? item.precio);
      return [
        item.nombre,
        item.qty,
        '$' + pu.toLocaleString('es-MX', { minimumFractionDigits: 2 }),
        '$' + (pu * item.qty).toLocaleString('es-MX', { minimumFractionDigits: 2 })
      ];
    }),
    styles: { fontSize: 10, cellPadding: 4, font: 'helvetica' },
    headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255], fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [252, 252, 252] },
    columnStyles: {
      1: { halign: 'center', cellWidth: 18 },
      2: { halign: 'right',  cellWidth: 38 },
      3: { halign: 'right',  cellWidth: 38 }
    }
  });

  // ── TOTALES ──
  let ty = doc.lastAutoTable.finalY + 6;
  const lx = 130, rx = 196;

  doc.setDrawColor(220, 220, 220);
  doc.line(lx, ty, rx, ty);
  ty += 6;

  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(80, 80, 80);
  doc.text('Subtotal (sin IVA):', lx, ty);
  doc.text('$' + subtotalSinIva.toLocaleString('es-MX', { minimumFractionDigits: 2 }), rx, ty, { align: 'right' });
  ty += 7;
  doc.text('IVA (16%):', lx, ty);
  doc.text('$' + iva.toLocaleString('es-MX', { minimumFractionDigits: 2 }), rx, ty, { align: 'right' });
  ty += 2;
  doc.setDrawColor(249, 115, 22);
  doc.line(lx, ty, rx, ty);
  ty += 6;
  doc.setFontSize(13);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(34, 120, 60);
  doc.text('TOTAL:', lx, ty);
  doc.text('$' + venta.total.toLocaleString('es-MX', { minimumFractionDigits: 2 }), rx, ty, { align: 'right' });

  // ── PIE ──
  ty += 14;
  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(120, 120, 120);
  doc.text('¡Gracias por su compra!', 105, ty, { align: 'center' });
  const ph = doc.internal.pageSize.height;
  doc.setFontSize(7);
  doc.text('Ferretería Las Torres — Sistema de Inventario v1.0', 105, ph - 8, { align: 'center' });

  doc.save(`nota-venta-${folio}.pdf`);
}

// ===== EXPORTAR EXCEL VENTAS =====
async function exportarExcelVentas(filtro) {
  try { await prepararExcel(); } catch { toast('No se pudo cargar la librería Excel. Verifica tu conexión.', 'error'); return; }

  const hoyDate  = new Date();
  const pad2 = n => String(n).padStart(2, '0');
  const mesActual = `${hoyDate.getFullYear()}-${pad2(hoyDate.getMonth()+1)}`;
  const diasSem   = semanaActual();

  const ventasFiltradas = ventas.filter(v =>
    filtro === 'semanal'
      ? diasSem.includes(v.fecha.slice(0, 10))
      : v.fecha.slice(0, 7) === mesActual
  );

  if (ventasFiltradas.length === 0) {
    toast('No hay ventas en el período seleccionado.', 'info');
    return;
  }

  const mpLabels = { efectivo: 'Efectivo', tarjeta: 'Tarjeta', transferencia: 'Transferencia' };
  const encabezado = ['Folio', 'Fecha', 'Hora', 'Producto', 'Código', 'Cant.', 'P. Unitario', 'Subtotal s/IVA', 'IVA 16%', 'Total c/IVA', 'Método de Pago'];

  const filas = [encabezado];
  let grandTotal = 0, grandIva = 0;

  ventasFiltradas.forEach(v => {
    const d         = new Date(v.fecha);
    const fechaStr  = d.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const horaStr   = d.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
    const mpLabel   = mpLabels[v.metodoPago] || 'Efectivo';

    v.items.forEach((item, idx) => {
      const pu          = Number(item.precioVenta ?? item.precio);
      const subtotalItem = parseFloat((pu * item.qty).toFixed(2));
      const sinIva      = parseFloat((subtotalItem / 1.16).toFixed(2));
      const ivaItem     = parseFloat((subtotalItem - sinIva).toFixed(2));
      grandIva   += ivaItem;
      filas.push([
        `#${v.folio || 'S/N'}`,
        idx === 0 ? fechaStr : '',
        idx === 0 ? horaStr  : '',
        item.nombre,
        item.codigo || '',
        item.qty,
        pu,
        sinIva,
        ivaItem,
        subtotalItem,
        idx === 0 ? mpLabel : ''
      ]);
    });

    const totalSinIva = parseFloat((v.total / 1.16).toFixed(2));
    const ivaVenta    = parseFloat((v.total - totalSinIva).toFixed(2));
    grandTotal += v.total;
    filas.push(['', '', '', '↳ SUBTOTAL VENTA', '', '', '', totalSinIva, ivaVenta, parseFloat(v.total.toFixed(2)), '']);
    filas.push([]);
  });

  // Fila de gran total
  const grandSinIva = parseFloat((grandTotal / 1.16).toFixed(2));
  filas.push(['TOTAL PERÍODO', '', '', '', '', '', '', grandSinIva, parseFloat(grandIva.toFixed(2)), parseFloat(grandTotal.toFixed(2)), '']);

  const ws = XLSX.utils.aoa_to_sheet(filas);
  ws['!cols'] = [
    { wch: 8 }, { wch: 12 }, { wch: 7 }, { wch: 32 }, { wch: 10 },
    { wch: 6 }, { wch: 12 }, { wch: 15 }, { wch: 10 }, { wch: 12 }, { wch: 16 }
  ];

  const wb = XLSX.utils.book_new();
  const sheetName = filtro === 'semanal' ? 'Ventas Semana' : 'Ventas Mes';
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  const periodo = filtro === 'semanal'
    ? `${diasSem[0]}_al_${diasSem[diasSem.length - 1]}`
    : mesActual;

  XLSX.writeFile(wb, `ventas_${filtro}_${periodo}.xlsx`);
}

// ===== CORTE DE CAJA =====
function hoy() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}

function ventasDeHoy() {
  return ventas.filter(v => v.fecha.slice(0, 10) === hoy());
}

function pagosProvDeHoy() {
  return pagosProveedor.filter(p => p.fecha === hoy());
}

function renderCaja() {
  const vh = ventasDeHoy();
  const totalVentas = vh.reduce((s, v) => s + v.total, 0);
  const articulosDia = vh.reduce((s, v) => s + v.items.reduce((a, i) => a + i.qty, 0), 0);
  const totalPagos = pagosProvDeHoy().reduce((s, p) => s + p.monto, 0);
  const neto = totalVentas - totalPagos;

  const cortesHoy = cortes.filter(c => c.fecha === hoy());
  const yaCerrado = cortesHoy.length > 0;

  document.getElementById('caja-resumen-hoy').innerHTML = `
    <div class="caja-stats">
      <div class="caja-stat">
        <p class="caja-stat-val">${vh.length}</p>
        <p class="caja-stat-label">Ventas del día</p>
      </div>
      <div class="caja-stat">
        <p class="caja-stat-val">${articulosDia}</p>
        <p class="caja-stat-label">Artículos vendidos</p>
      </div>
      <div class="caja-stat verde">
        <p class="caja-stat-val">$${totalVentas.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</p>
        <p class="caja-stat-label">Total ventas</p>
      </div>
      <div class="caja-stat rojo">
        <p class="caja-stat-val">−$${totalPagos.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</p>
        <p class="caja-stat-label">Pagos a proveed.</p>
      </div>
      <div class="caja-stat ${neto >= 0 ? 'verde' : 'rojo'}">
        <p class="caja-stat-val">$${neto.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</p>
        <p class="caja-stat-label">Neto del día</p>
      </div>
    </div>
    ${yaCerrado ? '<p class="caja-cerrada">✅ Caja cerrada hoy — ' + cortesHoy[cortesHoy.length-1].hora + '</p>' : ''}`;

  const btnCerrar = document.getElementById('btn-cerrar-caja');
  btnCerrar.disabled = yaCerrado;
  btnCerrar.textContent = yaCerrado ? '✅ Caja ya cerrada' : 'Cerrar caja del día';

  // Ventas de hoy
  const listaHoy = document.getElementById('caja-ventas-hoy');
  listaHoy.innerHTML = vh.length ? vh.map(v => {
    const hora = new Date(v.fecha).toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
    const prods = v.items.map(i => `${i.nombre} x${i.qty}`).join(', ');
    return `
      <div class="dash-item">
        <div>
          <p class="dash-item-nombre">${hora} — ${v.items.length} producto${v.items.length !== 1 ? 's' : ''}</p>
          <p class="dash-item-cat" style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${prods}">${prods}</p>
        </div>
        <span class="dash-item-stock stock-ok">$${v.total.toFixed(2)}</span>
      </div>`;
  }).join('') : '<p class="dash-vacío">No hay ventas hoy.</p>';

  renderPagosProveedor();
  renderResumenSemanal();
  renderHistorialCortes();
}

// ===== PAGOS A PROVEEDORES =====
function agregarPagoProveedor() {
  const nombre = document.getElementById('prov-nombre').value.trim();
  const monto  = parseFloat(document.getElementById('prov-monto').value);
  if (!nombre) { toast('Ingresa el nombre del proveedor.', 'error'); return; }
  if (isNaN(monto) || monto <= 0) { toast('Ingresa un monto válido.', 'error'); return; }

  const ahora = new Date();
  pagosProveedor.push({
    id: Date.now().toString(),
    fecha: hoy(),
    hora: ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
    proveedor: nombre,
    monto
  });

  document.getElementById('prov-nombre').value = '';
  document.getElementById('prov-monto').value  = '';
  guardar();
  renderCaja();
}

function eliminarPagoProveedor(id) {
  pagosProveedor = pagosProveedor.filter(p => p.id !== id);
  guardar();
  renderCaja();
}

function renderPagosProveedor() {
  const pagosHoy = pagosProvDeHoy();
  const total = pagosHoy.reduce((s, p) => s + p.monto, 0);
  const el = document.getElementById('lista-pagos-prov');

  if (pagosHoy.length === 0) {
    el.innerHTML = '<p class="dash-vacío" style="margin-top:10px">Sin pagos a proveedores hoy.</p>';
    return;
  }

  el.innerHTML = `
    <div class="tabla-wrap" style="border:none;border-radius:0;margin-top:10px">
      <table class="tabla">
        <thead><tr><th>Hora</th><th>Proveedor</th><th>Monto</th><th></th></tr></thead>
        <tbody>
          ${pagosHoy.map(p => `
            <tr>
              <td style="color:var(--texto-gris);font-size:11px">${p.hora}</td>
              <td><strong>${p.proveedor}</strong></td>
              <td style="color:var(--rojo);font-weight:700">−$${p.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</td>
              <td><button class="btn-eliminar" onclick="eliminarPagoProveedor('${p.id}')">✕</button></td>
            </tr>`).join('')}
          <tr style="border-top:2px solid var(--borde)">
            <td colspan="2" style="font-weight:700;color:var(--texto)">Total pagado hoy</td>
            <td style="color:var(--rojo);font-weight:700">−$${total.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>`;
}

// ===== RESUMEN SEMANAL =====
function semanaActual() {
  const hoyDate    = new Date();
  const dow        = hoyDate.getDay();
  const diasDesdeL = dow === 0 ? 6 : dow - 1;
  const pad = n => String(n).padStart(2, '0');
  const dias = [];
  for (let i = 0; i <= diasDesdeL; i++) {
    const d = new Date(hoyDate);
    d.setDate(hoyDate.getDate() - diasDesdeL + i);
    dias.push(`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`);
  }
  return dias;
}

function renderResumenSemanal() {
  const dias = semanaActual();
  const nombres = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
  let acumulado = 0;
  let totalVentasSem = 0;
  let totalPagosSem = 0;

  const filas = dias.map(fecha => {
    const vDia = ventas.filter(v => v.fecha.slice(0, 10) === fecha)
      .reduce((s, v) => s + v.total, 0);
    const pDia = pagosProveedor.filter(p => p.fecha === fecha)
      .reduce((s, p) => s + p.monto, 0);
    const neto = vDia - pDia;
    acumulado += neto;
    totalVentasSem += vDia;
    totalPagosSem  += pDia;
    const d = new Date(fecha + 'T12:00:00');
    const esHoy = fecha === hoy();
    return { fecha, dia: nombres[d.getDay()], vDia, pDia, neto, acumulado, esHoy };
  });

  const netoSem = totalVentasSem - totalPagosSem;
  const colorNeto = n => n >= 0 ? 'var(--verde)' : 'var(--rojo)';

  document.getElementById('resumen-semanal').innerHTML = `
    <div class="tabla-wrap" style="border:none;border-radius:0">
      <table class="tabla">
        <thead>
          <tr>
            <th>Día</th>
            <th>Ventas</th>
            <th>Pagos Proveed.</th>
            <th>Neto del día</th>
            <th>Acumulado semana</th>
          </tr>
        </thead>
        <tbody>
          ${filas.map(f => `
            <tr ${f.esHoy ? 'style="background:rgba(249,115,22,0.07)"' : ''}>
              <td><strong>${f.dia}</strong> <span style="color:var(--texto-gris);font-size:10px">${f.fecha.slice(5)}</span>${f.esHoy ? ' <span style="color:var(--naranja);font-size:10px">hoy</span>' : ''}</td>
              <td style="color:var(--verde)">$${f.vDia.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</td>
              <td style="color:${f.pDia > 0 ? 'var(--rojo)' : 'var(--texto-gris)'}">
                ${f.pDia > 0 ? '−$' + f.pDia.toLocaleString('es-MX', { minimumFractionDigits: 2 }) : '—'}
              </td>
              <td style="color:${colorNeto(f.neto)};font-weight:700">
                ${f.neto >= 0 ? '' : '−'}$${Math.abs(f.neto).toLocaleString('es-MX', { minimumFractionDigits: 2 })}
              </td>
              <td style="color:${colorNeto(f.acumulado)};font-weight:700">
                ${f.acumulado >= 0 ? '' : '−'}$${Math.abs(f.acumulado).toLocaleString('es-MX', { minimumFractionDigits: 2 })}
              </td>
            </tr>`).join('')}
          <tr style="border-top:2px solid var(--borde)">
            <td><strong>Total semana</strong></td>
            <td style="color:var(--verde);font-weight:700">$${totalVentasSem.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</td>
            <td style="color:var(--rojo);font-weight:700">${totalPagosSem > 0 ? '−$' + totalPagosSem.toLocaleString('es-MX', { minimumFractionDigits: 2 }) : '—'}</td>
            <td colspan="2" style="color:${colorNeto(netoSem)};font-weight:700;font-size:14px">
              ${netoSem >= 0 ? '' : '−'}$${Math.abs(netoSem).toLocaleString('es-MX', { minimumFractionDigits: 2 })}
            </td>
          </tr>
        </tbody>
      </table>
    </div>`;
}

function cerrarCaja() {
  const vh = ventasDeHoy();
  if (vh.length === 0) { toast('No hay ventas hoy para registrar en el corte.', 'error'); return; }

  const cortesHoy = cortes.filter(c => c.fecha === hoy());
  if (cortesHoy.length > 0) { toast('La caja ya fue cerrada hoy.', 'info'); return; }

  const totalVentas = vh.reduce((s, v) => s + v.total, 0);
  const articulosDia = vh.reduce((s, v) => s + v.items.reduce((a, i) => a + i.qty, 0), 0);
  const totalPagos = pagosProvDeHoy().reduce((s, p) => s + p.monto, 0);
  const neto = totalVentas - totalPagos;
  const ahora = new Date();

  const corte = {
    id: Date.now().toString(),
    fecha: hoy(),
    hora: ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
    numVentas: vh.length,
    articulos: articulosDia,
    total: totalVentas,
    pagosProveedor: totalPagos,
    neto,
    cerradoPor: 'Administrador'
  };

  cortes.unshift(corte);
  guardar();
  renderCaja();

  const esDomingo = new Date().getDay() === 0;
  toast(
    `✅ Caja cerrada — Ventas: $${totalVentas.toFixed(2)} · Pagos: $${totalPagos.toFixed(2)} · Neto: $${neto.toFixed(2)}${esDomingo ? ' · Generando PDF semanal…' : ''}`,
    'ok', 6000
  );

  if (esDomingo) generarReporteSemanal();
}

// ===== REPORTE SEMANAL PDF =====
async function generarReporteSemanal() {
  try { await prepararPDF(); } catch { toast('No se pudo cargar la librería PDF. Verifica tu conexión.', 'error'); return; }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  const dias = semanaActual();
  const nombresDias = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];

  // Recopilar datos
  let acumulado = 0;
  let totalVentasSem = 0;
  let totalPagosSem = 0;

  const filasDias = dias.map(fecha => {
    const vDia = ventas.filter(v => v.fecha.slice(0, 10) === fecha).reduce((s, v) => s + v.total, 0);
    const pDia = pagosProveedor.filter(p => p.fecha === fecha).reduce((s, p) => s + p.monto, 0);
    const neto = vDia - pDia;
    acumulado += neto;
    totalVentasSem += vDia;
    totalPagosSem  += pDia;
    const d = new Date(fecha + 'T12:00:00');
    return [
      nombresDias[d.getDay()] + ' ' + fecha.slice(5),
      '$' + vDia.toLocaleString('es-MX', { minimumFractionDigits: 2 }),
      pDia > 0 ? '-$' + pDia.toLocaleString('es-MX', { minimumFractionDigits: 2 }) : '—',
      (neto < 0 ? '-' : '') + '$' + Math.abs(neto).toLocaleString('es-MX', { minimumFractionDigits: 2 }),
      (acumulado < 0 ? '-' : '') + '$' + Math.abs(acumulado).toLocaleString('es-MX', { minimumFractionDigits: 2 })
    ];
  });

  const netoSem = totalVentasSem - totalPagosSem;

  // ── ENCABEZADO ──
  doc.setFillColor(249, 115, 22);
  doc.rect(0, 0, 210, 36, 'F');
  if (typeof LOGO_B64 !== 'undefined') {
    doc.addImage(LOGO_B64, 'JPEG', 174, 4, 28, 28);
  }
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(20);
  doc.setFont('helvetica', 'bold');
  doc.text('Ferretería Las Torres', 14, 13);
  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.text('Tlapalería y Ferretería', 14, 20);
  doc.setFontSize(13);
  doc.setFont('helvetica', 'bold');
  doc.text('Reporte Semanal de Caja', 14, 30);

  // Período y fecha de generación
  doc.setTextColor(60, 60, 60);
  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.text(`Período: ${dias[0]}  al  ${dias[dias.length - 1]}`, 14, 44);
  doc.text(`Generado: ${new Date().toLocaleString('es-MX')}`, 14, 50);

  // ── CAJAS DE RESUMEN ──
  let y = 58;
  const cajas = [
    { label: 'Total Ventas', val: '$' + totalVentasSem.toLocaleString('es-MX', { minimumFractionDigits: 2 }), rgb: [34, 197, 94] },
    { label: 'Pagos a Proveedores', val: '-$' + totalPagosSem.toLocaleString('es-MX', { minimumFractionDigits: 2 }), rgb: [239, 68, 68] },
    { label: 'Neto de la Semana', val: (netoSem < 0 ? '-' : '') + '$' + Math.abs(netoSem).toLocaleString('es-MX', { minimumFractionDigits: 2 }), rgb: netoSem >= 0 ? [34, 197, 94] : [239, 68, 68] }
  ];
  const bw = 58, gap = 5;
  cajas.forEach((b, i) => {
    const x = 14 + i * (bw + gap);
    doc.setFillColor(248, 248, 248);
    doc.setDrawColor(220, 220, 220);
    doc.roundedRect(x, y, bw, 22, 3, 3, 'FD');
    doc.setFontSize(8);
    doc.setTextColor(120, 120, 120);
    doc.setFont('helvetica', 'normal');
    doc.text(b.label, x + bw / 2, y + 8, { align: 'center' });
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(...b.rgb);
    doc.text(b.val, x + bw / 2, y + 18, { align: 'center' });
  });
  y += 30;

  // ── TABLA DÍA A DÍA ──
  doc.setFontSize(11);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(40, 40, 40);
  doc.text('Desglose diario', 14, y);
  y += 2;

  doc.autoTable({
    startY: y,
    head: [['Día', 'Ventas', 'Pagos Prov.', 'Neto del día', 'Acumulado semana']],
    body: filasDias,
    foot: [[
      'Total semana',
      '$' + totalVentasSem.toLocaleString('es-MX', { minimumFractionDigits: 2 }),
      totalPagosSem > 0 ? '-$' + totalPagosSem.toLocaleString('es-MX', { minimumFractionDigits: 2 }) : '—',
      (netoSem < 0 ? '-' : '') + '$' + Math.abs(netoSem).toLocaleString('es-MX', { minimumFractionDigits: 2 }),
      ''
    ]],
    styles: { fontSize: 9, cellPadding: 3.5, font: 'helvetica' },
    headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255], fontStyle: 'bold' },
    footStyles: { fillColor: [235, 235, 235], textColor: [40, 40, 40], fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [252, 252, 252] },
    columnStyles: { 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right', fontStyle: 'bold' }, 4: { halign: 'right', fontStyle: 'bold' } },
    didParseCell(data) {
      if (data.section === 'body' && (data.column.index === 3 || data.column.index === 4)) {
        const v = String(data.cell.raw);
        data.cell.styles.textColor = v.startsWith('-') ? [220, 50, 50] : [30, 160, 80];
      }
    }
  });

  y = doc.lastAutoTable.finalY + 8;

  // ── DETALLE DE PAGOS A PROVEEDORES ──
  const pagosDelaSemana = pagosProveedor.filter(p => dias.includes(p.fecha));
  if (pagosDelaSemana.length > 0) {
    doc.setFontSize(11);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(40, 40, 40);
    doc.text('Detalle de pagos a proveedores', 14, y);
    y += 2;

    doc.autoTable({
      startY: y,
      head: [['Fecha', 'Hora', 'Proveedor', 'Monto']],
      body: pagosDelaSemana.map(p => [
        p.fecha, p.hora, p.proveedor,
        '-$' + p.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })
      ]),
      styles: { fontSize: 9, cellPadding: 3 },
      headStyles: { fillColor: [239, 68, 68], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: { 3: { halign: 'right', textColor: [220, 50, 50], fontStyle: 'bold' } }
    });

    y = doc.lastAutoTable.finalY + 8;
  }

  // ── DETALLE DE VENTAS ──
  const ventasSemana = ventas.filter(v => dias.includes(v.fecha.slice(0, 10)));
  if (ventasSemana.length > 0) {
    doc.setFontSize(11);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(40, 40, 40);
    doc.text('Detalle de ventas', 14, y);
    y += 2;

    const iconoPago = { efectivo: 'Efectivo', tarjeta: 'Tarjeta', transferencia: 'Transferencia' };
    doc.autoTable({
      startY: y,
      head: [['Fecha', 'Hora', 'Artículos', 'Método', 'Total']],
      body: ventasSemana.map(v => {
        const d = new Date(v.fecha);
        return [
          d.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' }),
          d.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' }),
          v.items.reduce((s, i) => s + i.qty, 0),
          iconoPago[v.metodoPago] || 'Efectivo',
          '$' + v.total.toLocaleString('es-MX', { minimumFractionDigits: 2 })
        ];
      }),
      styles: { fontSize: 8, cellPadding: 2.5 },
      headStyles: { fillColor: [34, 197, 94], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: { 4: { halign: 'right', fontStyle: 'bold', textColor: [30, 160, 80] } },
      alternateRowStyles: { fillColor: [252, 252, 252] }
    });
  }

  // ── PIE DE PÁGINA ──
  const ph = doc.internal.pageSize.height;
  doc.setFontSize(7);
  doc.setTextColor(170, 170, 170);
  doc.setFont('helvetica', 'normal');
  doc.text('Ferretería Las Torres — Sistema de Inventario v1.0', 105, ph - 8, { align: 'center' });

  // Descargar
  const nombre = `reporte-semanal_${dias[0]}_al_${dias[dias.length - 1]}.pdf`;
  doc.save(nombre);
}

function renderHistorialCortes() {
  const tbody = document.getElementById('historial-cortes');
  const vacio = document.getElementById('cortes-vacio');

  if (cortes.length === 0) {
    tbody.innerHTML = '';
    vacio.classList.remove('oculto');
    return;
  }

  vacio.classList.add('oculto');
  tbody.innerHTML = cortes.map(c => {
    const neto = c.neto !== undefined ? c.neto : c.total;
    const pagos = c.pagosProveedor || 0;
    return `
    <tr>
      <td>${c.fecha} <span style="color:var(--texto-gris);font-size:11px">${c.hora}</span></td>
      <td>${c.numVentas}</td>
      <td>${c.articulos}</td>
      <td style="color:var(--verde);font-weight:700">$${c.total.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</td>
      <td style="color:${pagos > 0 ? 'var(--rojo)' : 'var(--texto-gris)'}">
        ${pagos > 0 ? '−$' + pagos.toLocaleString('es-MX', { minimumFractionDigits: 2 }) : '—'}
      </td>
      <td style="color:${neto >= 0 ? 'var(--verde)' : 'var(--rojo)'};font-weight:700">
        ${neto >= 0 ? '' : '−'}$${Math.abs(neto).toLocaleString('es-MX', { minimumFractionDigits: 2 })}
      </td>
      <td>${c.cerradoPor}</td>
    </tr>`;
  }).join('');
}

// ===== ÓRDENES DE COMPRA =====
let proveedorActivoCompras = '';

function filtrarComprasProv(prov) {
  proveedorActivoCompras = prov;
  renderCompras();
}

function renderCompras() {
  const necesitan = productos.filter(p => p.cantidad < p.stockMin)
    .sort((a, b) => a.cantidad - b.cantidad);

  // Botones de filtro por proveedor
  const todosProv = [...new Set(necesitan.map(p => p.proveedor || 'Sin proveedor'))].sort();
  const filtroEl  = document.getElementById('compras-prov-filtro');
  if (filtroEl) {
    filtroEl.innerHTML = todosProv.length ? `
      <div class="prov-filtro-wrap">
        <span class="prov-filtro-label">Proveedor:</span>
        <div class="prov-filtro-btns">
          <button class="prov-filtro-btn ${proveedorActivoCompras === '' ? 'activo' : ''}"
            onclick="filtrarComprasProv('')">🏪 Todos</button>
          ${todosProv.map(p => `
            <button class="prov-filtro-btn ${proveedorActivoCompras === p ? 'activo' : ''}"
              onclick="filtrarComprasProv('${p.replace(/'/g, "\\'")}')">
              ${p}
            </button>`).join('')}
        </div>
      </div>` : '';
  }

  // Título dinámico
  const titulo = document.getElementById('compras-card-titulo');
  if (titulo) titulo.textContent = proveedorActivoCompras
    ? `⚠️ Reabastecimiento — ${proveedorActivoCompras}`
    : '⚠️ Productos que necesitan reabastecimiento';

  // Filtrar
  const filtrados = proveedorActivoCompras === ''
    ? necesitan
    : necesitan.filter(p => (p.proveedor || 'Sin proveedor') === proveedorActivoCompras);

  // Inicializa estado sólo para productos nuevos; conserva cambios del usuario
  filtrados.forEach(p => {
    if (!_reabastoState[p.id]) {
      _reabastoState[p.id] = {
        checked: true,
        cantidad: Math.max(p.stockMin * 2 - p.cantidad, p.stockMin)
      };
    }
  });
  _reabastoFiltrados = filtrados;
  pagReabasto = 1;

  const vacio = document.getElementById('reabasto-vacio');
  if (filtrados.length === 0) {
    document.getElementById('cuerpo-reabasto').innerHTML = '';
    document.getElementById('pag-reabasto').innerHTML   = '';
    vacio.classList.remove('oculto');
  } else {
    vacio.classList.add('oculto');
    renderReasbastoPage();
  }

  renderListaOrdenes();
}

function renderReasbastoPage() {
  const tbody = document.getElementById('cuerpo-reabasto');
  const pagEl = document.getElementById('pag-reabasto');
  const ini   = (pagReabasto - 1) * POR_PAG_REABASTO;
  const pagina = _reabastoFiltrados.slice(ini, ini + POR_PAG_REABASTO);

  tbody.innerHTML = pagina.map(p => {
    const est        = _reabastoState[p.id] || { checked: true, cantidad: Math.max(p.stockMin * 2 - p.cantidad, p.stockMin) };
    const provNombre = p.proveedor || 'Sin proveedor';
    return `
      <tr>
        <td><strong>${p.nombre}</strong>${p.codigo ? `<br><span style="font-family:monospace;color:var(--naranja);font-size:10px">${p.codigo}</span>` : ''}</td>
        <td><span class="cat-badge">${p.categoria}</span></td>
        <td><strong class="${estadoStock(p) === 'critico' ? 'stock-critico' : 'stock-bajo'}">${p.cantidad} ${p.unidad}</strong></td>
        <td>${p.stockMin}</td>
        <td><input type="number" class="input-cantidad-pedir" id="pedir-${p.id}"
              value="${est.cantidad}" min="1" style="width:70px"
              oninput="updateReasbastoQty('${p.id}', this.value)"></td>
        <td><span class="prov-badge">${provNombre}</span></td>
        <td><input type="checkbox" class="checkbox-incluir" id="check-${p.id}"
              ${est.checked ? 'checked' : ''}
              onchange="toggleReasbastoCheck('${p.id}')"></td>
      </tr>`;
  }).join('');

  const total = Math.ceil(_reabastoFiltrados.length / POR_PAG_REABASTO);
  if (total <= 1) { pagEl.innerHTML = ''; return; }

  let h = `<button class="btn-pag-i" onclick="irPagReabasto(${pagReabasto-1})" ${pagReabasto===1?'disabled':''}>‹</button>`;
  for (let i = 1; i <= total; i++) {
    if (total > 8 && Math.abs(i - pagReabasto) > 2 && i !== 1 && i !== total) {
      if (i === pagReabasto - 3 || i === pagReabasto + 3) h += `<span style="color:var(--texto-gris)">…</span>`;
      continue;
    }
    h += `<button class="btn-pag-i ${i===pagReabasto?'activo':''}" onclick="irPagReabasto(${i})">${i}</button>`;
  }
  h += `<button class="btn-pag-i" onclick="irPagReabasto(${pagReabasto+1})" ${pagReabasto===total?'disabled':''}>›</button>`;
  h += `<span class="pag-info-i">${_reabastoFiltrados.length} productos · pág ${pagReabasto}/${total}</span>`;
  pagEl.innerHTML = h;
}

function irPagReabasto(n) {
  const total = Math.ceil(_reabastoFiltrados.length / POR_PAG_REABASTO);
  if (n < 1 || n > total) return;
  pagReabasto = n;
  renderReasbastoPage();
  document.getElementById('tabla-reabasto').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function toggleReasbastoCheck(id) {
  if (_reabastoState[id]) _reabastoState[id].checked = !_reabastoState[id].checked;
}

function updateReasbastoQty(id, val) {
  if (_reabastoState[id]) _reabastoState[id].cantidad = Math.max(1, Number(val) || 1);
}

async function generarOrden() {
  if (_reabastoFiltrados.length === 0) { toast('Todos los productos tienen stock suficiente.', 'info'); return; }

  const seleccionados = _reabastoFiltrados.filter(p => _reabastoState[p.id]?.checked !== false);

  if (seleccionados.length === 0) { toast('Selecciona al menos un producto para la orden.', 'error'); return; }

  const items = seleccionados.map(p => ({
    id: p.id, nombre: p.nombre, codigo: p.codigo || '',
    categoria: p.categoria, cantidadActual: p.cantidad,
    cantidadPedir: Math.max(_reabastoState[p.id]?.cantidad || p.stockMin, 1),
    unidad: p.unidad, proveedor: p.proveedor || 'Sin proveedor'
  }));

  const proveedores = [...new Set(items.map(i => i.proveedor))];
  const orden = {
    id: Date.now().toString(),
    numOrden: siguienteOrdenFolio(),
    fecha: fechaISOLocal(),
    items, proveedores, estado: 'pendiente'
  };

  ordenes.unshift(orden);
  guardar();
  renderCompras();
  toast(`Orden ${orden.numOrden} generada — ${items.length} producto(s)`, 'ok', 4000);

  try { await prepararPDF(); } catch { toast('No se pudo cargar la librería PDF. Verifica tu conexión.', 'error'); return; }
  generarOrdenPDF(orden);
}

function generarOrdenPDF(orden) {
  const { jsPDF } = window.jspdf;
  const doc   = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const fecha = new Date(orden.fecha);
  const fechaStr = fecha.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const horaStr  = fecha.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });

  // Agrupar por proveedor
  const porProv = {};
  orden.items.forEach(item => {
    if (!porProv[item.proveedor]) porProv[item.proveedor] = [];
    porProv[item.proveedor].push(item);
  });

  let primera = true;
  Object.entries(porProv).forEach(([prov, items]) => {
    if (!primera) doc.addPage();
    primera = false;

    // ── ENCABEZADO ──
    doc.setFillColor(249, 115, 22);
    doc.rect(0, 0, 210, 40, 'F');
    if (typeof LOGO_B64 !== 'undefined') doc.addImage(LOGO_B64, 'JPEG', 174, 4, 28, 28);
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(20); doc.setFont('helvetica', 'bold');
    doc.text('Ferretería Las Torres', 14, 13);
    doc.setFontSize(9);  doc.setFont('helvetica', 'normal');
    doc.text('Tlapalería y Ferretería', 14, 20);
    doc.setFontSize(15); doc.setFont('helvetica', 'bold');
    doc.text('ORDEN DE COMPRA', 14, 32);

    // ── CAJA PROVEEDOR ──
    doc.setFillColor(245, 245, 245);
    doc.setDrawColor(220, 220, 220);
    doc.roundedRect(14, 46, 182, 20, 3, 3, 'FD');
    doc.setFontSize(7); doc.setFont('helvetica', 'normal'); doc.setTextColor(130, 130, 130);
    doc.text('PROVEEDOR', 20, 53);
    doc.setFontSize(15); doc.setFont('helvetica', 'bold'); doc.setTextColor(30, 30, 30);
    doc.text(prov, 20, 63);

    doc.setFontSize(8); doc.setFont('helvetica', 'normal'); doc.setTextColor(100, 100, 100);
    doc.text(`Fecha: ${fechaStr}   Hora: ${horaStr}`, 14, 74);
    doc.text(`Total de artículos en esta orden: ${items.length}`, 14, 80);

    // ── TABLA ──
    doc.autoTable({
      startY: 84,
      head: [['#', 'Código', 'Producto', 'Categoría', 'Stock', 'A pedir', 'Und.']],
      body: items.map((item, i) => [
        i + 1,
        item.codigo || '—',
        item.nombre,
        item.categoria,
        item.cantidadActual,
        item.cantidadPedir,
        item.unidad
      ]),
      foot: [['', '', `${items.length} producto${items.length !== 1 ? 's' : ''}`, '', '',
        items.reduce((s, i) => s + i.cantidadPedir, 0), 'uds. totales']],
      styles: { fontSize: 9, cellPadding: 3.5, font: 'helvetica' },
      headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255], fontStyle: 'bold' },
      footStyles: { fillColor: [240, 240, 240], fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [252, 252, 252] },
      columnStyles: {
        0: { cellWidth: 8,  halign: 'center' },
        1: { cellWidth: 22, halign: 'center' },
        4: { cellWidth: 18, halign: 'center' },
        5: { cellWidth: 18, halign: 'center', fontStyle: 'bold', textColor: [249, 115, 22] },
        6: { cellWidth: 14, halign: 'center' }
      }
    });

    // ── NOTAS ──
    const fy = doc.lastAutoTable.finalY + 8;
    doc.setDrawColor(200, 200, 200);
    doc.roundedRect(14, fy, 182, 24, 3, 3, 'D');
    doc.setFontSize(8); doc.setFont('helvetica', 'normal'); doc.setTextColor(140, 140, 140);
    doc.text('Notas / instrucciones especiales:', 19, fy + 7);

    // ── PIE ──
    const ph = doc.internal.pageSize.height;
    doc.setFontSize(7); doc.setTextColor(180, 180, 180);
    doc.text('Ferretería Las Torres — Sistema de Inventario v1.0', 105, ph - 8, { align: 'center' });
  });

  const provSlug = orden.proveedores.length === 1
    ? orden.proveedores[0].replace(/\s+/g, '-').toLowerCase()
    : 'varios-proveedores';
  doc.save(`orden-compra_${provSlug}_${fecha.toISOString().slice(0, 10)}.pdf`);
}

function marcarOrdenEnviada(id) {
  const orden = ordenes.find(o => o.id === id);
  if (orden) { orden.estado = 'enviada'; guardar(); renderListaOrdenes(); }
}

function renderListaOrdenes() {
  const lista = document.getElementById('lista-ordenes');
  const vacio = document.getElementById('ordenes-vacio');

  if (ordenes.length === 0) {
    lista.innerHTML = '';
    vacio.classList.remove('oculto');
    return;
  }

  vacio.classList.add('oculto');
  const textoEstado = { pendiente: '📤 Pendiente', enviada: '✅ Enviada', recibida: '📦 Recibida' };
  lista.innerHTML = ordenes.map(o => {
    const fecha = new Date(o.fecha).toLocaleString('es-MX', {
      day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit'
    });
    const resumen = o.items.map(i => `${i.nombre} ×${i.cantidadPedir}`).join(', ');
    const accion  = o.estado === 'pendiente' ? `marcarOrdenEnviada('${o.id}')` : 'void(0)';
    return `
      <div class="orden-item">
        <div class="orden-info">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:2px">
            <span style="font-family:monospace;color:var(--naranja);font-weight:700;font-size:13px">${o.numOrden || '—'}</span>
            <span class="orden-fecha">${fecha}</span>
          </div>
          <p class="orden-productos" title="${resumen}">${o.items.length} producto${o.items.length !== 1 ? 's' : ''}: ${resumen}</p>
          <div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:4px">
            ${o.proveedores.map(p => `<span class="prov-badge">${p}</span>`).join('')}
          </div>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;align-items:flex-end">
          <button class="btn-nota" onclick="rebajarOrdenPDF('${o.id}')">⬇ PDF</button>
          <button class="btn-orden-estado ${o.estado}" onclick="${accion}">
            ${textoEstado[o.estado] || o.estado}
          </button>
        </div>
      </div>`;
  }).join('');
}

async function rebajarOrdenPDF(id) {
  const orden = ordenes.find(o => o.id === id);
  if (!orden) return;
  try { await prepararPDF(); } catch { toast('No se pudo cargar la librería PDF.', 'error'); return; }
  generarOrdenPDF(orden);
}

// ===== RECIBIR MERCANCÍA =====
function buscarOrdenRecibir() {
  const raw    = document.getElementById('input-recibir-orden').value.trim().toUpperCase();
  const detalle = document.getElementById('recibir-detalle');

  if (!raw) {
    detalle.innerHTML = '<p style="color:var(--rojo);font-size:12px;margin-top:10px">Ingresa un número de orden.</p>';
    return;
  }

  // Acepta formatos: "OC-0001", "OC0001", "1", "0001"
  const numRaw = raw.replace(/^OC[-]?/, '').replace(/^0+/, '') || '0';
  const orden  = ordenes.find(o => {
    const oNum = (o.numOrden || '').replace(/^OC[-]?/, '').replace(/^0+/, '') || '0';
    return oNum === numRaw;
  });

  if (!orden) {
    detalle.innerHTML = `<p style="color:var(--rojo);font-size:12px;margin-top:10px">❌ No se encontró la orden "${raw}".</p>`;
    return;
  }

  renderDetalleRecibir(orden);
}

function renderDetalleRecibir(orden) {
  const detalle = document.getElementById('recibir-detalle');

  if (orden.estado === 'recibida') {
    const fechaRec = orden.fechaRecepcion
      ? new Date(orden.fechaRecepcion).toLocaleString('es-MX', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' })
      : '—';
    detalle.innerHTML = `
      <div style="margin-top:12px;padding:12px 14px;background:rgba(34,197,94,0.1);border:1px solid rgba(34,197,94,0.3);border-radius:8px">
        <p style="color:var(--verde);font-size:13px;font-weight:700">✅ La orden ${orden.numOrden} ya fue recibida</p>
        <p style="color:var(--texto-gris);font-size:12px;margin-top:4px">Recibida el: ${fechaRec}</p>
      </div>`;
    return;
  }

  const fecha = new Date(orden.fecha).toLocaleString('es-MX', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' });

  const filas = orden.items.map(item => {
    const prod = productos.find(p => p.id === item.id);
    const stockActual = prod ? prod.cantidad : null;
    const stockMostrado = stockActual !== null ? `${stockActual} ${item.unidad}` : '—';
    const nuevoStock    = stockActual !== null ? stockActual + item.cantidadPedir : '—';
    return `
      <tr>
        <td>
          <strong>${item.nombre}</strong>
          ${item.codigo ? `<br><span style="font-family:monospace;color:var(--naranja);font-size:10px">${item.codigo}</span>` : ''}
        </td>
        <td><span class="cat-badge">${item.categoria || '—'}</span></td>
        <td><strong>${item.cantidadPedir}</strong> ${item.unidad}</td>
        <td style="color:var(--texto-gris)">${stockMostrado}</td>
        <td>
          <input type="number" class="input-cantidad-pedir" id="recv-qty-${item.id}"
            value="${item.cantidadPedir}" min="0" style="width:80px"
            oninput="actualizarPreviewStock('${item.id}', ${stockActual ?? 0})">
        </td>
        <td id="recv-preview-${item.id}" style="color:var(--verde);font-weight:700">
          ${stockActual !== null ? nuevoStock : '—'}
        </td>
      </tr>`;
  }).join('');

  detalle.innerHTML = `
    <div style="margin-top:14px">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:8px">
        <div>
          <p style="font-size:13px;font-weight:700;color:var(--texto)">Orden ${orden.numOrden}</p>
          <p style="font-size:11px;color:var(--texto-gris)">Generada el ${fecha} · Proveedores: ${orden.proveedores.join(', ')}</p>
        </div>
        <button class="btn-agregar" onclick="confirmarRecepcion('${orden.id}')">✅ Confirmar recepción</button>
      </div>
      <div class="tabla-wrap" style="border:none;border-radius:0">
        <table class="tabla">
          <thead>
            <tr>
              <th>Producto</th>
              <th>Categoría</th>
              <th>Cant. pedida</th>
              <th>Stock actual</th>
              <th>Cant. recibida</th>
              <th>Nuevo stock</th>
            </tr>
          </thead>
          <tbody>${filas}</tbody>
        </table>
      </div>
    </div>`;
}

function actualizarPreviewStock(itemId, stockActual) {
  const input   = document.getElementById(`recv-qty-${itemId}`);
  const preview = document.getElementById(`recv-preview-${itemId}`);
  if (!input || !preview) return;
  const recibido = Math.max(0, Number(input.value) || 0);
  preview.textContent = stockActual + recibido;
}

function confirmarRecepcion(ordenId) {
  const orden = ordenes.find(o => o.id === ordenId);
  if (!orden || orden.estado === 'recibida') return;

  let actualizados = 0;
  orden.items.forEach(item => {
    const input        = document.getElementById(`recv-qty-${item.id}`);
    const cantRecibida = Math.max(0, Number(input?.value ?? item.cantidadPedir));
    item.cantidadRecibida = cantRecibida;

    const prod = productos.find(p => p.id === item.id);
    if (prod && cantRecibida > 0) {
      prod.cantidad += cantRecibida;
      actualizados++;
    }
  });

  orden.estado          = 'recibida';
  orden.fechaRecepcion  = fechaISOLocal();

  guardar();
  renderTodo();
  renderCompras();

  document.getElementById('input-recibir-orden').value = '';
  document.getElementById('recibir-detalle').innerHTML = `
    <div style="margin-top:12px;padding:12px 14px;background:rgba(34,197,94,0.1);border:1px solid rgba(34,197,94,0.3);border-radius:8px">
      <p style="color:var(--verde);font-size:13px;font-weight:700">✅ Recepción registrada — ${actualizados} producto${actualizados !== 1 ? 's' : ''} actualizado${actualizados !== 1 ? 's' : ''} en inventario</p>
    </div>`;

  toast(`Mercancía recibida — ${actualizados} producto(s) actualizados en inventario`, 'ok', 4000);
}

// ===== TICKET TÉRMICO (80 mm) =====
function imprimirTicket(ventaId) {
  const venta = ventas.find(v => v.id === ventaId);
  if (!venta) return;

  const folio    = venta.folio || 'S/N';
  const fecha    = new Date(venta.fecha);
  const fechaStr = fecha.toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const horaStr  = fecha.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
  const mpLabels = { efectivo: 'Efectivo', tarjeta: 'Tarjeta', transferencia: 'Transferencia' };
  const mpLabel  = mpLabels[venta.metodoPago] || 'Efectivo';

  const subtotalSinIva = venta.total / 1.16;
  const iva            = venta.total - subtotalSinIva;

  const itemsHtml = venta.items.map(item => {
    const pu       = Number(item.precioVenta ?? item.precio);
    const subtotal = pu * item.qty;
    return `
      <div class="t-item">
        <div class="t-nombre">${item.nombre}</div>
        <div class="t-fila t-det">
          <span>${item.qty} x $${pu.toFixed(2)}</span>
          <span>$${subtotal.toFixed(2)}</span>
        </div>
      </div>`;
  }).join('');

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Ticket #${folio}</title>
<style>
  @page { size: 80mm auto; margin: 2mm 3mm 6mm 3mm; }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Courier New', Courier, monospace;
    font-size: 11px;
    width: 74mm;
    color: #000;
    background: #fff;
  }
  .t-centro   { text-align: center; }
  .t-logo     { font-size: 14px; font-weight: bold; margin-bottom: 1px; }
  .t-sub      { font-size: 9px; }
  .t-sep      { border-top: 1px dashed #000; margin: 5px 0; }
  .t-fila     { display: flex; justify-content: space-between; font-size: 11px; margin: 2px 0; }
  .t-det      { font-size: 10px; color: #333; }
  .t-fila.grande {
    font-size: 14px; font-weight: bold;
    border-top: 1px solid #000; padding-top: 4px; margin-top: 2px;
  }
  .t-item     { margin: 5px 0; }
  .t-nombre   { font-weight: bold; font-size: 11px; word-break: break-word; }
  .t-pie      { text-align: center; font-size: 9px; color: #444; margin-top: 6px; }
  @media print { button { display: none !important; } }
</style>
</head>
<body>
  <div class="t-centro">
    <div class="t-logo">Ferreteria Las Torres</div>
    <div class="t-sub">Tlapaleria y Ferreteria</div>
  </div>
  <div class="t-sep"></div>
  <div class="t-fila"><span>Folio:</span><span><strong>#${folio}</strong></span></div>
  <div class="t-fila"><span>Fecha:</span><span>${fechaStr}</span></div>
  <div class="t-fila"><span>Hora:</span><span>${horaStr}</span></div>
  <div class="t-fila"><span>Pago:</span><span>${mpLabel}</span></div>
  <div class="t-sep"></div>
  ${itemsHtml}
  <div class="t-sep"></div>
  <div class="t-fila t-det"><span>Subtotal s/IVA</span><span>$${subtotalSinIva.toFixed(2)}</span></div>
  <div class="t-fila t-det"><span>IVA 16%</span><span>$${iva.toFixed(2)}</span></div>
  <div class="t-fila grande"><span>TOTAL</span><span>$${venta.total.toFixed(2)}</span></div>
  <div class="t-pie">
    <div class="t-sep"></div>
    <p>Gracias por su compra!</p>
    <p>Conserve su ticket</p>
  </div>
  <br><br><br>
</body>
</html>`;

  const w = window.open('', '_blank', 'width=380,height=600,scrollbars=yes');
  if (!w) {
    toast('El navegador bloqueó la ventana emergente. Permite ventanas emergentes para este sitio.', 'error', 5000);
    return;
  }
  w.document.write(html);
  w.document.close();
  w.focus();
  setTimeout(() => w.print(), 600);
}
