/* ===========================================================
   KITTY BOUTIQUE CRM — Capa de datos (localStorage)
   Estructura basada en el Excel real "Kitty_Boutique_100_Clientes.xlsx":
   Fecha, Cliente, Empresa, Producto, Valor USD, Estado, Fuente,
   Prioridad, Mes, Notas, Fecha Seguimiento
   =========================================================== */

const KB_STORAGE_KEY = 'kb_leads_v1';
const KB_SYNC_URL_KEY = 'kb_sync_url';

const KB_ESTADOS = ['Pendiente', 'Contactada', 'En negociación', 'Apartado', 'Vendido'];
const KB_ESTADO_NEXT = {
  'Pendiente': 'Contactada',
  'Contactada': 'En negociación',
  'En negociación': 'Apartado',
  'Apartado': 'Vendido',
  'Vendido': 'Vendido'
};
const KB_FUENTES = ['Página Web', 'Instagram', 'Facebook', 'TikTok', 'Referido', 'WhatsApp'];
const KB_PRODUCTOS = ['Falda Jeans', 'Top Elegante', 'Accesorios', 'Conjunto Casual', 'Bolso Fashion',
  'Blazer Mujer', 'Vestido Floral', 'Blusa Casual', 'Pantalón Jeans', 'Blusa Suplex'];
const KB_PRIORIDADES = ['Alta', 'Media', 'Baja'];
const KB_MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio'];
const KB_EMPRESA = 'Kitty Boutique';

const KB_BADGE_MAP = {
  'Pendiente': 'badge-Pendiente',
  'Contactada': 'badge-Contactada',
  'En negociación': 'badge-EnNegociacion',
  'Apartado': 'badge-Apartado',
  'Vendido': 'badge-Vendido'
};
function kbBadgeClass(estado) { return KB_BADGE_MAP[estado] || 'badge-Pendiente'; }

/* ---------- Generador de nombres realistas (estilo del Excel) ---------- */
const KB_NOMBRES = ['Ana', 'Cinthya', 'Nicole', 'Verónica', 'María José', 'Gabriela', 'Daniela', 'Karla',
  'Fernanda', 'Valeria', 'Andrea', 'Camila', 'Lucía', 'Mishell', 'Paola', 'Ximena', 'Briggitte', 'Yadira',
  'Estefanía', 'Mayra', 'Johanna', 'Doménica', 'Melanie', 'Jazmín', 'Liseth'];
const KB_APELLIDOS = ['Panchana', 'Mendoza', 'Rodríguez', 'Gómez', 'Alvarado', 'Suárez', 'Zambrano',
  'Cedeño', 'Vélez', 'Macías', 'Torres', 'Pincay', 'Loor', 'Quiñonez', 'Yagual', 'Tomalá', 'Bravo',
  'Reyes', 'Chávez', 'Borbor'];

function kbRand(arr) { return arr[Math.floor(Math.random() * arr.length)]; }
function kbRandFloat(min, max) { return Math.round((Math.random() * (max - min) + min) * 100) / 100; }
function kbRandInt(min, max) { return Math.floor(Math.random() * (max - min + 1)) + min; }
function kbPad(n) { return n < 10 ? '0' + n : '' + n; }

function kbFormatDate(d) {
  return kbPad(d.getDate()) + '/' + kbPad(d.getMonth() + 1) + '/' + d.getFullYear();
}

function kbMesNombre(d) { return KB_MESES[d.getMonth() % KB_MESES.length] || KB_MESES[(d.getMonth()) % 6]; }

const KB_PRODUCT_NOTE = {
  'Falda Jeans': 'falda jeans', 'Top Elegante': 'top elegante', 'Accesorios': 'accesorios',
  'Conjunto Casual': 'conjunto casual', 'Bolso Fashion': 'bolso fashion', 'Blazer Mujer': 'blazer mujer',
  'Vestido Floral': 'vestido floral', 'Blusa Casual': 'blusa casual', 'Pantalón Jeans': 'pantalón jeans',
  'Blusa Suplex': 'blusa suplex'
};

function kbGenerarLeads(cantidad) {
  const leads = [];
  const base = new Date(2026, 0, 1);
  for (let i = 0; i < cantidad; i++) {
    const fecha = new Date(base.getTime());
    fecha.setDate(fecha.getDate() + Math.floor(i * (180 / cantidad)) + kbRandInt(0, 2));
    const seguimiento = new Date(fecha.getTime());
    seguimiento.setDate(seguimiento.getDate() + kbRandInt(3, 25));

    const producto = kbRand(KB_PRODUCTOS);
    const estado = kbRand(KB_ESTADOS);
    leads.push({
      id: 'L' + (1000 + i),
      fecha: kbFormatDate(fecha),
      cliente: kbRand(KB_NOMBRES) + ' ' + kbRand(KB_APELLIDOS),
      empresa: KB_EMPRESA,
      producto: producto,
      valor: kbRandFloat(6.5, 25),
      estado: estado,
      fuente: kbRand(KB_FUENTES),
      prioridad: kbRand(KB_PRIORIDADES),
      mes: kbMesNombre(fecha),
      notas: 'Interesada en ' + KB_PRODUCT_NOTE[producto],
      fechaSeguimiento: kbFormatDate(seguimiento)
    });
  }
  return leads;
}

function kbGetLeads() {
  const raw = localStorage.getItem(KB_STORAGE_KEY);
  if (!raw) {
    const seed = kbGenerarLeads(126);
    localStorage.setItem(KB_STORAGE_KEY, JSON.stringify(seed));
    return seed;
  }
  try { return JSON.parse(raw); } catch (e) { return []; }
}

function kbSaveLeads(leads) {
  localStorage.setItem(KB_STORAGE_KEY, JSON.stringify(leads));
}

function kbAddLead(lead) {
  const leads = kbGetLeads();
  lead.id = 'L' + Date.now();
  leads.unshift(lead);
  kbSaveLeads(leads);
  return lead;
}

function kbUpdateLead(id, data) {
  const leads = kbGetLeads();
  const idx = leads.findIndex(l => l.id === id);
  if (idx === -1) return null;
  leads[idx] = Object.assign({}, leads[idx], data, { id });
  kbSaveLeads(leads);
  return leads[idx];
}

function kbDeleteLead(id) {
  const leads = kbGetLeads().filter(l => l.id !== id);
  kbSaveLeads(leads);
}

/* ---------- Parsing de fechas dd/mm/yyyy ---------- */
function kbParseDate(str) {
  if (!str) return null;
  const parts = str.split('/');
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
}

function kbToast(msg, type) {
  let t = document.getElementById('kb-toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'kb-toast';
    t.className = 'toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.className = 'toast show' + (type ? ' ' + type : '');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 3200);
}

function kbFormatUSD(n) {
  return '$' + Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
