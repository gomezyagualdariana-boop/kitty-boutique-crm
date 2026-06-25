/* ===========================================================
   KITTY BOUTIQUE CRM — Barra lateral común + modal de sincronización
   =========================================================== */

function kbRenderSidebar(activePage) {
  const links = [
    { href: 'dashboard.html', icon: '📊', label: 'Resumen', key: 'dashboard' },
    { href: 'pipeline.html', icon: '🗂️', label: 'Pipeline', key: 'pipeline' },
    { href: 'fuente.html', icon: '📣', label: 'Fuente / Canal', key: 'fuente' },
    { href: 'leads.html', icon: '👥', label: 'Leads', key: 'leads' },
    { href: 'seguimientos.html', icon: '⏰', label: 'Seguimientos', key: 'seguimientos' }
  ];

  const linksHtml = links.map(l =>
    `<a class="nav-link ${l.key === activePage ? 'active' : ''}" href="${l.href}">
       <span class="ico">${l.icon}</span> ${l.label}
     </a>`).join('');

  const sidebar = document.createElement('aside');
  sidebar.className = 'sidebar';
  sidebar.innerHTML = `
    <div class="logo-wrap">
      <img src="assets/logo.png" alt="Kitty Boutique"
           onerror="this.onerror=null;this.outerHTML='<svg width=180 height=56 viewBox=\\'0 0 180 56\\'><circle cx=28 cy=28 r=26 fill=%23ffd6e8 stroke=%23e91e8c stroke-width=2/><text x=28 y=33 font-size=20 text-anchor=middle fill=%23e91e8c font-family=Georgia>KB</text><text x=62 y=24 font-size=15 fill=%23e91e8c font-weight=bold font-family=Segoe UI>KITTY</text><text x=62 y=42 font-size=12 fill=%233a2a33 font-family=Segoe UI>BOUTIQUE</text></svg>'">
    </div>
    <nav>${linksHtml}</nav>
    <button class="nav-link sync-btn" id="kb-sync-open" style="width:100%;border:1px dashed var(--kb-fuxia);background:none;cursor:pointer;">
      <span class="ico">🔗</span> Sincronizar
    </button>
    <div class="sidebar-footer">
      <button class="logout-btn" id="kb-logout-btn">⎋ Cerrar sesión</button>
    </div>
  `;

  const appShell = document.querySelector('.app-shell');
  appShell.insertBefore(sidebar, appShell.firstChild);

  document.getElementById('kb-logout-btn').addEventListener('click', kbLogout);
  document.getElementById('kb-sync-open').addEventListener('click', kbOpenSyncModal);
  kbInjectSyncModal();
}

/* ---------- Modal de sincronización con Google Sheets ---------- */
function kbInjectSyncModal() {
  if (document.getElementById('kb-sync-modal')) return;
  const wrap = document.createElement('div');
  wrap.id = 'kb-sync-modal';
  wrap.className = 'modal-overlay hidden';
  wrap.innerHTML = `
    <div class="modal-box">
      <h2>🔗 Sincronizar con Google Sheets</h2>
      <p class="modal-sub">
        Sincronización manual de reemplazo completo, en una sola dirección por clic.
        No es fusión automática ni tiempo real (no es viable sin un servidor 24/7
        detrás de una página HTML estática).
      </p>
      <div class="form-field full">
        <label>URL de tu Apps Script desplegado (Code.gs → Implementar → Aplicación web)</label>
        <input type="text" id="kb-sync-url" placeholder="https://script.google.com/macros/s/XXXX/exec">
      </div>
      <div id="kb-sync-status" style="font-size:12.5px;margin-top:8px;min-height:18px;"></div>
      <div class="modal-actions" style="justify-content:space-between;flex-wrap:wrap;">
        <button class="btn btn-outline" id="kb-sync-save">Guardar URL</button>
        <div style="display:flex;gap:10px;">
          <button class="btn btn-secondary" id="kb-sync-pull">⬇ Traer de Sheets</button>
          <button class="btn btn-primary" id="kb-sync-push">⬆ Subir a Sheets</button>
          <button class="btn btn-outline" id="kb-sync-close">Cerrar</button>
        </div>
      </div>
    </div>`;
  document.body.appendChild(wrap);

  const urlInput = document.getElementById('kb-sync-url');
  urlInput.value = localStorage.getItem(KB_SYNC_URL_KEY) || '';

  document.getElementById('kb-sync-close').addEventListener('click', kbCloseSyncModal);
  document.getElementById('kb-sync-save').addEventListener('click', () => {
    localStorage.setItem(KB_SYNC_URL_KEY, urlInput.value.trim());
    kbSyncStatus('URL guardada.', false);
    kbToast('URL de sincronización guardada', 'success');
  });
  document.getElementById('kb-sync-pull').addEventListener('click', kbSyncPull);
  document.getElementById('kb-sync-push').addEventListener('click', kbSyncPush);
}

function kbSyncStatus(msg, isError) {
  const el = document.getElementById('kb-sync-status');
  el.textContent = msg;
  el.style.color = isError ? 'var(--kb-danger)' : 'var(--kb-success)';
}

function kbOpenSyncModal() { document.getElementById('kb-sync-modal').classList.remove('hidden'); }
function kbCloseSyncModal() { document.getElementById('kb-sync-modal').classList.add('hidden'); }

function kbGetSyncUrl() {
  const url = (document.getElementById('kb-sync-url') || {}).value || localStorage.getItem(KB_SYNC_URL_KEY);
  return (url || '').trim();
}

async function kbSyncPull() {
  const url = kbGetSyncUrl();
  if (!url) { kbSyncStatus('Primero guarda la URL de tu Apps Script.', true); return; }
  kbSyncStatus('Conectando con Google Sheets...', false);
  try {
    const res = await fetch(url, { method: 'GET' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const json = await res.json();
    const leads = Array.isArray(json) ? json : json.leads;
    if (!Array.isArray(leads)) throw new Error('Respuesta inesperada del servidor');
    kbSaveLeads(leads);
    kbSyncStatus(`Listo: se reemplazaron los datos locales con ${leads.length} filas de Sheets.`, false);
    kbToast('Datos traídos de Google Sheets', 'success');
  } catch (err) {
    kbSyncStatus('Error al traer datos: ' + err.message + '. Revisa la URL y que el script esté implementado como "Cualquier usuario".', true);
    kbToast('Error de sincronización', 'error');
  }
}

async function kbSyncPush() {
  const url = kbGetSyncUrl();
  if (!url) { kbSyncStatus('Primero guarda la URL de tu Apps Script.', true); return; }
  kbSyncStatus('Subiendo datos a Google Sheets...', false);
  try {
    const leads = kbGetLeads();
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ accion: 'reemplazarTodo', leads: leads })
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const json = await res.json();
    if (json.ok === false) throw new Error(json.error || 'el servidor respondió con error');
    kbSyncStatus(`Listo: se subieron ${leads.length} filas, reemplazando todo en la hoja.`, false);
    kbToast('Datos subidos a Google Sheets', 'success');
  } catch (err) {
    kbSyncStatus('Error al subir datos: ' + err.message + '. Revisa la URL y los permisos de implementación.', true);
    kbToast('Error de sincronización', 'error');
  }
}
