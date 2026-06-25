/* ===========================================================
   KITTY BOUTIQUE CRM — Autenticación (DEMO / LOCAL ONLY)

   IMPORTANTE: Esto es protección de demostración para uso local
   en un solo navegador. La contraseña se verifica comparando un
   hash SHA-256 calculado en el cliente con un hash guardado en
   este mismo archivo, y la sesión se guarda en sessionStorage.
   Esto NO es seguridad real para un sitio multiusuario expuesto
   a internet: cualquiera que abra las herramientas de desarrollador
   puede leer este archivo. Para producción real se necesitaría un
   backend con autenticación de servidor.
   =========================================================== */

const KB_SESSION_KEY = 'kb_session_ok';

// Usuario: admin   Contraseña: kitty2026
// Hash SHA-256 de "kitty2026"
const KB_USER = 'admin';
const KB_PASS_HASH = '7ec5e5e9be7cea127299d67395aae1d8c42c3cc2440174f02788fafeb1124125';

async function kbSha256(text) {
  const enc = new TextEncoder().encode(text);
  const buf = await crypto.subtle.digest('SHA-256', enc);
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
}

function kbIsLoggedIn() {
  return sessionStorage.getItem(KB_SESSION_KEY) === '1';
}

function kbRequireAuth() {
  if (!kbIsLoggedIn()) {
    window.location.href = 'login.html';
  }
}

function kbLogout() {
  sessionStorage.removeItem(KB_SESSION_KEY);
  window.location.href = 'login.html';
}
