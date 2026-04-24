/**
 * TerrierGPT Lite — Chat Widget
 * -----------------------------
 * Embeddable chat widget targeting the new NEXUS Agent backend
 * (de-nexus-chat, FastAPI). Two endpoints are supported out of the box:
 *
 *   * /api/v1/lite-chat  — widget-facing endpoint that gates BU-specific
 *                          queries behind an openid (emits `auth_required`
 *                          SSE frames the widget handles below).
 *
 * Embedding:
 *   <script src="http://localhost:4001/static/widgets/chat.js"></script>
 *
 * Public API (window.TerrierChatWidget):
 *   .show() / .hide() / .toggle() / .unmount() / .reset()
 *   .mounted                     — boolean
 *   .conversationId              — server-assigned conversation id (or null)
 *   .getOpenid()                 — current openid (or null)
 *   .getOpenidSource()           — where we got it from, for debugging
 *   .getEmail()                  — current email (or null)
 *   .setOpenid(value[, opts])    — host pushes openid (opts: {source, email})
 *   .setEmail(email)             — host pushes email only
 *   .refreshOpenidFromHost()     — re-run host-page probes (for SPAs)
 *   .signInSilent()              — MSAL ssoSilent()
 *   .signInPopup()               — MSAL loginPopup()
 *   .signIn()                    — silent first, popup fallback
 *
 * Authentication strategies (all optional, all additive):
 *   1. Host-page probing — scans window globals, localStorage/sessionStorage,
 *      and cookies for an existing Entra openid.
 *   2. Host-API retrieval — before an auth-gated request succeeds, calls a
 *      host "who am I" endpoint (e.g. /about/me) and extracts openid/email
 *      via `hostApi.openidPath` / `hostApi.emailPath` dot-paths or a
 *      `hostApi.mapper(response)` function.
 *   3. MSAL silent login — loads msal-browser from a CDN and calls
 *      ssoSilent() when the backend signals `auth_required`.
 *   4. MSAL popup login — interactive popup when silent is unavailable
 *      or disabled.
 *
 * Configuration via `window.TerrierChatWidgetConfig` set BEFORE the script
 * loads. Full shape (all optional):
 *   window.TerrierChatWidgetConfig = {
 *     backendUrl: 'http://localhost:9090/api/v1/lite-chat',
 *     chatDefaults: { provider: 'azure_openai', model: null, ... },
 *     extractOpenidFromHost: true,
 *     allowMsalFallback: true,
 *     allowMsalPopupLogin: true,
 *     autoRetryAfterAuth: true,
 *     msal: { clientId, authority, redirectUri, scopes, loginHint, scriptUrl },
 *     allowToRetrieveFromApi: false,
 *     hostApi: { endpoint, method, credentials, headers, openidPath, emailPath, mapper },
 *   };
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Configuration
  // ---------------------------------------------------------------------------
  const userConfig = (typeof window !== 'undefined' && window.TerrierChatWidgetConfig) || {};
  const CONFIG = Object.assign({
    backendUrl: 'http://localhost:9090/api/v1/chat',
    title: 'TerrierGPT Lite',
    placeholder: 'Ask me anything…',
    startHidden: true,
    showLauncher: true,
    launcherLabel: '💬',
    // When true, surface tool-use events as faint system-style messages.
    // Set to false if you want a clean user/assistant-only transcript.
    showToolEvents: true,

    // --- Entra ID / openid options (ported from chat.js) ----------------
    // Scan the host page at mount time for an existing Entra openid.
    extractOpenidFromHost: true,
    // On `auth_required`, enable MSAL silent sign-in (hidden iframe).
    // Requires `msal.clientId` and `msal.authority`.
    allowMsalFallback: true,
    // On `auth_required`, allow an interactive MSAL popup login.
    // Requires `msal.clientId` and `msal.authority`. When both silent
    // and popup are enabled, silent is tried first.
    allowMsalPopupLogin: true,
    // After a successful login (or host-API retrieval), automatically
    // re-send the last query.
    autoRetryAfterAuth: true,

    // --- Host-API retrieval strategy ------------------------------------
    // When true, on `auth_required` the widget calls the host REST
    // endpoint (usually a cookie-auth'd "who am I" route) and extracts
    // openid/email via `hostApi` mapping. Takes precedence over MSAL
    // when enabled.
    allowToRetrieveFromApi: false,
  }, userConfig);

  // Defaults merged into every POST body. Hosts can override any subset
  // (e.g. just `provider`) without wiping the rest. `null`/`undefined`
  // values are stripped at send time so Pydantic uses its own defaults.
  CONFIG.chatDefaults = Object.assign({
    provider: 'azure_openai',
    model: null,
    deployment: null,
    system_message_id: null,
  }, (userConfig && userConfig.chatDefaults) || {});

  // MSAL config, deep-merged so hosts can override only what they need.
  CONFIG.msal = Object.assign({
    clientId: null,
    authority: null,
    redirectUri: null,
    scopes: ['openid', 'profile'],
    loginHint: null,
    scriptUrl: 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.30.0/lib/msal-browser.min.js',
  }, (userConfig && userConfig.msal) || {});

  // Host-API retrieval config.
  CONFIG.hostApi = Object.assign({
    endpoint: null,
    method: 'GET',
    credentials: 'same-origin',
    headers: null,
    openidPath: 'openid',
    emailPath: 'email',
    // Optional escape hatch: function(responseJson) => { openid, email }.
    mapper: null,
  }, (userConfig && userConfig.hostApi) || {});

  const WIDGET_ID = 'tgv2-chat-widget-root';
  const LAUNCHER_ID = 'tgv2-chat-widget-launcher';
  const STYLE_ID = 'tgv2-chat-widget-style';

  // Conversation state — retained across turns so multi-turn history works
  // on the server. Reset via window.TerrierChatWidget.reset().
  let conversationId = null;

  // Auth state.
  let currentOpenid = null;
  let openidSource = null;   // 'global:...' | 'localStorage:...' | 'cookie:...' | 'msal:...' | 'host-api' | 'host'
  let currentEmail = null;
  // Stash of the last sent user query, for auto-retry after authentication.
  let lastQuery = null;

  // Prevent double-injection if the script is loaded twice.
  if (typeof document === 'undefined') return;
  if (document.getElementById(WIDGET_ID)) return;

  // ---------------------------------------------------------------------------
  // Styles (scoped by root id — class names intentionally match chat.js so
  // both widgets can coexist without colliding).
  // ---------------------------------------------------------------------------
  const CSS = `
#${WIDGET_ID} {
  position: fixed;
  bottom: 90px;
  right: 20px;
  width: 360px;
  height: 520px;
  max-height: calc(100vh - 120px);
  background: #ffffff;
  border: 1px solid #e5e5e5;
  border-radius: 12px;
  box-shadow: 0 12px 36px rgba(0, 0, 0, 0.18);
  display: flex;
  flex-direction: column;
  overflow: hidden;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 14px;
  color: #111;
  z-index: 2147483647;
}
#${WIDGET_ID}.tg-hidden { display: none; }

#${WIDGET_ID} .tg-header {
  padding: 12px 16px;
  background: #cc0000;
  color: #fff;
  font-weight: 600;
  display: flex;
  align-items: center;
  justify-content: space-between;
  user-select: none;
}
#${WIDGET_ID} .tg-close {
  background: transparent;
  border: none;
  color: #fff;
  font-size: 22px;
  line-height: 1;
  cursor: pointer;
  padding: 0 4px;
}
#${WIDGET_ID} .tg-close:hover { opacity: 0.8; }

#${WIDGET_ID} .tg-conv {
  flex: 1;
  overflow-y: auto;
  padding: 12px;
  background: #fafafa;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

#${WIDGET_ID} .tg-msg {
  padding: 8px 12px;
  border-radius: 12px;
  max-width: 85%;
  white-space: pre-wrap;
  word-wrap: break-word;
  line-height: 1.4;
}
#${WIDGET_ID} .tg-msg-user {
  background: #cc0000;
  color: #fff;
  align-self: flex-end;
  border-bottom-right-radius: 4px;
}
#${WIDGET_ID} .tg-msg-assistant {
  background: #ececec;
  color: #111;
  align-self: flex-start;
  border-bottom-left-radius: 4px;
}
#${WIDGET_ID} .tg-msg-system {
  background: transparent;
  color: #888;
  align-self: center;
  font-size: 12px;
  font-style: italic;
}
#${WIDGET_ID} .tg-msg-tool {
  background: #f3f3f3;
  color: #555;
  align-self: flex-start;
  font-size: 12px;
  border-left: 3px solid #cc0000;
  border-radius: 4px;
  max-width: 95%;
}
#${WIDGET_ID} .tg-msg.tg-streaming::after {
  content: '▍';
  display: inline-block;
  margin-left: 2px;
  animation: tgv2-blink 1s steps(2, start) infinite;
}
#${WIDGET_ID} .tg-round-break {
  align-self: stretch;
  border-top: 1px dashed #ddd;
  margin: 4px 0;
}
@keyframes tgv2-blink { to { visibility: hidden; } }

#${WIDGET_ID} .tg-auth-link {
  display: inline-block;
  margin-top: 6px;
  padding: 6px 10px;
  background: #cc0000;
  color: #fff;
  border-radius: 6px;
  cursor: pointer;
  font-weight: 600;
  font-size: 13px;
  user-select: none;
}
#${WIDGET_ID} .tg-auth-link:hover { background: #a60000; }
#${WIDGET_ID} .tg-auth-link.tg-auth-busy { opacity: 0.6; pointer-events: none; }
#${WIDGET_ID} .tg-auth-note { display: block; margin-top: 4px; font-size: 12px; color: #666; }

#${WIDGET_ID} .tg-input-row {
  display: flex;
  gap: 8px;
  padding: 10px;
  border-top: 1px solid #eee;
  background: #fff;
}
#${WIDGET_ID} .tg-input {
  flex: 1;
  resize: none;
  border: 1px solid #ddd;
  border-radius: 8px;
  padding: 8px 10px;
  font-family: inherit;
  font-size: 14px;
  height: 40px;
  max-height: 120px;
  outline: none;
  box-sizing: border-box;
}
#${WIDGET_ID} .tg-input:focus { border-color: #cc0000; }
#${WIDGET_ID} .tg-send {
  padding: 0 16px;
  background: #cc0000;
  color: #fff;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  font-weight: 600;
  font-size: 14px;
}
#${WIDGET_ID} .tg-send:disabled { opacity: 0.5; cursor: not-allowed; }

#${LAUNCHER_ID} {
  position: fixed;
  bottom: 20px;
  right: 20px;
  width: 56px;
  height: 56px;
  border-radius: 50%;
  border: none;
  background: #cc0000;
  color: #fff;
  font-size: 24px;
  line-height: 1;
  cursor: pointer;
  box-shadow: 0 6px 18px rgba(0, 0, 0, 0.25);
  z-index: 2147483646;
  transition: transform 0.15s ease;
  display: flex;
  align-items: center;
  justify-content: center;
}
#${LAUNCHER_ID}:hover { transform: scale(1.05); }
#${LAUNCHER_ID}:active { transform: scale(0.95); }
#${LAUNCHER_ID}.tg-hidden { display: none; }
`;

  const styleEl = document.createElement('style');
  styleEl.id = STYLE_ID;
  styleEl.setAttribute('data-tgv2-chat-widget', 'true');
  styleEl.textContent = CSS;
  document.head.appendChild(styleEl);

  // ---------------------------------------------------------------------------
  // DOM
  // ---------------------------------------------------------------------------
  const root = document.createElement('div');
  root.id = WIDGET_ID;
  if (CONFIG.startHidden) root.classList.add('tg-hidden');
  root.innerHTML = `
    <div class="tg-header">
      <span class="tg-title"></span>
      <button type="button" class="tg-close" aria-label="Close">×</button>
    </div>
    <div class="tg-conv" role="log" aria-live="polite"></div>
    <div class="tg-input-row">
      <textarea class="tg-input" rows="1"></textarea>
      <button type="button" class="tg-send">Send</button>
    </div>
  `;

  let launcherEl = null;
  if (CONFIG.showLauncher) {
    launcherEl = document.createElement('button');
    launcherEl.id = LAUNCHER_ID;
    launcherEl.type = 'button';
    launcherEl.setAttribute('aria-label', 'Open chat');
    launcherEl.textContent = CONFIG.launcherLabel;
    if (!CONFIG.startHidden) launcherEl.classList.add('tg-hidden');
  }

  const mount = () => {
    document.body.appendChild(root);
    if (launcherEl) document.body.appendChild(launcherEl);
  };
  if (document.body) mount();
  else document.addEventListener('DOMContentLoaded', mount, { once: true });

  const titleEl = root.querySelector('.tg-title');
  const convEl = root.querySelector('.tg-conv');
  const inputEl = root.querySelector('.tg-input');
  const sendBtn = root.querySelector('.tg-send');
  const closeBtn = root.querySelector('.tg-close');

  titleEl.textContent = CONFIG.title;
  inputEl.setAttribute('placeholder', CONFIG.placeholder);

  addMessage('system', 'Connected. Type a message below.');

  // ---------------------------------------------------------------------------
  // Conversation helpers
  // ---------------------------------------------------------------------------
  function addMessage(role, text) {
    const el = document.createElement('div');
    el.className = 'tg-msg tg-msg-' + role;
    el.textContent = text || '';
    convEl.appendChild(el);
    convEl.scrollTop = convEl.scrollHeight;
    return el;
  }

  function appendTo(el, text) {
    el.textContent += text;
    convEl.scrollTop = convEl.scrollHeight;
  }

  function addRoundBreak() {
    const el = document.createElement('div');
    el.className = 'tg-round-break';
    convEl.appendChild(el);
    convEl.scrollTop = convEl.scrollHeight;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // Openid helpers (ported from chat.js)
  // ═══════════════════════════════════════════════════════════════════════════

  function looksLikeOpenid(val) {
    if (typeof val !== 'string') return false;
    const s = val.trim();
    if (s.length < 8 || s.length > 8192) return false;
    // GUID 8-4-4-4-12 hex
    if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s)) return true;
    // JWT: header.payload.signature
    if (/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(s)) return true;
    // Last resort: long opaque string.
    return s.length >= 16 && !/\s/.test(s);
  }

  function decodeJwtPayload(jwt) {
    try {
      const parts = String(jwt).split('.');
      if (parts.length !== 3) return null;
      const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      const json = decodeURIComponent(
        atob(b64).split('').map((c) =>
          '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2)
        ).join('')
      );
      return JSON.parse(json);
    } catch (_) {
      return null;
    }
  }

  function oidFromJwt(jwt) {
    const claims = decodeJwtPayload(jwt);
    if (!claims) return null;
    return claims.oid || claims.sub || null;
  }

  // ---------------------------------------------------------------------------
  // Strategy 1 — host-page probing (globals, storage, cookies)
  // ---------------------------------------------------------------------------

  function tryExtractFromGlobals() {
    const candidates = [
      'openid', '__openid__',
      'userOpenid', 'currentOpenid',
      'idToken', '__idToken__',
      'userIdToken',
    ];
    for (const name of candidates) {
      const v = window[name];
      if (looksLikeOpenid(v)) return { value: v, source: 'global:' + name };
    }
    const containers = ['user', 'currentUser', 'authContext', 'msalAccount'];
    for (const name of containers) {
      const obj = window[name];
      if (obj && typeof obj === 'object') {
        const v = obj.oid || obj.openid || obj.id_token || obj.idToken || obj.sub;
        if (looksLikeOpenid(v)) return { value: v, source: 'global:' + name };
      }
    }
    const msalHandles = ['msalInstance', '_msalInstance', '__msal__'];
    for (const name of msalHandles) {
      const inst = window[name];
      if (!inst || typeof inst.getAllAccounts !== 'function') continue;
      try {
        const accounts = inst.getAllAccounts() || [];
        for (const acct of accounts) {
          const oid = acct.idTokenClaims && (acct.idTokenClaims.oid || acct.idTokenClaims.sub);
          if (looksLikeOpenid(oid)) return { value: oid, source: 'global:msal:' + name };
          if (acct.idToken && looksLikeOpenid(acct.idToken)) {
            return { value: oidFromJwt(acct.idToken) || acct.idToken, source: 'global:msal:' + name };
          }
        }
      } catch (_) { /* defensive */ }
    }
    return null;
  }

  function tryExtractFromStorage(storage, label) {
    if (!storage) return null;
    let best = null;
    try {
      for (let i = 0; i < storage.length; i++) {
        const key = storage.key(i);
        if (!key) continue;
        const keyLower = key.toLowerCase();

        const looksMsal = keyLower.includes('idtoken')
          || keyLower.startsWith('msal.')
          || keyLower.includes('login.windows.net')
          || keyLower.includes('login.microsoftonline.com');
        if (!looksMsal) continue;

        const raw = storage.getItem(key);
        if (!raw) continue;

        try {
          const parsed = JSON.parse(raw);
          if (parsed && parsed.credentialType === 'IdToken' && looksLikeOpenid(parsed.secret)) {
            const oid = oidFromJwt(parsed.secret) || parsed.secret;
            best = { value: oid, source: label + ':' + key };
            break;
          }
          if (parsed && (parsed.oid || parsed.openid)) {
            const v = parsed.oid || parsed.openid;
            if (looksLikeOpenid(v)) {
              best = { value: v, source: label + ':' + key };
              break;
            }
          }
        } catch (_) {
          if (looksLikeOpenid(raw)) {
            best = { value: oidFromJwt(raw) || raw, source: label + ':' + key };
            break;
          }
        }
      }
    } catch (_) { /* storage access may throw in sandboxed iframes */ }
    return best;
  }

  function tryExtractFromCookies() {
    if (typeof document === 'undefined' || !document.cookie) return null;
    const wanted = ['openid', 'id_token', 'idtoken', 'user_oid', 'auth_token'];
    const pairs = document.cookie.split(';').map((p) => p.trim());
    for (const pair of pairs) {
      const eq = pair.indexOf('=');
      if (eq < 0) continue;
      const name = pair.slice(0, eq).toLowerCase();
      const val = decodeURIComponent(pair.slice(eq + 1));
      if (wanted.includes(name) && looksLikeOpenid(val)) {
        return { value: oidFromJwt(val) || val, source: 'cookie:' + name };
      }
    }
    return null;
  }

  function extractOpenidFromHost() {
    if (!CONFIG.extractOpenidFromHost) return;
    const found =
      tryExtractFromGlobals()
      || tryExtractFromStorage(window.localStorage, 'localStorage')
      || tryExtractFromStorage(window.sessionStorage, 'sessionStorage')
      || tryExtractFromCookies();
    if (found) {
      currentOpenid = found.value;
      openidSource = found.source;
      try { console.debug('[TerrierChatWidget] openid detected via', openidSource); } catch (_) {}
    }
  }

  // ---------------------------------------------------------------------------
  // Strategy 2 — host-API retrieval
  //
  // Returns {ok: true} on success (openid populated) or
  // {ok: false, error: <human-readable string>} on any failure.
  // ---------------------------------------------------------------------------

  function readJsonPath(obj, dotPath) {
    if (!dotPath || typeof dotPath !== 'string') return undefined;
    return dotPath.split('.').reduce(
      (acc, key) => (acc != null && typeof acc === 'object') ? acc[key] : undefined,
      obj
    );
  }

  function isHostApiStrategyEnabled() {
    return !!(CONFIG.allowToRetrieveFromApi && CONFIG.hostApi && CONFIG.hostApi.endpoint);
  }

  async function retrieveFromHostApi() {
    if (!CONFIG.allowToRetrieveFromApi) {
      return { ok: false, error: 'Host-API strategy is disabled.' };
    }
    const api = CONFIG.hostApi || {};
    if (!api.endpoint) {
      return { ok: false, error: 'Host-API endpoint is not configured.' };
    }

    let resp;
    try {
      resp = await fetch(api.endpoint, {
        method: api.method || 'GET',
        credentials: api.credentials || 'same-origin',
        headers: api.headers || undefined,
      });
    } catch (err) {
      const msg = err && err.message ? err.message : String(err);
      return { ok: false, error: 'Network error calling host API: ' + msg };
    }

    if (!resp.ok) {
      return { ok: false, error: 'Host API returned HTTP ' + resp.status + '.' };
    }

    let data;
    try {
      data = await resp.json();
    } catch (_) {
      return { ok: false, error: 'Host API response was not valid JSON.' };
    }

    let openid, email;
    if (typeof api.mapper === 'function') {
      const mapped = api.mapper(data) || {};
      openid = mapped.openid;
      email = mapped.email;
    } else {
      openid = readJsonPath(data, api.openidPath);
      email = readJsonPath(data, api.emailPath);
    }

    const openidStr = typeof openid === 'string' ? openid.trim() : '';
    if (!openidStr) {
      return {
        ok: false,
        error: 'Host API response did not include an openid at path "'
               + (api.openidPath || '(mapper)') + '".',
      };
    }

    currentOpenid = openidStr;
    openidSource = 'host-api';
    if (typeof email === 'string' && email.trim()) {
      currentEmail = email.trim();
    }
    return { ok: true };
  }

  // ---------------------------------------------------------------------------
  // Strategy 3+4 — MSAL silent / popup login
  // ---------------------------------------------------------------------------

  let msalLoadPromise = null;
  let msalInstance = null;

  function loadMsalScript() {
    if (window.msal && typeof window.msal.PublicClientApplication === 'function') {
      return Promise.resolve(window.msal);
    }
    if (msalLoadPromise) return msalLoadPromise;
    msalLoadPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = CONFIG.msal.scriptUrl;
      s.async = true;
      s.onload = () => {
        if (window.msal && typeof window.msal.PublicClientApplication === 'function') {
          resolve(window.msal);
        } else {
          reject(new Error('msal-browser loaded but window.msal is unavailable'));
        }
      };
      s.onerror = () => reject(new Error('Failed to load msal-browser from ' + CONFIG.msal.scriptUrl));
      document.head.appendChild(s);
    });
    return msalLoadPromise;
  }

  async function getMsalInstance() {
    if (msalInstance) return msalInstance;
    const { clientId, authority, redirectUri } = CONFIG.msal;
    if (!clientId || !authority) {
      throw new Error(
        'MSAL is not configured. Set window.TerrierChatWidgetConfig.msal.'
        + '{clientId, authority} before the widget loads.'
      );
    }
    const msal = await loadMsalScript();
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId,
        authority,
        redirectUri: redirectUri || window.location.origin,
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
      },
    });
    if (typeof msalInstance.initialize === 'function') {
      await msalInstance.initialize();  // msal-browser v3+ requires this
    }
    return msalInstance;
  }

  async function silentEntraLogin() {
    if (!CONFIG.allowMsalFallback) {
      throw new Error('Silent Entra login is disabled (allowMsalFallback=false).');
    }
    const inst = await getMsalInstance();
    const result = await inst.ssoSilent({
      scopes: CONFIG.msal.scopes,
      loginHint: CONFIG.msal.loginHint || undefined,
    });
    return extractOpenidFromMsalResult(result, 'msal:ssoSilent');
  }

  async function popupEntraLogin() {
    if (!CONFIG.allowMsalPopupLogin) {
      throw new Error('Popup Entra login is disabled (allowMsalPopupLogin=false).');
    }
    const inst = await getMsalInstance();
    const result = await inst.loginPopup({
      scopes: CONFIG.msal.scopes,
      loginHint: CONFIG.msal.loginHint || undefined,
    });
    return extractOpenidFromMsalResult(result, 'msal:loginPopup');
  }

  function extractOpenidFromMsalResult(result, source) {
    const claims = (result && result.idTokenClaims) || {};
    const account = (result && result.account) || {};

    const oid = claims.oid || claims.sub || (result && result.idToken ? oidFromJwt(result.idToken) : null);
    const value = oid || (result && result.idToken) || null;
    if (!value) throw new Error('Login succeeded but no openid/oid claim was returned.');

    let email = claims.email || claims.preferred_username || claims.upn || account.username || null;
    if (!email && result && result.idToken) {
      const decoded = decodeJwtPayload(result.idToken) || {};
      email = decoded.email || decoded.preferred_username || decoded.upn || null;
    }

    currentOpenid = value;
    openidSource = source;
    currentEmail = email || null;
    return value;
  }

  // ---------------------------------------------------------------------------
  // Request body builder
  //
  // Merges chatDefaults + conversation_id + current openid/email into the
  // per-turn fields, then strips null/undefined so Pydantic uses its own
  // defaults server-side. Empty strings for openid/email are also dropped
  // so the lite-chat endpoint sees `None` and evaluates authenticated=False.
  // ---------------------------------------------------------------------------
  function buildRequestBody(message) {
    const body = Object.assign(
      {},
      CONFIG.chatDefaults,
      {
        message,
        conversation_id: conversationId,
        openid: currentOpenid,
        email: currentEmail,
      },
    );
    for (const k of Object.keys(body)) {
      const v = body[k];
      if (v === null || v === undefined || v === '') delete body[k];
    }
    return body;
  }

  // ---------------------------------------------------------------------------
  // SSE streaming over POST (manual parser)
  //
  // EventSource is GET-only, so we POST the ChatRequest and read
  // response.body as a stream, splitting on the SSE event delimiter.
  // Tolerant of both LF and CRLF line endings (sse-starlette uses CRLF).
  // ---------------------------------------------------------------------------
  async function streamChat(message, handlers) {
    // handlers: { onDelta, onToolEvent, onRoundBreak, onEnd, onError, onAuthRequired }
    let resp;
    try {
      resp = await fetch(CONFIG.backendUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'text/event-stream',
        },
        body: JSON.stringify(buildRequestBody(message)),
      });
    } catch (err) {
      handlers.onError(err);
      return;
    }

    if (!resp.ok || !resp.body) {
      let detail = '';
      try {
        detail = ' — ' + (await resp.text());
      } catch (_) { /* ignore */ }
      handlers.onError(new Error('Backend returned HTTP ' + resp.status + detail));
      return;
    }

    const reader = resp.body.getReader();
    const decoder = new TextDecoder('utf-8');
    // sse-starlette defaults to CRLF — event blocks end with "\r\n\r\n".
    // Match either "\n\n" or "\r\n\r\n" (and mixed).
    const BLOCK_SEP = /\r?\n\r?\n/;
    let buffer = '';

    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        let m;
        while ((m = BLOCK_SEP.exec(buffer)) !== null) {
          const rawEvent = buffer.slice(0, m.index);
          buffer = buffer.slice(m.index + m[0].length);
          handleEvent(rawEvent, handlers);
        }
      }
      if (buffer.trim()) handleEvent(buffer, handlers);
      handlers.onEnd();
    } catch (err) {
      handlers.onError(err);
    }
  }

  function handleEvent(rawEvent, handlers) {
    let eventName = 'message';
    let dataStr = '';
    // Accept either LF or CRLF inside an event block.
    for (const line of rawEvent.split(/\r?\n/)) {
      if (line.startsWith('event:')) {
        eventName = line.slice(6).trim();
      } else if (line.startsWith('data:')) {
        const chunk = line.slice(5);
        dataStr += chunk.startsWith(' ') ? chunk.slice(1) : chunk;
      }
    }

    let data = {};
    if (dataStr) {
      try { data = JSON.parse(dataStr); } catch (_) { data = { text: dataStr }; }
    }

    switch (eventName) {
      case 'message_start':
        if (data && typeof data.conversation_id === 'string') {
          conversationId = data.conversation_id;
        }
        break;
      case 'content_delta':
        if (typeof data.text === 'string') handlers.onDelta(data.text);
        break;
      case 'round_break':
        handlers.onRoundBreak();
        break;
      case 'tool_call_start':
        handlers.onToolEvent('→ calling tool: ' + (data.tool_name || '(unknown)'));
        break;
      case 'tool_call_input':
        // Noisy per-chunk fragments — skip.
        break;
      case 'tool_result':
        {
          const preview = typeof data.result === 'string'
            ? (data.result.length > 240 ? data.result.slice(0, 240) + '…' : data.result)
            : JSON.stringify(data.result);
          handlers.onToolEvent((data.is_error ? '✗ tool error: ' : '← tool result: ') + preview);
        }
        break;
      case 'message_end':
        handlers.onEnd();
        break;
      case 'auth_required':
        // Widget-facing gate (emitted by /api/v1/lite-chat).
        handlers.onAuthRequired(data || {});
        break;
      case 'error':
        handlers.onError(new Error((data && data.message) || 'Unknown server error'));
        break;
      default:
        // Unknown event — ignore (forward-compatible).
        break;
    }
  }

  // ---------------------------------------------------------------------------
  // Send handler
  // ---------------------------------------------------------------------------
  let inFlight = false;

  async function sendMessage(queryOverride) {
    if (inFlight) return;
    // Accept either a typed input or a programmatic retry (auto-retry
    // after authentication). Never echo the retry into the input box.
    const fromRetry = typeof queryOverride === 'string';
    const text = fromRetry ? queryOverride.trim() : inputEl.value.trim();
    if (!text) return;

    if (!fromRetry) {
      inputEl.value = '';
      inputEl.style.height = '40px';
    }
    addMessage('user', text);
    lastQuery = text;

    const respEl = addMessage('assistant', '');
    respEl.classList.add('tg-streaming');

    inFlight = true;
    sendBtn.disabled = true;

    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      respEl.classList.remove('tg-streaming');
      inFlight = false;
      sendBtn.disabled = false;
      inputEl.focus();
    };

    await streamChat(text, {
      onDelta: (chunk) => appendTo(respEl, chunk),
      onToolEvent: (label) => {
        if (!CONFIG.showToolEvents) return;
        addMessage('tool', label);
      },
      onRoundBreak: () => addRoundBreak(),
      onEnd: () => finish(),
      onError: (err) => {
        const msg = err && err.message ? err.message : String(err);
        appendTo(respEl, (respEl.textContent ? '\n' : '') + '[error] ' + msg);
        finish();
      },
      onAuthRequired: (payload) => {
        respEl.classList.remove('tg-streaming');

        // Host-API strategy takes precedence over MSAL when enabled —
        // the host already knows who the user is, so we verify silently
        // instead of making the user click "Authenticate".
        if (isHostApiStrategyEnabled()) {
          handleHostApiAuth(respEl, finish);
          return;
        }

        respEl.textContent = payload.message || 'Please sign in to continue.';
        renderAuthenticateLink(respEl, payload);
        finish();
      },
    });
  }

  // ---------------------------------------------------------------------------
  // Auth UI — rendered inside an assistant bubble on `auth_required`
  // ---------------------------------------------------------------------------

  async function handleHostApiAuth(respEl, finish) {
    respEl.textContent = 'Verifying session…';
    const result = await retrieveFromHostApi();
    // Release the input BEFORE triggering the retry so sendMessage's
    // `if (inFlight) return` guard doesn't swallow it.
    finish();

    if (result.ok) {
      if (CONFIG.autoRetryAfterAuth && lastQuery) {
        respEl.textContent = 'Signed in ✓ — retrying your question…';
        setTimeout(() => sendMessage(lastQuery), 150);
      } else {
        respEl.textContent = 'Signed in ✓ — ask your question again.';
      }
    } else {
      respEl.textContent = '[error] ' + (result.error || 'Host API authentication failed.');
    }
  }

  function renderAuthenticateLink(respEl, payload) {
    const silentEnabled = CONFIG.allowMsalFallback && CONFIG.msal.clientId && CONFIG.msal.authority;
    const popupEnabled  = CONFIG.allowMsalPopupLogin && CONFIG.msal.clientId && CONFIG.msal.authority;

    const link = document.createElement('span');
    link.className = 'tg-auth-link';
    link.setAttribute('role', 'button');
    link.setAttribute('tabindex', '0');
    link.textContent = payload.label || 'Authenticate';

    const note = document.createElement('span');
    note.className = 'tg-auth-note';
    if (!CONFIG.msal.clientId || !CONFIG.msal.authority) {
      note.textContent = 'MSAL is not configured on this page. Ask the host to set window.TerrierChatWidgetConfig.msal.';
    } else if (silentEnabled && popupEnabled) {
      note.textContent = 'Will try silent sign-in first, then open a login popup if needed.';
    } else if (silentEnabled) {
      note.textContent = 'Signs you in silently using your existing BU Entra session.';
    } else if (popupEnabled) {
      note.textContent = 'Opens a popup to sign you in with your BU account.';
    } else {
      note.textContent = 'Both silent and popup login are disabled in widget config.';
    }

    respEl.appendChild(document.createElement('br'));
    respEl.appendChild(link);
    respEl.appendChild(note);

    const onSuccess = () => {
      link.textContent = 'Signed in ✓';
      if (CONFIG.autoRetryAfterAuth && lastQuery) {
        setTimeout(() => sendMessage(lastQuery), 150);
      } else {
        note.textContent = 'You can now ask your question again.';
      }
    };

    // Track which phase the link is in. After a silent failure, the link
    // switches to "popup-ready" so the next click goes straight to popup
    // with a fresh user gesture (ssoSilent's iframe timeout eats the
    // original click gesture, so popups opened after would be blocked).
    let phase = (silentEnabled && popupEnabled) ? 'silent-first'
              : silentEnabled                   ? 'silent-only'
              : popupEnabled                    ? 'popup-only'
              :                                   'disabled';

    const activate = async () => {
      if (link.classList.contains('tg-auth-busy')) return;

      if (phase === 'disabled') {
        note.textContent = 'Cannot authenticate — MSAL is not configured or both login modes are disabled.';
        return;
      }

      link.classList.add('tg-auth-busy');

      // --- Silent path ---
      if (phase === 'silent-first' || phase === 'silent-only') {
        link.textContent = 'Signing in silently…';
        try {
          await silentEntraLogin();
          onSuccess();
          return;
        } catch (silentErr) {
          if (phase === 'silent-first') {
            phase = 'popup-only';
            link.classList.remove('tg-auth-busy');
            link.textContent = 'Sign in with BU Login';
            note.textContent = 'Silent sign-in failed. Click above to sign in via popup.';
            return;
          }
          link.classList.remove('tg-auth-busy');
          link.textContent = payload.label || 'Authenticate';
          note.textContent = 'Silent sign-in failed: ' + (silentErr && silentErr.message ? silentErr.message : String(silentErr));
          return;
        }
      }

      // --- Popup path (standalone or after silent failure) ---
      if (phase === 'popup-only') {
        link.textContent = 'Opening login…';
        try {
          await popupEntraLogin();
          onSuccess();
        } catch (popupErr) {
          link.classList.remove('tg-auth-busy');
          link.textContent = 'Sign in with BU Login';
          note.textContent = 'Sign-in failed: ' + (popupErr && popupErr.message ? popupErr.message : String(popupErr));
        }
      }
    };

    link.addEventListener('click', activate);
    link.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        activate();
      }
    });
  }

  // Probe the host page for any Entra ID signal we can reuse. Runs once
  // at mount — cheap, synchronous, best-effort.
  extractOpenidFromHost();

  // ---------------------------------------------------------------------------
  // Event wiring
  // ---------------------------------------------------------------------------
  sendBtn.addEventListener('click', () => sendMessage());
  closeBtn.addEventListener('click', () => hide());

  inputEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  inputEl.addEventListener('input', () => {
    inputEl.style.height = '40px';
    inputEl.style.height = Math.min(inputEl.scrollHeight, 120) + 'px';
  });

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------
  function show() {
    root.classList.remove('tg-hidden');
    if (launcherEl) launcherEl.classList.add('tg-hidden');
    setTimeout(() => inputEl.focus(), 0);
  }
  function hide() {
    root.classList.add('tg-hidden');
    if (launcherEl) launcherEl.classList.remove('tg-hidden');
  }
  function toggle() {
    if (root.classList.contains('tg-hidden')) show();
    else hide();
  }
  function unmount() {
    root.remove();
    if (launcherEl) launcherEl.remove();
    const s = document.getElementById(STYLE_ID);
    if (s) s.remove();
    try { delete window.TerrierChatWidget; } catch (_) { window.TerrierChatWidget = undefined; }
  }
  function reset() {
    conversationId = null;
  }

  if (launcherEl) launcherEl.addEventListener('click', toggle);

  window.TerrierChatWidget = {
    show,
    hide,
    toggle,
    unmount,
    reset,
    get mounted() { return true; },
    get conversationId() { return conversationId; },

    // Openid / email controls — useful for hosts that already know the user.
    getOpenid() { return currentOpenid; },
    getOpenidSource() { return openidSource; },
    getEmail() { return currentEmail; },
    // `opts` may be a string (legacy source label) or an options object
    // `{ source?: string, email?: string }` so hosts can push the email
    // alongside the openid without a separate call.
    setOpenid(value, opts) {
      if (!value) {
        currentOpenid = null;
        openidSource = null;
        currentEmail = null;
        return;
      }
      if (!looksLikeOpenid(value)) {
        try { console.warn('[TerrierChatWidget] setOpenid: value does not look like an openid/oid/JWT'); } catch (_) {}
      }
      currentOpenid = value;
      if (opts && typeof opts === 'object') {
        openidSource = opts.source || 'host';
        if (typeof opts.email === 'string') currentEmail = opts.email || null;
      } else {
        openidSource = opts || 'host';
      }
    },
    setEmail(email) { currentEmail = email || null; },
    refreshOpenidFromHost() { extractOpenidFromHost(); return currentOpenid; },
    signInSilent: silentEntraLogin,
    signInPopup: popupEntraLogin,
    async signIn() {
      if (CONFIG.allowMsalFallback && CONFIG.msal.clientId && CONFIG.msal.authority) {
        try { return await silentEntraLogin(); } catch (_) { /* fall through */ }
      }
      return popupEntraLogin();
    },
  };

  try {
    window.dispatchEvent(new CustomEvent('terrier-chat-v2-ready'));
  } catch (_) { /* old browsers */ }
})();
