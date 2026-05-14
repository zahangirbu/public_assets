/**
 * NEXUS Chat Widget — nexus-auth edition
 * --------------------------------------
 * Authentication is handled exclusively by either:
 *   - `allowNexusAuthLogin` — popup OIDC against the nexus-auth BFF
 *     (postMessage delivers an agent-issued JWT to the widget).
 *   - `allowToRetrieveFromApi` — silent host-API retrieval of
 *     openid/email for the anonymous lite-chat path.
 *
 * Both strategies coexist with a single `authToken` page state: when a
 * JWT is captured, /chat goes to /api/v1/chat with `Authorization:
 * Bearer <jwt>`; otherwise it goes to /api/v1/lite-chat with optional
 * `openid` / `email` body fields.
 *
 * Config via window.TerrierChatWidgetConfig BEFORE the script loads.
 */
(function () {
  'use strict';

  // ═══════════════════════════════════════════════════════════════
  // CONFIG
  // ═══════════════════════════════════════════════════════════════
  const uc = (typeof window !== 'undefined' && window.TerrierChatWidgetConfig) || {};
  const C = Object.assign({
    // Single source of truth for the backend host. Every endpoint —
    // chat, lite-chat, feedback, deployments, and nexus-auth /start —
    // is composed as `apiBaseUrl + <constant path>`. Path constants
    // live in the `// PATH CONSTANTS` block right after `STATE`.
    apiBaseUrl: 'http://localhost:8080',
    title: 'TerrierGPT Lite',
    // Optional sub-line under the title in the header.
    subtitle: '',
    placeholder: 'Type a message\u2026',
    startHidden: true,
    // When true, render the floating launcher (chat-bubble) button.
    showLauncher: true,
    launcherLabel: '\uD83D\uDCAC',
    // When true, surface tool-use events as collapsible tool-call cards
    // in the transcript. Set to false for a clean user/assistant-only
    // transcript.
    showToolEvents: true,
    // Greeting message shown before the user sends their first turn.
    greeting: 'Hi! I\u2019m here to help with BU questions. What can I look up for you?',
    // Optional array of canned prompt strings rendered as quick-reply
    // buttons under the greeting. `null` hides the section.
    samplePrompts: null,
    // Optional HTML/string body for the slide-down "info" panel
    // (toggled from the header). `null` hides the info button.
    infoContent: null,

    // After a successful login (or host-API retrieval), automatically
    // re-send the last query.
    autoRetryAfterAuth: true,

    // --- Host-API retrieval strategy ------------------------------------
    // When true, on `auth_required` the widget calls the host REST
    // endpoint (usually a cookie-auth'd "who am I" route) and extracts
    // openid/email via `hostApi` mapping. Takes precedence over MSAL
    // when enabled.
    allowToRetrieveFromApi: false,

    // --- Nexus-auth popup OIDC strategy --------------------------------
    // When true, on `auth_required` the widget renders a "Sign in"
    // link that opens the popup at `${apiBaseUrl}/api/v1/auth/start`.
    // On success the widget captures the returned JWT (postMessage'd
    // from the popup) and switches subsequent /api/v1/lite-chat calls
    // to the Bearer-protected /api/v1/chat endpoint.
    allowNexusAuthLogin: false,

    // --- Theming / runtime extras (v2-only) -----------------------------
    // Primary brand color used for the header, send button, links, and
    // citation chips. Any valid CSS color string.
    accentColor: '#cc0000',
    // When true, attempt to GET runtime config (provider/deployment/etc.)
    // from the backend at mount time so hosts don't have to hard-code it.
    autoFetchConfig: true,
    // Custom avatar URL for assistant bubbles. When null, falls back to
    // the embedded BU shield SVG.
    avatarUrl: null,
    // Play a soft notification beep when a reply arrives while the
    // widget is hidden / minimized.
    notificationSound: true,
  }, uc);

  // Defaults merged into every POST body. Hosts can override any subset
  // (e.g. just `provider`) without wiping the rest. `null`/`undefined`
  // values are stripped at send time so Pydantic uses its own defaults.
  C.chatDefaults = Object.assign({
    provider: 'azure_openai',
    model: null,
    deployment: null,
    system_message_id: null,
  }, (uc && uc.chatDefaults) || {});

  // Host-API retrieval config.
  C.hostApi = Object.assign({
    endpoint: null,
    method: 'GET',
    credentials: 'same-origin',
    headers: null,
    openidPath: 'openid',
    emailPath: 'email',
    // Optional escape hatch: function(responseJson) => { openid, email }.
    mapper: null,
  }, (uc && uc.hostApi) || {});

  // Nexus-auth popup OIDC config (paired with `allowNexusAuthLogin`).
  // host contract: the widget POSTs the user to
  // `${apiBaseUrl}${AUTH_START_PATH}?return_origin=...&nonce=...` in a
  // popup, and the popup posts back `{ type: 'nexus-auth', payload: {
  // token, nonce } }` via window.postMessage. The path is a constant
  // (see AUTH_START_PATH below) — not configurable, since the endpoint
  // is owned by the agent backend.
  C.nexusAuth = Object.assign({
    // window.open features string for the popup.
    popupFeatures: 'width=500,height=700',
    // Optional extra query params appended to the start URL (passthrough).
    extraParams: null,
    // Set true to drop the conversation id when auth state changes — the
    // anon convo lives under user_id="anonymous" on the agent side and
    // continuing it as a signed-in user would cross identities.
    resetConversationOnAuth: true,
  }, (uc && uc.nexusAuth) || {});

  // ═══════════════════════════════════════════════════════════════
  // STATE
  // ═══════════════════════════════════════════════════════════════
  const WID = 'nx-chat', LID = 'nx-chat-launcher', SID = 'nx-chat-style';
  const STORE_KEY = 'nx_chat_convId';
  let convId = null, openid = null, oidSrc = null, email = null, lastQ = null;
  // ── Nexus-auth state (active only when `allowNexusAuthLogin` = true) ──
  // `authToken` is the JWT issued by nexus-auth; presence of a token
  // routes /chat through `/api/v1/chat` with `Authorization: Bearer …`
  // and suppresses sending openid/email in the request body. `authClaims`
  // is the decoded payload (read-only — display only). `pendingAuth`
  // tracks the in-flight popup promise so a superseding sign-in
  // attempt can reject the prior one cleanly.
  let authToken = null, authClaims = null, pendingAuth = null;
  let mode = 'panel', busy = false, widgetVisible = false;
  if (typeof document === 'undefined') return;
  if (document.getElementById(WID)) return;

  // Restore conversation from sessionStorage
  try { convId = sessionStorage.getItem(STORE_KEY) || null; } catch (_) {}

  const ac = C.accentColor;
  function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

  // Script base URL for font loading
  const scriptEl = document.currentScript || document.querySelector('script[src*="chat.js"]');
  const scriptBase = scriptEl ? scriptEl.src.replace(/\/[^/]*$/, '') : '.';

  // ── PATH CONSTANTS ─────────────────────────────────────────
  // Backend host comes from `C.apiBaseUrl`. Every endpoint is
  // `C.apiBaseUrl + <PATH constant>` so there is exactly one place to
  // configure the host. Trim a trailing slash defensively so
  // `apiUrl(p)` never produces a double-slash in the middle.
  const API_BASE = (C.apiBaseUrl || '').replace(/\/+$/, '');
  const LITE_CHAT_PATH = '/api/v1/lite-chat';
  const CHAT_PATH = '/api/v1/chat';
  const DEPLOYMENTS_PATH = '/api/v1/deployments';
  const FEEDBACK_PATH_PREFIX = '/api/v1/conversations/';   // + <conversationId> + '/feedback'
  const AUTH_START_PATH = '/api/v1/auth/start';
  const apiUrl = (path) => API_BASE + path;

  // Origin of `apiBaseUrl`, used to validate `postMessage` events
  // arriving from the nexus-auth popup. Computed once at boot.
  const apiOrigin = (() => {
    try { return new URL(API_BASE).origin; } catch (_) { return null; }
  })();

  // Avatar image source — custom URL or BU shield fallback
  function getAvatarSrc() { return C.avatarUrl || BU_SHIELD; }

  // BU sub-brand plate SVG data URI
  const BU_SHIELD = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 67.2 47"><rect fill="#c00" width="67.2" height="47"/><path fill="#fff" d="M2.8,2.8v41.5h61.7V2.8H2.8ZM63.3,43.1H3.9V3.9h59.4v39.2Z"/><path fill="#fff" d="M25.4,22.4c2.1-.4,3.6-2.1,3.6-4s-.7-2.7-2-3.6c-1.2-.9-3.7-1.3-6.5-1.3-.6,0-3.1,0-4.8.1-.9,0-3.2.2-4.1.2h0v.8h1.2c2.2,0,2.4.9,2.4,2.7v13.4c0,2-.4,2.5-3.1,2.6h-.8v.8h11.9c5.2,0,7-3.6,7-6.2,0-2.5-1.8-4.5-4.7-5.4h0ZM21.4,22.2h-3.5v-7.8c.8,0,2.1-.1,2.8-.1,3.9,0,5.2,2,5.2,4.2,0,2.5-1.5,3.7-4.6,3.7h0ZM27,27.9c0,2.5-1.2,5.1-5.4,5.1-2.7,0-3.6-.8-3.6-3.2v-6.8h3.4c3.7,0,5.6,2.5,5.6,5h0Z"/><path fill="#fff" d="M48.6,13.7v.8h0c2.7.2,3,.6,3.1,3.7v7.6c0,5.5-3.7,6.9-7.3,6.9-4.8,0-7.3-2.6-7.3-6.9v-9c0-1.8.7-2.3,2.5-2.3h.4v-.8h-8.5v.8h.1c2.2,0,2.7.3,2.7,2.3v9.5c0,4.7,3.2,8.3,9.5,8.3,7.1,0,9.7-4.4,9.7-8.4v-8.8c0-2.1.7-2.8,2.8-2.9h0v-.8s-7.8,0-7.8,0Z"/></svg>');

  // ═══════════════════════════════════════════════════════════════
  // MARKDOWN RENDERER (lightweight, no dependencies)
  // ═══════════════════════════════════════════════════════════════
  function md(text) {
    if (!text) return '';
    let s = esc(text);
    // Code blocks ```...```
    s = s.replace(/```(\w*)\n?([\s\S]*?)```/g, '<pre style="background:#f5f3f1;padding:8px 10px;border-radius:6px;overflow-x:auto;font-size:12px;margin:4px 0"><code>$2</code></pre>');
    // Inline code
    s = s.replace(/`([^`]+)`/g, '<code style="background:#f5f3f1;padding:1px 4px;border-radius:3px;font-size:12px">$1</code>');
    // Bold
    s = s.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    // Italic
    s = s.replace(/\*(.+?)\*/g, '<em>$1</em>');
    // Headings (tight margins — surrounding <br> cleaned up below)
    s = s.replace(/^### (.+)$/gm, '<strong style="font-size:14px;display:block;margin:4px 0 2px">$1</strong>');
    s = s.replace(/^## (.+)$/gm, '<strong style="font-size:15px;display:block;margin:4px 0 2px">$1</strong>');
    // Horizontal rules
    s = s.replace(/^---+$/gm, '<hr style="border:none;border-top:1px solid #ddd;margin:6px 0">');
    // Markdown links [text](url)
    s = s.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, '<a href="$2" target="_blank" rel="noopener" style="color:' + ac + ';text-decoration:underline">$1</a>');
    // Unordered lists
    s = s.replace(/^[-*] (.+)$/gm, '<li style="margin-left:16px;list-style:disc">$1</li>');
    // Ordered lists
    s = s.replace(/^\d+\. (.+)$/gm, '<li style="margin-left:16px;list-style:decimal">$1</li>');
    // Pipe tables
    s = s.replace(/((?:^\|.+\|$\n?)+)/gm, function(block) {
      const rows = block.trim().split('\n').filter(r => r.trim());
      if (rows.length < 2) return block;
      // Skip separator row (|---|---|)
      const dataRows = rows.filter(r => !/^\|[\s-:|]+\|$/.test(r));
      if (!dataRows.length) return block;
      let html = '<table style="border-collapse:collapse;width:100%;margin:6px 0;font-size:13px">';
      dataRows.forEach((row, i) => {
        const cells = row.split('|').filter((c, j, a) => j > 0 && j < a.length - 1).map(c => c.trim());
        const tag = i === 0 ? 'th' : 'td';
        const bg = i === 0 ? 'background:#f5f3f1;font-weight:600;' : (i % 2 === 0 ? 'background:#fafafa;' : '');
        html += '<tr>' + cells.map(c => '<' + tag + ' style="padding:6px 10px;border:1px solid #e0e0e0;text-align:left;' + bg + '">' + c + '</' + tag + '>').join('') + '</tr>';
      });
      html += '</table>';
      return html;
    });
    // Auto-linkify bare URLs (that aren't already in an href)
    s = s.replace(/(?<!href="|">)(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener" style="color:' + ac + ';text-decoration:underline;word-break:break-all">$1</a>');
    // Line breaks — collapse 3+ consecutive newlines into 2
    s = s.replace(/\n{3,}/g, '\n\n');
    s = s.replace(/\n/g, '<br>');
    // Strip <br> around block elements (headings, tables, pre, hr, li)
    s = s.replace(/(<br>)+(<(?:strong[^>]*style|table|pre|hr|li))/gi, '$2');
    s = s.replace(/(<\/(?:strong|table|pre|li)>)(<br>)+/gi, '$1');
    s = s.replace(/(<hr[^>]*>)(<br>)+/gi, '$1');
    s = s.replace(/(<br>)+(<hr)/gi, '$2');
    // Collapse remaining triple+ <br> to double
    s = s.replace(/(<br>){3,}/g, '<br><br>');
    return s;
  }

  // Extract URLs from text for source citations
  function extractUrls(text) {
    if (!text) return [];
    const m = text.match(/https?:\/\/[^\s)]+/g);
    return m ? [...new Set(m)] : [];
  }

  // ═══════════════════════════════════════════════════════════════
  // NOTIFICATION SOUND
  // ═══════════════════════════════════════════════════════════════
  let audioCtx = null;
  function playNotif() {
    if (!C.notificationSound || widgetVisible) return;
    try {
      if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.connect(gain); gain.connect(audioCtx.destination);
      osc.frequency.value = 880; gain.gain.value = 0.08;
      osc.start(); gain.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + 0.15);
      osc.stop(audioCtx.currentTime + 0.15);
    } catch (_) {}
  }

  // ═══════════════════════════════════════════════════════════════
  // RELATIVE TIME
  // ═══════════════════════════════════════════════════════════════
  function relTime(ts) {
    const diff = Math.floor((Date.now() - ts) / 1000);
    if (diff < 10) return 'just now';
    if (diff < 60) return diff + 's ago';
    if (diff < 3600) return Math.floor(diff / 60) + 'm ago';
    if (diff < 86400) return Math.floor(diff / 3600) + 'h ago';
    return new Date(ts).toLocaleDateString();
  }

  // ═══════════════════════════════════════════════════════════════
  // CSS
  // ═══════════════════════════════════════════════════════════════
  const css = `
@font-face{font-family:'Whitney';src:url('${scriptBase}/fonts/Whitney-Book.otf') format('opentype');font-weight:400;font-style:normal;font-display:swap}
@font-face{font-family:'Whitney';src:url('${scriptBase}/fonts/Whitney-Semibold.otf') format('opentype');font-weight:600;font-style:normal;font-display:swap}
@keyframes nx-in{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
@keyframes nx-blink{to{visibility:hidden}}
@keyframes nx-dots{0%,80%,100%{transform:scale(0)}40%{transform:scale(1)}}
@keyframes nx-msg-l{from{opacity:0;transform:translateX(-12px)}to{opacity:1;transform:translateX(0)}}
@keyframes nx-msg-r{from{opacity:0;transform:translateX(12px)}to{opacity:1;transform:translateX(0)}}
@keyframes nx-pulse{0%{box-shadow:0 0 0 0 rgba(204,0,0,.5)}70%{box-shadow:0 0 0 14px rgba(204,0,0,0)}100%{box-shadow:0 0 0 0 rgba(204,0,0,0)}}
@keyframes nx-stagger{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}

#${WID}{position:fixed;z-index:2147483647;display:flex;flex-direction:column;background:#fff;color:#2D2926;
  font-family:'Whitney',-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;font-size:14px;line-height:1.5;
  border:1px solid #e0e0e0;box-shadow:0 8px 32px rgba(0,0,0,.18);overflow:hidden;transition:all .3s cubic-bezier(.4,0,.2,1);animation:nx-in .3s cubic-bezier(.4,0,.2,1)}
#${WID}.nx-hidden{display:none!important}

/* Panel */
#${WID}.nx-panel{bottom:84px;right:20px;width:380px;height:560px;max-height:calc(100vh - 110px);border-radius:16px;min-width:300px;min-height:360px}
/* Resize handles */
#${WID} .nx-rh{position:absolute;z-index:10}
#${WID} .nx-rh-tl{top:0;left:0;width:12px;height:12px;cursor:nw-resize}
#${WID} .nx-rh-tr{top:0;right:0;width:12px;height:12px;cursor:ne-resize}
#${WID} .nx-rh-bl{bottom:0;left:0;width:12px;height:12px;cursor:sw-resize}
#${WID} .nx-rh-br{bottom:0;right:0;width:12px;height:12px;cursor:se-resize}
#${WID} .nx-rh-t{top:0;left:12px;right:12px;height:4px;cursor:n-resize}
#${WID} .nx-rh-b{bottom:0;left:12px;right:12px;height:4px;cursor:s-resize}
#${WID} .nx-rh-l{top:12px;bottom:12px;left:0;width:4px;cursor:w-resize}
#${WID} .nx-rh-r{top:12px;bottom:12px;right:0;width:4px;cursor:e-resize}
#${WID}:not(.nx-panel):not(.nx-sidebar) .nx-rh{display:none}
#${WID}.nx-sidebar .nx-rh:not(.nx-rh-l){display:none}
#${WID}.nx-sidebar .nx-rh-l{top:0;bottom:0;width:6px;cursor:ew-resize}

/* Sidebar */
#${WID}.nx-sidebar{top:0;right:0;bottom:0;width:420px;min-width:300px;max-width:80vw;border-radius:0;border-right:none;border-top:none;border-bottom:none}
/* Modal */
#${WID}.nx-modal{top:50%;left:50%;transform:translate(-50%,-50%);width:700px;height:80vh;max-width:calc(100vw - 40px);max-height:calc(100vh - 40px);border-radius:16px}
.nx-bk{position:fixed;inset:0;z-index:2147483646;background:rgba(0,0,0,.4);backdrop-filter:blur(2px)}
.nx-bk.nx-hidden{display:none}

/* Header */
#${WID} .nx-hdr{display:flex;align-items:center;gap:8px;padding:10px 14px;background:${ac};color:#fff;user-select:none;flex-shrink:0;cursor:grab}
#${WID} .nx-hdr.nx-dragging{cursor:grabbing}
#${WID}.nx-dragging{transition:none!important}
#${WID} .nx-hdr-ic{width:22px;height:22px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
#${WID} .nx-hdr-ic svg{width:18px;height:18px;fill:none;stroke:#fff;stroke-width:2}
#${WID} .nx-hdr-t{flex:1;font-weight:600;font-size:14px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#${WID} .nx-hb{background:none;border:none;color:#fff;cursor:pointer;padding:4px;border-radius:4px;display:flex;align-items:center;justify-content:center;opacity:.85;transition:opacity .15s,background .15s}
#${WID} .nx-hb svg{width:16px;height:16px}
#${WID} .nx-hb:hover{opacity:1;background:rgba(255,255,255,.15)}
#${WID} .nx-hb.nx-active{opacity:1;background:rgba(255,255,255,.25)}
#${WID} .nx-ha{display:flex;gap:2px;align-items:center}

/* Toolbar */
#${WID} .nx-tb{display:flex;align-items:center;justify-content:flex-end;padding:4px 10px;border-bottom:1px solid #eee;flex-shrink:0;background:#fafafa;gap:2px}
#${WID} .nx-tbb{background:none;border:none;cursor:pointer;padding:4px 6px;border-radius:4px;color:#767676;display:flex;align-items:center;transition:color .15s,background .15s}
#${WID} .nx-tbb:hover{background:#eee;color:#333}
#${WID} .nx-tbb svg{width:16px;height:16px}

/* Info panel */
#${WID} .nx-info{display:none;padding:12px 16px;background:#f8f8f8;border-bottom:1px solid #eee;font-size:13px;color:#595550;max-height:200px;overflow-y:auto;flex-shrink:0}
#${WID} .nx-info.nx-open{display:block}
#${WID} .nx-info p{margin:0 0 6px}#${WID} .nx-info p:last-child{margin:0}

/* Conversation */
#${WID} .nx-cv{flex:1;overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:8px;background:#fff}

/* Messages — row layout with avatar beside bubble */
#${WID} .nx-mw{display:flex;gap:8px;max-width:92%;animation:nx-msg-l .25s ease}
#${WID} .nx-mw-a{align-self:flex-start;align-items:flex-start}
#${WID} .nx-mw-u{align-self:flex-end;flex-direction:row-reverse;animation:nx-msg-r .25s ease}
#${WID} .nx-col{display:flex;flex-direction:column;flex:1;min-width:0}
#${WID} .nx-m{border-radius:16px;word-wrap:break-word;position:relative}
#${WID} .nx-m-u{background:${ac};color:#fff;border-bottom-right-radius:4px;padding:8px 12px}
#${WID} .nx-m-a{background:#f0f0f0;color:#2D2926;border-bottom-left-radius:4px;padding:8px 12px}
#${WID} .nx-m-t{animation:nx-msg-l .2s ease}
#${WID} .nx-m-s{animation:nx-msg-l .2s ease}
#${WID} .nx-m-a .nx-mc{white-space:normal}
#${WID} .nx-m-a .nx-mc a{color:${ac}}
#${WID} .nx-m-a .nx-mc pre{white-space:pre-wrap}
#${WID} .nx-m-a .nx-mc li{margin-left:20px}
#${WID} .nx-m-a .nx-mc table{border-collapse:collapse}
#${WID} .nx-av{width:28px;height:28px;flex-shrink:0;border-radius:50%;overflow:hidden;margin-top:2px}
#${WID} .nx-av img{width:100%;height:100%;object-fit:cover}
#${WID} .nx-av-u{width:28px;height:28px;flex-shrink:0;border-radius:50%;overflow:hidden;margin-top:2px;background:#e0dcd8;display:flex;align-items:center;justify-content:center}
#${WID} .nx-av-u svg{width:16px;height:16px;fill:#8a8480}
#${WID} .nx-m-s{background:transparent;color:#767676;align-self:center;font-size:12px;font-style:italic;text-align:center;max-width:100%;padding:4px 10px}
#${WID} .nx-rb{align-self:stretch;border-top:1px dashed #e0e0e0;margin:4px 0}

/* Tool call collapsible */
#${WID} .nx-tc{align-self:flex-start;max-width:95%;animation:nx-msg-l .2s ease}
#${WID} .nx-tc-hdr{display:flex;align-items:center;gap:6px;padding:6px 10px;background:#f6f6f6;border-radius:6px;cursor:pointer;font-size:12px;color:#595550;border:1px solid #eee;transition:background .15s;user-select:none}
#${WID} .nx-tc-hdr:hover{background:#eee}
#${WID} .nx-tc-hdr svg{width:10px;height:10px;transition:transform .2s}
#${WID} .nx-tc.nx-open .nx-tc-hdr svg{transform:rotate(90deg)}
#${WID} .nx-tc-dot{width:6px;height:6px;border-radius:50%;background:${ac};animation:nx-blink 1s steps(2,start) infinite}
#${WID} .nx-tc.nx-done .nx-tc-dot{animation:none;background:#4caf50}
#${WID} .nx-tc-body{display:none;padding:6px 10px;font-size:11px;color:#595550;background:#fafafa;border:1px solid #eee;border-top:none;border-radius:0 0 6px 6px;max-height:200px;overflow-y:auto}
#${WID} .nx-tc.nx-open .nx-tc-body{display:block}
#${WID} .nx-tc-line{padding:2px 0;border-bottom:1px solid #f0f0f0}
#${WID} .nx-tc-line:last-child{border-bottom:none}

/* Typing indicator */
#${WID} .nx-typing{display:flex;gap:4px;align-items:center;padding:12px 14px}
#${WID} .nx-typing-dot{width:8px;height:8px;border-radius:50%;background:#767676;animation:nx-dots 1.4s infinite ease-in-out both}
#${WID} .nx-typing-dot:nth-child(1){animation-delay:-.32s}
#${WID} .nx-typing-dot:nth-child(2){animation-delay:-.16s}

/* Timestamp */
#${WID} .nx-ts{font-size:10px;color:#767676;margin-top:2px;visibility:hidden;user-select:none}
#${WID} .nx-mw:hover .nx-ts{visibility:visible}

/* Message actions (copy, feedback) */
#${WID} .nx-ma{display:flex;gap:4px;margin-top:2px;visibility:hidden}
#${WID} .nx-mw:hover .nx-ma{visibility:visible}
#${WID} .nx-ma button{background:none;border:none;cursor:pointer;padding:2px 4px;border-radius:4px;color:#767676;font-size:12px;display:flex;align-items:center;gap:3px;transition:color .15s,background .15s}
#${WID} .nx-ma button:hover{background:#e8e8e8;color:#333}
#${WID} .nx-ma button.nx-voted{color:${ac}}
#${WID} .nx-ma svg{width:12px;height:12px}

/* Source citations */
#${WID} .nx-sources{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}
#${WID} .nx-src{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:6px;background:#f5f3f1;color:${ac};font-size:11px;text-decoration:none;border:1px solid #e5e0db;transition:background .15s}
#${WID} .nx-src:hover{background:#ece8e4}
#${WID} .nx-src svg{width:10px;height:10px;flex-shrink:0}

/* Greeting */
#${WID} .nx-gr{font-size:14px;color:#2D2926;line-height:1.5}
#${WID} .nx-sp{display:flex;flex-direction:column;gap:6px;margin-top:8px}
#${WID} .nx-spb{background:#fff;border:1px solid #ddd;border-radius:10px;padding:8px 12px;cursor:pointer;text-align:left;font-size:13px;color:#2D2926;transition:border-color .15s,background .15s;opacity:0;animation:nx-stagger .3s ease forwards}
#${WID} .nx-spb:hover{border-color:${ac};background:#fef8f8}

/* Auth */
#${WID} .nx-al{display:inline-block;margin-top:6px;padding:6px 12px;background:${ac};color:#fff;border-radius:8px;cursor:pointer;font-weight:600;font-size:13px;user-select:none}
#${WID} .nx-al:hover{filter:brightness(.9)}#${WID} .nx-al.nx-bsy{opacity:.6;pointer-events:none}
#${WID} .nx-an{display:block;margin-top:4px;font-size:12px;color:#595550}

/* Input */
#${WID} .nx-ia{display:flex;align-items:flex-end;gap:8px;padding:10px 12px;border-top:1px solid #eee;background:#fff;flex-shrink:0}
#${WID} .nx-in-wrap{flex:1;display:flex;align-items:flex-end;border:1px solid #ddd;border-radius:12px;background:#fff;transition:border-color .15s,box-shadow .15s}
#${WID} .nx-in-wrap:focus-within{border-color:${ac};box-shadow:0 0 0 2px ${ac}22}
#${WID} .nx-mic{background:none;border:none;cursor:pointer;padding:8px 4px 8px 10px;color:#767676;display:flex;align-items:center;flex-shrink:0;transition:color .15s}
#${WID} .nx-mic:hover{color:#2D2926}
#${WID} .nx-mic:disabled{cursor:not-allowed;opacity:.4}
#${WID} .nx-mic:disabled:hover{color:#767676}
#${WID} .nx-mic.nx-mic-on{color:${ac};animation:nx-pulse 1.5s infinite}
#${WID} .nx-mic.nx-mic-loading{cursor:progress;color:#767676;animation:nx-pulse 2s infinite}
#${WID} .nx-mic.nx-mic-busy{cursor:wait;color:${ac};opacity:.7}
#${WID} .nx-mic svg{width:16px;height:16px}
#${WID} .nx-mic.nx-mic-hide{display:none}
#${WID} .nx-in{flex:1;resize:none;border:none;border-radius:0;padding:10px 14px 10px 4px;font-family:inherit;font-size:14px;height:44px;max-height:120px;outline:none;box-sizing:border-box;background:transparent}
#${WID} .nx-sn{width:36px;height:36px;border-radius:50%;border:none;background:${ac};color:#fff;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:opacity .15s}
#${WID} .nx-sn:disabled{opacity:.4;cursor:not-allowed}
#${WID} .nx-sn svg{width:16px;height:16px;fill:currentColor}

/* Input hint */
#${WID} .nx-hint{padding:0 14px 4px;font-size:10px;color:#767676;flex-shrink:0}

/* Footer */
#${WID} .nx-ft{padding:4px 12px;text-align:center;font-size:11px;color:#767676;border-top:1px solid #f0f0f0;background:#fafafa;flex-shrink:0}

/* Launcher */
#${LID}{position:fixed;bottom:20px;right:20px;width:56px;height:56px;border-radius:50%;border:none;background:${ac};color:#fff;font-size:24px;line-height:1;cursor:pointer;box-shadow:0 6px 18px rgba(0,0,0,.25);z-index:2147483646;transition:transform .2s cubic-bezier(.4,0,.2,1);display:flex;align-items:center;justify-content:center}
#${LID}:hover{transform:scale(1.08)}#${LID}:active{transform:scale(.95)}#${LID}.nx-hidden{display:none}
#${LID}.nx-pulse{animation:nx-pulse 2s infinite}

/* Focus-visible — keyboard-only focus rings (WCAG 2.4.7) */
#${WID} button:focus-visible,#${WID} textarea:focus-visible,#${WID} a:focus-visible,#${WID} [role="button"]:focus-visible{outline:2px solid ${ac};outline-offset:2px;border-radius:4px}
#${WID} .nx-hb:focus-visible{outline-color:#fff;outline-offset:1px}
#${LID}:focus-visible{outline:2px solid ${ac};outline-offset:3px}

/* Inline feedback comment field */
#${WID} .nx-fb-comment{display:flex;gap:4px;margin-top:4px;animation:nx-msg-l .2s ease}
#${WID} .nx-fb-input{flex:1;border:1px solid #ddd;border-radius:6px;padding:4px 8px;font-family:inherit;font-size:12px;outline:none}
#${WID} .nx-fb-input:focus{border-color:${ac}}
#${WID} .nx-fb-send{background:none;border:none;color:${ac};cursor:pointer;font-size:12px;font-weight:600;padding:4px 6px}
#${WID} .nx-fb-send:hover{text-decoration:underline}
`;

  const sEl = document.createElement('style'); sEl.id = SID; sEl.textContent = css; document.head.appendChild(sEl);

  // ═══════════════════════════════════════════════════════════════
  // SVG ICONS
  // ═══════════════════════════════════════════════════════════════
  const ic = {
    chat: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>',
    expand: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/></svg>',
    sidebar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="9" y1="3" x2="9" y2="21"/></svg>',
    chevDown: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>',
    send: '<svg viewBox="0 0 24 24"><path d="M2.01 21L23 12 2.01 3 2 10l15 2-15 2z"/></svg>',
    info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><circle cx="12" cy="8" r=".5" fill="currentColor" stroke="none"/></svg>',
    plus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>',
    copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>',
    thumbUp: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 9V5a3 3 0 0 0-3-3l-4 9v11h11.28a2 2 0 0 0 2-1.7l1.38-9a2 2 0 0 0-2-2.3H14zM7 22H4a2 2 0 0 1-2-2v-7a2 2 0 0 1 2-2h3"/></svg>',
    thumbDown: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 15v4a3 3 0 0 0 3 3l4-9V2H5.72a2 2 0 0 0-2 1.7l-1.38 9a2 2 0 0 0 2 2.3H10zM17 2h2.67A2.31 2.31 0 0 1 22 4v7a2.31 2.31 0 0 1-2.33 2H17"/></svg>',
    link: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>',
    person: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>',
    mic: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="23"/><line x1="8" y1="23" x2="16" y2="23"/></svg>',
  };

  // ═══════════════════════════════════════════════════════════════
  // DOM
  // ═══════════════════════════════════════════════════════════════
  const root = document.createElement('div');
  root.id = WID;
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-label', C.title + ' chat widget');
  root.className = 'nx-panel' + (C.startHidden ? ' nx-hidden' : '');

  const bk = document.createElement('div');
  bk.className = 'nx-bk nx-hidden';
  bk.addEventListener('click', () => setMode('panel'));

  root.innerHTML = `
    <div class="nx-rh nx-rh-tl"></div><div class="nx-rh nx-rh-tr"></div>
    <div class="nx-rh nx-rh-bl"></div><div class="nx-rh nx-rh-br"></div>
    <div class="nx-rh nx-rh-t"></div><div class="nx-rh nx-rh-b"></div>
    <div class="nx-rh nx-rh-l"></div><div class="nx-rh nx-rh-r"></div>
    <div class="nx-hdr">
      <div class="nx-hdr-ic">${ic.chat}</div>
      <span class="nx-hdr-t">${esc(C.title)}</span>
      <div class="nx-ha">
        <button type="button" class="nx-hb" data-a="expand" title="Expand" aria-label="Toggle full screen">${ic.expand}</button>
        <button type="button" class="nx-hb" data-a="sidebar" title="Sidebar" aria-label="Toggle sidebar">${ic.sidebar}</button>
        <button type="button" class="nx-hb" data-a="close" title="Minimize" aria-label="Minimize chat">${ic.chevDown}</button>
      </div>
    </div>
    <div class="nx-tb">
      <button type="button" class="nx-tbb" data-a="new" title="New conversation" aria-label="Start new conversation">${ic.plus}</button>
      <button type="button" class="nx-tbb" data-a="info" title="Info" aria-label="Show info">${ic.info}</button>
    </div>
    <div class="nx-info" role="complementary" aria-label="Chat information"></div>
    <div class="nx-cv" role="log" aria-live="polite" aria-label="Chat messages"></div>
    <div class="nx-ia">
      <div class="nx-in-wrap">
        <!-- Voice dictation button. Wired up by the VOICE INPUT block
             below (transformers.js + Whisper). Hidden automatically on
             browsers without MediaRecorder/getUserMedia. -->
        <button type="button" class="nx-mic" title="Voice input" aria-label="Voice input">${ic.mic}</button>
        <textarea class="nx-in" id="nx-chat-input" name="nx-chat-input" rows="1" placeholder="${esc(C.placeholder)}" aria-label="Chat message input" autocomplete="off"></textarea>
      </div>
      <button type="button" class="nx-sn" title="Send" aria-label="Send message">${ic.send}</button>
    </div>
    <div class="nx-hint">Enter to send \u00b7 Shift+Enter for new line</div>
    <div class="nx-ft">Always verify AI-generated responses with the accompanying sources.</div>
  `;

  let launcher = null;
  if (C.showLauncher) {
    launcher = document.createElement('button');
    launcher.id = LID; launcher.type = 'button';
    launcher.setAttribute('aria-label', 'Open chat assistant');
    launcher.textContent = C.launcherLabel;
    if (!C.startHidden) launcher.classList.add('nx-hidden');
    else launcher.classList.add('nx-pulse');
  }

  const doMount = () => { document.body.appendChild(bk); document.body.appendChild(root); if (launcher) document.body.appendChild(launcher); };
  if (document.body) doMount(); else document.addEventListener('DOMContentLoaded', doMount, { once: true });

  const cv = root.querySelector('.nx-cv');
  const inp = root.querySelector('.nx-in');
  const snb = root.querySelector('.nx-sn');
  const inf = root.querySelector('.nx-info');

  // Info content
  inf.innerHTML = C.infoContent || ('<p><strong>' + esc(C.title) + '</strong></p><p>AI-powered assistant. Verify important information with official sources.</p>' + (C.subtitle ? '<p>' + esc(C.subtitle) + '</p>' : ''));

  // ═══════════════════════════════════════════════════════════════
  // GREETING
  // ═══════════════════════════════════════════════════════════════
  function greet() {
    if (!C.greeting && !(C.samplePrompts && C.samplePrompts.length)) return;
    const w = document.createElement('div'); w.style.cssText = 'display:flex;gap:8px;align-items:flex-start;';
    const av = document.createElement('div'); av.className = 'nx-av'; av.innerHTML = '<img src="' + getAvatarSrc() + '" alt="BU">'; w.appendChild(av);
    const inner = document.createElement('div'); inner.style.cssText = 'flex:1;min-width:0;display:flex;flex-direction:column;gap:8px;';
    if (C.greeting) { const g = document.createElement('div'); g.className = 'nx-gr'; g.textContent = C.greeting; inner.appendChild(g); }
    if (C.samplePrompts && C.samplePrompts.length) {
      const s = document.createElement('div'); s.className = 'nx-sp';
      C.samplePrompts.forEach((p, i) => { const b = document.createElement('button'); b.type = 'button'; b.className = 'nx-spb'; b.style.animationDelay = (i * 80) + 'ms'; b.textContent = p; b.addEventListener('click', (e) => { e.stopPropagation(); sendMsg(p); }); s.appendChild(b); });
      inner.appendChild(s);
    }
    w.appendChild(inner);
    cv.appendChild(w);
  }
  greet();

  // Auto-fetch deployment config from the agent backend
  if (C.autoFetchConfig && C.chatDefaults.deployment) {
    fetch(apiUrl(DEPLOYMENTS_PATH + '?deploy_type=web-widget'))
      .then(r => r.ok ? r.json() : null)
      .then(data => {
        if (!data || !data.deployments) return;
        const dep = data.deployments.find(d => d.slug === C.chatDefaults.deployment) || data.deployments[0];
        if (!dep) return;
        const name = dep.display_name || (dep.agent && dep.agent.display_name);
        if (name) { C.title = name; root.querySelector('.nx-hdr-t').textContent = name; }
        if (dep.info_content) { inf.innerHTML = dep.info_content; }
        if (dep.avatar_url) { C.avatarUrl = dep.avatar_url; }
        // Apply greeting/prompts and re-render (avatar may have changed)
        let needsRerender = false;
        if (dep.greeting) { C.greeting = dep.greeting; needsRerender = true; }
        if (dep.sample_prompts && dep.sample_prompts.length) { C.samplePrompts = dep.sample_prompts; needsRerender = true; }
        if (dep.avatar_url || needsRerender) { cv.innerHTML = ''; greet(); }
      }).catch(() => {}); // Fail silently — static config works as fallback
  }

  // ═══════════════════════════════════════════════════════════════
  // MODE SWITCHING
  // ═══════════════════════════════════════════════════════════════
  const expandBtn = root.querySelector('[data-a="expand"]');
  const sidebarBtn = root.querySelector('[data-a="sidebar"]');

  function setMode(m) {
    root.classList.remove('nx-panel', 'nx-sidebar', 'nx-modal');
    root.style.left = ''; root.style.top = ''; root.style.right = ''; root.style.bottom = '';
    root.style.width = ''; root.style.height = ''; root.style.maxHeight = '';
    bk.classList.add('nx-hidden');
    expandBtn.classList.remove('nx-active');
    sidebarBtn.classList.remove('nx-active');
    if (m === 'sidebar') { root.classList.add('nx-sidebar'); sidebarBtn.classList.add('nx-active'); }
    else if (m === 'modal') { root.classList.add('nx-modal'); bk.classList.remove('nx-hidden'); expandBtn.classList.add('nx-active'); trapFocus(true); }
    else { root.classList.add('nx-panel'); }
    if (m !== 'modal') trapFocus(false);
    mode = m;
  }

  expandBtn.addEventListener('click', () => { mode === 'modal' ? setMode('panel') : setMode('modal'); });
  sidebarBtn.addEventListener('click', () => { mode === 'sidebar' ? setMode('panel') : setMode('sidebar'); });
  root.querySelector('[data-a="close"]').addEventListener('click', () => hide());
  root.querySelector('[data-a="info"]').addEventListener('click', () => inf.classList.toggle('nx-open'));
  root.querySelector('[data-a="new"]').addEventListener('click', () => { convId = null; try { sessionStorage.removeItem(STORE_KEY); } catch (_) {} cv.innerHTML = ''; greet(); });

  // ═══════════════════════════════════════════════════════════════
  // DRAG TO MOVE
  // ═══════════════════════════════════════════════════════════════
  const hdr = root.querySelector('.nx-hdr');
  let dragState = null;
  hdr.addEventListener('mousedown', (e) => {
    if (e.target.closest('.nx-hb') || e.target.closest('.nx-ha') || mode !== 'panel') return;
    e.preventDefault();
    const rect = root.getBoundingClientRect();
    dragState = { sx: e.clientX, sy: e.clientY, ol: rect.left, ot: rect.top };
    hdr.classList.add('nx-dragging'); root.classList.add('nx-dragging');
  });
  document.addEventListener('mousemove', (e) => {
    if (!dragState) return;
    root.style.left = Math.max(0, Math.min(window.innerWidth - 100, dragState.ol + e.clientX - dragState.sx)) + 'px';
    root.style.top = Math.max(0, Math.min(window.innerHeight - 60, dragState.ot + e.clientY - dragState.sy)) + 'px';
    root.style.right = 'auto'; root.style.bottom = 'auto';
  });
  document.addEventListener('mouseup', () => { if (dragState) { dragState = null; hdr.classList.remove('nx-dragging'); root.classList.remove('nx-dragging'); } });

  // ═══════════════════════════════════════════════════════════════
  // RESIZE
  // ═══════════════════════════════════════════════════════════════
  let resizeState = null;
  const MIN_W = 300, MIN_H = 360;
  root.querySelectorAll('.nx-rh').forEach(handle => {
    handle.addEventListener('mousedown', (e) => {
      if (mode !== 'panel' && mode !== 'sidebar') return;
      e.preventDefault(); e.stopPropagation();
      const rect = root.getBoundingClientRect(); const cls = handle.className;
      resizeState = { sx: e.clientX, sy: e.clientY, ow: rect.width, oh: rect.height, ol: rect.left, ot: rect.top,
        n: cls.includes('-tl') || cls.includes('-tr') || cls.includes('-t'),
        s: cls.includes('-bl') || cls.includes('-br') || cls.includes('-b'),
        w: cls.includes('-tl') || cls.includes('-bl') || cls.includes('-l'),
        e: cls.includes('-tr') || cls.includes('-br') || cls.includes('-r'),
        sb: mode === 'sidebar' };
      root.classList.add('nx-dragging');
    });
  });
  document.addEventListener('mousemove', (e) => {
    if (!resizeState) return;
    const r = resizeState, dx = e.clientX - r.sx;
    if (r.sb) { root.style.width = Math.max(MIN_W, Math.min(window.innerWidth * 0.8, r.ow - dx)) + 'px'; return; }
    const dy = e.clientY - r.sy;
    let nw = r.ow, nh = r.oh, nl = r.ol, nt = r.ot;
    if (r.e) nw = Math.max(MIN_W, r.ow + dx);
    if (r.s) nh = Math.max(MIN_H, r.oh + dy);
    if (r.w) { const p = r.ow - dx; if (p >= MIN_W) { nw = p; nl = r.ol + dx; } }
    if (r.n) { const p = r.oh - dy; if (p >= MIN_H) { nh = p; nt = r.ot + dy; } }
    root.style.width = nw + 'px'; root.style.height = nh + 'px';
    root.style.left = nl + 'px'; root.style.top = nt + 'px';
    root.style.right = 'auto'; root.style.bottom = 'auto'; root.style.maxHeight = 'none';
  });
  document.addEventListener('mouseup', () => { if (resizeState) { resizeState = null; root.classList.remove('nx-dragging'); } });

  // ═══════════════════════════════════════════════════════════════
  // FOCUS TRAP (modal mode accessibility)
  // ═══════════════════════════════════════════════════════════════
  let trapActive = false;
  function trapFocus(on) {
    trapActive = on;
    if (on) setTimeout(() => inp.focus(), 50);
  }
  root.addEventListener('keydown', (e) => {
    if (!trapActive || e.key !== 'Tab') return;
    const focusable = root.querySelectorAll('button:not([disabled]),textarea,[tabindex]:not([tabindex="-1"])');
    if (!focusable.length) return;
    const first = focusable[0], last = focusable[focusable.length - 1];
    if (e.shiftKey) { if (document.activeElement === first) { e.preventDefault(); last.focus(); } }
    else { if (document.activeElement === last) { e.preventDefault(); first.focus(); } }
  });

  // ═══════════════════════════════════════════════════════════════
  // MESSAGES
  // ═══════════════════════════════════════════════════════════════
  let msgCounter = 0;

  function addMsg(role, text, opts) {
    opts = opts || {};
    const id = 'nx-m-' + (++msgCounter);

    // Wrapper: avatar + bubble in a row
    const wrap = document.createElement('div');
    wrap.className = 'nx-mw nx-mw-' + role[0];
    wrap.id = id;
    wrap.setAttribute('data-ts', Date.now());

    const el = document.createElement('div');
    el.className = 'nx-m nx-m-' + role[0];

    if (role === 'assistant') {
      // BU shield avatar
      const av = document.createElement('div'); av.className = 'nx-av'; av.innerHTML = '<img src="' + getAvatarSrc() + '" alt="BU">';
      wrap.appendChild(av);

      // Column: bubble + meta below
      const col = document.createElement('div'); col.className = 'nx-col';
      const mc = document.createElement('div'); mc.className = 'nx-mc'; mc.innerHTML = md(text || ''); el.appendChild(mc);
      col.appendChild(el);
      const ts = document.createElement('div'); ts.className = 'nx-ts'; ts.textContent = 'just now'; col.appendChild(ts);

      // Actions bar — outside bubble, inside column
      const ma = document.createElement('div'); ma.className = 'nx-ma';
      const cpBtn = document.createElement('button'); cpBtn.type = 'button'; cpBtn.title = 'Copy'; cpBtn.setAttribute('aria-label', 'Copy message');
      cpBtn.innerHTML = ic.copy + ' Copy';
      cpBtn.addEventListener('click', () => {
        const txt = mc.textContent || mc.innerText;
        navigator.clipboard.writeText(txt).then(() => { cpBtn.innerHTML = ic.copy + ' Copied!'; setTimeout(() => { cpBtn.innerHTML = ic.copy + ' Copy'; }, 1500); }).catch(() => {});
      });
      ma.appendChild(cpBtn);
      const upBtn = document.createElement('button'); upBtn.type = 'button'; upBtn.title = 'Helpful'; upBtn.setAttribute('aria-label', 'Mark as helpful');
      upBtn.innerHTML = ic.thumbUp;
      function showCommentField(rating) {
        // Remove existing comment field if any
        const parent = el.closest('.nx-col') || el;
        const existing = parent.querySelector('.nx-fb-comment');
        if (existing) existing.remove();
        const row = document.createElement('div'); row.className = 'nx-fb-comment';
        const input = document.createElement('input'); input.type = 'text'; input.className = 'nx-fb-input';
        input.placeholder = rating === 'positive' ? 'Any additional feedback?' : 'What could be improved?';
        input.setAttribute('aria-label', 'Feedback comment');
        const sendLink = document.createElement('button'); sendLink.type = 'button'; sendLink.className = 'nx-fb-send'; sendLink.textContent = 'Send';
        const submit = () => { submitFeedback(id, rating, input.value.trim() || undefined); row.innerHTML = '<span style="font-size:11px;color:#767676">Thanks for your feedback!</span>'; };
        sendLink.addEventListener('click', submit);
        input.addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); });
        row.appendChild(input); row.appendChild(sendLink);
        parent.appendChild(row);
        input.focus();
      }
      upBtn.addEventListener('click', () => {
        upBtn.classList.add('nx-voted'); downBtn.classList.remove('nx-voted');
        submitFeedback(id, 'positive');
        showCommentField('positive');
      });
      ma.appendChild(upBtn);
      const downBtn = document.createElement('button'); downBtn.type = 'button'; downBtn.title = 'Not helpful'; downBtn.setAttribute('aria-label', 'Mark as not helpful');
      downBtn.innerHTML = ic.thumbDown;
      downBtn.addEventListener('click', () => {
        downBtn.classList.add('nx-voted'); upBtn.classList.remove('nx-voted');
        submitFeedback(id, 'negative');
        showCommentField('negative');
      });
      ma.appendChild(downBtn);
      col.appendChild(ma);

      wrap.appendChild(col);
    } else if (role === 'user') {
      // Column: bubble + timestamp below
      const col = document.createElement('div'); col.className = 'nx-col';
      el.textContent = text || '';
      col.appendChild(el);
      const ts = document.createElement('div'); ts.className = 'nx-ts'; ts.textContent = 'just now'; col.appendChild(ts);
      wrap.appendChild(col);

      // User avatar (gender-neutral silhouette) — appended after bubble,
      // row-reverse puts it on the right visually
      const uav = document.createElement('div'); uav.className = 'nx-av-u'; uav.setAttribute('aria-hidden', 'true'); uav.innerHTML = ic.person;
      wrap.insertBefore(uav, col);
    } else {
      // system/tool messages — no wrapper, just the element
      el.textContent = text || '';
      cv.appendChild(el); cv.lastElementChild && cv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'end' });
      return el;
    }

    cv.appendChild(wrap); cv.lastElementChild && cv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'end' });
    // Return the inner bubble element for appTo/finishMsg compatibility
    el._wrap = wrap;
    return el;
  }

  function appTo(el, text) {
    const mc = el.querySelector('.nx-mc');
    if (mc) {
      // Accumulate raw text, re-render markdown
      if (!el._rawText) el._rawText = '';
      el._rawText += text;
      mc.innerHTML = md(el._rawText);
    } else {
      el.textContent += text;
    }
    cv.lastElementChild && cv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'end' });
  }

  function finishMsg(el) {
    // Add source citation chips for any URLs in the message
    if (el._rawText) {
      const urls = extractUrls(el._rawText);
      if (urls.length > 0) {
        const src = document.createElement('div'); src.className = 'nx-sources';
        for (const url of urls.slice(0, 5)) { // max 5 sources
          const a = document.createElement('a'); a.className = 'nx-src'; a.href = url; a.target = '_blank'; a.rel = 'noopener';
          // Show last meaningful path segment instead of just hostname
          let label;
          try {
            const u = new URL(url);
            const segs = u.pathname.split('/').filter(Boolean);
            const last = segs[segs.length - 1] || '';
            label = last ? last.replace(/[-_]/g, ' ').replace(/\/$/, '') : u.hostname.replace(/^www\./, '');
          } catch (_) { label = url.slice(0, 30); }
          a.innerHTML = ic.link + ' ' + esc(label);
          src.appendChild(a);
        }
        const mc = el.querySelector('.nx-mc');
        if (mc) mc.after(src); else el.appendChild(src);
      }
    }
  }

  function addTyping() {
    const wrap = document.createElement('div'); wrap.className = 'nx-mw nx-mw-a'; wrap.id = 'nx-typing-msg';
    const av = document.createElement('div'); av.className = 'nx-av'; av.innerHTML = '<img src="' + getAvatarSrc() + '" alt="BU">'; wrap.appendChild(av);
    const el = document.createElement('div'); el.className = 'nx-m nx-m-a';
    const dots = document.createElement('div'); dots.className = 'nx-typing'; dots.setAttribute('role', 'status'); dots.setAttribute('aria-label', 'Assistant is typing');
    dots.innerHTML = '<div class="nx-typing-dot"></div><div class="nx-typing-dot"></div><div class="nx-typing-dot"></div>';
    el.appendChild(dots);
    wrap.appendChild(el);
    cv.appendChild(wrap); cv.lastElementChild && cv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'end' });
    return wrap;
  }

  function removeTyping() {
    const t = root.querySelector('#nx-typing-msg');
    if (t) t.remove();
  }

  function addBrk() { const el = document.createElement('div'); el.className = 'nx-rb'; cv.appendChild(el); }

  // Tool call collapsible group
  let currentToolGroup = null;
  let toolCount = 0;

  function getToolGroup() {
    if (currentToolGroup) return currentToolGroup;
    toolCount = 0;
    const wrap = document.createElement('div'); wrap.className = 'nx-tc';
    const hdrEl = document.createElement('div'); hdrEl.className = 'nx-tc-hdr';
    const chevSvg = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>';
    const dot = document.createElement('span'); dot.className = 'nx-tc-dot';
    const label = document.createElement('span'); label.className = 'nx-tc-label'; label.textContent = 'Using tools\u2026';
    hdrEl.innerHTML = chevSvg; hdrEl.appendChild(dot); hdrEl.appendChild(label);
    hdrEl.addEventListener('click', () => wrap.classList.toggle('nx-open'));
    const body = document.createElement('div'); body.className = 'nx-tc-body';
    wrap.appendChild(hdrEl); wrap.appendChild(body);
    cv.appendChild(wrap);
    cv.lastElementChild && cv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'end' });
    currentToolGroup = { wrap, label, body };
    return currentToolGroup;
  }

  let toolLineMap = {}; // tool_call_id → DOM element

  function addToolStart(text, toolCallId) {
    const g = getToolGroup();
    toolCount++;
    g.label.textContent = toolCount === 1 ? 'Searching\u2026' : 'Searching (' + toolCount + ' sources)\u2026';
    const line = document.createElement('div'); line.className = 'nx-tc-line'; line.textContent = text;
    g.body.appendChild(line);
    if (toolCallId) toolLineMap[toolCallId] = line;
    cv.lastElementChild && cv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'end' });
  }

  function addToolResult(text, toolCallId) {
    // Append result to the matching tool start line
    const startLine = toolCallId && toolLineMap[toolCallId];
    if (startLine) {
      const result = document.createElement('span');
      result.style.cssText = 'margin-left:6px;color:#767676;font-size:11px;';
      result.textContent = text;
      startLine.appendChild(result);
    } else {
      const g = getToolGroup();
      const line = document.createElement('div'); line.className = 'nx-tc-line'; line.textContent = text;
      g.body.appendChild(line);
    }
  }

  function finishToolGroup() {
    if (!currentToolGroup) return;
    currentToolGroup.wrap.classList.add('nx-done');
    currentToolGroup.label.textContent = toolCount === 1 ? 'Searched 1 source' : 'Searched ' + toolCount + ' sources';
    currentToolGroup = null; toolCount = 0; toolLineMap = {};
  }

  // Update timestamps periodically
  setInterval(() => {
    root.querySelectorAll('.nx-ts').forEach(ts => {
      const msg = ts.closest('.nx-mw') || ts.closest('.nx-m');
      if (msg) { const t = parseInt(msg.getAttribute('data-ts')); if (t) ts.textContent = relTime(t); }
    });
  }, 30000);

  // ═══════════════════════════════════════════════════════════════
  // FEEDBACK
  // ═══════════════════════════════════════════════════════════════
  function submitFeedback(msgId, rating, comment) {
    if (!convId) return;
    const url = apiUrl(FEEDBACK_PATH_PREFIX + convId + '/feedback');
    const body = { rating };
    if (comment) body.comment = comment;
    fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    }).catch(() => {});
  }

  // ═══════════════════════════════════════════════════════════════
  // HOST-API RETRIEVAL STRATEGY  (allowToRetrieveFromApi)
  // ═══════════════════════════════════════════════════════════════
  // Silent path: when enabled and an `auth_required` arrives, the
  // widget GETs the host-supplied "who am I" endpoint and pulls
  // openid / email out of the JSON. Populates the page state used
  // by the lite-chat body fields (no Bearer token is produced).
  function rp(o,p){if(!p)return undefined;return p.split('.').reduce((a2,k)=>(a2!=null&&typeof a2==='object')?a2[k]:undefined,o)}
  function haEnabled(){return!!(C.allowToRetrieveFromApi&&C.hostApi&&C.hostApi.endpoint)}
  async function hostApiFetch(){if(!C.allowToRetrieveFromApi)return{ok:false};const api=C.hostApi||{};if(!api.endpoint)return{ok:false};let r;try{r=await fetch(api.endpoint,{method:api.method||'GET',credentials:api.credentials||'same-origin',headers:api.headers||undefined})}catch(e){return{ok:false,error:String(e)}}if(!r.ok)return{ok:false,error:'HTTP '+r.status};let d;try{d=await r.json()}catch(_){return{ok:false,error:'Bad JSON.'}}let oid2,em;if(typeof api.mapper==='function'){const m=api.mapper(d)||{};oid2=m.openid;em=m.email}else{oid2=rp(d,api.openidPath);em=rp(d,api.emailPath)}if(!oid2)return{ok:false,error:'No openid.'};openid=String(oid2).trim();oidSrc='host-api';if(em)email=String(em).trim();return{ok:true}}

  // ═══════════════════════════════════════════════════════════════
  // NEXUS-AUTH POPUP OIDC STRATEGY
  // ═══════════════════════════════════════════════════════════════
  // When allowNexusAuthLogin: true, on `auth_required` the widget
  // opens an OAuth popup at `${apiBaseUrl}/api/v1/auth/start`,
  // listens for a `postMessage` carrying the JWT, and switches
  // /api/v1/lite-chat calls to /api/v1/chat with `Authorization:
  // Bearer <jwt>`.
  function nxEnabled() { return !!(C.allowNexusAuthLogin && API_BASE); }

  // Bearer-protected chat endpoint (used when an authToken is held).
  function authedBackendUrl() { return apiUrl(CHAT_PATH); }

  // Decode a JWT payload (no signature verification — display only).
  function nxDecodeJwt(token) {
    if (!token) return null;
    const b64 = String(token).split('.')[1];
    if (!b64) return null;
    const s = b64.replace(/-/g, '+').replace(/_/g, '/');
    const pad = '==='.slice((s.length + 3) % 4);
    return JSON.parse(decodeURIComponent(atob(s + pad).split('').map(c =>
      '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2)
    ).join('')));
  }

  /**
   * Open the nexus-auth popup and wait for postMessage. Returns a
   * Promise that resolves with the captured JWT, or rejects if the
   * popup is superseded by another sign-in attempt or is misconfigured.
   *
   * MUST be invoked synchronously inside a user gesture (e.g. click)
   * because window.open(...) is the call popup blockers gate on. The
   * `renderNexusAuth` button below preserves that gesture.
   */
  function nexusSignIn() {
    return new Promise((resolve, reject) => {
      if (!C.allowNexusAuthLogin) { reject(new Error('Nexus auth is disabled.')); return; }
      if (!API_BASE)              { reject(new Error('apiBaseUrl is not configured.')); return; }

      // Reject any prior in-flight popup — superseded by this new click.
      if (pendingAuth) { try { pendingAuth.reject(new Error('superseded')); } catch (_) {} }

      const nonce = (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function')
        ? crypto.randomUUID()
        : (Math.random().toString(36).slice(2) + Date.now().toString(36));
      pendingAuth = { nonce, resolve, reject };

      const params = new URLSearchParams({
        return_origin: window.location.origin,
        nonce,
      });
      if (C.nexusAuth.extraParams && typeof C.nexusAuth.extraParams === 'object') {
        for (const k of Object.keys(C.nexusAuth.extraParams)) {
          const v = C.nexusAuth.extraParams[k];
          if (v != null) params.set(k, String(v));
        }
      }
      // Always appended LAST so the host can't accidentally drop it via
      // extraParams. Lets the auth BFF (/api/v1/auth/start on the agent)
      // know which deployment the popup should sign in against — the
      // same deployment slug the widget sends in /chat bodies.
      const deploy = C.chatDefaults && C.chatDefaults.deployment;
      if (typeof deploy === 'string' && deploy.trim()) {
        params.set('deployment', deploy.trim());
      }
      const url = apiUrl(AUTH_START_PATH) + '?' + params.toString();
      // window.open MUST be sync inside the click handler or popup
      // blockers will eat it.
      window.open(url, 'nexus-auth', C.nexusAuth.popupFeatures || 'width=500,height=700');
    });
  }

  /**
   * Clear in-memory nexus-auth state. There is no server-side session
   * to revoke; refresh = signed out by design (token never persisted).
   */
  function nexusSignOut() {
    authToken = null;
    authClaims = null;
    if (C.nexusAuth.resetConversationOnAuth) {
      convId = null;
      try { sessionStorage.removeItem(STORE_KEY); } catch (_) {}
    }
  }

  // Install the postMessage listener exactly once. Three-layer envelope
  // check: origin must match the apiBaseUrl's origin (the popup is
  // served by the agent itself), type must be 'nexus-auth', and the
  // payload's nonce must equal the in-flight pendingAuth.nonce.
  // Anything else is silently ignored — the widget shares the global
  // `message` event with everything else on the page.
  if (typeof window !== 'undefined') {
    window.addEventListener('message', (e) => {
      if (!C.allowNexusAuthLogin) return;
      if (!apiOrigin) return;
      if (e.origin !== apiOrigin) return;
      if (!e.data || e.data.type !== 'nexus-auth') return;
      const p = e.data.payload || {};
      if (!pendingAuth || p.nonce !== pendingAuth.nonce) return;

      const tok = p.token || null;
      authToken = tok;
      try { authClaims = nxDecodeJwt(tok); } catch (_) { authClaims = {}; }
      // Surface email from the JWT for chat.js's existing display paths
      // (header avatar tooltip, etc.). Doesn't replace any host-supplied
      // value if the claim is missing.
      if (authClaims) {
        const e2 = authClaims.email || authClaims.preferred_username || authClaims.upn;
        if (e2) email = e2;
      }
      // Drop conversation_id on auth-state change so we don't continue an
      // anonymous conversation under a now-known identity.
      if (C.nexusAuth.resetConversationOnAuth) {
        convId = null;
        try { sessionStorage.removeItem(STORE_KEY); } catch (_) {}
      }

      const r = pendingAuth.resolve;
      pendingAuth = null;
      try { r(tok); } catch (_) {}
    });
  }

  // ═══════════════════════════════════════════════════════════════
  // SSE STREAMING
  // ═══════════════════════════════════════════════════════════════
  // Build the JSON body for /chat. When ANY strategy has produced a
  // JWT into `authToken` (nexus-auth popup, MSAL silent, or MSAL popup),
  // we deliberately omit `openid` / `email` — the agent verifies
  // identity from the Bearer JWT instead, and forwarding unverified
  // claims would be security theatre. The legacy host-API / host-page
  // probing strategies (which populate `openid` / `email` directly)
  // keep sending those fields as before since they don't produce a
  // bearer token.
  function buildBody(msg) {
    const authed = !!authToken;
    const b = Object.assign(
      {},
      C.chatDefaults,
      authed
        ? { message: msg, conversation_id: convId }
        : { message: msg, conversation_id: convId, openid, email },
    );
    for (const k of Object.keys(b)) { if (b[k] == null || b[k] === '') delete b[k]; }
    return b;
  }

  async function sseStream(msg, H) {
    // Endpoint + headers vary with auth state. Both URLs are derived
    // from `apiBaseUrl` + a path constant — there's no per-endpoint
    // configuration anywhere else in the file. Anonymous requests go
    // to /api/v1/lite-chat; when we hold a JWT in `authToken` (from
    // the nexus-auth popup) we switch to the Bearer-protected
    // /api/v1/chat.
    const authed = !!authToken;
    const url = authed ? apiUrl(CHAT_PATH) : apiUrl(LITE_CHAT_PATH);
    const headers = { 'Content-Type': 'application/json', 'Accept': 'text/event-stream' };
    if (authed) headers['Authorization'] = 'Bearer ' + authToken;

    let r;
    try { r = await fetch(url, { method: 'POST', headers, body: JSON.stringify(buildBody(msg)) }); }
    catch (e) { H.err(e); return; }
    if (!r.ok || !r.body) { let d = ''; try { d = ' \u2014 ' + (await r.text()); } catch (_) {} H.err(new Error('HTTP ' + r.status + d)); return; }
    const rd = r.body.getReader(), dc = new TextDecoder('utf-8'), SP = /\r?\n\r?\n/;
    let bf = '';
    try {
      while (true) { const { value, done } = await rd.read(); if (done) break; bf += dc.decode(value, { stream: true }); let m; while ((m = SP.exec(bf)) !== null) { parseEvt(bf.slice(0, m.index), H); bf = bf.slice(m.index + m[0].length); } }
      if (bf.trim()) parseEvt(bf, H);
      H.end();
    } catch (e) { H.err(e); }
  }

  function parseEvt(raw, H) {
    let ev = 'message', ds = '';
    for (const l of raw.split(/\r?\n/)) { if (l.startsWith('event:')) ev = l.slice(6).trim(); else if (l.startsWith('data:')) { const c2 = l.slice(5); ds += c2.startsWith(' ') ? c2.slice(1) : c2; } }
    let d = {}; if (ds) { try { d = JSON.parse(ds); } catch (_) { d = { text: ds }; } }
    switch (ev) {
      case 'message_start': if (d.conversation_id) { convId = d.conversation_id; try { sessionStorage.setItem(STORE_KEY, convId); } catch (_) {} } break;
      case 'content_delta': if (typeof d.text === 'string') H.delta(d.text); break;
      case 'round_break': H.brk(); break;
      case 'tool_call_start': {
        const friendly = (d.tool_name || '?').replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
        H.toolStart('Searching: ' + friendly, d.tool_call_id);
      } break;
      case 'tool_result': {
        if (d.is_error) H.toolResult('\u2717 Error', d.tool_call_id);
        else H.toolResult('\u2713', d.tool_call_id);
      } break;
      case 'message_end': H.end(); break;
      case 'auth_required': H.auth(d || {}); break;
      case 'error': H.err(new Error((d && d.message) || 'Server error')); break;
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // SEND MESSAGE
  // ═══════════════════════════════════════════════════════════════
  async function sendMsg(qo) {
    if (busy) return;
    const retry = typeof qo === 'string';
    const text = retry ? qo.trim() : inp.value.trim();
    if (!text) return;
    if (!retry) { inp.value = ''; inp.style.height = '44px'; }

    addMsg('user', text);
    lastQ = text;

    // Show typing indicator
    const typingEl = addTyping();
    busy = true; snb.disabled = true;
    let gotContent = false;
    let respEl = null;
    let settled = false;

    const finish = () => {
      if (settled) return; settled = true;
      removeTyping();
      finishToolGroup();
      if (respEl) {
        respEl.classList.remove('nx-str');
        finishMsg(respEl);
      }
      busy = false; snb.disabled = false; inp.focus();
      playNotif();
    };

    await sseStream(text, {
      delta: (chunk) => {
        if (!gotContent) { removeTyping(); finishToolGroup(); respEl = addMsg('assistant', ''); respEl.classList.add('nx-str'); gotContent = true; }
        appTo(respEl, chunk);
      },
      toolStart: (label, id) => { if (C.showToolEvents) addToolStart(label, id); },
      toolResult: (label, id) => { if (C.showToolEvents) addToolResult(label, id); },
      brk: () => {
        finishToolGroup();
        // Close current bubble so post-tool content gets a fresh one
        if (respEl) { respEl.classList.remove('nx-str'); finishMsg(respEl); }
        respEl = null; gotContent = false;
      },
      end: () => finish(),
      err: (e) => {
        if (!gotContent) { removeTyping(); respEl = addMsg('assistant', ''); gotContent = true; }
        appTo(respEl, (respEl._rawText ? '\n' : '') + '[error] ' + (e.message || e));
        finish();
      },
      auth: (p) => {
        removeTyping();
        // Collect elements to remove on successful auth: preamble, tool group, user message
        const toRemove = [];
        if (currentToolGroup) { toRemove.push(currentToolGroup.wrap); currentToolGroup = null; toolCount = 0; }
        if (gotContent && respEl) { toRemove.push(respEl._wrap || respEl); }
        // Also remove the user message that triggered this (retry will re-add it)
        const userBubbles = cv.querySelectorAll('.nx-mw-u');
        if (userBubbles.length > 0) toRemove.push(userBubbles[userBubbles.length - 1]);

        gotContent = false; respEl = null;
        // Create a clean system-style auth message
        const authEl = addMsg('system', p.message || 'Please sign in to continue.');
        // Mark as done so retry can proceed
        settled = true; busy = false; snb.disabled = false;

        // Strategy precedence on `auth_required`:
        //   1. host-API   — silent retrieval, no UI
        //   2. nexus-auth — popup OIDC
        // No third path: MSAL was removed in this fork. If neither
        // strategy is configured, the auth bubble is left as a plain
        // message — the host has misconfigured the widget.
        if (haEnabled()) { handleHostAuth(authEl, () => {}); return; }
        if (nxEnabled()) {
          renderNexusAuth(authEl, p, () => {
            authEl.remove();
            toRemove.forEach(el => { if (el && el.parentNode) el.remove(); });
          });
          return;
        }
        // No strategy configured — surface a hint inline so the host
        // sees what's wrong.
        const hint = document.createElement('span');
        hint.className = 'nx-an';
        hint.textContent = 'No sign-in strategy configured (set allowNexusAuthLogin or allowToRetrieveFromApi).';
        authEl.appendChild(document.createElement('br'));
        authEl.appendChild(hint);
      },
    });
  }

  async function handleHostAuth(el, fin) {
    appTo(el, 'Verifying\u2026');
    const r = await hostApiFetch(); fin();
    if (r.ok) { if (C.autoRetryAfterAuth && lastQ) { appTo(el, ' \u2713'); setTimeout(() => sendMsg(lastQ), 150); } else appTo(el, ' \u2713 ask again.'); }
    else appTo(el, ' [error] ' + (r.error || 'failed'));
  }

  /**
   * Render the auth bubble UI for the nexus-auth strategy. Mirrors the
   * shape of `renderAuth` (uses the same `nx-al` / `nx-an` / `nx-bsy`
   * classes) but dispatches to `nexusSignIn()` instead of MSAL. Only
   * called when `nxEnabled()` returns true.
   *
   * The click handler runs `window.open(...)` synchronously so popup
   * blockers don't interfere; the postMessage listener installed at
   * boot resolves the returned promise on success.
   */
  function renderNexusAuth(el, p, onAuthSuccess) {
    const lk = document.createElement('span');
    lk.className = 'nx-al';
    lk.setAttribute('role', 'button');
    lk.setAttribute('tabindex', '0');
    lk.textContent = p.label || 'Sign in';

    const nt = document.createElement('span');
    nt.className = 'nx-an';
    nt.textContent = 'Opens a popup to sign you in.';

    el.appendChild(document.createElement('br'));
    el.appendChild(lk);
    el.appendChild(nt);

    function afterLogin() {
      if (onAuthSuccess) { try { onAuthSuccess(); } catch (_) {} }
      if (C.autoRetryAfterAuth && lastQ) setTimeout(() => sendMsg(lastQ), 150);
    }

    const go = async () => {
      if (lk.classList.contains('nx-bsy')) return;
      lk.classList.add('nx-bsy');
      lk.textContent = 'Opening…';
      try {
        await nexusSignIn();
        lk.textContent = 'Signed in ✓';
        afterLogin();
      } catch (e) {
        lk.classList.remove('nx-bsy');
        lk.textContent = p.label || 'Sign in';
        // 'superseded' fires when the user clicked twice; keep silent.
        const m = (e && e.message) ? e.message : String(e);
        if (m !== 'superseded') nt.textContent = 'Failed: ' + m;
      }
    };
    lk.addEventListener('click', go);
    lk.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); go(); }
    });
  }

  // ═══════════════════════════════════════════════════════════════
  // EVENT WIRING
  // ═══════════════════════════════════════════════════════════════
  snb.addEventListener('click', () => sendMsg());
  inp.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMsg(); } });
  inp.addEventListener('input', () => { inp.style.height = '44px'; inp.style.height = Math.min(inp.scrollHeight, 120) + 'px'; });
  if (launcher) launcher.addEventListener('click', toggle);

  // ═══════════════════════════════════════════════════════════════
  // VOICE INPUT — transformers.js + Xenova/whisper-base.en (in-browser)
  // ═══════════════════════════════════════════════════════════════
  // Hold-to-talk dictation. The model (~150 MB of ONNX weights) is fetched
  // from the HF CDN on first use and cached in IndexedDB by transformers.js;
  // no audio leaves the browser. Inference runs on WebGPU when available
  // and falls back to WASM.
  //
  // UX: press-and-hold the mic to record, release to transcribe + auto-send
  // (the gesture implies commitment — the user kept their finger down). The
  // button reuses chat.js's existing `.nx-mic` styles plus three additive
  // state classes: `.nx-mic-loading` (model download), `.nx-mic-on` (live
  // recording — same red pulse the input row already uses), and
  // `.nx-mic-busy` (transcribing).
  //
  // Browsers without `MediaRecorder` / `getUserMedia` (rare today) get the
  // mic hidden via the existing `.nx-mic-hide` class.
  const micBtn = root.querySelector('.nx-mic');

  if (!micBtn) {
    /* no mic button rendered — nothing to wire */
  } else if (!navigator.mediaDevices || !window.MediaRecorder) {
    micBtn.classList.add('nx-mic-hide');
  } else {
    const VOICE_MODEL = 'Xenova/whisper-base.en';
    const VOICE_LIB   = 'https://cdn.jsdelivr.net/npm/@xenova/transformers@2.17.2/dist/transformers.min.js';

    // ── State ────────────────────────────────────────────────
    // `recording` = MediaRecorder is active.
    // `pressHeld` = the user's pointer is currently down on the button.
    // We track them separately so a release during the async getUserMedia
    // permission prompt cancels cleanly instead of leaving an orphan stream.
    let transcriber = null;
    let recording = false;
    let pressHeld = false;
    let mediaStream = null;
    let mediaRecorder = null;
    let recordedChunks = [];

    // Stable "home" for the input placeholder. onRecordingStop restores
    // to this — never to a transient value (e.g. "Transcribing…")
    // that a fast second press might capture.
    let basePlaceholder = C.placeholder || '';
    inp.placeholder = basePlaceholder;

    // Original mic title for restore (chat.js's existing `Voice input`).
    const baseTitle = micBtn.getAttribute('title') || 'Voice input';

    /**
     * Mic button state machine. All state classes flow through here so
     * transitions are atomic — no chance of `nx-mic-on` and `nx-mic-busy`
     * coexisting on a race. Only the additive nx-mic-* classes are
     * touched; `.nx-mic` itself and the SVG icon are preserved.
     *
     * @param {'idle'|'loading'|'listening'|'busy'|'error'|'unavailable'} state
     * @param {string} [title] tooltip override
     */
    function setMicState(state, title) {
      micBtn.classList.remove('nx-mic-on', 'nx-mic-loading', 'nx-mic-busy', 'nx-mic-hide');
      let disabled = false;
      let resolvedTitle = title;
      switch (state) {
        case 'loading':
          micBtn.classList.add('nx-mic-loading');
          disabled = true;
          if (resolvedTitle === undefined) resolvedTitle = 'Loading voice model…';
          break;
        case 'listening':
          micBtn.classList.add('nx-mic-on');
          if (resolvedTitle === undefined) resolvedTitle = 'Listening — release to send';
          break;
        case 'busy':
          micBtn.classList.add('nx-mic-busy');
          disabled = true;
          if (resolvedTitle === undefined) resolvedTitle = 'Transcribing…';
          break;
        case 'unavailable':
          micBtn.classList.add('nx-mic-hide');
          disabled = true;
          break;
        case 'error':
          disabled = true;
          if (resolvedTitle === undefined) resolvedTitle = 'Voice input unavailable';
          break;
        case 'idle':
        default:
          if (resolvedTitle === undefined) resolvedTitle = 'Hold to talk';
          break;
      }
      micBtn.disabled = disabled;
      micBtn.setAttribute('title', resolvedTitle || baseTitle);
      micBtn.setAttribute('aria-label', resolvedTitle || baseTitle);
    }

    /**
     * Lazily pulls the transformers.js ESM bundle and instantiates the
     * Whisper pipeline. Called once at startup. While the download runs
     * the mic button stays in `loading` state and shows rough overall
     * progress in its tooltip.
     * @returns {Promise<void>}
     */
    async function loadVoiceModel() {
      try {
        setMicState('loading');
        // Use the package's prebuilt ESM bundle. jsdelivr's `+esm` auto-
        // transform rewrites imports in a way that breaks onnxruntime-web's
        // backend registration (null.registerBackend) — go direct.
        const { pipeline, env } = await import(VOICE_LIB);
        env.allowLocalModels = false;   // hit HF CDN; rely on IndexedDB cache
        // Aggregate per-file download progress into a single percentage.
        const fileProgress = {};
        const onProgress = (p) => {
          if (p.status === 'progress' && p.file) {
            fileProgress[p.file] = p.progress || 0;
            const vals = Object.values(fileProgress);
            const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
            setMicState('loading', 'Loading voice model… ' + Math.round(avg) + '%');
          } else if (p.status === 'done') {
            setMicState('loading', 'Loading voice model… finalizing');
          }
        };
        transcriber = await pipeline(
          'automatic-speech-recognition',
          VOICE_MODEL,
          { progress_callback: onProgress }
        );
        setMicState('idle');
        try { console.log('[voice] model ready'); } catch (_) {}
      } catch (err) {
        try { console.error('[voice] model load failed:', err); } catch (_) {}
        setMicState('error', 'Voice input unavailable (model failed to load)');
      }
    }

    /**
     * Decode the MediaRecorder blob, downmix to mono, resample to 16 kHz —
     * the exact shape Whisper expects.
     * @param {Blob} blob
     * @returns {Promise<Float32Array>}
     */
    async function blobToFloat32_16kMono(blob) {
      const arrayBuf = await blob.arrayBuffer();
      const tmpCtx = new (window.AudioContext || window.webkitAudioContext)();
      const decoded = await tmpCtx.decodeAudioData(arrayBuf);
      tmpCtx.close();

      const target = 16000;
      const offline = new OfflineAudioContext(
        1,
        Math.ceil(decoded.duration * target),
        target
      );
      const src = offline.createBufferSource();
      src.buffer = decoded;
      src.connect(offline.destination);
      src.start();
      const rendered = await offline.startRendering();
      return rendered.getChannelData(0);
    }

    /** Request mic permission and start a MediaRecorder session. */
    async function startRecording() {
      if (!transcriber || recording) return;
      let stream;
      try {
        stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      } catch (err) {
        try { console.warn('[voice] mic permission denied:', err); } catch (_) {}
        pressHeld = false;
        setMicState('idle', 'Microphone blocked — check browser permissions');
        return;
      }
      // The user may have released (or a second press may have already
      // resolved) while the permission prompt was open. Bail cleanly.
      if (!pressHeld || recording) {
        stream.getTracks().forEach((t) => t.stop());
        return;
      }
      mediaStream = stream;
      recordedChunks = [];
      mediaRecorder = new MediaRecorder(mediaStream);
      mediaRecorder.addEventListener('dataavailable', (e) => {
        if (e.data && e.data.size > 0) recordedChunks.push(e.data);
      });
      mediaRecorder.addEventListener('stop', onRecordingStop);
      mediaRecorder.start();
      recording = true;
      setMicState('listening');
    }

    /** Stop the MediaRecorder; the `stop` event drives transcription. */
    function stopRecording() {
      if (!recording) return;
      try { mediaRecorder.stop(); } catch (_) {}
      try { mediaStream.getTracks().forEach((t) => t.stop()); } catch (_) {}
      recording = false;
      setMicState('busy');
    }

    /**
     * MediaRecorder `stop` handler. Assembles captured chunks, runs Whisper
     * over them, and inserts the transcript into the input. Hold-to-talk
     * implies intent, so we auto-submit on success.
     */
    async function onRecordingStop() {
      inp.placeholder = 'Transcribing…';
      try {
        const blob = new Blob(recordedChunks, { type: mediaRecorder.mimeType });
        if (blob.size >= 500) { // anything smaller is silence / no audio captured
          const audio = await blobToFloat32_16kMono(blob);
          const result = await transcriber(audio);
          const text = (result && result.text ? result.text : '').trim();
          if (text) {
            inp.value = inp.value ? inp.value + ' ' + text : text;
            // Match chat.js's input auto-resize behaviour after programmatic value change.
            inp.style.height = '44px';
            inp.style.height = Math.min(inp.scrollHeight, 120) + 'px';
            inp.focus();
            // Hold-to-talk implies commitment — auto-send on release.
            sendMsg();
          }
        }
        inp.placeholder = basePlaceholder;
      } catch (err) {
        try { console.error('[voice] transcription failed:', err); } catch (_) {}
        inp.placeholder = 'Voice transcription failed';
        setTimeout(() => {
          if (inp.placeholder === 'Voice transcription failed') {
            inp.placeholder = basePlaceholder;
          }
        }, 4000);
      } finally {
        setMicState('idle');
      }
    }

    // ── Press-and-hold gesture handlers ─────────────────────
    // Mouse + touch handled symmetrically. preventDefault() avoids the
    // double-fire (focus + click) sequence on touch devices.
    const pressStart = (e) => {
      if (micBtn.disabled || recording || pressHeld) return;
      e.preventDefault();
      pressHeld = true;
      startRecording();
    };
    const pressEnd = (e) => {
      if (!pressHeld) return;
      e.preventDefault();
      pressHeld = false;
      // If startRecording is still mid-permission-prompt, the `pressHeld`
      // check inside it abandons the stream cleanly.
      if (recording) stopRecording();
    };
    micBtn.addEventListener('mousedown', pressStart);
    micBtn.addEventListener('touchstart', pressStart, { passive: false });
    micBtn.addEventListener('mouseup', pressEnd);
    micBtn.addEventListener('mouseleave', pressEnd);
    micBtn.addEventListener('touchend', pressEnd);
    micBtn.addEventListener('touchcancel', pressEnd);

    // Eagerly start the model download so most of the ~150 MB transfer
    // happens in the background while the user is still reading the page.
    loadVoiceModel();
  }
  // ═══════════════════════════════════════════════════════════════
  // END VOICE INPUT
  // ═══════════════════════════════════════════════════════════════

  // ═══════════════════════════════════════════════════════════════
  // PUBLIC API
  // ═══════════════════════════════════════════════════════════════
  function show() { root.classList.remove('nx-hidden'); if (launcher) { launcher.classList.add('nx-hidden'); launcher.classList.remove('nx-pulse'); } widgetVisible = true; setTimeout(() => inp.focus(), 0); }
  function hide() { root.classList.add('nx-hidden'); bk.classList.add('nx-hidden'); if (launcher) launcher.classList.remove('nx-hidden'); widgetVisible = false; trapFocus(false); }
  function toggle() { root.classList.contains('nx-hidden') ? show() : hide(); }
  function unmount() { root.remove(); bk.remove(); if (launcher) launcher.remove(); const s = document.getElementById(SID); if (s) s.remove(); try { delete window.TerrierChatWidget; } catch (_) {} }
  function reset() { convId = null; try { sessionStorage.removeItem(STORE_KEY); } catch (_) {} cv.innerHTML = ''; greet(); }

  window.TerrierChatWidget = {
    show, hide, toggle, unmount, reset, setMode,
    get mounted() { return true },
    get conversationId() { return convId },
    get mode() { return mode },
    // Email display (set by nexus-auth claims, host-API mapping, or
    // explicit setEmail() from the host).
    getEmail() { return email },
    setEmail(e2) { email = e2 || null; },
    // Nexus-auth strategy (popup OIDC; switches /chat from
    // /api/v1/lite-chat to /api/v1/chat when a token is captured).
    // Active only when `allowNexusAuthLogin` is true.
    signInNexus: nexusSignIn,
    signOutNexus: nexusSignOut,
    getAuthToken() { return authToken; },
    getAuthClaims() { return authClaims; },
  };
  try { window.dispatchEvent(new CustomEvent('terrier-chat-v2-ready')); } catch (_) {}
})();
