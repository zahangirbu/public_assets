/**
 * NEXUS Chat Widget v2
 * --------------------
 * Premium embeddable chat for the NEXUS Agent backend.
 *
 * Features:
 *   - Panel / sidebar / modal modes with drag-move + edge-resize
 *   - Markdown rendering (bold, italic, links, lists, code, headings)
 *   - Auto-linkified URLs + source citation chips
 *   - Typing indicator (animated dots) before first content
 *   - Message timestamps on hover
 *   - Copy-to-clipboard on assistant messages
 *   - Thumbs up/down feedback per message
 *   - Session persistence (survives page refresh)
 *   - ARIA accessibility (focus trap in modal, live region, labels)
 *   - Smooth slide-up open animation
 *   - BU branding (Whitney font, BU Red, sub-brand shield avatar)
 *   - MSAL / host-API / cookie auth strategies
 *   - SSE streaming with tool events
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
    backendUrl: 'http://localhost:8080/api/v1/lite-chat',
    // Endpoint for thumbs-up / thumbs-down feedback POSTs. When null,
    // it is derived from the backendUrl base + '/api/v1/feedback'.
    feedbackUrl: null,
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

    // --- Theming / runtime extras (v2-only) -----------------------------
    // Primary brand color used for the header, send button, links, and
    // citation chips. Any valid CSS color string.
    accentColor: '#cc0000',
    // When true, attempt to GET runtime config (provider/deployment/etc.)
    // from the backend at mount time so hosts don't have to hard-code it.
    autoFetchConfig: true,
    // Override URL for the runtime-config endpoint. When null, it is
    // derived from the backendUrl base.
    configUrl: null,
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

  // MSAL config, deep-merged so hosts can override only what they need.
  C.msal = Object.assign({
    clientId: null,
    authority: null,
    redirectUri: null,
    scopes: ['openid', 'profile'],
    loginHint: null,
    scriptUrl: 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.30.0/lib/msal-browser.min.js',
  }, (uc && uc.msal) || {});

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

  // ═══════════════════════════════════════════════════════════════
  // STATE
  // ═══════════════════════════════════════════════════════════════
  const WID = 'nx-chat', LID = 'nx-chat-launcher', SID = 'nx-chat-style';
  const STORE_KEY = 'nx_chat_convId';
  let convId = null, openid = null, oidSrc = null, email = null, lastQ = null;
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
  const apiBase = C.configUrl || C.backendUrl.replace(/\/api\/v1\/.*$/, '');

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
#${WID} .nx-cv{flex:1;overflow-y:auto;padding:16px;display:flex;flex-direction:column;gap:12px;background:#fff}

/* Messages — row layout with avatar beside bubble */
#${WID} .nx-mw{display:flex;gap:8px;max-width:92%;animation:nx-msg-l .25s ease}
#${WID} .nx-mw-a{align-self:flex-start;align-items:flex-start}
#${WID} .nx-mw-u{align-self:flex-end;flex-direction:row-reverse;animation:nx-msg-r .25s ease}
#${WID} .nx-m{border-radius:16px;word-wrap:break-word;position:relative;flex:1;min-width:0}
#${WID} .nx-m-u{background:${ac};color:#fff;border-bottom-right-radius:4px;padding:10px 14px}
#${WID} .nx-m-a{background:#f0f0f0;color:#2D2926;border-bottom-left-radius:4px;padding:10px 14px}
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
#${WID} .nx-ts{font-size:10px;color:#767676;margin-top:2px;opacity:0;transition:opacity .15s;user-select:none}
#${WID} .nx-mw:hover .nx-ts,#${WID} .nx-m:hover .nx-ts{opacity:1}

/* Message actions (copy, feedback) */
#${WID} .nx-ma{display:flex;gap:4px;margin-top:4px;opacity:0;transition:opacity .15s}
#${WID} .nx-mw:hover .nx-ma{opacity:1}
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
#${WID} .nx-mic.nx-mic-on{color:${ac};animation:nx-pulse 1.5s infinite}
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
        <!-- Voice dictation button (commented out — requires HTTPS + speech service access)
        <button type="button" class="nx-mic" title="Voice input" aria-label="Start voice dictation">${ic.mic}</button>
        -->
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
      C.samplePrompts.forEach((p, i) => { const b = document.createElement('button'); b.type = 'button'; b.className = 'nx-spb'; b.style.animationDelay = (i * 80) + 'ms'; b.textContent = p; b.addEventListener('click', () => { inp.value = p; sendMsg(); }); s.appendChild(b); });
      inner.appendChild(s);
    }
    w.appendChild(inner);
    cv.appendChild(w);
  }
  greet();

  // Auto-fetch deployment config from the agent backend
  if (C.autoFetchConfig && C.chatDefaults.deployment) {
    fetch(apiBase + '/api/v1/deployments?deploy_type=web-widget')
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

      const mc = document.createElement('div'); mc.className = 'nx-mc'; mc.innerHTML = md(text || ''); el.appendChild(mc);
      const ts = document.createElement('div'); ts.className = 'nx-ts'; ts.textContent = 'just now'; el.appendChild(ts);

      // Actions bar
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
        const existing = el.querySelector('.nx-fb-comment');
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
        el.appendChild(row);
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
      el.appendChild(ma);

      wrap.appendChild(el);
    } else if (role === 'user') {
      el.textContent = text || '';
      const ts = document.createElement('div'); ts.className = 'nx-ts'; ts.textContent = 'just now'; el.appendChild(ts);
      wrap.appendChild(el);

      // User avatar (gender-neutral silhouette) — appended after bubble,
      // row-reverse puts it on the right visually
      const uav = document.createElement('div'); uav.className = 'nx-av-u'; uav.setAttribute('aria-hidden', 'true'); uav.innerHTML = ic.person;
      wrap.insertBefore(uav, el);
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

  function addToolLine(text) {
    const g = getToolGroup();
    toolCount++;
    g.label.textContent = 'Using tools\u2026 (' + toolCount + ')';
    const line = document.createElement('div'); line.className = 'nx-tc-line'; line.textContent = text;
    g.body.appendChild(line);
    cv.lastElementChild && cv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'end' });
  }

  function finishToolGroup() {
    if (!currentToolGroup) return;
    currentToolGroup.wrap.classList.add('nx-done');
    currentToolGroup.label.textContent = 'Used ' + toolCount + ' tool' + (toolCount !== 1 ? 's' : '');
    currentToolGroup = null; toolCount = 0;
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
    const url = C.feedbackUrl || (apiBase + '/api/v1/conversations/' + convId + '/feedback');
    const body = { rating };
    if (comment) body.comment = comment;
    fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    }).catch(() => {});
  }

  // ═══════════════════════════════════════════════════════════════
  // AUTH (compact — all strategies from original)
  // ═══════════════════════════════════════════════════════════════
  function isOid(v){if(typeof v!=='string')return false;const s=v.trim();if(s.length<8||s.length>8192)return false;if(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s))return true;if(/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(s))return true;return s.length>=16&&!/\s/.test(s)}
  function djwt(j){try{const p=String(j).split('.');if(p.length!==3)return null;const b=p[1].replace(/-/g,'+').replace(/_/g,'/');return JSON.parse(decodeURIComponent(atob(b).split('').map(c=>'%'+('00'+c.charCodeAt(0).toString(16)).slice(-2)).join('')))}catch(_){return null}}
  function joid(j){const c=djwt(j);return c?(c.oid||c.sub||null):null}
  function probeGlobals(){for(const n of['openid','__openid__','userOpenid','idToken','__idToken__']){const v=window[n];if(isOid(v))return{v,s:'g:'+n}}for(const n of['user','currentUser','authContext','msalAccount']){const o=window[n];if(o&&typeof o==='object'){const v=o.oid||o.openid||o.id_token||o.idToken||o.sub;if(isOid(v))return{v,s:'g:'+n}}}for(const n of['msalInstance','_msalInstance','__msal__']){const i=window[n];if(!i||typeof i.getAllAccounts!=='function')continue;try{for(const a2 of(i.getAllAccounts()||[])){const oid=a2.idTokenClaims&&(a2.idTokenClaims.oid||a2.idTokenClaims.sub);if(isOid(oid))return{v:oid,s:'msal:'+n};if(a2.idToken&&isOid(a2.idToken))return{v:joid(a2.idToken)||a2.idToken,s:'msal:'+n}}}catch(_){}}return null}
  function probeStorage(st,lb){if(!st)return null;try{for(let i=0;i<st.length;i++){const k=st.key(i);if(!k)continue;const kl=k.toLowerCase();if(!(kl.includes('idtoken')||kl.startsWith('msal.')||kl.includes('login.windows.net')||kl.includes('login.microsoftonline.com')))continue;const r=st.getItem(k);if(!r)continue;try{const p=JSON.parse(r);if(p&&p.credentialType==='IdToken'&&isOid(p.secret))return{v:joid(p.secret)||p.secret,s:lb+':'+k};if(p&&(p.oid||p.openid)){const v2=p.oid||p.openid;if(isOid(v2))return{v:v2,s:lb+':'+k}}}catch(_){if(isOid(r))return{v:joid(r)||r,s:lb+':'+k}}}}catch(_){}return null}
  function probeCookies(){if(!document.cookie)return null;for(const p of document.cookie.split(';').map(p2=>p2.trim())){const eq=p.indexOf('=');if(eq<0)continue;const n=p.slice(0,eq).toLowerCase(),v2=decodeURIComponent(p.slice(eq+1));if(['openid','id_token','idtoken','user_oid','auth_token'].includes(n)&&isOid(v2))return{v:joid(v2)||v2,s:'c:'+n}}return null}
  function probeHost(){if(!C.extractOpenidFromHost)return;const f=probeGlobals()||probeStorage(window.localStorage,'ls')||probeStorage(window.sessionStorage,'ss')||probeCookies();if(f){openid=f.v;oidSrc=f.s}}
  function rp(o,p){if(!p)return undefined;return p.split('.').reduce((a2,k)=>(a2!=null&&typeof a2==='object')?a2[k]:undefined,o)}
  function haEnabled(){return!!(C.allowToRetrieveFromApi&&C.hostApi&&C.hostApi.endpoint)}
  async function hostApiFetch(){if(!C.allowToRetrieveFromApi)return{ok:false};const api=C.hostApi||{};if(!api.endpoint)return{ok:false};let r;try{r=await fetch(api.endpoint,{method:api.method||'GET',credentials:api.credentials||'same-origin',headers:api.headers||undefined})}catch(e){return{ok:false,error:String(e)}}if(!r.ok)return{ok:false,error:'HTTP '+r.status};let d;try{d=await r.json()}catch(_){return{ok:false,error:'Bad JSON.'}}let oid2,em;if(typeof api.mapper==='function'){const m=api.mapper(d)||{};oid2=m.openid;em=m.email}else{oid2=rp(d,api.openidPath);em=rp(d,api.emailPath)}if(!oid2)return{ok:false,error:'No openid.'};openid=String(oid2).trim();oidSrc='host-api';if(em)email=String(em).trim();return{ok:true}}
  let msP=null,msI=null;
  function loadMs(){if(window.msal&&typeof window.msal.PublicClientApplication==='function')return Promise.resolve(window.msal);if(msP)return msP;msP=new Promise((r,j)=>{const s=document.createElement('script');s.src=C.msal.scriptUrl;s.async=true;s.onload=()=>window.msal?r(window.msal):j(new Error('!'));s.onerror=()=>j(new Error('msal load failed'));document.head.appendChild(s)});return msP}
  async function getMs(){if(msI)return msI;const{clientId,authority,redirectUri}=C.msal;if(!clientId||!authority)throw new Error('MSAL not configured.');const m=await loadMs();msI=new m.PublicClientApplication({auth:{clientId,authority,redirectUri:redirectUri||window.location.origin},cache:{cacheLocation:'sessionStorage'}});if(typeof msI.initialize==='function')await msI.initialize();return msI}
  async function silentAuth(){if(!C.allowMsalFallback)throw new Error('off');const i=await getMs();return exMs(await i.ssoSilent({scopes:C.msal.scopes,loginHint:C.msal.loginHint||undefined}),'msal:silent')}
  async function popupAuth(){if(!C.allowMsalPopupLogin)throw new Error('off');const i=await getMs();return exMs(await i.loginPopup({scopes:C.msal.scopes}),'msal:popup')}
  function exMs(r,src){const c2=(r&&r.idTokenClaims)||{},ac2=(r&&r.account)||{};const oid2=c2.oid||c2.sub||(r&&r.idToken?joid(r.idToken):null);const v2=oid2||(r&&r.idToken)||null;if(!v2)throw new Error('No oid.');openid=v2;oidSrc=src;email=c2.email||c2.preferred_username||c2.upn||ac2.username||null;return v2}
  probeHost();

  // ═══════════════════════════════════════════════════════════════
  // SSE STREAMING
  // ═══════════════════════════════════════════════════════════════
  function buildBody(msg){const b=Object.assign({},C.chatDefaults,{message:msg,conversation_id:convId,openid,email});for(const k of Object.keys(b)){if(b[k]==null||b[k]==='')delete b[k]}return b}

  async function sseStream(msg, H) {
    let r;
    try { r = await fetch(C.backendUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Accept': 'text/event-stream' }, body: JSON.stringify(buildBody(msg)) }); }
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
      case 'tool_call_start': H.tool('\u2192 ' + (d.tool_name || '?')); break;
      case 'tool_result': { const p = typeof d.result === 'string' ? (d.result.length > 200 ? d.result.slice(0, 200) + '\u2026' : d.result) : JSON.stringify(d.result); H.tool((d.is_error ? '\u2717 ' : '\u2190 ') + p); } break;
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
      tool: (label) => { if (C.showToolEvents) addToolLine(label); },
      brk: () => { finishToolGroup(); addBrk(); },
      end: () => finish(),
      err: (e) => {
        if (!gotContent) { removeTyping(); respEl = addMsg('assistant', ''); gotContent = true; }
        appTo(respEl, (respEl._rawText ? '\n' : '') + '[error] ' + (e.message || e));
        finish();
      },
      auth: (p) => {
        removeTyping();
        if (!gotContent) { respEl = addMsg('assistant', ''); gotContent = true; }
        if (haEnabled()) { handleHostAuth(respEl, finish); return; }
        const mc = respEl.querySelector('.nx-mc');
        if (mc) mc.innerHTML = md(p.message || 'Please sign in to continue.');
        renderAuth(respEl, p);
        finish();
      },
    });
  }

  async function handleHostAuth(el, fin) {
    appTo(el, 'Verifying\u2026');
    const r = await hostApiFetch(); fin();
    if (r.ok) { if (C.autoRetryAfterAuth && lastQ) { appTo(el, ' \u2713'); setTimeout(() => sendMsg(lastQ), 150); } else appTo(el, ' \u2713 ask again.'); }
    else appTo(el, ' [error] ' + (r.error || 'failed'));
  }

  function renderAuth(el, p) {
    const sOk = C.allowMsalFallback && C.msal.clientId && C.msal.authority;
    const pOk = C.allowMsalPopupLogin && C.msal.clientId && C.msal.authority;
    const lk = document.createElement('span'); lk.className = 'nx-al'; lk.setAttribute('role', 'button'); lk.textContent = p.label || 'Sign in with BU Login';
    const nt = document.createElement('span'); nt.className = 'nx-an'; nt.textContent = sOk ? 'Will try silent first.' : pOk ? 'Opens a popup.' : 'MSAL not configured.';
    el.appendChild(document.createElement('br')); el.appendChild(lk); el.appendChild(nt);
    let ph = sOk && pOk ? 'sf' : sOk ? 'so' : pOk ? 'po' : 'x';
    const go = async () => {
      if (lk.classList.contains('nx-bsy')) return;
      lk.classList.add('nx-bsy');
      if (ph === 'sf' || ph === 'so') {
        lk.textContent = 'Signing in\u2026';
        try { await silentAuth(); lk.textContent = 'Signed in \u2713'; if (C.autoRetryAfterAuth && lastQ) setTimeout(() => sendMsg(lastQ), 150); return; }
        catch (_) { if (ph === 'sf') { ph = 'po'; lk.classList.remove('nx-bsy'); lk.textContent = 'Sign in with BU Login'; nt.textContent = 'Silent failed. Click for popup.'; return; } lk.classList.remove('nx-bsy'); lk.textContent = 'Retry'; return; }
      }
      if (ph === 'po') {
        lk.textContent = 'Opening\u2026';
        try { await popupAuth(); lk.textContent = 'Signed in \u2713'; if (C.autoRetryAfterAuth && lastQ) setTimeout(() => sendMsg(lastQ), 150); }
        catch (e) { lk.classList.remove('nx-bsy'); lk.textContent = 'Sign in with BU Login'; nt.textContent = 'Failed: ' + (e.message || e); }
      }
    };
    lk.addEventListener('click', go);
  }

  // ═══════════════════════════════════════════════════════════════
  // EVENT WIRING
  // ═══════════════════════════════════════════════════════════════
  snb.addEventListener('click', () => sendMsg());
  inp.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMsg(); } });
  inp.addEventListener('input', () => { inp.style.height = '44px'; inp.style.height = Math.min(inp.scrollHeight, 120) + 'px'; });
  if (launcher) launcher.addEventListener('click', toggle);

  // ── Voice dictation (Web Speech API) ──────────────────────
  // Commented out — requires HTTPS + working speech service.
  // Uncomment when deploying to production with valid certs.
  // See: https://developer.mozilla.org/en-US/docs/Web/API/SpeechRecognition
  /*
  const micBtn = root.querySelector('.nx-mic');
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

  if (!SpeechRecognition) {
    micBtn.classList.add('nx-mic-hide');
  } else {
    let recognition = null;
    let isListening = false;

    micBtn.addEventListener('click', () => {
      if (isListening) { recognition.stop(); return; }

      recognition = new SpeechRecognition();
      recognition.continuous = true;
      recognition.interimResults = true;
      recognition.lang = 'en-US';

      const beforeText = inp.value;
      let finalTranscript = '', hadFatalError = false;

      recognition.onstart = () => { isListening = true; micBtn.classList.add('nx-mic-on'); inp.placeholder = 'Listening\u2026'; };
      recognition.onresult = (event) => {
        let interim = ''; finalTranscript = '';
        for (let i = 0; i < event.results.length; i++) {
          if (event.results[i].isFinal) finalTranscript += event.results[i][0].transcript;
          else interim += event.results[i][0].transcript;
        }
        inp.value = beforeText + finalTranscript + interim;
        inp.style.height = '44px'; inp.style.height = Math.min(inp.scrollHeight, 120) + 'px';
      };
      recognition.onend = () => {
        if (isListening && !hadFatalError) { try { recognition.start(); return; } catch (_) {} }
        isListening = false; micBtn.classList.remove('nx-mic-on'); inp.placeholder = C.placeholder; inp.focus();
      };
      recognition.onerror = (event) => {
        if (['not-allowed','service-not-allowed','network','audio-capture'].includes(event.error)) {
          hadFatalError = true; isListening = false; micBtn.classList.remove('nx-mic-on');
          inp.placeholder = event.error === 'network' ? 'Voice requires HTTPS' : 'Voice unavailable';
          setTimeout(() => { inp.placeholder = C.placeholder; }, 3000);
          if (event.error === 'network') micBtn.classList.add('nx-mic-hide');
        }
      };
      recognition.start();
    });
  }
  */

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
    getOpenid() { return openid },
    getOpenidSource() { return oidSrc },
    getEmail() { return email },
    setOpenid(v, o) { if (!v) { openid = null; oidSrc = null; email = null; return; } openid = v; if (o && typeof o === 'object') { oidSrc = o.source || 'host'; if (o.email) email = o.email; } else oidSrc = o || 'host'; },
    setEmail(e2) { email = e2 || null; },
    refreshOpenidFromHost() { probeHost(); return openid; },
    signInSilent: silentAuth,
    signInPopup: popupAuth,
    async signIn() { if (C.allowMsalFallback && C.msal.clientId) { try { return await silentAuth(); } catch (_) {} } return popupAuth(); },
  };
  try { window.dispatchEvent(new CustomEvent('terrier-chat-v2-ready')); } catch (_) {}
})();
