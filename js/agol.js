// ═══════════════════════════════════════════════════════════════════════════
//  GAIA — ArcGIS Online Integration  (agol.js)
//
//  OAuth 2.0 Implicit Flow (no server required, safe for static hosting)
//
//  SETUP CHECKLIST (org admin):
//    1. Go to https://www.arcgis.com/home/organization.html → App Registration
//    2. Register new app → Redirect URI = your GitHub Pages URL
//       e.g.  https://your-org.github.io/gaia/
//    3. Copy the Client ID and paste it in gaia-config.js  (see below)
//    4. Ensure "Allow implicit grant" is ticked on the registration
//
//  This file handles:
//    - Sign-in / token exchange
//    - Token refresh / expiry detection
//    - Content search (own items + org items)
//    - Folder browsing
//    - Group browsing
//    - Layer loading (FeatureService → GeoJSON, MapService → tiles)
//    - Sign-out / token cleanup
// ═══════════════════════════════════════════════════════════════════════════

// ── Configuration ──────────────────────────────────────────────────────────
// Loaded from gaia-config.js.  Falls back to placeholder so the app
// still loads even without a config file; sign-in will fail gracefully.
const AGOL_CONFIG = (typeof GAIA_CONFIG !== 'undefined' && GAIA_CONFIG.agol)
  ? GAIA_CONFIG.agol
  : {
      clientId:   'YOUR_CLIENT_ID',         // set in gaia-config.js
      portalUrl:  'https://www.arcgis.com', // or your org portal URL
      redirectUri: window.location.href.split('?')[0].split('#')[0],
    };

// ── Module state ────────────────────────────────────────────────────────────
const agol = {
  token:      null,   // current OAuth token string
  expires:    null,   // token expiry timestamp (ms since epoch)
  username:   null,   // AGOL username
  orgId:      null,   // AGOL org ID
  fullName:   null,   // display name
  orgName:    null,   // org display name
  searchQuery:'',
  searchStart: 1,
  searchTotal: 0,
  pageSize:   24,
  currentFolder: null,  // null = root, string = folder id
  currentScope: 'mine', // 'mine' | 'org' | 'groups'
  groups: [],
  currentGroup: null,
};

// ─────────────────────────────────────────────────────────────────────────
//  TOKEN MANAGEMENT
// ─────────────────────────────────────────────────────────────────────────

/** Kick off the OAuth implicit-flow redirect */
function agolSignIn() {
  if (!AGOL_CONFIG.clientId || AGOL_CONFIG.clientId === 'YOUR_CLIENT_ID') {
    agolShowError(
      'ArcGIS Online is not yet configured.\n\n' +
      'Please set your Client ID in gaia-config.js and reload Gaia.\n' +
      'See js/agol.js for the setup checklist.'
    );
    return;
  }
  const params = new URLSearchParams({
    client_id:     AGOL_CONFIG.clientId,
    response_type: 'token',
    redirect_uri:  AGOL_CONFIG.redirectUri,
    expiration:    120,   // token lifetime in minutes (max 21600 = 15 days for named users)
  });
  const authUrl = AGOL_CONFIG.portalUrl.replace(/\/+$/, '') +
    '/sharing/rest/oauth2/authorize?' + params.toString();
  window.location.href = authUrl;
}

/** Sign out — clear token and update UI */
function agolSignOut() {
  agol.token    = null;
  agol.expires  = null;
  agol.username = null;
  agol.orgId    = null;
  agol.fullName = null;
  agol.orgName  = null;
  try { sessionStorage.removeItem('gaia_agol_token'); } catch(e) {}
  _agolUpdateUI();
  toast('Signed out of ArcGIS Online', 'info');
}

/** Called once on page load — picks up token from hash or sessionStorage */
function agolInit() {
  // 1. Check URL hash (post-OAuth redirect)
  const hash = window.location.hash;
  if (hash && hash.includes('access_token')) {
    const params = new URLSearchParams(hash.replace('#', ''));
    const token   = params.get('access_token');
    const expires = params.get('expires_in');
    if (token) {
      agol.token   = token;
      agol.expires = Date.now() + parseInt(expires || 7200) * 1000;
      // Persist in sessionStorage so reload keeps the session
      try {
        sessionStorage.setItem('gaia_agol_token', JSON.stringify({
          token:   agol.token,
          expires: agol.expires,
        }));
      } catch(e) {}
      // Clean hash from URL so it doesn't leak or confuse the app
      history.replaceState(null, '', window.location.pathname + window.location.search);
      // Fetch user info in the background
      _agolFetchSelf().then(() => _agolUpdateUI());
      return;
    }
  }
  // 2. Restore from sessionStorage
  try {
    const stored = sessionStorage.getItem('gaia_agol_token');
    if (stored) {
      const data = JSON.parse(stored);
      if (data.expires > Date.now() + 60000) {  // still valid for >1 min
        agol.token   = data.token;
        agol.expires = data.expires;
        _agolFetchSelf().then(() => _agolUpdateUI());
        return;
      }
    }
  } catch(e) {}
  _agolUpdateUI();
}

/** Fetch /sharing/rest/community/self to get username, org etc. */
async function _agolFetchSelf() {
  if (!agol.token) return;
  try {
    const url = AGOL_CONFIG.portalUrl.replace(/\/+$/, '') +
      '/sharing/rest/community/self?f=json&token=' + agol.token;
    const resp = await fetch(url);
    const data = await resp.json();
    if (data.error) throw new Error(data.error.message);
    agol.username = data.username;
    agol.orgId    = data.orgId;
    agol.fullName = data.fullName || data.username;
    // fetch org info for display name
    if (data.orgId) {
      const orgResp = await fetch(
        AGOL_CONFIG.portalUrl.replace(/\/+$/, '') +
        '/sharing/rest/portals/' + data.orgId + '?f=json&token=' + agol.token
      );
      const orgData = await orgResp.json();
      agol.orgName = orgData.name || '';
    }
  } catch(e) {
    console.warn('AGOL self-fetch failed:', e.message);
    // Don't sign out — token might still be valid for API calls
  }
}

// ─────────────────────────────────────────────────────────────────────────
//  UI RENDERING
// ─────────────────────────────────────────────────────────────────────────

function _agolUpdateUI() {
  const pane = document.getElementById('agol-pane');
  if (!pane) return;

  if (!agol.token) {
    pane.innerHTML = _agolRenderSignIn();
    return;
  }
  // Signed in — render browser
  pane.innerHTML = _agolRenderBrowser();
  // Auto-load content
  agolBrowse('mine');
}

function _agolRenderSignIn() {
  const configured = AGOL_CONFIG.clientId && AGOL_CONFIG.clientId !== 'YOUR_CLIENT_ID';
  return `
    <div style="padding:24px 16px;text-align:center;">
      <div style="font-size:32px;margin-bottom:12px;">🌐</div>
      <div style="font-family:var(--mono);font-size:12px;font-weight:600;color:var(--text);margin-bottom:6px;">
        ArcGIS Online
      </div>
      <div style="font-family:var(--mono);font-size:10px;color:var(--text3);margin-bottom:16px;line-height:1.6;">
        ${configured
          ? 'Sign in with your ArcGIS Online or SAML credentials to browse and load layers.'
          : '⚠ Not configured — set your Client ID in <code>gaia-config.js</code>'
        }
      </div>
      ${configured ? `
        <button class="btn btn-primary" onclick="agolSignIn()"
          style="padding:8px 24px;font-family:var(--mono);font-size:11px;">
          Sign In to ArcGIS Online
        </button>
        <div style="font-family:var(--mono);font-size:9px;color:var(--text3);margin-top:10px;line-height:1.6;">
          Your organisation's SAML login will open in this tab.<br/>
          Gaia never sees your password.
        </div>` : `
        <a href="js/agol.js" target="_blank"
          style="font-family:var(--mono);font-size:10px;color:var(--teal);">
          View setup instructions →
        </a>`
      }
    </div>`;
}

function _agolRenderBrowser() {
  const expiryMins = agol.expires
    ? Math.max(0, Math.round((agol.expires - Date.now()) / 60000))
    : null;
  const expiryStr = expiryMins !== null
    ? (expiryMins > 60 ? `~${Math.round(expiryMins/60)}h` : `${expiryMins}m`)
    : '?';

  return `
    <!-- User bar -->
    <div style="display:flex;align-items:center;gap:8px;padding:6px 10px;background:var(--bg3);
                border-bottom:1px solid var(--border);font-family:var(--mono);font-size:9px;">
      <span style="color:var(--teal);">●</span>
      <div style="flex:1;overflow:hidden;">
        <div style="color:var(--text);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
          ${escHtml(agol.fullName || agol.username || 'Signed in')}
        </div>
        <div style="color:var(--text3);">
          ${escHtml(agol.orgName || AGOL_CONFIG.portalUrl)} · token ~${expiryStr}
        </div>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="agolSignOut()"
        style="font-size:9px;padding:2px 7px;flex-shrink:0;">Sign Out</button>
    </div>

    <!-- Scope tabs -->
    <div style="display:flex;border-bottom:1px solid var(--border);background:var(--bg2);">
      ${['mine','org','groups'].map(s => `
        <div onclick="agolBrowse('${s}')" id="agol-scope-${s}"
          style="padding:7px 12px;font-family:var(--mono);font-size:9px;cursor:pointer;
                 border-bottom:2px solid ${agol.currentScope===s?'var(--teal)':'transparent'};
                 color:${agol.currentScope===s?'var(--teal)':'var(--text3)'};">
          ${s === 'mine' ? '👤 My Content' : s === 'org' ? '🏢 Organisation' : '👥 Groups'}
        </div>`).join('')}
    </div>

    <!-- Search bar -->
    <div style="padding:6px 8px;border-bottom:1px solid var(--border);display:flex;gap:6px;">
      <input type="text" id="agol-search-input" value="${escHtml(agol.searchQuery)}"
        placeholder="Search layers…"
        style="flex:1;padding:4px 8px;font-family:var(--mono);font-size:10px;"
        onkeydown="if(event.key==='Enter')agolSearch()"/>
      <button class="btn btn-ghost btn-sm" onclick="agolSearch()" style="font-size:10px;padding:3px 8px;">🔍</button>
      ${agol.searchQuery ? `<button class="btn btn-ghost btn-sm" onclick="agolClearSearch()" style="font-size:10px;padding:3px 8px;">✕</button>` : ''}
    </div>

    <!-- Group selector (shown when scope = groups) -->
    <div id="agol-group-row" style="display:${agol.currentScope==='groups'?'block':'none'};
      padding:6px 8px;border-bottom:1px solid var(--border);">
      <select id="agol-group-select" onchange="agolSelectGroup(this.value)"
        style="width:100%;padding:4px 8px;font-family:var(--mono);font-size:10px;">
        <option value="">— loading groups…</option>
      </select>
    </div>

    <!-- Folder breadcrumb (shown when scope = mine) -->
    <div id="agol-folder-row" style="display:${agol.currentScope==='mine'?'flex':'none'};
      align-items:center;gap:6px;padding:5px 10px;border-bottom:1px solid var(--border);
      font-family:var(--mono);font-size:9px;color:var(--text3);">
      <span onclick="agolGoHome()" style="cursor:pointer;color:var(--teal);">🏠 My Content</span>
      <span id="agol-folder-name"></span>
    </div>

    <!-- Results -->
    <div id="agol-results" style="overflow-y:auto;flex:1;max-height:340px;">
      <div style="padding:16px;text-align:center;font-family:var(--mono);font-size:10px;color:var(--text3);">
        Loading…
      </div>
    </div>

    <!-- Pagination -->
    <div id="agol-pagination" style="display:none;padding:6px 10px;border-top:1px solid var(--border);
      display:flex;align-items:center;gap:8px;font-family:var(--mono);font-size:9px;color:var(--text3);">
    </div>
    `;
}

// ─────────────────────────────────────────────────────────────────────────
//  CONTENT BROWSING
// ─────────────────────────────────────────────────────────────────────────

async function agolBrowse(scope) {
  agol.currentScope = scope;
  agol.searchStart  = 1;
  agol.currentFolder = null;
  const pane = document.getElementById('agol-pane');
  if (!pane) return;
  pane.innerHTML = _agolRenderBrowser();

  if (scope === 'groups') {
    await _agolLoadGroups();
    return;
  }
  await _agolFetchContent();
}

async function _agolLoadGroups() {
  const groupSel = document.getElementById('agol-group-select');
  if (!groupSel) return;

  try {
    const url = AGOL_CONFIG.portalUrl.replace(/\/+$/, '') +
      '/sharing/rest/community/self?f=json&token=' + agol.token;
    const resp = await fetch(url);
    const data = await resp.json();
    agol.groups = data.groups || [];

    if (!agol.groups.length) {
      groupSel.innerHTML = '<option value="">No groups found</option>';
      return;
    }
    groupSel.innerHTML = '<option value="">— select a group —</option>' +
      agol.groups.map(g =>
        `<option value="${escHtml(g.id)}">${escHtml(g.title)}</option>`
      ).join('');

    // Auto-select first group
    if (agol.groups[0]) {
      groupSel.value = agol.groups[0].id;
      agol.currentGroup = agol.groups[0].id;
      await _agolFetchContent();
    }
  } catch(e) {
    _agolShowResults(`<div style="padding:12px;font-family:var(--mono);font-size:10px;color:var(--red);">
      Failed to load groups: ${escHtml(e.message)}</div>`);
  }
}

async function agolSelectGroup(groupId) {
  agol.currentGroup = groupId;
  agol.searchStart  = 1;
  if (groupId) await _agolFetchContent();
}

function agolSearch() {
  const input = document.getElementById('agol-search-input');
  agol.searchQuery = input ? input.value.trim() : '';
  agol.searchStart = 1;
  _agolFetchContent();
}

function agolClearSearch() {
  agol.searchQuery = '';
  agol.searchStart = 1;
  _agolFetchContent();
}

function agolGoHome() {
  agol.currentFolder = null;
  agol.searchStart   = 1;
  _agolFetchContent();
  const nameEl = document.getElementById('agol-folder-name');
  if (nameEl) nameEl.textContent = '';
}

async function agolOpenFolder(folderId, folderName) {
  agol.currentFolder = folderId;
  agol.searchStart   = 1;
  const nameEl = document.getElementById('agol-folder-name');
  if (nameEl) nameEl.textContent = '› ' + folderName;
  await _agolFetchContent();
}

async function _agolFetchContent() {
  const resultsEl = document.getElementById('agol-results');
  if (!resultsEl) return;
  resultsEl.innerHTML =
    '<div style="padding:16px;text-align:center;font-family:var(--mono);font-size:10px;color:var(--text3);">Loading…</div>';

  try {
    const items = await _agolQueryItems();
    _agolRenderItems(items);
  } catch(e) {
    resultsEl.innerHTML =
      `<div style="padding:12px;font-family:var(--mono);font-size:10px;color:var(--red);">
        ${escHtml(e.message)}
      </div>`;
  }
}

async function _agolQueryItems() {
  const portal = AGOL_CONFIG.portalUrl.replace(/\/+$/, '');
  const layerTypes = [
    'Feature Service', 'Map Service', 'Vector Tile Service',
    'WMS', 'WFS', 'Feature Collection'
  ].map(t => `type:"${t}"`).join(' OR ');

  let q = `(${layerTypes})`;
  if (agol.searchQuery) q += ` AND (title:"${agol.searchQuery}" OR tags:"${agol.searchQuery}")`;

  let endpoint;
  const params = {
    f:        'json',
    token:    agol.token,
    num:      agol.pageSize,
    start:    agol.searchStart,
    sortField:'modified',
    sortOrder:'desc',
  };

  if (agol.currentScope === 'mine') {
    if (agol.currentFolder) {
      // Browse a specific folder — use user content endpoint
      endpoint = `${portal}/sharing/rest/content/users/${agol.username}/` +
        `${agol.currentFolder}/items`;
      // Folder endpoint doesn't use a search query
      delete params.num; delete params.start;
    } else {
      params.q = q + ` AND owner:${agol.username}`;
      endpoint = `${portal}/sharing/rest/search`;
    }
  } else if (agol.currentScope === 'org') {
    params.q = q + ` AND orgid:${agol.orgId}`;
    endpoint = `${portal}/sharing/rest/search`;
  } else if (agol.currentScope === 'groups') {
    if (!agol.currentGroup) return { results:[], total:0 };
    params.q = q;
    params.groups = agol.currentGroup;
    endpoint = `${portal}/sharing/rest/search`;
  }

  const resp  = await fetch(endpoint + '?' + new URLSearchParams(params));
  const data  = await resp.json();
  if (data.error) throw new Error(data.error.message || JSON.stringify(data.error));

  agol.searchTotal = data.total || (data.items ? data.items.length : 0);

  // Folder endpoint returns { items } directly; search returns { results }
  const rawItems = data.results || data.items || [];

  // If browsing root of "my content", also fetch folders
  let folders = [];
  if (agol.currentScope === 'mine' && !agol.currentFolder && !agol.searchQuery) {
    try {
      const fResp = await fetch(
        `${portal}/sharing/rest/content/users/${agol.username}?f=json&token=${agol.token}`
      );
      const fData = await fResp.json();
      folders = (fData.folders || []).map(f => ({ ...f, _isFolder: true }));
    } catch(e) { /* folders are optional */ }
  }

  return { results: [...folders, ...rawItems], total: agol.searchTotal };
}

function _agolRenderItems({ results, total }) {
  const resultsEl = document.getElementById('agol-results');
  const pagEl     = document.getElementById('agol-pagination');
  if (!resultsEl) return;

  if (!results.length) {
    resultsEl.innerHTML =
      '<div style="padding:16px;text-align:center;font-family:var(--mono);font-size:10px;color:var(--text3);">No layers found.</div>';
    if (pagEl) pagEl.style.display = 'none';
    return;
  }

  resultsEl.innerHTML = results.map(item => {
    if (item._isFolder) return _agolRenderFolder(item);
    return _agolRenderItem(item);
  }).join('');

  // Pagination
  if (pagEl) {
    const hasPrev = agol.searchStart > 1;
    const hasNext = agol.searchStart + agol.pageSize - 1 < total;
    if (hasPrev || hasNext) {
      pagEl.style.display = 'flex';
      const pageNum = Math.ceil(agol.searchStart / agol.pageSize);
      const totalPages = Math.ceil(total / agol.pageSize);
      pagEl.innerHTML =
        `${hasPrev ? `<button class="btn btn-ghost btn-sm" onclick="agolPage(-1)" style="font-size:9px;padding:2px 7px;">‹ Prev</button>` : '<span></span>'}
         <span style="flex:1;text-align:center;color:var(--text3);">
           Page ${pageNum} of ${totalPages} · ${total} items
         </span>
         ${hasNext ? `<button class="btn btn-ghost btn-sm" onclick="agolPage(1)" style="font-size:9px;padding:2px 7px;">Next ›</button>` : '<span></span>'}`;
    } else {
      pagEl.style.display = 'none';
    }
  }
}

function _agolRenderFolder(f) {
  return `
    <div onclick="agolOpenFolder('${escHtml(f.id)}','${escHtml(f.title)}')"
      style="display:flex;align-items:center;gap:10px;padding:8px 12px;
             border-bottom:1px solid var(--border);cursor:pointer;"
      onmouseover="this.style.background='var(--bg3)'"
      onmouseout="this.style.background=''">
      <span style="font-size:16px;flex-shrink:0;">📁</span>
      <div style="flex:1;min-width:0;">
        <div style="font-family:var(--mono);font-size:10px;font-weight:600;color:var(--text);
                    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
          ${escHtml(f.title)}
        </div>
        <div style="font-family:var(--mono);font-size:8px;color:var(--text3);">Folder</div>
      </div>
      <span style="font-family:var(--mono);font-size:10px;color:var(--text3);">›</span>
    </div>`;
}

function _agolRenderItem(item) {
  const typeIcon = _agolTypeIcon(item.type);
  const typeLabel = item.type || 'Unknown';
  const modified = item.modified
    ? new Date(item.modified).toLocaleDateString()
    : '';
  const snippet = (item.snippet || item.description || '').substring(0, 80);
  const canLoad = _agolCanLoad(item.type);

  return `
    <div style="display:flex;align-items:flex-start;gap:10px;padding:8px 12px;
                border-bottom:1px solid var(--border);"
      onmouseover="this.style.background='var(--bg3)'"
      onmouseout="this.style.background=''">
      <span style="font-size:18px;flex-shrink:0;margin-top:1px;">${typeIcon}</span>
      <div style="flex:1;min-width:0;">
        <div style="font-family:var(--mono);font-size:10px;font-weight:600;color:var(--text);
                    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"
             title="${escHtml(item.title)}">
          ${escHtml(item.title)}
        </div>
        <div style="font-family:var(--mono);font-size:8px;color:var(--text3);margin-top:1px;">
          ${escHtml(typeLabel)} ${modified ? '· ' + modified : ''} ${item.owner ? '· ' + escHtml(item.owner) : ''}
        </div>
        ${snippet ? `<div style="font-family:var(--mono);font-size:8px;color:var(--text3);margin-top:2px;
                         white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
          ${escHtml(snippet)}</div>` : ''}
      </div>
      ${canLoad ? `
        <button class="btn btn-ghost btn-sm"
          onclick="agolLoadItem('${escHtml(item.id)}','${escHtml(item.title).replace(/'/g,"\\'")}','${escHtml(item.type)}')"
          style="flex-shrink:0;font-size:9px;padding:3px 8px;white-space:nowrap;">
          + Add
        </button>` : `
        <span style="font-family:var(--mono);font-size:8px;color:var(--text3);flex-shrink:0;padding-top:3px;">
          Not supported
        </span>`
      }
    </div>`;
}

function agolPage(dir) {
  agol.searchStart = Math.max(1, agol.searchStart + dir * agol.pageSize);
  _agolFetchContent();
}

// ─────────────────────────────────────────────────────────────────────────
//  LAYER LOADING
// ─────────────────────────────────────────────────────────────────────────

async function agolLoadItem(itemId, title, itemType) {
  const portal = AGOL_CONFIG.portalUrl.replace(/\/+$/, '');
  toast('Loading "' + title + '"…', 'info');
  openURLModal(); // close the AGOL modal so user sees progress

  try {
    // Fetch item details to get the service URL
    const detailResp = await fetch(
      `${portal}/sharing/rest/content/items/${itemId}?f=json&token=${agol.token}`
    );
    const detail = await detailResp.json();
    if (detail.error) throw new Error(detail.error.message);

    const serviceUrl = detail.url;
    if (!serviceUrl) throw new Error('Item has no associated service URL');

    if (itemType === 'Feature Service') {
      await _agolLoadFeatureService(serviceUrl, title, itemId);
    } else if (itemType === 'Map Service') {
      _agolLoadMapServiceTiles(serviceUrl, title);
    } else if (itemType === 'Vector Tile Service') {
      _agolLoadVectorTiles(serviceUrl, title);
    } else if (itemType === 'WMS') {
      _agolLoadWMSFromItem(detail, title);
    } else if (itemType === 'Feature Collection') {
      await _agolLoadFeatureCollection(itemId, title);
    } else {
      throw new Error('Unsupported item type: ' + itemType);
    }
  } catch(e) {
    toast('AGOL Error: ' + e.message, 'error');
    console.error('AGOL load error:', e);
  }
}

/** Load all sub-layers from a FeatureService, or a specific layer if URL ends in /N */
async function _agolLoadFeatureService(serviceUrl, title, itemId) {
  const portal = AGOL_CONFIG.portalUrl.replace(/\/+$/, '');
  const cleanUrl = serviceUrl.replace(/\/+$/, '');

  // Does the URL already point at a specific layer?
  const layerMatch = cleanUrl.match(/\/(\d+)$/);
  if (layerMatch) {
    // Single layer
    await _agolDownloadFeatureLayer(cleanUrl, title);
    return;
  }

  // Root FeatureService — discover sub-layers
  const infoResp = await fetch(cleanUrl + '?f=json&token=' + agol.token);
  const info = await infoResp.json();
  if (info.error) throw new Error(info.error.message);

  const layers = info.layers || [];
  if (!layers.length) throw new Error('Feature Service has no layers');

  if (layers.length === 1) {
    // Only one layer — load it directly
    await _agolDownloadFeatureLayer(cleanUrl + '/' + layers[0].id, layers[0].name || title);
    return;
  }

  // Multiple layers — show picker
  _agolShowLayerPicker(cleanUrl, title, layers);
}

function _agolShowLayerPicker(serviceUrl, title, layers) {
  const modal = document.createElement('div');
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:10500;display:flex;align-items:center;justify-content:center;';
  modal.innerHTML = `
    <div style="background:var(--bg2);border-radius:10px;width:420px;max-width:95vw;
                box-shadow:0 8px 40px rgba(0,0,0,0.3);overflow:hidden;max-height:80vh;display:flex;flex-direction:column;">
      <div style="background:linear-gradient(135deg,#0C2E44,#113c64);border-bottom:2px solid #14b1e7;
                  padding:12px 16px;display:flex;align-items:center;justify-content:space-between;">
        <span style="font-family:var(--mono);font-weight:700;font-size:11px;color:#e8f4fb;">
          Select Layers — ${escHtml(title)}
        </span>
        <span style="cursor:pointer;color:#7a96aa;" onclick="this.closest('[style]').remove()">✕</span>
      </div>
      <div style="padding:10px 14px;overflow-y:auto;flex:1;">
        ${layers.map(l => `
          <label style="display:flex;align-items:center;gap:8px;padding:6px 4px;cursor:pointer;
                         font-family:var(--mono);font-size:10px;color:var(--text);"
            onmouseover="this.style.background='var(--bg3)'" onmouseout="this.style.background=''">
            <input type="checkbox" checked value="${l.id}" style="accent-color:var(--teal);"/>
            ${escHtml(l.name || 'Layer ' + l.id)}
            <span style="font-size:8px;color:var(--text3);margin-left:auto;">${escHtml(l.type||'')}</span>
          </label>`).join('')}
      </div>
      <div style="padding:10px 14px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px;">
        <button class="btn btn-ghost btn-sm" onclick="this.closest('[style]').remove()">Cancel</button>
        <button class="btn btn-primary btn-sm" onclick="agolLoadPickedLayers('${escHtml(serviceUrl)}',this)">
          Load Selected
        </button>
      </div>
    </div>`;
  document.body.appendChild(modal);
}

async function agolLoadPickedLayers(serviceUrl, btn) {
  const modal = btn.closest('[style]');
  const checked = Array.from(modal.querySelectorAll('input[type=checkbox]:checked'));
  if (!checked.length) { toast('Select at least one layer', 'error'); return; }
  modal.remove();
  for (const cb of checked) {
    const layerUrl = serviceUrl + '/' + cb.value;
    const label = cb.closest('label');
    const name = label ? label.textContent.trim() : 'Layer ' + cb.value;
    await _agolDownloadFeatureLayer(layerUrl, name);
  }
}

async function _agolDownloadFeatureLayer(layerUrl, name) {
  // Zoom warning
  if (state.map) {
    const zoom = state.map.getZoom();
    if (zoom < 10) {
      const ok = confirm(
        `You are at zoom level ${zoom}.\n\nAt this scale, many features may be returned. ` +
        `Only the first 2,000 features will be downloaded.\n\nContinue?`
      );
      if (!ok) return;
    }
  }

  toast('Downloading features from ' + name + '…', 'info');

  const MAX = 2000;
  const params = {
    where: '1=1',
    outFields: '*',
    outSR: '4326',
    f: 'geojson',
    resultRecordCount: MAX,
    returnGeometry: 'true',
    token: agol.token,
  };

  // Clip to current extent if zoomed in
  if (state.map && state.map.getZoom() >= 10) {
    const b = state.map.getBounds();
    params.geometry = [b.getWest(), b.getSouth(), b.getEast(), b.getNorth()].join(',');
    params.geometryType = 'esriGeometryEnvelope';
    params.inSR = '4326';
    params.spatialRel = 'esriSpatialRelIntersects';
  }

  const resp = await fetch(layerUrl + '/query?' + new URLSearchParams(params));
  const geojson = await resp.json();
  if (geojson.error) throw new Error(geojson.error.message);
  if (!geojson.features) throw new Error('No features returned');

  const n = geojson.features.length;
  addLayer(geojson, name, 'EPSG:4326', 'ArcGIS Online');
  toast(`"${name}" loaded — ${n} feature${n!==1?'s':''}${n>=MAX?' (limit reached)':''}`, 'success');
}

function _agolLoadMapServiceTiles(serviceUrl, name) {
  const tileUrl = serviceUrl.replace(/\/+$/, '') + '/tile/{z}/{y}/{x}';
  const tileL = L.tileLayer(tileUrl, { attribution: name, opacity: 0.85, token: agol.token });
  tileL.addTo(state.map);
  state.layers.push({
    name, format:'ArcGIS Map Service', color:'#f0883e',
    leafletLayer:tileL, visible:true, isTile:true,
    fields:{}, geojson:{features:[]}, geomType:'Tile', sourceCRS:'EPSG:4326',
  });
  updateLayerList(); updateExportLayerList();
  toast('Map Service tiles added: ' + name, 'success');
}

function _agolLoadVectorTiles(serviceUrl, name) {
  // Leaflet doesn't natively support vector tiles — add as ArcGIS REST tile overlay
  const tileUrl = serviceUrl.replace(/\/+$/, '') + '/tile/{z}/{y}/{x}.pbf';
  toast('"' + name + '" is a Vector Tile Service. Adding as a visual tile overlay (no attribute data).', 'info');
  const tileL = L.tileLayer(tileUrl, { attribution: name, opacity: 0.85 });
  tileL.addTo(state.map);
  state.layers.push({
    name, format:'Vector Tile Service', color:'#bc8cff',
    leafletLayer:tileL, visible:true, isTile:true,
    fields:{}, geojson:{features:[]}, geomType:'Tile', sourceCRS:'EPSG:4326',
  });
  updateLayerList(); updateExportLayerList();
  toast('Vector tile layer added: ' + name, 'success');
}

function _agolLoadWMSFromItem(detail, name) {
  const url = detail.url;
  const wmsL = L.tileLayer.wms(url, {
    layers: detail.servicePropertiesJson ? JSON.parse(detail.servicePropertiesJson).currentVersion : '',
    format: 'image/png', transparent: true, attribution: name, opacity: 0.85,
  });
  wmsL.addTo(state.map);
  state.layers.push({
    name, format:'WMS (AGOL)', color:'#5ab4f0',
    leafletLayer:wmsL, visible:true, isTile:true,
    fields:{}, geojson:{features:[]}, geomType:'Tile', sourceCRS:'EPSG:4326',
  });
  updateLayerList(); updateExportLayerList();
  toast('WMS layer added: ' + name, 'success');
}

async function _agolLoadFeatureCollection(itemId, title) {
  const portal = AGOL_CONFIG.portalUrl.replace(/\/+$/, '');
  const dataResp = await fetch(
    `${portal}/sharing/rest/content/items/${itemId}/data?f=json&token=${agol.token}`
  );
  const data = await dataResp.json();
  if (data.error) throw new Error(data.error.message);

  // Feature Collection has a "layers" array
  const layers = data.layers || [];
  for (const l of layers) {
    const fc = l.featureSet;
    if (!fc || !fc.features) continue;
    // Convert Esri geometry format to GeoJSON
    const geojson = _esriToGeoJSON(fc);
    if (geojson.features.length) {
      addLayer(geojson, l.layerDefinition?.name || title, 'EPSG:4326', 'Feature Collection');
    }
  }
  toast('Feature Collection loaded: ' + title, 'success');
}

/** Minimal Esri FeatureSet → GeoJSON converter */
function _esriToGeoJSON(featureSet) {
  const features = (featureSet.features || []).map(f => ({
    type: 'Feature',
    properties: f.attributes || {},
    geometry: _esriGeomToGeoJSON(featureSet.geometryType, f.geometry),
  })).filter(f => f.geometry);
  return { type: 'FeatureCollection', features };
}

function _esriGeomToGeoJSON(geomType, geom) {
  if (!geom) return null;
  switch (geomType) {
    case 'esriGeometryPoint':
      return { type:'Point', coordinates:[geom.x, geom.y] };
    case 'esriGeometryMultipoint':
      return { type:'MultiPoint', coordinates:geom.points||[] };
    case 'esriGeometryPolyline':
      return geom.paths && geom.paths.length === 1
        ? { type:'LineString', coordinates:geom.paths[0] }
        : { type:'MultiLineString', coordinates:geom.paths||[] };
    case 'esriGeometryPolygon':
      return geom.rings && geom.rings.length === 1
        ? { type:'Polygon', coordinates:[geom.rings[0]] }
        : { type:'MultiPolygon', coordinates:(geom.rings||[]).map(r=>[r]) };
    default: return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────

function _agolTypeIcon(type) {
  const icons = {
    'Feature Service':      '📍',
    'Map Service':          '🗺',
    'Vector Tile Service':  '🔷',
    'WMS':                  '🌐',
    'WFS':                  '🌐',
    'Feature Collection':   '📦',
    'Folder':               '📁',
  };
  return icons[type] || '📄';
}

function _agolCanLoad(type) {
  return ['Feature Service','Map Service','Vector Tile Service',
          'WMS','Feature Collection'].includes(type);
}

function _agolShowResults(html) {
  const el = document.getElementById('agol-results');
  if (el) el.innerHTML = html;
}

function agolShowError(msg) {
  // eslint-disable-next-line no-alert
  alert(msg);
}

// ─────────────────────────────────────────────────────────────────────────
//  Bootstrap on page load
// ─────────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', agolInit);
