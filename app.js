import { PublicClientApplication } from 'https://esm.sh/@azure/msal-browser@4.30.0';

// ── Config ────────────────────────────────────────────────────────────────────

const CLIENT_ID   = '31c0c6ff-5580-41c8-9753-b94d991cfc66';
const AUTHORITY   = 'https://login.microsoftonline.com/consumers';
const REDIRECT_URI = 'https://flyinthissky.github.io/project-dashboard';
const SCOPES      = ['Files.ReadWrite.AppFolder'];
const FILE_URL    = 'https://graph.microsoft.com/v1.0/me/drive/special/approot:/projects.json:/content';

// ── MSAL ──────────────────────────────────────────────────────────────────────

let msalInstance = null;

async function initMsal() {
  msalInstance = new PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: AUTHORITY, redirectUri: REDIRECT_URI },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
  });
  await msalInstance.initialize();
  await msalInstance.handleRedirectPromise();
}

async function getAccessToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error('Non connecté');
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ scopes: SCOPES });
    return r.accessToken;
  }
}

async function login() {
  clearLoginError();
  try {
    const r = await msalInstance.loginPopup({ scopes: SCOPES });
    onAuthenticated(r.account);
  } catch (err) {
    showLoginError(err.message || String(err));
  }
}

async function logout() {
  try { await msalInstance.logoutPopup(); } catch { /* ignore */ }
  showLogin();
}

// ── Graph helpers ─────────────────────────────────────────────────────────────

async function graphFetchWithRetry(url, options = {}, attempt = 0) {
  const token = await getAccessToken();
  const res = await fetch(url, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...(options.headers || {}) },
  });
  if (res.status === 429 && attempt < 4) {
    const wait = Math.pow(2, attempt) * 1000;
    await new Promise(r => setTimeout(r, wait));
    return graphFetchWithRetry(url, options, attempt + 1);
  }
  return res;
}

async function loadProjects() {
  setSyncState('syncing');
  const res = await graphFetchWithRetry(FILE_URL);
  if (res.status === 404) { setSyncState('ok'); return { version: 1, projects: [] }; }
  if (!res.ok) throw new Error(`Erreur lecture : HTTP ${res.status}`);
  const data = await res.json();
  setSyncState('ok');
  return data;
}

async function saveProjects(projects) {
  setSyncState('syncing');
  const body = JSON.stringify({ version: 1, projects });
  const res = await graphFetchWithRetry(FILE_URL, { method: 'PUT', body });
  if (!res.ok) throw new Error(`Erreur sauvegarde : HTTP ${res.status}`);
  setSyncState('ok');
}

// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  projects: [],
  search: '',
  filterStatus: 'all',
  filterPriority: 'all',
  filterTag: 'all',
  sort: 'updated',
};

let saveTimer = null;

function setState(patch) {
  Object.assign(state, patch);
  render();
}

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    try {
      await saveProjects(state.projects);
    } catch (err) {
      showToast(err.message, 'error');
      setSyncState('error');
    }
  }, 1000);
}

function upsertProject(proj) {
  const idx = state.projects.findIndex(p => p.id === proj.id);
  const projects = [...state.projects];
  if (idx >= 0) projects[idx] = proj;
  else projects.unshift(proj);
  setState({ projects });
  scheduleSave();
}

function deleteProject(id) {
  setState({ projects: state.projects.filter(p => p.id !== id) });
  scheduleSave();
}

function uuid() {
  return ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, c =>
    (c ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))).toString(16)
  );
}

// ── Status helpers ────────────────────────────────────────────────────────────

const STATUS_META = {
  active:    { emoji: '🟢', label: 'Actif',     color: 'var(--status-active)' },
  paused:    { emoji: '🟡', label: 'En pause',  color: 'var(--status-paused)' },
  abandoned: { emoji: '🔴', label: 'Abandonné', color: 'var(--status-abandoned)' },
  done:      { emoji: '✅', label: 'Terminé',   color: 'var(--status-done)' },
};

const PRIORITY_LABEL = { low: 'Basse', medium: 'Moyenne', high: 'Haute' };

function relativeDate(iso) {
  if (!iso) return '';
  const diff = Date.now() - new Date(iso).getTime();
  const d = Math.floor(diff / 86400000);
  if (d === 0) return "Aujourd'hui";
  if (d === 1) return 'Hier';
  if (d < 7)  return `Il y a ${d} j`;
  if (d < 30) return `Il y a ${Math.floor(d / 7)} sem`;
  if (d < 365) return `Il y a ${Math.floor(d / 30)} mois`;
  return `Il y a ${Math.floor(d / 365)} an${Math.floor(d / 365) > 1 ? 's' : ''}`;
}

// ── Sync indicator ────────────────────────────────────────────────────────────

function setSyncState(s) {
  const el = document.getElementById('sync-indicator');
  if (!el) return;
  el.className = `sync-indicator ${s === 'syncing' ? 'is-syncing' : s === 'error' ? 'is-error' : ''}`;
  const dot = el.querySelector('.dot');
  const txt = el.querySelector('.sync-text');
  if (dot && txt) {
    txt.textContent = s === 'syncing' ? 'Sync…' : s === 'error' ? 'Erreur' : 'Synced';
  }
}

// ── Render: toolbar counts ────────────────────────────────────────────────────

function getFilteredProjects() {
  let list = [...state.projects];
  if (state.search) {
    const q = state.search.toLowerCase();
    list = list.filter(p =>
      p.name.toLowerCase().includes(q) ||
      (p.description || '').toLowerCase().includes(q) ||
      (p.tech || '').toLowerCase().includes(q) ||
      (p.tags || []).some(t => t.toLowerCase().includes(q))
    );
  }
  if (state.filterStatus !== 'all') list = list.filter(p => p.status === state.filterStatus);
  if (state.filterPriority !== 'all') list = list.filter(p => p.priority === state.filterPriority);
  if (state.filterTag !== 'all') list = list.filter(p => (p.tags || []).includes(state.filterTag));

  list.sort((a, b) => {
    if (state.sort === 'name') return a.name.localeCompare(b.name);
    if (state.sort === 'progress') return (b.progress || 0) - (a.progress || 0);
    if (state.sort === 'created') return new Date(b.createdAt) - new Date(a.createdAt);
    return new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt);
  });
  return list;
}

function getAllTags() {
  const tags = new Set();
  state.projects.forEach(p => (p.tags || []).forEach(t => tags.add(t)));
  return [...tags].sort();
}

// ── Render: project list ──────────────────────────────────────────────────────

function renderProjects() {
  const container = document.getElementById('projects-list');
  const meta = document.getElementById('toolbar-meta');
  if (!container) return;

  const list = getFilteredProjects();
  if (meta) meta.textContent = `${list.length} / ${state.projects.length} projet${state.projects.length !== 1 ? 's' : ''}`;

  if (!list.length) {
    container.innerHTML = `
      <div class="empty-state">
        <h2>${state.projects.length === 0 ? 'Aucun projet pour l\'instant' : 'Aucun résultat'}</h2>
        <p>${state.projects.length === 0 ? 'Crée ton premier projet avec le bouton ci-dessus.' : 'Essaie de modifier les filtres ou la recherche.'}</p>
      </div>`;
    return;
  }

  container.innerHTML = list.map(p => {
    const sm = STATUS_META[p.status] || STATUS_META.active;
    const tags = (p.tags || []).map(t => `<span class="tag-chip">${escHtml(t)}</span>`).join('');
    const prog = Math.max(0, Math.min(100, p.progress || 0));
    return `
    <article class="project-card" data-id="${escHtml(p.id)}" role="button" tabindex="0" aria-label="Projet ${escHtml(p.name)}">
      <div class="project-card-head">
        <div class="project-card-name">
          <span class="project-status" title="${escHtml(sm.label)}">${sm.emoji}</span>
          <span class="name-text">${escHtml(p.name)}</span>
        </div>
        <div style="display:flex;align-items:center;gap:0.4rem">
          <span class="priority-pill priority-${escHtml(p.priority || 'low')}">${escHtml(PRIORITY_LABEL[p.priority] || 'Basse')}</span>
        </div>
      </div>
      ${p.description ? `<p style="font-size:0.82rem;color:var(--text-secondary);line-height:1.45;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical">${escHtml(p.description)}</p>` : ''}
      <div class="project-card-meta">
        ${p.tech ? `<span class="tech">${escHtml(p.tech)}</span>` : ''}
        <span>${relativeDate(p.updatedAt || p.createdAt)}</span>
        ${p.githubUrl ? `<a href="${escHtml(p.githubUrl)}" target="_blank" rel="noopener" onclick="event.stopPropagation()">GitHub ↗</a>` : ''}
      </div>
      <div class="project-progress" title="${prog}%">
        <div class="progress-track"><div class="progress-fill" style="width:${prog}%"></div></div>
        <span class="progress-value">${prog}%</span>
      </div>
      ${tags ? `<div class="tag-row">${tags}</div>` : ''}
      <div class="project-card-actions">
        <button class="btn-icon" data-action="edit" data-id="${escHtml(p.id)}" title="Modifier" aria-label="Modifier ${escHtml(p.name)}">✏️</button>
        <button class="btn-icon" data-action="delete" data-id="${escHtml(p.id)}" title="Supprimer" aria-label="Supprimer ${escHtml(p.name)}">🗑️</button>
      </div>
    </article>`;
  }).join('');

  container.querySelectorAll('.project-card').forEach(card => {
    card.addEventListener('click', e => {
      if (e.target.closest('[data-action]')) return;
      openProjectDialog(state.projects.find(p => p.id === card.dataset.id));
    });
    card.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (!e.target.closest('[data-action]'))
          openProjectDialog(state.projects.find(p => p.id === card.dataset.id));
      }
    });
    card.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        const id = btn.dataset.id;
        if (btn.dataset.action === 'edit')
          openProjectDialog(state.projects.find(p => p.id === id));
        else if (btn.dataset.action === 'delete')
          confirmDelete(id);
      });
    });
  });
}

// ── Render: tag filter options ────────────────────────────────────────────────

function renderTagFilter() {
  const sel = document.getElementById('filter-tag');
  if (!sel) return;
  const cur = sel.value;
  const tags = getAllTags();
  sel.innerHTML = `<option value="all">Tous les tags</option>` +
    tags.map(t => `<option value="${escHtml(t)}">${escHtml(t)}</option>`).join('');
  sel.value = tags.includes(cur) ? cur : 'all';
}

// ── Render: donut chart ───────────────────────────────────────────────────────

function renderStatusChart() {
  const el = document.getElementById('chart-status');
  if (!el) return;

  const counts = { active: 0, paused: 0, abandoned: 0, done: 0 };
  state.projects.forEach(p => { if (counts[p.status] !== undefined) counts[p.status]++; });
  const total = Object.values(counts).reduce((a, b) => a + b, 0);

  const colors = {
    active: '#55efc4', paused: '#fdcb6e', abandoned: '#ff6b6b', done: '#74b9ff',
  };

  const R = 70, cx = 90, cy = 90, stroke = 22;
  const circumference = 2 * Math.PI * R;

  let offset = -Math.PI / 2;
  let arcs = '';

  if (total === 0) {
    arcs = `<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="rgba(212,165,116,0.1)" stroke-width="${stroke}"/>`;
  } else {
    for (const [key, count] of Object.entries(counts)) {
      if (!count) continue;
      const angle = (count / total) * 2 * Math.PI;
      const x1 = cx + R * Math.cos(offset);
      const y1 = cy + R * Math.sin(offset);
      const x2 = cx + R * Math.cos(offset + angle);
      const y2 = cy + R * Math.sin(offset + angle);
      const large = angle > Math.PI ? 1 : 0;
      arcs += `<path d="M ${x1} ${y1} A ${R} ${R} 0 ${large} 1 ${x2} ${y2}"
        fill="none" stroke="${colors[key]}" stroke-width="${stroke}"
        stroke-linecap="butt"
        style="filter:drop-shadow(0 0 4px ${colors[key]}44)"/>`;
      offset += angle;
    }
  }

  const legend = Object.entries(STATUS_META).map(([key, meta]) => `
    <div class="legend-row">
      <div class="swatch" style="background:${colors[key]};box-shadow:0 0 6px ${colors[key]}66"></div>
      <span class="legend-label">${meta.label}</span>
      <span class="legend-count">${counts[key]}</span>
    </div>`).join('');

  el.innerHTML = `
    <div class="donut-wrap">
      <svg viewBox="0 0 180 180" width="160" height="160" role="img" aria-label="Répartition par statut">
        ${arcs}
        <text x="${cx}" y="${cy - 6}" text-anchor="middle" class="donut-center">${total}</text>
        <text x="${cx}" y="${cy + 14}" text-anchor="middle" class="donut-center-label">PROJETS</text>
      </svg>
      <div class="donut-legend">${legend}</div>
    </div>`;
}

// ── Render: activity bar chart ────────────────────────────────────────────────

function renderActivityChart() {
  const el = document.getElementById('chart-activity');
  if (!el) return;

  const W = 520, H = 150, pad = { top: 10, right: 10, bottom: 30, left: 24 };
  const innerW = W - pad.left - pad.right;
  const innerH = H - pad.top - pad.bottom;
  const weeks = 12;

  const now = new Date();
  const buckets = Array.from({ length: weeks }, (_, i) => {
    const d = new Date(now);
    d.setDate(d.getDate() - (weeks - 1 - i) * 7);
    return { week: d, count: 0, label: '' };
  });

  const getWeekStart = d => {
    const s = new Date(d);
    s.setHours(0, 0, 0, 0);
    s.setDate(s.getDate() - s.getDay());
    return s.getTime();
  };

  const weekStarts = buckets.map(b => getWeekStart(b.week));

  state.projects.forEach(p => {
    const d = new Date(p.updatedAt || p.createdAt);
    const ws = getWeekStart(d);
    const idx = weekStarts.indexOf(ws);
    if (idx >= 0) buckets[idx].count++;
  });

  const maxCount = Math.max(...buckets.map(b => b.count), 1);
  const barW = innerW / weeks;
  const barPad = barW * 0.2;

  const months = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Jun', 'Jul', 'Aoû', 'Sep', 'Oct', 'Nov', 'Déc'];
  let lastMonth = -1;

  const bars = buckets.map((b, i) => {
    const bH = (b.count / maxCount) * innerH;
    const x = pad.left + i * barW + barPad / 2;
    const y = pad.top + innerH - bH;
    const w = barW - barPad;
    const mo = b.week.getMonth();
    let label = '';
    if (mo !== lastMonth) { label = months[mo]; lastMonth = mo; }
    return { x, y, w, bH, count: b.count, label, week: b.week };
  });

  const yTick = maxCount === 1 ? 1 : Math.ceil(maxCount / 3);
  const yTicks = [];
  for (let v = 0; v <= maxCount; v += yTick) yTicks.push(v);

  const defs = `<defs>
    <linearGradient id="bar-grad" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#d4a574"/>
      <stop offset="100%" stop-color="#8a6440"/>
    </linearGradient>
  </defs>`;

  const grid = yTicks.map(v => {
    const y = pad.top + innerH - (v / maxCount) * innerH;
    return `<line class="axis" x1="${pad.left}" x2="${pad.left + innerW}" y1="${y}" y2="${y}"/>
      <text class="tick" x="${pad.left - 4}" y="${y + 3}" text-anchor="end">${v}</text>`;
  }).join('');

  const barsEl = bars.map(b => `
    <rect class="bar" x="${b.x}" y="${b.y}" width="${b.w}" height="${Math.max(b.bH, b.count ? 2 : 0)}"
      rx="3" fill="url(#bar-grad)" opacity="${b.count ? 1 : 0.2}">
      <title>${b.count} projet${b.count !== 1 ? 's' : ''} — sem. du ${b.week.toLocaleDateString('fr-FR')}</title>
    </rect>
    ${b.label ? `<text class="tick" x="${b.x + b.w / 2}" y="${pad.top + innerH + 18}" text-anchor="middle">${b.label}</text>` : ''}`
  ).join('');

  el.innerHTML = `
    <div class="activity-chart">
      <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Activité sur 12 semaines" class="bar-chart">
        ${defs}${grid}${barsEl}
        <line class="axis" x1="${pad.left}" x2="${pad.left + innerW}" y1="${pad.top + innerH}" y2="${pad.top + innerH}"/>
      </svg>
    </div>`;
}

// ── Main render ───────────────────────────────────────────────────────────────

function render() {
  renderTagFilter();
  renderProjects();
  renderStatusChart();
  renderActivityChart();
}

// ── Project dialog ────────────────────────────────────────────────────────────

function openProjectDialog(existing = null) {
  const dlg = document.getElementById('project-dialog');
  if (!dlg) return;

  const isNew = !existing;
  dlg.querySelector('.dialog-title').textContent = isNew ? 'Nouveau projet' : 'Modifier le projet';

  const f = {
    name:        dlg.querySelector('#f-name'),
    description: dlg.querySelector('#f-description'),
    status:      dlg.querySelector('#f-status'),
    priority:    dlg.querySelector('#f-priority'),
    progress:    dlg.querySelector('#f-progress'),
    progressVal: dlg.querySelector('#f-progress-val'),
    tech:        dlg.querySelector('#f-tech'),
    tags:        dlg.querySelector('#f-tags'),
    notes:       dlg.querySelector('#f-notes'),
    github:      dlg.querySelector('#f-github'),
  };

  if (existing) {
    f.name.value        = existing.name || '';
    f.description.value = existing.description || '';
    f.status.value      = existing.status || 'active';
    f.priority.value    = existing.priority || 'low';
    f.progress.value    = existing.progress ?? 0;
    f.progressVal.textContent = (existing.progress ?? 0) + '%';
    f.tech.value        = existing.tech || '';
    f.tags.value        = (existing.tags || []).join(', ');
    f.notes.value       = existing.notes || '';
    f.github.value      = existing.githubUrl || '';
  } else {
    f.name.value = f.description.value = f.tech.value = f.tags.value = f.notes.value = f.github.value = '';
    f.status.value   = 'active';
    f.priority.value = 'low';
    f.progress.value = 0;
    f.progressVal.textContent = '0%';
  }

  const btnDelete = dlg.querySelector('#btn-dlg-delete');
  btnDelete.hidden = isNew;
  btnDelete.dataset.id = existing?.id || '';

  const btnSave = dlg.querySelector('#btn-dlg-save');
  btnSave.dataset.id = existing?.id || '';
  btnSave.dataset.isNew = isNew ? '1' : '';

  dlg.showModal();
  f.name.focus();
}

function setupProjectDialog() {
  const dlg = document.getElementById('project-dialog');
  if (!dlg) return;

  dlg.querySelector('#f-progress').addEventListener('input', e => {
    dlg.querySelector('#f-progress-val').textContent = e.target.value + '%';
  });

  dlg.querySelector('#btn-dlg-cancel').addEventListener('click', () => dlg.close());
  dlg.querySelector('.dialog-close').addEventListener('click', () => dlg.close());

  dlg.querySelector('#btn-dlg-delete').addEventListener('click', e => {
    dlg.close();
    confirmDelete(e.currentTarget.dataset.id);
  });

  dlg.querySelector('#btn-dlg-save').addEventListener('click', () => {
    const name = dlg.querySelector('#f-name').value.trim();
    if (!name) { dlg.querySelector('#f-name').focus(); return; }

    const btn = dlg.querySelector('#btn-dlg-save');
    const isNew = btn.dataset.isNew === '1';
    const id = isNew ? uuid() : btn.dataset.id;
    const existing = isNew ? null : state.projects.find(p => p.id === id);
    const now = new Date().toISOString();

    const rawTags = dlg.querySelector('#f-tags').value;
    const tags = rawTags.split(',').map(t => t.trim()).filter(Boolean);

    const project = {
      id,
      name,
      description: dlg.querySelector('#f-description').value.trim(),
      status:      dlg.querySelector('#f-status').value,
      priority:    dlg.querySelector('#f-priority').value,
      progress:    Number(dlg.querySelector('#f-progress').value),
      tech:        dlg.querySelector('#f-tech').value.trim(),
      tags,
      notes:       dlg.querySelector('#f-notes').value.trim(),
      githubUrl:   dlg.querySelector('#f-github').value.trim(),
      createdAt:   existing?.createdAt || now,
      updatedAt:   now,
    };

    dlg.close();
    upsertProject(project);
    showToast(isNew ? `Projet « ${name} » créé` : `Projet « ${name} » mis à jour`, 'ok');
  });

  dlg.addEventListener('keydown', e => { if (e.key === 'Escape') dlg.close(); });
  trapFocus(dlg);
}

// ── Confirm dialog ────────────────────────────────────────────────────────────

function confirmDelete(id) {
  const proj = state.projects.find(p => p.id === id);
  if (!proj) return;

  const dlg = document.getElementById('confirm-dialog');
  dlg.querySelector('.confirm-message').innerHTML =
    `Supprimer le projet <strong>${escHtml(proj.name)}</strong> ?<br>Cette action est irréversible.`;

  dlg.querySelector('#btn-confirm-ok').onclick = () => {
    dlg.close();
    deleteProject(id);
    showToast(`Projet « ${proj.name} » supprimé`, 'ok');
  };
  dlg.querySelector('#btn-confirm-cancel').onclick = () => dlg.close();
  dlg.querySelector('.dialog-close').onclick = () => dlg.close();
  dlg.showModal();
}

// ── Toast ─────────────────────────────────────────────────────────────────────

function showToast(text, level = 'ok') {
  const region = document.getElementById('toast-region');
  const toast = document.createElement('div');
  toast.className = `toast toast-${level}`;
  toast.setAttribute('role', 'status');
  const icon = level === 'error' ? '⚠️' : '✓';
  toast.innerHTML = `<span class="toast-icon">${icon}</span><span class="toast-text">${escHtml(text)}</span>`;
  region.appendChild(toast);
  setTimeout(() => {
    toast.classList.add('fade-out');
    toast.addEventListener('animationend', () => toast.remove(), { once: true });
    setTimeout(() => toast.remove(), 400);
  }, 4000);
}

// ── Accessibility: focus trap ─────────────────────────────────────────────────

function trapFocus(el) {
  el.addEventListener('keydown', e => {
    if (e.key !== 'Tab') return;
    const focusable = [...el.querySelectorAll(
      'button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex="-1"])'
    )].filter(n => !n.closest('[hidden]'));
    if (!focusable.length) return;
    const first = focusable[0], last = focusable[focusable.length - 1];
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  });
}

// ── Toolbar bindings ──────────────────────────────────────────────────────────

function setupToolbar() {
  const search = document.getElementById('search-input');
  const filterStatus   = document.getElementById('filter-status');
  const filterPriority = document.getElementById('filter-priority');
  const filterTag      = document.getElementById('filter-tag');
  const sortSel        = document.getElementById('sort-select');

  search?.addEventListener('input', e => setState({ search: e.target.value }));
  filterStatus?.addEventListener('change', e => setState({ filterStatus: e.target.value }));
  filterPriority?.addEventListener('change', e => setState({ filterPriority: e.target.value }));
  filterTag?.addEventListener('change', e => setState({ filterTag: e.target.value }));
  sortSel?.addEventListener('change', e => setState({ sort: e.target.value }));

  document.getElementById('btn-new-project')?.addEventListener('click', () => openProjectDialog());
  document.getElementById('btn-refresh')?.addEventListener('click', refreshData);
  document.getElementById('btn-logout')?.addEventListener('click', logout);
}

// ── Login UI helpers ──────────────────────────────────────────────────────────

function showLoginError(msg) {
  const el = document.getElementById('login-error');
  if (el) el.textContent = msg;
}
function clearLoginError() {
  const el = document.getElementById('login-error');
  if (el) el.textContent = '';
}

// ── Show / hide screens ───────────────────────────────────────────────────────

function showApp() {
  document.getElementById('login-screen').hidden = true;
  document.getElementById('app').hidden = false;
}
function showLogin() {
  document.getElementById('app').hidden = true;
  document.getElementById('login-screen').hidden = false;
}

// ── Data loading ──────────────────────────────────────────────────────────────

async function refreshData() {
  const btn = document.getElementById('btn-refresh');
  if (btn) { btn.disabled = true; }
  try {
    const data = await loadProjects();
    setState({ projects: data.projects || [] });
  } catch (err) {
    showToast(err.message, 'error');
    setSyncState('error');
  } finally {
    if (btn) { btn.disabled = false; }
  }
}

function onAuthenticated(account) {
  const nameEl = document.getElementById('account-name');
  if (nameEl) nameEl.textContent = account.username || account.name || '';
  showApp();
  refreshData();
}

// ── Keyboard shortcuts ────────────────────────────────────────────────────────

function setupKeyboardShortcuts() {
  document.addEventListener('keydown', e => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable) return;
    if ((e.metaKey || e.ctrlKey) && e.key === 'k') {
      e.preventDefault();
      document.getElementById('search-input')?.focus();
    }
    if (e.key === 'n' && !e.metaKey && !e.ctrlKey) {
      const dlg = document.getElementById('project-dialog');
      if (!dlg?.open && !document.getElementById('confirm-dialog')?.open) openProjectDialog();
    }
    if (e.key === 'r' && !e.metaKey && !e.ctrlKey) refreshData();
  });
}

// ── Utils ─────────────────────────────────────────────────────────────────────

function escHtml(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// ── Init ──────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('btn-login')?.addEventListener('click', login);

  try {
    await initMsal();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length) {
      onAuthenticated(accounts[0]);
    } else {
      showLogin();
    }
  } catch (err) {
    showLoginError(err.message || String(err));
    showLogin();
  }

  setupProjectDialog();
  setupToolbar();
  setupKeyboardShortcuts();
  setupConfirmDialog();
});

function setupConfirmDialog() {
  const dlg = document.getElementById('confirm-dialog');
  if (!dlg) return;
  dlg.addEventListener('keydown', e => { if (e.key === 'Escape') dlg.close(); });
  trapFocus(dlg);
}
