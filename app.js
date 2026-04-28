import { PublicClientApplication } from 'https://esm.sh/@azure/msal-browser@4.30.0';
import { marked } from 'https://esm.sh/marked@13';

marked.setOptions({ breaks: true, gfm: true });

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
  const cleaned = projects.map(p => ({
    ...p,
    tasks: (p.tasks || []).map(({ _open, ...t }) => t),
  }));
  const body = JSON.stringify({ version: 1, projects: cleaned });
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
  currentProjectId: null,
  currentSprintTab: 'backlog',
  notesMode: 'edit',
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

function togglePin(id) {
  mutateProject(id, p => { p.pinned = !p.pinned; });
}

function uuid() {
  return ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, c =>
    (c ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))).toString(16)
  );
}

// ── Status helpers ────────────────────────────────────────────────────────────

const PROJECT_COLORS = [
  '#74b9ff', '#a29bfe', '#fd79a8', '#00cec9', '#e17055',
  '#55efc4', '#fdcb6e', '#6c5ce7', '#e84393', '#00b894',
];
function getProjectColor(index) {
  return PROJECT_COLORS[index % PROJECT_COLORS.length];
}

const STATUS_META = {
  active:    { emoji: '🟢', label: 'Actif',     color: 'var(--status-active)' },
  paused:    { emoji: '🟡', label: 'En pause',  color: 'var(--status-paused)' },
  abandoned: { emoji: '🔴', label: 'Abandonné', color: 'var(--status-abandoned)' },
  done:      { emoji: '✅', label: 'Terminé',   color: 'var(--status-done)' },
};

const PRIORITY_LABEL = { low: 'Basse', medium: 'Moyenne', high: 'Haute' };

const TASK_STATUS_ORDER = ['todo', 'doing', 'done'];
const TASK_STATUS_LABEL = { todo: 'À faire', doing: 'En cours', done: 'Terminée' };

function taskWeight(task) {
  if (task.subtasks?.length) {
    const done = task.subtasks.filter(s => s.done).length;
    return done / task.subtasks.length;
  }
  return task.status === 'done' ? 1 : task.status === 'doing' ? 0.5 : 0;
}

function computeAutoProgress(project) {
  const tasks = project.tasks || [];
  if (!tasks.length) return 0;
  const sum = tasks.reduce((acc, t) => acc + taskWeight(t), 0);
  return Math.round((sum / tasks.length) * 100);
}

function displayProgress(project) {
  if (project.progressMode === 'auto') return computeAutoProgress(project);
  return Math.max(0, Math.min(100, project.progress || 0));
}

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
      (p.notes || '').toLowerCase().includes(q) ||
      (p.tags || []).some(t => t.toLowerCase().includes(q))
    );
  }
  if (state.filterStatus !== 'all') list = list.filter(p => p.status === state.filterStatus);
  if (state.filterPriority !== 'all') list = list.filter(p => p.priority === state.filterPriority);
  if (state.filterTag !== 'all') list = list.filter(p => (p.tags || []).includes(state.filterTag));

  list.sort((a, b) => {
    if (a.pinned && !b.pinned) return -1;
    if (!a.pinned && b.pinned) return 1;
    if (state.sort === 'name') return a.name.localeCompare(b.name);
    if (state.sort === 'progress') return displayProgress(b) - displayProgress(a);
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
    const prog = displayProgress(p);
    const status = p.status || 'active';
    const taskCount = (p.tasks || []).length;
    return `
    <article class="project-card status-${escHtml(status)} ${p.pinned ? 'is-pinned' : ''}" data-id="${escHtml(p.id)}" role="button" tabindex="0" aria-label="Projet ${escHtml(p.name)}">
      ${p.pinned ? `<span class="pin-badge" aria-label="Épinglé" title="Épinglé">📌</span>` : ''}
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
        ${taskCount ? `<span title="Tâches">✓ ${taskCount}</span>` : ''}
        ${p.githubUrl ? `<a href="${escHtml(p.githubUrl)}" target="_blank" rel="noopener" onclick="event.stopPropagation()">GitHub ↗</a>` : ''}
      </div>
      <div class="project-progress" title="${prog}%${p.progressMode === 'auto' ? ' (auto)' : ''}">
        <div class="progress-track"><div class="progress-fill" style="width:${prog}%"></div></div>
        <span class="progress-value">${prog}%</span>
      </div>
      ${tags ? `<div class="tag-row">${tags}</div>` : ''}
      <div class="project-card-actions">
        <button class="btn-icon ${p.pinned ? 'is-pinned' : ''}" data-action="pin" data-id="${escHtml(p.id)}" title="${p.pinned ? 'Désépingler' : 'Épingler'}" aria-label="${p.pinned ? 'Désépingler' : 'Épingler'} ${escHtml(p.name)}">📌</button>
        <button class="btn-icon" data-action="edit" data-id="${escHtml(p.id)}" title="Modifier" aria-label="Modifier ${escHtml(p.name)}">✏️</button>
        <button class="btn-icon" data-action="delete" data-id="${escHtml(p.id)}" title="Supprimer" aria-label="Supprimer ${escHtml(p.name)}">🗑️</button>
      </div>
    </article>`;
  }).join('');

  container.querySelectorAll('.project-card').forEach(card => {
    card.addEventListener('click', e => {
      if (e.target.closest('[data-action]')) return;
      openProjectView(card.dataset.id);
    });
    card.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (!e.target.closest('[data-action]'))
          openProjectView(card.dataset.id);
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
        else if (btn.dataset.action === 'pin')
          togglePin(id);
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

// ── Per-project charts ────────────────────────────────────────────────────────

function renderProjectCharts(p) {
  const grid = document.getElementById('pv-charts-grid');
  if (!grid) return;
  const sprints = p.sprints || [];
  const tasks = p.tasks || [];
  const activeSprint = sprints.find(s => s.status === 'active' && tasks.some(t => t.sprintId === s.id))
    || sprints.find(s => s.status === 'active');
  const ganttSprints = sprints.filter(s => s.startDate && s.endDate);

  let cards = `
    <div class="chart-card">
      <h3>Tâches par statut</h3>
      <div id="pv-chart-donut"></div>
    </div>`;

  if (sprints.length > 0) cards += `
    <div class="chart-card">
      <h3>Vélocité par sprint</h3>
      <div id="pv-chart-velocity"></div>
    </div>`;

  if (activeSprint) cards += `
    <div class="chart-card">
      <h3>Burndown · ${escHtml(activeSprint.name)}</h3>
      <div id="pv-chart-burndown"></div>
    </div>`;

  if (ganttSprints.length > 0) cards += `
    <div class="chart-card pv-chart-wide">
      <h3>Planning des sprints</h3>
      <div id="pv-chart-gantt"></div>
    </div>`;

  grid.innerHTML = cards;
  renderProjectTaskDonut(p);
  if (sprints.length > 0) renderSprintVelocity(p);
  if (activeSprint) renderBurndownChart(p, activeSprint);
  if (ganttSprints.length > 0) renderProjectGantt(p);
}

function renderProjectTaskDonut(p) {
  const el = document.getElementById('pv-chart-donut');
  if (!el) return;
  const tasks = p.tasks || [];
  const counts = { todo: 0, doing: 0, done: 0 };
  tasks.forEach(t => { if (counts[t.status || 'todo'] !== undefined) counts[t.status || 'todo']++; });
  const total = tasks.length;
  const colors = { todo: '#636e72', doing: '#fdcb6e', done: '#55efc4' };
  const labels = { todo: 'À faire', doing: 'En cours', done: 'Terminées' };
  const R = 55, cx = 70, cy = 70, stroke = 18;
  let offset = -Math.PI / 2, arcs = '';
  if (total === 0) {
    arcs = `<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="rgba(212,165,116,0.1)" stroke-width="${stroke}"/>`;
  } else {
    for (const [key, count] of Object.entries(counts)) {
      if (!count) continue;
      const angle = (count / total) * 2 * Math.PI;
      const x1 = cx + R * Math.cos(offset), y1 = cy + R * Math.sin(offset);
      const x2 = cx + R * Math.cos(offset + angle), y2 = cy + R * Math.sin(offset + angle);
      arcs += `<path d="M ${x1} ${y1} A ${R} ${R} 0 ${angle > Math.PI ? 1 : 0} 1 ${x2} ${y2}" fill="none" stroke="${colors[key]}" stroke-width="${stroke}" stroke-linecap="butt" style="filter:drop-shadow(0 0 4px ${colors[key]}44)"/>`;
      offset += angle;
    }
  }
  const legend = Object.entries(labels).map(([key, label]) => `
    <div class="legend-row">
      <div class="swatch" style="background:${colors[key]};box-shadow:0 0 6px ${colors[key]}66"></div>
      <span class="legend-label">${label}</span>
      <span class="legend-count">${counts[key]}</span>
    </div>`).join('');
  el.innerHTML = `
    <div class="donut-wrap">
      <svg viewBox="0 0 140 140" width="120" height="120" role="img" aria-label="Répartition des tâches">
        ${arcs}
        <text x="${cx}" y="${cy - 5}" text-anchor="middle" class="donut-center" style="font-size:1.2rem">${total}</text>
        <text x="${cx}" y="${cy + 13}" text-anchor="middle" class="donut-center-label">TÂCHES</text>
      </svg>
      <div class="donut-legend">${legend}</div>
    </div>`;
}

function renderBurndownChart(p, sprint) {
  const el = document.getElementById('pv-chart-burndown');
  if (!el) return;
  const tasks = (p.tasks || []).filter(t => t.sprintId === sprint.id);
  if (!tasks.length) {
    el.innerHTML = `<p style="color:var(--text-muted);font-size:0.8rem;padding:0.75rem 0">Aucune tâche dans ce sprint.</p>`;
    return;
  }
  const start = new Date(sprint.startDate); start.setHours(0, 0, 0, 0);
  const end   = new Date(sprint.endDate);   end.setHours(23, 59, 59, 999);
  const today = new Date();
  const plotEnd = today < end ? today : end;
  const days = [];
  for (let d = new Date(start); d <= plotEnd; d = new Date(d.getTime() + 86400000)) {
    const remaining = tasks.filter(t => !t.completedAt || new Date(t.completedAt) > d).length;
    days.push({ ts: d.getTime(), remaining });
  }
  const total = tasks.length;
  const W = 340, H = 140, pad = { top: 10, right: 12, bottom: 28, left: 28 };
  const iW = W - pad.left - pad.right, iH = H - pad.top - pad.bottom;
  const xS = ts => pad.left + ((ts - start.getTime()) / (end.getTime() - start.getTime())) * iW;
  const yS = v  => pad.top  + iH - (v / total) * iH;
  const idealPath = `M ${xS(start.getTime())} ${yS(total)} L ${xS(end.getTime())} ${yS(0)}`;
  const actualPath = days.length > 1 ? `M ${days.map(d => `${xS(d.ts)} ${yS(d.remaining)}`).join(' L ')}` : '';
  const yTicks = [0, Math.ceil(total / 2), total].map(v =>
    `<line x1="${pad.left}" x2="${pad.left+iW}" y1="${yS(v)}" y2="${yS(v)}" stroke="rgba(212,165,116,0.08)" stroke-width="1"/>
     <text x="${pad.left-4}" y="${yS(v)+4}" text-anchor="end" class="tick">${v}</text>`
  ).join('');
  const fmt = d => d.toLocaleDateString('fr-FR', { day:'numeric', month:'short' });
  el.innerHTML = `
    <svg viewBox="0 0 ${W} ${H}" width="100%" role="img" class="bar-chart">
      ${yTicks}
      <path d="${idealPath}" stroke="rgba(212,165,116,0.3)" stroke-width="1.5" fill="none" stroke-dasharray="4 3"/>
      ${actualPath ? `<path d="${actualPath}" stroke="#55efc4" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/>` : ''}
      <line x1="${pad.left}" x2="${pad.left+iW}" y1="${pad.top+iH}" y2="${pad.top+iH}" stroke="rgba(212,165,116,0.15)" stroke-width="1"/>
      <text x="${xS(start.getTime())}" y="${H-6}" class="tick" text-anchor="start">${fmt(start)}</text>
      <text x="${xS(end.getTime())}"   y="${H-6}" class="tick" text-anchor="end">${fmt(end)}</text>
    </svg>
    <div class="pv-chart-legend">
      <span style="color:rgba(212,165,116,0.5)">— Idéal</span>
      <span style="color:#55efc4">— Réel</span>
    </div>`;
}

function renderSprintVelocity(p) {
  const el = document.getElementById('pv-chart-velocity');
  if (!el) return;
  const sprints = p.sprints || [];
  const tasks   = p.tasks   || [];
  if (!sprints.length) return;
  const data = sprints.map(s => ({
    name:  s.name,
    done:  tasks.filter(t => t.sprintId === s.id && t.status === 'done').length,
    total: tasks.filter(t => t.sprintId === s.id).length,
    isDone: s.status === 'done',
  }));
  const W = 340, H = 140, pad = { top: 16, right: 12, bottom: 30, left: 24 };
  const iW = W - pad.left - pad.right, iH = H - pad.top - pad.bottom;
  const maxVal = Math.max(...data.map(d => d.total), 1);
  const barW = iW / data.length, barPad = barW * 0.25;
  const defs = data.map((d, i) => {
    const c = d.isDone ? '#74b9ff' : '#55efc4';
    const cd = d.isDone ? '#3a82c4' : '#2fbf95';
    return `<linearGradient id="vg${i}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${c}"/><stop offset="100%" stop-color="${cd}"/></linearGradient>`;
  }).join('');
  const bars = data.map((d, i) => {
    const bH    = (d.done  / maxVal) * iH;
    const totH  = (d.total / maxVal) * iH;
    const x = pad.left + i * barW + barPad / 2, w = barW - barPad;
    const c = d.isDone ? '#74b9ff' : '#55efc4';
    const label = d.name.length > 9 ? d.name.slice(0, 8) + '…' : d.name;
    return `
      <rect x="${x}" y="${pad.top+iH-totH}" width="${w}" height="${Math.max(totH,2)}" rx="3" fill="rgba(212,165,116,0.1)"/>
      <rect x="${x}" y="${pad.top+iH-bH}"   width="${w}" height="${Math.max(bH, d.done ? 2 : 0)}" rx="3" fill="url(#vg${i})" opacity="${d.done ? 1 : 0.3}">
        <title>${d.done}/${d.total} — ${d.name}</title>
      </rect>
      <text x="${x+w/2}" y="${H-6}" class="tick" text-anchor="middle">${label}</text>
      ${d.done ? `<text x="${x+w/2}" y="${pad.top+iH-bH-5}" class="tick" text-anchor="middle" style="fill:${c}">${d.done}</text>` : ''}`;
  }).join('');
  el.innerHTML = `
    <svg viewBox="0 0 ${W} ${H}" width="100%" role="img" class="bar-chart">
      <defs>${defs}</defs>
      ${bars}
      <line x1="${pad.left}" x2="${pad.left+iW}" y1="${pad.top+iH}" y2="${pad.top+iH}" stroke="rgba(212,165,116,0.15)" stroke-width="1"/>
    </svg>`;
}

function renderProjectGantt(p) {
  const el = document.getElementById('pv-chart-gantt');
  if (!el) return;
  const sprints = (p.sprints || []).filter(s => s.startDate && s.endDate);
  if (!sprints.length) return;
  const allDates = sprints.flatMap(s => [new Date(s.startDate), new Date(s.endDate)]);
  const minDate = new Date(Math.min(...allDates.map(d => d.getTime())));
  const maxDate = new Date(Math.max(...allDates.map(d => d.getTime())));
  const totalMs = Math.max(maxDate.getTime() - minDate.getTime(), 1);
  const rowH = 34, W = 600, pad = { top: 8, right: 16, bottom: 26, left: 16 };
  const labelW = 110, iW = W - pad.left - pad.right - labelW;
  const H = pad.top + sprints.length * rowH + pad.bottom;
  const today = new Date();
  const todayX = pad.left + labelW + ((today.getTime() - minDate.getTime()) / totalMs) * iW;
  const showToday = today >= minDate && today <= maxDate;
  const rows = sprints.map((s, i) => {
    const x1 = pad.left + labelW + ((new Date(s.startDate).getTime() - minDate.getTime()) / totalMs) * iW;
    const x2 = pad.left + labelW + ((new Date(s.endDate).getTime()   - minDate.getTime()) / totalMs) * iW;
    const bW  = Math.max(x2 - x1, 4);
    const y   = pad.top + i * rowH;
    const tasksDone  = (p.tasks || []).filter(t => t.sprintId === s.id && t.status === 'done').length;
    const tasksTotal = (p.tasks || []).filter(t => t.sprintId === s.id).length;
    const fill = tasksTotal ? tasksDone / tasksTotal : 0;
    const color = s.status === 'done' ? '#74b9ff' : '#55efc4';
    const name = s.name.length > 14 ? s.name.slice(0, 13) + '…' : s.name;
    return `
      <text x="${pad.left+labelW-6}" y="${y+rowH/2+4}" text-anchor="end" class="tick" style="font-size:10px">${escHtml(name)}</text>
      <rect x="${x1}" y="${y+6}" width="${bW}" height="${rowH-12}" rx="4" fill="${color}" opacity="0.15"/>
      <rect x="${x1}" y="${y+6}" width="${bW*fill}" height="${rowH-12}" rx="4" fill="${color}" opacity="0.8" style="filter:drop-shadow(0 0 4px ${color}55)"/>
      ${s.status === 'done' ? `<text x="${x1+bW+5}" y="${y+rowH/2+4}" class="tick" style="fill:${color}">✓</text>` : ''}`;
  }).join('');
  let monthTicks = '';
  const mfmt = d => d.toLocaleDateString('fr-FR', { month:'short', year:'2-digit' });
  for (let d = new Date(minDate.getFullYear(), minDate.getMonth(), 1); d <= maxDate; d = new Date(d.getFullYear(), d.getMonth()+1, 1)) {
    const x = pad.left + labelW + ((d.getTime() - minDate.getTime()) / totalMs) * iW;
    if (x >= pad.left + labelW)
      monthTicks += `<line x1="${x}" x2="${x}" y1="${pad.top}" y2="${H-pad.bottom}" stroke="rgba(212,165,116,0.07)" stroke-width="1"/>
        <text x="${x}" y="${H-6}" class="tick" text-anchor="middle">${mfmt(d)}</text>`;
  }
  el.innerHTML = `
    <div class="gantt-scroll">
      <svg viewBox="0 0 ${W} ${H}" width="100%" role="img" class="bar-chart">
        ${monthTicks}${rows}
        ${showToday ? `<line x1="${todayX}" x2="${todayX}" y1="${pad.top}" y2="${H-pad.bottom}" stroke="var(--accent-warm)" stroke-width="1.5" stroke-dasharray="3 2" opacity="0.6"/>` : ''}
        <line x1="${pad.left+labelW}" x2="${pad.left+labelW+iW}" y1="${H-pad.bottom}" y2="${H-pad.bottom}" stroke="rgba(212,165,116,0.15)" stroke-width="1"/>
      </svg>
    </div>`;
}

function renderActivityLog(p) {
  const el = document.getElementById('pv-activity-log');
  if (!el) return;
  const log = p.activityLog || [];
  if (!log.length) {
    el.innerHTML = `<div class="activity-empty">Aucune activité encore enregistrée.</div>`;
    return;
  }
  const icons = { task_done:'✓', sprint_started:'🚀', sprint_completed:'🏁', sprint_deleted:'🗑', status_changed:'🔄' };
  el.innerHTML = log.map(e => `
    <div class="activity-item">
      <span class="activity-icon">${icons[e.type] || '·'}</span>
      <span class="activity-text">${escHtml(e.text)}</span>
      <span class="activity-date">${relativeDate(e.date)}</span>
    </div>`).join('');
}

function renderFocusWidget() {
  const el = document.getElementById('focus-widget');
  if (!el) return;
  const groups = state.projects
    .filter(p => p.status === 'active')
    .map(p => ({ project: p, tasks: (p.tasks || []).filter(t => t.status === 'doing') }))
    .filter(g => g.tasks.length > 0);
  if (!groups.length) {
    el.innerHTML = `<div class="focus-empty">Aucune tâche "En cours" sur tes projets actifs.</div>`;
    return;
  }
  el.innerHTML = groups.map(g => `
    <div class="focus-group">
      <div class="focus-group-header">
        <span class="project-status">${STATUS_META[g.project.status]?.emoji || '🟢'}</span>
        <span class="focus-project-name">${escHtml(g.project.name)}</span>
        <span class="focus-task-count">${g.tasks.length} en cours</span>
      </div>
      <div class="focus-task-list">
        ${g.tasks.map(t => `
          <div class="focus-task-item" data-project-id="${escHtml(g.project.id)}">
            <span class="focus-task-dot"></span>
            <span class="focus-task-title">${escHtml(t.title)}</span>
          </div>`).join('')}
      </div>
    </div>`).join('');
  el.querySelectorAll('.focus-task-item').forEach(item => {
    item.addEventListener('click', () => openProjectView(item.dataset.projectId));
  });
}

function renderGlobalGantt() {
  const el = document.getElementById('chart-global-gantt');
  if (!el) return;
  const allSprints = [];
  state.projects.forEach((proj, pi) => {
    const color = getProjectColor(pi);
    (proj.sprints || []).filter(s => s.startDate && s.endDate).forEach(s => {
      allSprints.push({ ...s, projectName: proj.name, projectId: proj.id, color });
    });
  });
  if (!allSprints.length) {
    el.innerHTML = `<div style="color:var(--text-muted);font-size:0.82rem;padding:1rem;text-align:center">Aucun sprint avec des dates défini. Crée des sprints dans tes projets.</div>`;
    return;
  }
  allSprints.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
  const allDates = allSprints.flatMap(s => [new Date(s.startDate), new Date(s.endDate)]);
  const minDate = new Date(Math.min(...allDates.map(d => d.getTime())));
  const maxDate = new Date(Math.max(...allDates.map(d => d.getTime())));
  const totalMs = Math.max(maxDate.getTime() - minDate.getTime(), 1);
  const rowH = 36, W = 700, pad = { top: 8, right: 16, bottom: 28, left: 16 };
  const labelW = 130, iW = W - pad.left - pad.right - labelW;
  const H = pad.top + allSprints.length * rowH + pad.bottom;
  const today = new Date();
  const todayX = pad.left + labelW + ((today.getTime() - minDate.getTime()) / totalMs) * iW;
  const showToday = today >= minDate && today <= maxDate;
  const rows = allSprints.map((s, i) => {
    const x1 = pad.left + labelW + ((new Date(s.startDate).getTime() - minDate.getTime()) / totalMs) * iW;
    const x2 = pad.left + labelW + ((new Date(s.endDate).getTime()   - minDate.getTime()) / totalMs) * iW;
    const bW  = Math.max(x2 - x1, 6);
    const y   = pad.top + i * rowH;
    const label = `${s.projectName} · ${s.name}`;
    const displayLabel = label.length > 20 ? label.slice(0, 19) + '…' : label;
    return `
      <g class="gantt-row" data-project-id="${escHtml(s.projectId)}">
        <text x="${pad.left+labelW-6}" y="${y+rowH/2+4}" text-anchor="end" class="tick" style="font-size:9.5px">${escHtml(displayLabel)}</text>
        <rect x="${x1}" y="${y+8}" width="${bW}" height="${rowH-16}" rx="4" fill="${s.color}" opacity="${s.status === 'done' ? 0.3 : 0.7}" style="filter:drop-shadow(0 0 4px ${s.color}44)">
          <title>${escHtml(s.projectName)} — ${escHtml(s.name)}</title>
        </rect>
        ${s.status === 'done' ? `<text x="${x1+bW/2}" y="${y+rowH/2+4}" text-anchor="middle" style="font-size:9px;fill:#fff;pointer-events:none">✓</text>` : ''}
      </g>`;
  }).join('');
  let monthTicks = '';
  const mfmt = d => d.toLocaleDateString('fr-FR', { month:'short', year:'2-digit' });
  for (let d = new Date(minDate.getFullYear(), minDate.getMonth(), 1); d <= maxDate; d = new Date(d.getFullYear(), d.getMonth()+1, 1)) {
    const x = pad.left + labelW + ((d.getTime() - minDate.getTime()) / totalMs) * iW;
    if (x >= pad.left + labelW)
      monthTicks += `<line x1="${x}" x2="${x}" y1="${pad.top}" y2="${H-pad.bottom}" stroke="rgba(212,165,116,0.07)" stroke-width="1"/>
        <text x="${x}" y="${H-6}" class="tick" text-anchor="middle">${mfmt(d)}</text>`;
  }
  el.innerHTML = `
    <div class="gantt-scroll">
      <svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="Gantt global" class="bar-chart">
        ${monthTicks}${rows}
        ${showToday ? `<line x1="${todayX}" x2="${todayX}" y1="${pad.top}" y2="${H-pad.bottom}" stroke="var(--accent-warm)" stroke-width="1.5" stroke-dasharray="3 2" opacity="0.7"/>` : ''}
      </svg>
    </div>`;
  el.querySelectorAll('.gantt-row').forEach(row => {
    row.addEventListener('click', () => { if (row.dataset.projectId) openProjectView(row.dataset.projectId); });
  });
}

// ── Main render ───────────────────────────────────────────────────────────────

function render() {
  const dashboard = document.getElementById('dashboard-view');
  const projectView = document.getElementById('project-view');
  const toolbar = document.querySelector('.toolbar');

  if (state.currentProjectId) {
    const proj = state.projects.find(p => p.id === state.currentProjectId);
    if (!proj) {
      state.currentProjectId = null;
    } else {
      if (dashboard) dashboard.hidden = true;
      if (toolbar) toolbar.hidden = true;
      if (projectView) projectView.hidden = false;
      renderProjectView(proj);
      return;
    }
  }

  if (dashboard) dashboard.hidden = false;
  if (toolbar) toolbar.hidden = false;
  if (projectView) projectView.hidden = true;

  renderTagFilter();
  renderProjects();
  renderFocusWidget();
  renderGlobalGantt();
}

// ── Project detail view ──────────────────────────────────────────────────────

function openProjectView(id) {
  const proj = state.projects.find(p => p.id === id);
  const firstActiveSprint = (proj?.sprints || []).find(s => s.status === 'active');
  state.currentProjectId = id;
  state.currentSprintTab = firstActiveSprint?.id || 'backlog';
  state.notesMode = 'edit';
  render();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function closeProjectView() {
  state.currentProjectId = null;
  render();
}

function renderProjectView(p) {
  const root = document.getElementById('project-view');
  if (!root) return;

  const sm = STATUS_META[p.status] || STATUS_META.active;
  const status = p.status || 'active';
  const prog = displayProgress(p);
  const auto = p.progressMode === 'auto';
  const tags = (p.tags || []).map(t => `<span class="tag-chip">${escHtml(t)}</span>`).join('');
  const sprints = p.sprints || [];
  const allTasks = p.tasks || [];

  const currentTab = state.currentSprintTab || 'backlog';
  const currentSprint = sprints.find(s => s.id === currentTab);
  const tabTasks = currentTab === 'backlog'
    ? allTasks.filter(t => !t.sprintId)
    : allTasks.filter(t => t.sprintId === currentTab);
  const doneCount = tabTasks.filter(t => t.status === 'done').length;

  // Sprint tabs
  const backlogCount = allTasks.filter(t => !t.sprintId).length;
  const sprintTabsHtml = sprints.map(s => {
    const sCount = allTasks.filter(t => t.sprintId === s.id).length;
    const start = s.startDate ? new Date(s.startDate) : null;
    const end   = s.endDate   ? new Date(s.endDate)   : null;
    const dateRange = (start && end)
      ? `${start.toLocaleDateString('fr-FR', { day:'numeric', month:'short' })} → ${end.toLocaleDateString('fr-FR', { day:'numeric', month:'short' })}`
      : '';
    return `
      <button class="sprint-tab${s.status === 'done' ? ' is-done' : ''}${currentTab === s.id ? ' is-active' : ''}" data-sprint-tab="${escHtml(s.id)}">
        <span class="sprint-tab-name">${escHtml(s.name)}</span>
        ${dateRange ? `<span class="sprint-tab-dates">${dateRange}</span>` : ''}
        <span class="sprint-tab-count">${sCount}</span>
        ${s.status === 'done' ? '<span class="sprint-tab-done-badge">✓</span>' : ''}
      </button>`;
  }).join('');

  // Sprint info bar
  const sprintInfoHtml = currentSprint ? `
    <div class="sprint-info-bar">
      <span class="sprint-info-dates">📅 ${new Date(currentSprint.startDate).toLocaleDateString('fr-FR', { day:'numeric', month:'long' })} → ${new Date(currentSprint.endDate).toLocaleDateString('fr-FR', { day:'numeric', month:'long' })}</span>
      <span class="sprint-info-progress">${allTasks.filter(t => t.sprintId === currentSprint.id && t.status === 'done').length} / ${allTasks.filter(t => t.sprintId === currentSprint.id).length} tâches terminées</span>
      <div class="sprint-info-actions">
        <button class="btn btn-ghost" id="pv-edit-sprint" data-sprint-id="${escHtml(currentSprint.id)}">Modifier</button>
        <button class="btn ${currentSprint.status === 'done' ? 'btn-ghost' : 'btn-primary'}" id="pv-toggle-sprint" data-sprint-id="${escHtml(currentSprint.id)}">
          ${currentSprint.status === 'done' ? '↩ Réactiver' : '✓ Terminer'}
        </button>
      </div>
    </div>` : '';

  root.innerHTML = `
    <div class="pv-topbar">
      <button class="btn-back" id="pv-back" aria-label="Retour à la liste">← Retour</button>
      <div class="pv-topbar-spacer"></div>
      <button class="btn btn-ghost ${p.pinned ? 'is-pinned' : ''}" id="pv-pin">📌 ${p.pinned ? 'Épinglé' : 'Épingler'}</button>
      <button class="btn btn-ghost" id="pv-edit">Modifier</button>
      <button class="btn btn-danger" id="pv-delete">Supprimer</button>
    </div>

    <header class="pv-header status-${escHtml(status)}">
      <div class="pv-title-row">
        <h1 class="pv-title">
          <span class="project-status" title="${escHtml(sm.label)}">${sm.emoji}</span>
          <span>${escHtml(p.name)}</span>
        </h1>
        <div class="pv-actions">
          <span class="priority-pill priority-${escHtml(p.priority || 'low')}">${escHtml(PRIORITY_LABEL[p.priority] || 'Basse')}</span>
        </div>
      </div>
      <div class="pv-meta">
        <span><strong>Statut</strong> · ${escHtml(sm.label)}</span>
        ${p.tech ? `<span class="tech">${escHtml(p.tech)}</span>` : ''}
        <span>Créé · ${relativeDate(p.createdAt)}</span>
        <span>Modifié · ${relativeDate(p.updatedAt || p.createdAt)}</span>
        ${p.githubUrl ? `<a href="${escHtml(p.githubUrl)}" target="_blank" rel="noopener">GitHub ↗</a>` : ''}
      </div>
      ${p.description ? `<p class="pv-description">${escHtml(p.description)}</p>` : ''}
      <div class="pv-progress" title="${prog}%${auto ? ' (auto)' : ''}">
        <div class="progress-track"><div class="progress-fill" style="width:${prog}%"></div></div>
        <span class="progress-value">${prog}%${auto ? ' · auto' : ''}</span>
      </div>
      ${tags ? `<div class="pv-tags">${tags}</div>` : ''}
    </header>

    <div class="sprint-tabs-bar" id="pv-sprint-tabs">
      <button class="sprint-tab${currentTab === 'backlog' ? ' is-active' : ''}" data-sprint-tab="backlog">
        Backlog <span class="sprint-tab-count">${backlogCount}</span>
      </button>
      ${sprintTabsHtml}
      <button class="sprint-tab sprint-tab-add" id="pv-add-sprint">+ Sprint</button>
    </div>

    ${sprintInfoHtml}

    <div class="pv-body-grid">
      <section class="pv-card" aria-label="Tâches">
        <div class="pv-card-head">
          <h3>${currentTab === 'backlog' ? 'Backlog' : escHtml(currentSprint?.name || 'Sprint')}</h3>
          <span class="pv-card-meta">${doneCount} / ${tabTasks.length} terminée${tabTasks.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="task-list" id="pv-task-list"></div>
        <div class="task-add-row">
          <input type="text" class="task-add-input" id="pv-task-add" placeholder="Ajouter une tâche... (Entrée)" />
        </div>
      </section>

      <section class="pv-card" aria-label="Notes">
        <div class="pv-card-head">
          <h3>Notes</h3>
          <div class="notes-head-actions">
            <div class="notes-tabs" role="tablist">
              <button class="mode-btn ${state.notesMode === 'edit' ? 'is-active' : ''}" data-notes-mode="edit" role="tab" aria-selected="${state.notesMode === 'edit'}">Édition</button>
              <button class="mode-btn ${state.notesMode === 'preview' ? 'is-active' : ''}" data-notes-mode="preview" role="tab" aria-selected="${state.notesMode === 'preview'}">Aperçu</button>
            </div>
            <button class="btn-icon" id="pv-export-md" title="Exporter en .md" aria-label="Exporter les notes en markdown">⬇</button>
          </div>
        </div>
        ${state.notesMode === 'edit'
          ? `<textarea class="notes-textarea" id="pv-notes" placeholder="Markdown supporté : # Titres, **gras**, *italique*, [liens](https://...), - listes, \`code\`, > citations..."></textarea>`
          : `<div class="notes-preview ${(!p.notes || !p.notes.trim()) ? 'is-empty' : ''}" id="pv-notes-preview"></div>`
        }
      </section>
    </div>

    <section class="pv-charts-section" aria-label="Graphiques">
      <div class="section-title">Graphiques</div>
      <div class="pv-charts-grid" id="pv-charts-grid"></div>
    </section>

    <section class="pv-card pv-activity-card" aria-label="Activité récente">
      <div class="pv-card-head"><h3>Activité récente</h3></div>
      <div id="pv-activity-log" class="activity-log"></div>
    </section>
  `;

  document.getElementById('pv-back').addEventListener('click', closeProjectView);
  document.getElementById('pv-edit').addEventListener('click', () => openProjectDialog(p));
  document.getElementById('pv-delete').addEventListener('click', () => confirmDelete(p.id));
  document.getElementById('pv-pin').addEventListener('click', () => togglePin(p.id));
  document.getElementById('pv-export-md').addEventListener('click', () => exportNotesAsMd(p));
  document.getElementById('pv-add-sprint').addEventListener('click', () => openSprintDialog(p.id));

  if (currentSprint) {
    document.getElementById('pv-edit-sprint')?.addEventListener('click', () => openSprintDialog(p.id, currentSprint));
    document.getElementById('pv-toggle-sprint')?.addEventListener('click', () => {
      toggleSprintStatus(p.id, currentSprint.id);
    });
  }

  root.querySelectorAll('[data-sprint-tab]').forEach(btn => {
    btn.addEventListener('click', () => {
      state.currentSprintTab = btn.dataset.sprintTab;
      render();
    });
  });

  root.querySelectorAll('[data-notes-mode]').forEach(btn => {
    btn.addEventListener('click', () => {
      state.notesMode = btn.dataset.notesMode;
      render();
    });
  });

  renderTaskList(p, tabTasks);
  setupTaskAddInput(p, currentTab === 'backlog' ? null : currentTab);
  setupNotesArea(p);
  renderProjectCharts(p);
  renderActivityLog(p);
}

// ── Tasks ────────────────────────────────────────────────────────────────────

function renderTaskList(p, filteredTasks = null) {
  const list = document.getElementById('pv-task-list');
  if (!list) return;
  const tasks = filteredTasks !== null ? filteredTasks : (p.tasks || []);
  const sprints = p.sprints || [];

  if (!tasks.length) {
    list.innerHTML = `<div class="task-empty">Aucune tâche ici. Ajoutes-en une ci-dessous.</div>`;
    return;
  }

  const sprintOptions = `<option value="">Backlog</option>` +
    sprints.map(s => `<option value="${escHtml(s.id)}">${escHtml(s.name)}</option>`).join('');

  list.innerHTML = tasks.map(t => {
    const subs = t.subtasks || [];
    const isOpen = !!t._open;
    return `
    <div class="task-item status-${escHtml(t.status || 'todo')}" data-task-id="${escHtml(t.id)}">
      <div class="task-row">
        <button class="task-drag-handle" draggable="true" title="Glisser pour réordonner" aria-label="Réordonner la tâche">⋮⋮</button>
        ${subs.length ? `<button class="task-expand ${isOpen ? 'is-open' : ''}" data-action="toggle-expand" aria-label="Afficher les sous-tâches">▶</button>` : `<span style="width:18px;flex-shrink:0"></span>`}
        <button class="task-status-btn status-${escHtml(t.status || 'todo')}" data-action="cycle-status" title="Changer le statut (À faire → En cours → Terminée)" aria-label="Statut: ${escHtml(TASK_STATUS_LABEL[t.status || 'todo'])}"></button>
        <input type="text" class="task-title-input" data-action="edit-title" value="${escHtml(t.title || '')}" placeholder="Titre de la tâche" />
        <select class="task-sprint-select" data-action="change-sprint" aria-label="Sprint">
          ${sprintOptions}
        </select>
        <div class="task-actions">
          <button class="btn-icon" data-action="add-subtask" title="Ajouter une sous-tâche" aria-label="Ajouter une sous-tâche">+</button>
          <button class="btn-icon" data-action="delete-task" title="Supprimer la tâche" aria-label="Supprimer la tâche">🗑️</button>
        </div>
      </div>
      ${(subs.length && isOpen) ? `
        <div class="subtask-list">
          ${subs.map(s => `
            <div class="subtask-row" data-sub-id="${escHtml(s.id)}">
              <button class="subtask-checkbox ${s.done ? 'is-done' : ''}" data-action="toggle-sub" aria-label="${s.done ? 'Décocher' : 'Cocher'}"></button>
              <input type="text" class="subtask-title-input ${s.done ? 'is-done' : ''}" data-action="edit-sub-title" value="${escHtml(s.title || '')}" placeholder="Sous-tâche" />
              <button class="subtask-remove" data-action="delete-sub" aria-label="Supprimer la sous-tâche">×</button>
            </div>
          `).join('')}
          <button class="add-subtask-btn" data-action="new-sub">+ Sous-tâche</button>
        </div>
      ` : ''}
    </div>`;
  }).join('');

  list.querySelectorAll('.task-item').forEach(el => {
    const taskId = el.dataset.taskId;
    el.querySelectorAll('[data-action]').forEach(btn => {
      const action = btn.dataset.action;
      if (action === 'cycle-status') {
        btn.addEventListener('click', () => cycleTaskStatus(p.id, taskId));
      } else if (action === 'edit-title') {
        btn.addEventListener('change', e => updateTaskTitle(p.id, taskId, e.target.value));
        btn.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
      } else if (action === 'add-subtask') {
        btn.addEventListener('click', () => addSubtask(p.id, taskId));
      } else if (action === 'new-sub') {
        btn.addEventListener('click', () => addSubtask(p.id, taskId));
      } else if (action === 'delete-task') {
        btn.addEventListener('click', () => deleteTask(p.id, taskId));
      } else if (action === 'toggle-expand') {
        btn.addEventListener('click', () => toggleTaskExpand(p.id, taskId));
      } else if (action === 'change-sprint') {
        btn.value = (p.tasks || []).find(t => t.id === taskId)?.sprintId || '';
        btn.addEventListener('change', e => moveTaskToSprint(p.id, taskId, e.target.value || null));
      }
    });
    el.querySelectorAll('.subtask-row').forEach(row => {
      const subId = row.dataset.subId;
      row.querySelectorAll('[data-action]').forEach(btn => {
        const action = btn.dataset.action;
        if (action === 'toggle-sub') {
          btn.addEventListener('click', () => toggleSubtask(p.id, taskId, subId));
        } else if (action === 'edit-sub-title') {
          btn.addEventListener('change', e => updateSubtaskTitle(p.id, taskId, subId, e.target.value));
          btn.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
        } else if (action === 'delete-sub') {
          btn.addEventListener('click', () => deleteSubtask(p.id, taskId, subId));
        }
      });
    });
  });

  setupTaskDnd(p);
}

function setupTaskDnd(p) {
  const list = document.getElementById('pv-task-list');
  if (!list) return;

  list.querySelectorAll('.task-drag-handle').forEach(handle => {
    handle.addEventListener('dragstart', e => {
      const item = handle.closest('.task-item');
      if (!item) return;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', item.dataset.taskId);
      try { e.dataTransfer.setDragImage(item, 20, 20); } catch {}
      requestAnimationFrame(() => item.classList.add('is-dragging'));
    });
    handle.addEventListener('dragend', () => {
      const dragging = list.querySelector('.is-dragging');
      if (!dragging) return;
      dragging.classList.remove('is-dragging');
      const ids = [...list.querySelectorAll('.task-item')].map(el => el.dataset.taskId);
      const current = (state.projects.find(x => x.id === p.id)?.tasks || []).map(t => t.id);
      if (ids.join(',') !== current.join(',')) reorderTasks(p.id, ids);
    });
  });

  list.addEventListener('dragover', e => {
    const dragging = list.querySelector('.is-dragging');
    if (!dragging) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const after = getTaskAfterCursor(list, e.clientY);
    if (after == null) {
      list.appendChild(dragging);
    } else if (after !== dragging) {
      list.insertBefore(dragging, after);
    }
  });
}

function getTaskAfterCursor(list, y) {
  const items = [...list.querySelectorAll('.task-item:not(.is-dragging)')];
  let closest = { offset: -Infinity, element: null };
  items.forEach(child => {
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) closest = { offset, element: child };
  });
  return closest.element;
}

function reorderTasks(projectId, ids) {
  mutateProject(projectId, p => {
    const map = new Map((p.tasks || []).map(t => [t.id, t]));
    p.tasks = ids.map(id => map.get(id)).filter(Boolean);
  });
}

function setupTaskAddInput(p, sprintId = null) {
  const input = document.getElementById('pv-task-add');
  if (!input) return;
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      const title = input.value.trim();
      if (!title) return;
      input.value = '';
      addTask(p.id, title, sprintId);
      document.getElementById('pv-task-add')?.focus();
    }
  });
}

function mutateProject(id, mutator) {
  const idx = state.projects.findIndex(p => p.id === id);
  if (idx < 0) return;
  const next = { ...state.projects[idx] };
  mutator(next);
  next.updatedAt = new Date().toISOString();
  const projects = [...state.projects];
  projects[idx] = next;
  setState({ projects });
  scheduleSave();
}

function logActivity(p, type, text) {
  p.activityLog = p.activityLog || [];
  p.activityLog.unshift({ type, text, date: new Date().toISOString() });
  if (p.activityLog.length > 30) p.activityLog.length = 30;
}

function addSprint(projectId, sprint) {
  mutateProject(projectId, p => {
    p.sprints = [...(p.sprints || []), sprint];
    logActivity(p, 'sprint_started', `Sprint "${sprint.name}" créé`);
  });
}

function updateSprint(projectId, sprint) {
  mutateProject(projectId, p => {
    p.sprints = (p.sprints || []).map(s => s.id === sprint.id ? sprint : s);
  });
}

function deleteSprint(projectId, sprintId) {
  mutateProject(projectId, p => {
    const sprint = (p.sprints || []).find(s => s.id === sprintId);
    p.tasks = (p.tasks || []).map(t => t.sprintId === sprintId ? { ...t, sprintId: null } : t);
    p.sprints = (p.sprints || []).filter(s => s.id !== sprintId);
    if (sprint) logActivity(p, 'sprint_deleted', `Sprint "${sprint.name}" supprimé`);
  });
}

function toggleSprintStatus(projectId, sprintId) {
  mutateProject(projectId, p => {
    const sprint = (p.sprints || []).find(s => s.id === sprintId);
    if (!sprint) return;
    const newStatus = sprint.status === 'done' ? 'active' : 'done';
    p.sprints = (p.sprints || []).map(s => s.id === sprintId ? { ...s, status: newStatus } : s);
    if (newStatus === 'done') logActivity(p, 'sprint_completed', `Sprint "${sprint.name}" terminé 🏁`);
  });
}

function moveTaskToSprint(projectId, taskId, sprintId) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t =>
      t.id === taskId ? { ...t, sprintId: sprintId || null } : t
    );
  });
}

function addTask(projectId, title, sprintId = null) {
  mutateProject(projectId, p => {
    p.tasks = [...(p.tasks || []), {
      id: uuid(), title, status: 'todo', subtasks: [],
      sprintId: sprintId || null, completedAt: null,
    }];
  });
}

function updateTaskTitle(projectId, taskId, title) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t => t.id === taskId ? { ...t, title: title.trim() } : t);
  });
}

function cycleTaskStatus(projectId, taskId) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t => {
      if (t.id !== taskId) return t;
      const idx = TASK_STATUS_ORDER.indexOf(t.status || 'todo');
      const next = TASK_STATUS_ORDER[(idx + 1) % TASK_STATUS_ORDER.length];
      const completedAt = next === 'done' ? new Date().toISOString() : null;
      if (next === 'done') logActivity(p, 'task_done', `"${t.title}" terminée`);
      return { ...t, status: next, completedAt };
    });
  });
}

function deleteTask(projectId, taskId) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).filter(t => t.id !== taskId);
  });
}

function toggleTaskExpand(projectId, taskId) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t => t.id === taskId ? { ...t, _open: !t._open } : t);
  });
}

function addSubtask(projectId, taskId) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t => {
      if (t.id !== taskId) return t;
      const subtasks = [...(t.subtasks || []), { id: uuid(), title: '', done: false }];
      return { ...t, subtasks, _open: true };
    });
  });
}

function toggleSubtask(projectId, taskId, subId) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t => {
      if (t.id !== taskId) return t;
      const subtasks = (t.subtasks || []).map(s => s.id === subId ? { ...s, done: !s.done } : s);
      return { ...t, subtasks };
    });
  });
}

function updateSubtaskTitle(projectId, taskId, subId, title) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t => {
      if (t.id !== taskId) return t;
      const subtasks = (t.subtasks || []).map(s => s.id === subId ? { ...s, title: title.trim() } : s);
      return { ...t, subtasks };
    });
  });
}

function deleteSubtask(projectId, taskId, subId) {
  mutateProject(projectId, p => {
    p.tasks = (p.tasks || []).map(t => {
      if (t.id !== taskId) return t;
      return { ...t, subtasks: (t.subtasks || []).filter(s => s.id !== subId) };
    });
  });
}

// ── Notes (markdown) ─────────────────────────────────────────────────────────

function silentNotesUpdate(projectId, value) {
  const idx = state.projects.findIndex(p => p.id === projectId);
  if (idx < 0) return;
  state.projects[idx] = { ...state.projects[idx], notes: value, updatedAt: new Date().toISOString() };
}

function setupNotesArea(p) {
  if (state.notesMode === 'edit') {
    const ta = document.getElementById('pv-notes');
    if (!ta) return;
    ta.value = p.notes || '';
    ta.addEventListener('input', () => {
      silentNotesUpdate(p.id, ta.value);
      scheduleSave();
    });
  } else {
    const preview = document.getElementById('pv-notes-preview');
    if (!preview) return;
    if (!p.notes || !p.notes.trim()) {
      preview.textContent = 'Aucune note. Passe en mode Édition pour en ajouter.';
    } else {
      preview.innerHTML = marked.parse(p.notes);
      preview.querySelectorAll('a').forEach(a => {
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
      });
    }
  }
}

// ── Export notes as markdown ─────────────────────────────────────────────────

function slugify(s) {
  return String(s || 'projet')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    || 'projet';
}

function exportNotesAsMd(p) {
  const sm = STATUS_META[p.status] || STATUS_META.active;
  const tasks = p.tasks || [];
  const tasksMd = tasks.length ? tasks.map(t => {
    const box = t.status === 'done' ? '[x]' : t.status === 'doing' ? '[~]' : '[ ]';
    let line = `- ${box} ${t.title || ''}`;
    if (t.subtasks?.length) {
      line += '\n' + t.subtasks.map(s => `  - [${s.done ? 'x' : ' '}] ${s.title || ''}`).join('\n');
    }
    return line;
  }).join('\n') : '_Aucune tâche._';

  const meta = [
    `# ${p.name}`,
    '',
    `> **Statut** · ${sm.label} | **Priorité** · ${PRIORITY_LABEL[p.priority] || 'Basse'} | **Avancement** · ${displayProgress(p)}%`,
    p.tech ? `> **Techno** · ${p.tech}` : null,
    p.tags?.length ? `> **Tags** · ${p.tags.join(', ')}` : null,
    p.githubUrl ? `> **GitHub** · ${p.githubUrl}` : null,
    '',
    p.description ? p.description : null,
    '',
    '## Tâches',
    '',
    tasksMd,
    '',
    '## Notes',
    '',
    p.notes || '_Aucune note._',
    '',
  ].filter(l => l !== null).join('\n');

  const blob = new Blob([meta], { type: 'text/markdown;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${slugify(p.name)}.md`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  showToast(`Notes exportées : ${a.download}`, 'ok');
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

  const mode = existing?.progressMode === 'auto' ? 'auto' : 'manual';
  applyProgressMode(dlg, mode, existing);

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

function applyProgressMode(dlg, mode, existing = null) {
  const row = dlg.querySelector('#f-progress-row');
  const hint = dlg.querySelector('#f-progress-hint');
  const buttons = dlg.querySelectorAll('.progress-mode-toggle .mode-btn');
  buttons.forEach(b => {
    const active = b.dataset.mode === mode;
    b.classList.toggle('is-active', active);
    b.setAttribute('aria-selected', active ? 'true' : 'false');
  });
  if (mode === 'auto') {
    if (row) row.hidden = true;
    if (hint) {
      hint.hidden = false;
      const auto = existing ? computeAutoProgress(existing) : 0;
      const taskCount = existing?.tasks?.length || 0;
      hint.textContent = taskCount
        ? `Calculé automatiquement à partir des tâches (${auto}% — ${taskCount} tâche${taskCount > 1 ? 's' : ''}).`
        : `Calculé automatiquement à partir des tâches. Aucune tâche pour l'instant.`;
    }
  } else {
    if (row) row.hidden = false;
    if (hint) hint.hidden = true;
  }
  dlg.dataset.progressMode = mode;
}

function setupProjectDialog() {
  const dlg = document.getElementById('project-dialog');
  if (!dlg) return;

  dlg.querySelector('#f-progress').addEventListener('input', e => {
    dlg.querySelector('#f-progress-val').textContent = e.target.value + '%';
  });

  dlg.querySelectorAll('.progress-mode-toggle .mode-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = dlg.querySelector('#btn-dlg-save').dataset.id;
      const existing = id ? state.projects.find(p => p.id === id) : null;
      applyProgressMode(dlg, btn.dataset.mode, existing);
    });
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

    const progressMode = dlg.dataset.progressMode === 'auto' ? 'auto' : 'manual';
    const newStatus = dlg.querySelector('#f-status').value;

    const activityLog = [...(existing?.activityLog || [])];
    if (existing && existing.status !== newStatus) {
      const sm = STATUS_META[newStatus] || STATUS_META.active;
      activityLog.unshift({ type: 'status_changed', text: `Statut changé en "${sm.label}"`, date: now });
      if (activityLog.length > 30) activityLog.length = 30;
    }

    const project = {
      id,
      name,
      description: dlg.querySelector('#f-description').value.trim(),
      status:      newStatus,
      priority:    dlg.querySelector('#f-priority').value,
      progress:    Number(dlg.querySelector('#f-progress').value),
      progressMode,
      tech:        dlg.querySelector('#f-tech').value.trim(),
      tags,
      notes:       dlg.querySelector('#f-notes').value.trim(),
      githubUrl:   dlg.querySelector('#f-github').value.trim(),
      tasks:       existing?.tasks || [],
      sprints:     existing?.sprints || [],
      activityLog,
      pinned:      existing?.pinned ?? false,
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

// ── Sprint dialog ─────────────────────────────────────────────────────────────

function openSprintDialog(projectId, existing = null) {
  const dlg = document.getElementById('sprint-dialog');
  if (!dlg) return;
  const isNew = !existing;
  dlg.querySelector('.dialog-title').textContent = isNew ? 'Nouveau sprint' : 'Modifier le sprint';
  dlg.querySelector('#sf-name').value    = existing?.name      || '';
  dlg.querySelector('#sf-start').value   = existing?.startDate?.slice(0, 10) || '';
  dlg.querySelector('#sf-end').value     = existing?.endDate?.slice(0, 10)   || '';
  dlg.querySelector('#sf-status').value  = existing?.status    || 'active';
  const btnDelete = dlg.querySelector('#btn-sprint-delete');
  btnDelete.hidden = isNew;
  btnDelete.dataset.projectId = projectId;
  btnDelete.dataset.sprintId  = existing?.id || '';
  const btnSave = dlg.querySelector('#btn-sprint-save');
  btnSave.dataset.projectId = projectId;
  btnSave.dataset.sprintId  = existing?.id || '';
  btnSave.dataset.isNew     = isNew ? '1' : '';
  dlg.showModal();
  dlg.querySelector('#sf-name').focus();
}

function setupSprintDialog() {
  const dlg = document.getElementById('sprint-dialog');
  if (!dlg) return;
  dlg.querySelector('#btn-sprint-cancel').addEventListener('click', () => dlg.close());
  dlg.querySelector('.dialog-close').addEventListener('click',       () => dlg.close());
  dlg.querySelector('#btn-sprint-delete').addEventListener('click', e => {
    const { projectId, sprintId } = e.currentTarget.dataset;
    dlg.close();
    deleteSprint(projectId, sprintId);
    if (state.currentSprintTab === sprintId) state.currentSprintTab = 'backlog';
    render();
    showToast('Sprint supprimé', 'ok');
  });
  dlg.querySelector('#btn-sprint-save').addEventListener('click', () => {
    const name      = dlg.querySelector('#sf-name').value.trim();
    const startDate = dlg.querySelector('#sf-start').value;
    const endDate   = dlg.querySelector('#sf-end').value;
    const status    = dlg.querySelector('#sf-status').value;
    if (!name)      { dlg.querySelector('#sf-name').focus();  return; }
    if (!startDate) { dlg.querySelector('#sf-start').focus(); return; }
    if (!endDate)   { dlg.querySelector('#sf-end').focus();   return; }
    if (endDate < startDate) {
      showToast('La date de fin doit être après le début', 'error');
      dlg.querySelector('#sf-end').focus();
      return;
    }
    const btn = dlg.querySelector('#btn-sprint-save');
    const { projectId } = btn.dataset;
    const isNew    = btn.dataset.isNew === '1';
    const sprintId = isNew ? uuid() : btn.dataset.sprintId;
    const sprint   = { id: sprintId, name, startDate, endDate, status };
    dlg.close();
    if (isNew) {
      addSprint(projectId, sprint);
      state.currentSprintTab = sprintId;
    } else {
      updateSprint(projectId, sprint);
    }
    render();
    showToast(isNew ? `Sprint "${name}" créé` : `Sprint "${name}" mis à jour`, 'ok');
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
    if (e.key === 'Escape' && state.currentProjectId) {
      const dlg = document.getElementById('project-dialog');
      const cdlg = document.getElementById('confirm-dialog');
      if (!dlg?.open && !cdlg?.open) {
        e.preventDefault();
        closeProjectView();
        return;
      }
    }
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable) return;
    if (e.metaKey || e.ctrlKey || e.altKey) {
      if ((e.metaKey || e.ctrlKey) && e.key === 'k') {
        e.preventDefault();
        document.getElementById('search-input')?.focus();
      }
      return;
    }
    const dlgOpen = document.getElementById('project-dialog')?.open;
    const cdlgOpen = document.getElementById('confirm-dialog')?.open;
    if (dlgOpen || cdlgOpen) return;

    if (state.currentProjectId) {
      const proj = state.projects.find(p => p.id === state.currentProjectId);
      if (e.key === 't' || e.key === 'T') {
        e.preventDefault();
        document.getElementById('pv-task-add')?.focus();
      } else if (e.key === 'e' || e.key === 'E') {
        e.preventDefault();
        if (proj) openProjectDialog(proj);
      } else if (e.key === 'p' || e.key === 'P') {
        e.preventDefault();
        if (proj) togglePin(proj.id);
      } else if (e.key === 's' || e.key === 'S') {
        e.preventDefault();
        if (proj) openSprintDialog(proj.id);
      }
      return;
    }

    if (e.key === 'n') openProjectDialog();
    if (e.key === 'r') refreshData();
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
  setupSprintDialog();
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
