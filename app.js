'use strict';

/* ============================================================
   My Scheduler — app.js
   All state lives in localStorage. No external dependencies.
   ============================================================ */

// ── Constants ──────────────────────────────────────────────
const DAYS         = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
const START_HOUR   = 6;    // 6:00 AM
const END_HOUR     = 18;   // 6:00 PM
const SLOT_MIN     = 5;    // minutes per grid row
const TOTAL_SLOTS  = ((END_HOUR - START_HOUR) * 60) / SLOT_MIN;  // 144
const STORAGE_KEY   = 'cm_scheduler_v1';
const COL_WIDTH_KEY = 'cm_col_widths';
const COL_ZOOM_KEY  = 'cm_col_zoom';
const BASE_UNIT_PX  = 12;   // px per slot at 100 % zoom
const BASE_COL_PX   = 150;  // default column width at 100 % col-zoom
let   _idCounter   = Date.now();

const PERSON_COLORS = [
  '#c8e6d4','#b8d4e8','#e8d4c8','#d4c8e8','#e8e4c8',
  '#c8d4e8','#e8c8d4','#c8e8e4','#ddc8e8','#e4d0c8'
];

// ── State ──────────────────────────────────────────────────
let state = {
  people:        [],
  selectedDays:  ['Monday'],
  selectedPeople:[],
  cards:         [],
  // schedule[day][personName] = [ { instanceId, cardId, slotIndex } ]
  schedule:      {},
  personColors:  {}
};

let zoomLevel      = 100;    // percent (time / vertical zoom)
let colZoom        = 100;    // percent (person/day / horizontal zoom)
let dragState      = null;   // active drag info
let gridSort       = 'day';  // 'day' | 'person'
let _collapsePanel = null;   // set by initPanelResizer

const selectedCardIds = new Set(); // card IDs currently highlighted in the left panel

function clearCardSelection() {
  selectedCardIds.clear();
  document.querySelectorAll('.activity-card.card-selected')
    .forEach(el => el.classList.remove('card-selected'));
}

function toggleCardSelection(cardId) {
  const el = document.querySelector(`[data-card-id="${CSS.escape(cardId)}"]`);
  if (selectedCardIds.has(cardId)) {
    selectedCardIds.delete(cardId);
    el?.classList.remove('card-selected');
  } else {
    selectedCardIds.add(cardId);
    el?.classList.add('card-selected');
  }
}

const selectedPlacedIds = new Set(); // instanceIds of selected placed cards in the grid

// ── History (undo/redo) ─────────────────────────────────────
const MAX_HISTORY = 50;
const HISTORY_KEY = 'cm_history';
let undoStack = [];   // array of JSON strings (oldest → newest)
let redoStack = [];

function saveHistory() {
  try { sessionStorage.setItem(HISTORY_KEY, JSON.stringify({ u: undoStack, r: redoStack })); }
  catch(e) { /* quota exceeded — silently ignore */ }
}

function loadHistory() {
  try {
    const h = JSON.parse(sessionStorage.getItem(HISTORY_KEY) || '{}');
    undoStack = Array.isArray(h.u) ? h.u : [];
    redoStack = Array.isArray(h.r) ? h.r : [];
  } catch { undoStack = []; redoStack = []; }
}

function updateUndoButtons() {
  const u = document.getElementById('undo-btn');
  const r = document.getElementById('redo-btn');
  if (u) u.disabled = undoStack.length === 0;
  if (r) r.disabled = redoStack.length === 0;
}

// ── Column widths ───────────────────────────────────────────
let colWidths = {};  // keyed by person name → px width

function saveColWidths() {
  try { localStorage.setItem(COL_WIDTH_KEY, JSON.stringify(colWidths)); }
  catch(e) { /* ignore */ }
}

function loadColWidths() {
  try {
    const parsed = JSON.parse(localStorage.getItem(COL_WIDTH_KEY) || '{}');
    colWidths = (parsed && typeof parsed === 'object') ? parsed : {};
  } catch { colWidths = {}; }
}

function saveColZoom() {
  try { localStorage.setItem(COL_ZOOM_KEY, String(colZoom)); }
  catch(e) { /* ignore */ }
}

function loadColZoom() {
  try {
    const v = parseInt(localStorage.getItem(COL_ZOOM_KEY));
    colZoom = (!isNaN(v) && v >= 25 && v <= 400) ? v : 100;
  } catch { colZoom = 100; }
}

const _measureCtx = document.createElement('canvas').getContext('2d');
function measureTextPx(text, font) {
  _measureCtx.font = font;
  return _measureCtx.measureText(text).width;
}

function clearPlacedSelection() {
  selectedPlacedIds.clear();
  document.querySelectorAll('.placed-card.placed-selected')
    .forEach(el => el.classList.remove('placed-selected'));
}

function togglePlacedSelection(instanceId, el) {
  if (selectedPlacedIds.has(instanceId)) {
    selectedPlacedIds.delete(instanceId);
    el?.classList.remove('placed-selected');
  } else {
    selectedPlacedIds.add(instanceId);
    el?.classList.add('placed-selected');
  }
}

// ── Utilities ──────────────────────────────────────────────
function uid() { return 'id' + (++_idCounter); }

/** Escape a string for safe text-node insertion via textContent (no-op,
 *  but also used to sanitize before storing in state). */
function sanitize(raw) {
  return String(raw)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#x27;');
}

/** Strip tags / scripts entirely from user-supplied text before storage. */
function clean(raw) {
  return String(raw).replace(/<[^>]*>/g, '').trim();
}

function unitPx() { return (BASE_UNIT_PX * zoomLevel) / 100; }

function slotToPx(slots) { return slots * unitPx(); }

function minutesToLabel(minutesFromStart) {
  const total = START_HOUR * 60 + minutesFromStart;
  const h = Math.floor(total / 60);
  const m = total % 60;
  const ampm = h >= 12 ? 'PM' : 'AM';
  const dh   = h > 12 ? h - 12 : (h === 0 ? 12 : h);
  return `${dh}:${String(m).padStart(2, '0')} ${ampm}`;
}

function showConfirm(msg, onConfirm, okLabel = 'Confirm') {
  const overlay   = document.getElementById('confirm-modal');
  const msgEl     = document.getElementById('confirm-msg');
  const okBtn     = document.getElementById('confirm-ok');
  const cancelBtn = document.getElementById('confirm-cancel');

  msgEl.textContent = msg;
  okBtn.textContent = okLabel;
  overlay.hidden = false;
  okBtn.focus();

  const close = () => {
    overlay.hidden = true;
    document.removeEventListener('keydown', onKey);
  };
  const onKey = e => {
    if (e.key === 'Escape') { e.preventDefault(); close(); }
    if (e.key === 'Enter')  { e.preventDefault(); close(); onConfirm(); }
  };

  okBtn.onclick     = () => { close(); onConfirm(); };
  cancelBtn.onclick = close;
  overlay.onclick   = e => { if (e.target === overlay) close(); };
  document.addEventListener('keydown', onKey);
}

function showToast(msg) {
  const t = document.createElement('div');
  t.className = 'toast';
  t.setAttribute('role', 'status');
  t.setAttribute('aria-live', 'polite');
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3100);
}

// ── Persistence ────────────────────────────────────────────
function saveState() {
  undoStack.push(JSON.stringify(state));
  if (undoStack.length > MAX_HISTORY) undoStack.shift();
  redoStack = [];
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
  catch(e) { console.error('Save failed', e); }
  saveHistory();
  updateUndoButtons();
}

function undo() {
  if (!undoStack.length) return;
  redoStack.push(JSON.stringify(state));
  state = JSON.parse(undoStack.pop());
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
  catch(e) { console.error('Save failed', e); }
  selectedCardIds.clear();
  selectedPlacedIds.clear();
  saveHistory();
  renderAll();
  updateUndoButtons();
}

function redo() {
  if (!redoStack.length) return;
  undoStack.push(JSON.stringify(state));
  state = JSON.parse(redoStack.pop());
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
  catch(e) { console.error('Save failed', e); }
  selectedCardIds.clear();
  selectedPlacedIds.clear();
  saveHistory();
  renderAll();
  updateUndoButtons();
}

function parseCard(c) {
  return {
    id:          String(c.id   || uid()),
    name:        clean(c.name  || ''),
    description: clean(c.description || ''),
    time:        Math.max(5, Math.min(480, parseInt(c.time) || 30)),
    count:          Math.max(1, Math.min(5,   parseInt(c.count) || 1)),
    assignedPerson: clean(String(c.assignedPerson || '')),
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const p = JSON.parse(raw);
    if (!p || typeof p !== 'object') return;

    state.people = Array.isArray(p.people)
      ? p.people.map(x => clean(String(x))).filter(Boolean)
      : [];

    // Backward compat: old saves use selectedDay (string); new saves use selectedDays (array)
    if (Array.isArray(p.selectedDays)) {
      state.selectedDays = p.selectedDays.filter(d => DAYS.includes(d));
      if (state.selectedDays.length === 0) state.selectedDays = ['Monday'];
    } else if (typeof p.selectedDay === 'string' && DAYS.includes(p.selectedDay)) {
      state.selectedDays = [p.selectedDay];
    } else {
      state.selectedDays = ['Monday'];
    }

    state.selectedPeople = Array.isArray(p.selectedPeople)
      ? p.selectedPeople.map(x => clean(String(x))).filter(x => state.people.includes(x))
      : [];

    state.cards = Array.isArray(p.cards) ? p.cards.map(parseCard) : [];

    state.personColors = {};
    if (p.personColors && typeof p.personColors === 'object') {
      Object.entries(p.personColors).forEach(([k, v]) => {
        if (typeof k === 'string' && /^#[0-9a-fA-F]{3,6}$/.test(v))
          state.personColors[clean(k)] = v;
      });
    }

    // Reconstruct schedule, validating slot indices
    state.schedule = {};
    if (p.schedule && typeof p.schedule === 'object') {
      DAYS.forEach(day => {
        if (!p.schedule[day] || typeof p.schedule[day] !== 'object') return;
        state.schedule[day] = {};
        state.people.forEach(person => {
          const slots = p.schedule[day][person];
          if (!Array.isArray(slots)) return;
          state.schedule[day][person] = slots
            .filter(s => s && typeof s.cardId === 'string' && typeof s.slotIndex === 'number')
            .map(s => ({
              instanceId: String(s.instanceId || uid()),
              cardId:     String(s.cardId),
              slotIndex:  Math.max(0, Math.min(TOTAL_SLOTS - 1, Math.floor(s.slotIndex))),
            }));
        });
      });
    }
  } catch(e) { console.error('Load failed', e); }
  loadHistory();
  updateUndoButtons();
  loadColWidths();
  loadColZoom();
}

// ── Placed-count helpers ────────────────────────────────────
function totalPlaced(cardId) {
  let n = 0;
  Object.values(state.schedule).forEach(day =>
    Object.values(day).forEach(arr =>
      arr.forEach(s => { if (s.cardId === cardId) n++; })));
  return n;
}

// ── Render: Day buttons ─────────────────────────────────────
function renderDayButtons() {
  document.querySelectorAll('.day-btn').forEach(btn => {
    const active = state.selectedDays.includes(btn.dataset.day);
    btn.classList.toggle('active', active);
    btn.setAttribute('aria-pressed', String(active));
  });
}

function updateLayoutClass() {
  const multiDay = state.selectedDays.length > 1;
  const ca = document.querySelector('.content-area');
  if (ca) ca.classList.toggle('multi-day', multiDay);
  const sortSel = document.getElementById('grid-sort-select');
  if (sortSel) sortSel.hidden = !multiDay;
}

// ── Render: Name buttons ────────────────────────────────────
function renderNameButtons() {
  const list = document.getElementById('name-btn-list');
  list.innerHTML = '';
  state.people.forEach(person => {
    const btn = document.createElement('button');
    btn.className = 'name-btn';
    btn.textContent = person;
    const sel = state.selectedPeople.includes(person);
    btn.classList.toggle('selected', sel);
    btn.setAttribute('aria-pressed', String(sel));
    btn.setAttribute('aria-label', `${person} — click to toggle, double-click or long-press to rename`);
    const pColor = state.personColors[person];
    if (pColor) btn.style.backgroundColor = pColor;

    btn.addEventListener('click', () => {
      const idx = state.selectedPeople.indexOf(person);
      if (idx > -1) {
        state.selectedPeople.splice(idx, 1);
      } else {
        state.selectedPeople.push(person);
      }
      const nowSel = state.selectedPeople.includes(person);
      btn.classList.toggle('selected', nowSel);
      btn.setAttribute('aria-pressed', String(nowSel));
      saveState();
      renderGrid();
    });

    btn.addEventListener('dblclick', e => {
      e.preventDefault();
      inlineEditPerson(btn, person);
    });

    // Long-press (touch) triggers the same inline edit
    let longPressTimer = null;
    btn.addEventListener('touchstart', () => {
      longPressTimer = setTimeout(() => {
        longPressTimer = null;
        inlineEditPerson(btn, person);
      }, 500);
    }, { passive: true });
    btn.addEventListener('touchend',  () => { clearTimeout(longPressTimer); longPressTimer = null; });
    btn.addEventListener('touchmove', () => { clearTimeout(longPressTimer); longPressTimer = null; });

    list.appendChild(btn);
  });

  // Keep the person dropdown in sync with the current people list
  const personSel = document.getElementById('new-card-person');
  if (personSel) {
    const prev = personSel.value;
    personSel.innerHTML = '';
    const noneOpt = document.createElement('option');
    noneOpt.value = ''; noneOpt.textContent = '— None —';
    personSel.appendChild(noneOpt);
    state.people.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p; opt.textContent = p;
      personSel.appendChild(opt);
    });
    personSel.value = state.people.includes(prev) ? prev : '';
  }
}

function inlineEditPerson(btn, oldName) {
  // Wrapper holds the input + color picker + delete button side by side
  const wrapper = document.createElement('div');
  wrapper.className = 'name-btn-edit-wrapper';

  const input = document.createElement('input');
  input.type = 'text';
  input.value = oldName;
  input.className = 'name-btn-edit-input';
  input.maxLength = 40;
  input.setAttribute('aria-label', 'Rename person — press Enter to save, Escape to cancel');

  const colorPick = document.createElement('input');
  colorPick.type = 'color';
  colorPick.className = 'person-color-input';
  colorPick.value = state.personColors[oldName]
    || PERSON_COLORS[Math.max(0, state.people.indexOf(oldName)) % PERSON_COLORS.length];
  colorPick.title = 'Change person color';
  colorPick.setAttribute('aria-label', `Color for ${oldName}`);

  const delBtn = document.createElement('button');
  delBtn.className = 'name-btn-delete-btn';
  delBtn.textContent = '×';
  delBtn.setAttribute('aria-label', `Delete ${oldName}`);

  wrapper.appendChild(input);
  wrapper.appendChild(colorPick);
  wrapper.appendChild(delBtn);
  btn.replaceWith(wrapper);
  input.focus();
  input.select();

  // After a long-press the flex layout shifts while the finger is still down,
  // causing a spurious blur as soon as the touch lifts. Ignore blur for a
  // short window so that reflow-induced blurs can't immediately close the edit.
  let blurEnabled = false;
  setTimeout(() => { blurEnabled = true; }, 350);

  const commit = () => {
    const newName  = clean(input.value);
    const newColor = colorPick.value;
    if (newName && newName !== oldName && !state.people.includes(newName)) {
      const pi = state.people.indexOf(oldName);
      if (pi > -1) state.people[pi] = newName;
      const si = state.selectedPeople.indexOf(oldName);
      if (si > -1) state.selectedPeople[si] = newName;
      DAYS.forEach(day => {
        if (state.schedule[day] && state.schedule[day][oldName]) {
          state.schedule[day][newName] = state.schedule[day][oldName];
          delete state.schedule[day][oldName];
        }
      });
      state.personColors[newName] = newColor;
      delete state.personColors[oldName];
      state.cards.forEach(c => { if (c.assignedPerson === oldName) c.assignedPerson = newName; });
    } else if (newName) {
      state.personColors[oldName] = newColor;
    }
    saveState();
    renderNameButtons();
    renderCards();
    renderGrid();
  };

  // mousedown (desktop) and touchstart (touch devices) both fire before the
  // text input's blur event, so we suppress commit in both. change fires when
  // the user finishes picking; blur on colorPick re-enables for the cancel case.
  const suppressBlur = () => {
    blurEnabled = false;
    setTimeout(() => { blurEnabled = true; }, 5000); // safety if OS dialog is cancelled
  };
  colorPick.addEventListener('mousedown', suppressBlur);
  colorPick.addEventListener('touchstart', suppressBlur, { passive: true });
  colorPick.addEventListener('change', () => {
    blurEnabled = true;
    input.focus();
  });
  colorPick.addEventListener('blur', () => {
    blurEnabled = true; // re-enable if color dialog cancelled without picking
  });

  input.addEventListener('blur', () => { if (blurEnabled) commit(); });
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter')  commit();
    if (e.key === 'Escape') renderNameButtons();
  });

  // Prevent input blur when clicking delete so the click still fires
  delBtn.addEventListener('mousedown', e => e.preventDefault());
  delBtn.addEventListener('click', () => {
    blurEnabled = false; // prevent spurious commit when wrapper is removed
    wrapper.remove();
    deletePerson(oldName);
  });
}

function deletePerson(name) {
  state.people         = state.people.filter(p => p !== name);
  state.selectedPeople = state.selectedPeople.filter(p => p !== name);
  DAYS.forEach(day => {
    if (state.schedule[day]) delete state.schedule[day][name];
  });
  delete state.personColors[name];
  state.cards.forEach(c => { if (c.assignedPerson === name) c.assignedPerson = ''; });
  saveState();
  renderNameButtons();
  renderCards();
  renderGrid();
}

// ── Render: Activity cards (left panel) ────────────────────
function renderCards() {
  const list = document.getElementById('card-list');
  list.innerHTML = '';
  state.cards.forEach(card => {
    const el = buildCardEl(card);
    if (selectedCardIds.has(card.id)) el.classList.add('card-selected');
    list.appendChild(el);
  });
}

function buildCardEl(card) {
  const placed     = totalPlaced(card.id);
  const remaining  = card.count - placed;
  const fullyPlaced = remaining <= 0;
  const heightPx   = Math.max(72, slotToPx(card.time / SLOT_MIN));

  const el = document.createElement('div');
  el.className = 'activity-card' + (fullyPlaced ? ' fully-placed' : '');
  el.setAttribute('role', 'listitem');
  el.dataset.cardId = card.id;
  el.style.minHeight = heightPx + 'px';
  el.style.backgroundColor = state.personColors[card.assignedPerson] || '';
  el.setAttribute('aria-label',
    `${card.name}: ${card.description}. ${card.time} min. ${remaining} of ${card.count} remaining. Drag to schedule.`);

  // ── reorder handle (left edge) ──
  const handle = el.appendChild(document.createElement('div'));
  handle.className = 'card-drag-handle';
  handle.setAttribute('aria-hidden', 'true');
  handle.title = 'Drag to reorder';
  handle.textContent = '⠿';

  handle.addEventListener('mousedown', e => {
    e.stopPropagation();
    startReorder(e, card.id, el);
  });
  handle.addEventListener('touchstart', e => {
    e.stopPropagation();
    startReorderTouch(e, card.id, el);
  }, { passive: false });

  // ── header row ──
  const header = el.appendChild(document.createElement('div'));
  header.className = 'card-header';

  const nameSpan = header.appendChild(document.createElement('span'));
  nameSpan.className = 'card-name';
  nameSpan.textContent = card.name;
  nameSpan.setAttribute('aria-label', 'Double-click to edit name');
  nameSpan.addEventListener('dblclick', e => { e.stopPropagation(); inlineEditField(nameSpan, card, 'name'); });

  const badge = header.appendChild(document.createElement('span'));
  badge.className = 'card-count-badge';
  badge.textContent = `${remaining}/${card.count}`;
  badge.setAttribute('aria-label', `${remaining} of ${card.count} remaining`);

  // ── description ──
  const desc = el.appendChild(document.createElement('div'));
  desc.className = 'card-desc';
  desc.textContent = card.description || '';
  desc.setAttribute('aria-label', 'Double-click to edit description');
  desc.addEventListener('dblclick', e => { e.stopPropagation(); inlineEditField(desc, card, 'description'); });

  // ── time label ──
  const timeLabel = el.appendChild(document.createElement('div'));
  timeLabel.className = 'card-time-label';
  timeLabel.textContent = card.time + ' min';
  timeLabel.setAttribute('aria-label', 'Double-click to edit duration');
  timeLabel.addEventListener('dblclick', e => { e.stopPropagation(); inlineEditField(timeLabel, card, 'time'); });

  // ── actions ──
  const actions = el.appendChild(document.createElement('div'));
  actions.className = 'card-actions';

  const assignBtn = actions.appendChild(document.createElement('button'));
  assignBtn.className = 'card-action-btn';
  assignBtn.textContent = card.assignedPerson || 'Person';
  assignBtn.setAttribute('aria-label', `Assign person for ${card.name}`);
  assignBtn.addEventListener('click', e => {
    e.stopPropagation();
    if (state.people.length === 0) {
      showToast('Add people using + Person before assigning');
      return;
    }
    inlineEditAssignedPerson(assignBtn, card);
  });

  const countBtn = actions.appendChild(document.createElement('button'));
  countBtn.className = 'card-action-btn';
  countBtn.textContent = 'Count';
  countBtn.setAttribute('aria-label', `Edit count for ${card.name}`);
  countBtn.addEventListener('click', e => { e.stopPropagation(); inlineEditCount(badge, card); });

  const delBtn = actions.appendChild(document.createElement('button'));
  delBtn.className = 'card-action-btn delete-btn';
  delBtn.textContent = 'Delete';
  delBtn.setAttribute('aria-label', `Delete ${card.name}`);
  delBtn.addEventListener('click', e => { e.stopPropagation(); deleteCard(card.id); });

  // ── drag to grid listeners (body only, not handle/buttons/inputs) ──
  el.addEventListener('mousedown', e => {
    if (e.target.closest('button, input, select, .card-drag-handle')) return;
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      e.stopPropagation();
      toggleCardSelection(card.id);
      return;
    }
    clearCardSelection();
    startDrag(e, card.id, el);
  });
  el.addEventListener('touchstart', e => {
    if (e.target.closest('button, input, select, .card-drag-handle')) return;
    startDragTouch(e, card.id, el);
  }, { passive: false });

  return el;
}

function inlineEditField(span, card, field) {
  const input = document.createElement('input');
  input.type = field === 'time' ? 'number' : 'text';
  input.value = card[field];
  input.className = 'inline-edit-input';
  if (field === 'time') { input.min = '5'; input.max = '480'; input.step = '5'; }
  if (field === 'name') { input.maxLength = 60; }
  if (field === 'description') { input.maxLength = 120; }
  input.setAttribute('aria-label', `Edit ${field}`);
  span.replaceWith(input);
  input.focus(); input.select();

  const commit = () => {
    if (field === 'time') {
      const v = parseInt(input.value);
      if (!isNaN(v) && v >= 5) card.time = v;
    } else {
      const v = clean(input.value);
      if (v) card[field] = v;
    }
    saveState();
    renderCards();
    renderGrid();
  };

  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter')  input.blur();
    if (e.key === 'Escape') renderCards();
  });
}

function inlineEditCount(badge, card) {
  const isMobile = window.innerWidth < 768;

  if (isMobile) {
    const select = document.createElement('select');
    select.className = 'inline-edit-select';
    for (let i = 1; i <= 5; i++) {
      const opt = document.createElement('option');
      opt.value = String(i);
      opt.textContent = String(i);
      if (i === card.count) opt.selected = true;
      select.appendChild(opt);
    }
    badge.replaceWith(select);
    select.focus();

    let committed = false;
    const commit = () => {
      if (committed) return;
      committed = true;
      card.count = parseInt(select.value);
      saveState();
      renderCards();
      renderGrid();
    };
    select.addEventListener('change', commit);
    select.addEventListener('blur', () => setTimeout(commit, 0));
    select.addEventListener('keydown', e => {
      if (e.key === 'Escape') { committed = true; renderCards(); }
    });
  } else {
    const input = document.createElement('input');
    input.type = 'number';
    input.value = card.count;
    input.min = '1'; input.max = '5';
    input.className = 'inline-edit-input';
    input.style.width = '50px';
    input.setAttribute('aria-label', 'Edit count');
    badge.replaceWith(input);
    input.focus(); input.select();

    const clamp = () => {
      const v = parseInt(input.value);
      if (!isNaN(v)) input.value = Math.min(5, Math.max(1, v));
    };
    const commit = () => {
      clamp();
      const v = parseInt(input.value);
      if (!isNaN(v)) card.count = Math.min(5, Math.max(1, v));
      saveState();
      renderCards();
      renderGrid();
    };
    input.addEventListener('input', clamp);
    input.addEventListener('blur', commit);
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter')  input.blur();
      if (e.key === 'Escape') renderCards();
    });
  }
}

function inlineEditAssignedPerson(btn, card) {
  const select = document.createElement('select');
  select.className = 'inline-edit-select';

  const noneOpt = document.createElement('option');
  noneOpt.value = '';
  noneOpt.textContent = '— None —';
  select.appendChild(noneOpt);

  state.people.forEach(p => {
    const opt = document.createElement('option');
    opt.value = p;
    opt.textContent = p;
    if (p === card.assignedPerson) opt.selected = true;
    select.appendChild(opt);
  });

  select.value = card.assignedPerson;
  btn.replaceWith(select);
  select.focus();

  let committed = false;
  const commit = () => {
    if (committed) return;
    committed = true;
    card.assignedPerson = select.value;
    saveState();
    renderCards();
    renderGrid();
  };

  select.addEventListener('change', commit);
  // Defer blur-commit so 'change' always fires first (avoids browser ordering differences)
  select.addEventListener('blur', () => setTimeout(commit, 0));
  select.addEventListener('keydown', e => {
    if (e.key === 'Escape') { committed = true; renderCards(); }
  });
}

function deleteCard(cardId) {
  selectedCardIds.delete(cardId);
  state.cards = state.cards.filter(c => c.id !== cardId);
  DAYS.forEach(day => {
    if (!state.schedule[day]) return;
    Object.keys(state.schedule[day]).forEach(person => {
      state.schedule[day][person] = state.schedule[day][person]
        .filter(s => s.cardId !== cardId);
    });
  });
  saveState();
  renderCards();
  renderGrid();
}

function updateBadges() {
  document.querySelectorAll('.activity-card').forEach(el => {
    const card = state.cards.find(c => c.id === el.dataset.cardId);
    if (!card) return;
    const placed    = totalPlaced(card.id);
    const remaining = card.count - placed;
    const badge  = el.querySelector('.card-count-badge');
    if (badge) {
      badge.textContent = `${remaining}/${card.count}`;
      badge.setAttribute('aria-label', `${remaining} of ${card.count} remaining`);
    }
    el.classList.toggle('fully-placed', remaining <= 0);
    // Update height in case zoom changed
    const h = Math.max(72, slotToPx(card.time / SLOT_MIN));
    el.style.minHeight = h + 'px';
  });
}

// ── Render: Schedule Grid ───────────────────────────────────
function renderGrid() {
  // Update CSS custom property for grid unit
  document.documentElement.style.setProperty('--grid-unit', unitPx() + 'px');

  const timeCol  = document.getElementById('time-column');
  const gridEl   = document.getElementById('schedule-grid');
  const wrapper  = document.getElementById('schedule-wrapper');
  timeCol.innerHTML    = '';
  gridEl.innerHTML     = '';
  timeCol.style.paddingTop = '';
  timeCol.style.height     = '';
  wrapper.style.maxHeight  = '';

  updateLayoutClass();

  // ── Time labels ──
  for (let i = 0; i < TOTAL_SLOTS; i++) {
    const label = document.createElement('div');
    label.className = 'time-label';
    label.style.height = unitPx() + 'px';
    const mins = i * SLOT_MIN;
    label.textContent = minutesToLabel(mins);
    if (mins % 60 === 0) {
      label.classList.add('hour');
    } else if (i % 2 === 0) {
      label.classList.add('tick-bold');
    } else {
      label.classList.add('tick-light');
    }
    timeCol.appendChild(label);
  }
  // Terminal 6 PM marker: overlaps the last slot so it stays within the column height
  const endLabel = document.createElement('div');
  endLabel.className = 'time-label hour';
  endLabel.style.height  = unitPx() + 'px';
  endLabel.style.marginTop = (-unitPx()) + 'px';
  endLabel.textContent = minutesToLabel(TOTAL_SLOTS * SLOT_MIN);
  timeCol.appendChild(endLabel);

  // ── Grid columns ──
  const isMultiDay = state.selectedDays.length > 1;
  const colPairs = [];
  if (gridSort === 'person') {
    state.selectedPeople.forEach(person =>
      state.selectedDays.forEach(day => colPairs.push({ day, person }))
    );
  } else {
    state.selectedDays.forEach(day =>
      state.selectedPeople.forEach(person => colPairs.push({ day, person }))
    );
  }
  colPairs.forEach(({ day, person }) => {
    const col = document.createElement('div');
    col.className = 'grid-column';
    col.dataset.person = person;
    col.dataset.day    = day;
    col.setAttribute('role', 'gridcell');
    col.setAttribute('aria-label', `${person}'s schedule for ${day}`);

    // Column header
    const colHeader = col.appendChild(document.createElement('div'));
    colHeader.className = 'col-header';
    colHeader.setAttribute('aria-hidden', 'true');
    const personColor = state.personColors[person];
    if (personColor) {
      colHeader.style.background = personColor;
      colHeader.style.color = '#343232';
      colHeader.style.borderBottomColor = personColor;
    }
    if (isMultiDay) {
      const dayLabel = document.createElement('span');
      dayLabel.className = 'col-day-label';
      dayLabel.textContent = day.slice(0, 3);
      colHeader.appendChild(dayLabel);
    }
    const personSpan = document.createElement('span');
    personSpan.textContent = person;
    colHeader.appendChild(personSpan);

    // Resize handle
    const resizeHandle = document.createElement('div');
    resizeHandle.className = 'col-resize-handle';
    resizeHandle.setAttribute('aria-hidden', 'true');
    colHeader.appendChild(resizeHandle);

    // Apply column width: base width (stored or default) scaled by colZoom
    const baseW = colWidths[person] != null ? colWidths[person] : BASE_COL_PX;
    col.style.flex = 'none';
    col.style.minWidth = '0';
    col.style.width = Math.max(8, Math.round(baseW * colZoom / 100)) + 'px';

    // Drag-to-resize (min = handle width so column can be fully collapsed)
    resizeHandle.addEventListener('mousedown', e => {
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startW = col.offsetWidth;
      const MIN_W = 8; // just enough to show the handle
      document.body.classList.add('col-resizing');

      const onMove = ev => {
        const w = Math.max(MIN_W, startW + ev.clientX - startX);
        col.style.flex = 'none';
        col.style.minWidth = '0';
        col.style.width = w + 'px';
      };
      const onUp = () => {
        document.body.classList.remove('col-resizing');
        colWidths[person] = Math.round(col.offsetWidth * 100 / colZoom);
        saveColWidths();
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });

    // Double-click handle → auto-fit to widest content
    resizeHandle.addEventListener('dblclick', e => {
      e.stopPropagation();
      const headerFont = getComputedStyle(personSpan).font;
      let maxW = 0;
      colHeader.querySelectorAll('span:not(.col-resize-handle)').forEach(s => {
        maxW = Math.max(maxW, measureTextPx(s.textContent, headerFont));
      });
      maxW += 32; // header padding + handle space

      col.querySelectorAll('.placed-name').forEach(el => {
        const font = getComputedStyle(el).font;
        maxW = Math.max(maxW, measureTextPx(el.textContent, font) + 32);
      });

      const w = Math.max(80, Math.ceil(maxW));
      col.style.flex = 'none';
      col.style.minWidth = '0';
      col.style.width = w + 'px';
      colWidths[person] = Math.round(w * 100 / colZoom);
      saveColWidths();
    });

    // Slot container
    const slots = col.appendChild(document.createElement('div'));
    slots.className = 'col-slots';
    const totalH = TOTAL_SLOTS * unitPx();
    slots.style.height = totalH + 'px';

    // Draw dividing lines (not interactive, just visual)
    for (let i = 0; i < TOTAL_SLOTS; i++) {
      const line = document.createElement('div');
      line.className = 'slot-line' + (i * SLOT_MIN % 60 === 0 ? ' hour' : '');
      line.style.top = (i * unitPx()) + 'px';
      slots.appendChild(line);
    }

    // Drop listeners on the column
    col.addEventListener('dragover', e => e.preventDefault());

    // Mouse/touch drop is handled in drag logic; also support native HTML DnD as fallback
    col.addEventListener('drop', e => {
      e.preventDefault();
      if (!dragState) return;
      dropOnColumn(e.clientY, col, dragState.cardId);
    });

    gridEl.appendChild(col);

    // ── Render placed cards ──
    const daySchedule = (state.schedule[day] || {})[person] || [];
    daySchedule.forEach(placement => {
      const card = state.cards.find(c => c.id === placement.cardId);
      if (!card) return;
      placedCardEl(card, placement, person, slots, day);
    });
  });

  updateBadges();

  // After layout: align time column with col-slots and cap wrapper to exact grid height,
  // so no blank space appears below the 6 PM boundary.
  requestAnimationFrame(() => {
    const gc = gridEl.querySelector('.grid-column');
    const ch = gc?.querySelector('.col-header');
    if (gc && ch) {
      timeCol.style.paddingTop = ch.offsetHeight + 'px';
      timeCol.style.height     = gc.offsetHeight + 'px';
      wrapper.style.maxHeight  = gc.offsetHeight + 'px';
    }
  });
}

function placedCardEl(card, placement, person, slotsContainer, day) {
  const heightPx = Math.max(unitPx(), slotToPx(card.time / SLOT_MIN));
  const topPx    = slotToPx(placement.slotIndex);

  const el = document.createElement('div');
  el.className = 'placed-card';
  if (selectedPlacedIds.has(placement.instanceId)) el.classList.add('placed-selected');
  el.style.top    = topPx    + 'px';
  el.style.height = heightPx + 'px';
  const personColor = state.personColors[card.assignedPerson];
  if (personColor) el.style.backgroundColor = personColor;
  el.dataset.instanceId = placement.instanceId;
  el.dataset.day = day;
  el.setAttribute('aria-label',
    `${card.name} at ${minutesToLabel(placement.slotIndex * SLOT_MIN)}, ${card.time} min. Drag to move, click × to remove.`);

  const nameEl = el.appendChild(document.createElement('div'));
  nameEl.className = 'placed-name';
  nameEl.textContent = card.name;

  const durEl = el.appendChild(document.createElement('div'));
  durEl.className = 'placed-duration';
  durEl.textContent = card.time + ' min';

  const removeBtn = el.appendChild(document.createElement('button'));
  removeBtn.className = 'remove-placed-btn';
  removeBtn.textContent = '×';
  removeBtn.setAttribute('aria-label', `Remove ${card.name}`);
  removeBtn.addEventListener('click', e => {
    e.stopPropagation();
    removePlaced(day, person, placement.instanceId);
  });

  // Drag to reposition (with Ctrl+click for multi-select)
  el.addEventListener('mousedown', e => {
    if (e.target.closest('button')) return;
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      togglePlacedSelection(placement.instanceId, el);
      return;
    }
    // If clicking a non-selected card, clear existing selection first
    if (!selectedPlacedIds.has(placement.instanceId)) {
      clearPlacedSelection();
    }
    startGridCardDrag(e, card, placement, person, el);
  });
  el.addEventListener('touchstart', e => {
    if (e.target.closest('button')) return;
    clearPlacedSelection();
    startGridCardDragTouch(e, card, placement, person, el);
  }, { passive: false });

  slotsContainer.appendChild(el);
}

function removePlaced(day, person, instanceId) {
  selectedPlacedIds.delete(instanceId);
  const daySchedule = state.schedule[day];
  if (!daySchedule || !daySchedule[person]) return;
  daySchedule[person] = daySchedule[person].filter(s => s.instanceId !== instanceId);
  saveState();
  renderGrid();
}

// ── Place card into schedule ─────────────────────────────────
function placeCard(cardId, person, slotIndex, day) {
  const card = state.cards.find(c => c.id === cardId);
  if (!card) return;

  const slotsNeeded = card.time / SLOT_MIN;
  slotIndex = Math.max(0, Math.min(TOTAL_SLOTS - slotsNeeded, Math.floor(slotIndex)));
  const endSlot = slotIndex + slotsNeeded;

  if (!state.schedule[day]) state.schedule[day] = {};
  if (!state.schedule[day][person]) state.schedule[day][person] = [];
  const existing = state.schedule[day][person];

  // Total count exhausted check
  if (totalPlaced(cardId) >= card.count) {
    showToast(`${card.name} has no remaining placements (0/${card.count})`);
    return false;
  }

  // One-instance-per-person-per-day check
  const alreadyOnDay = existing.some(p => p.cardId === cardId);
  if (alreadyOnDay) {
    showToast(`${card.name} is already scheduled for ${person} on ${day} — only one per day allowed`);
    return false;
  }

  // Overlap check
  const overlaps = existing.some(p => {
    const ec = state.cards.find(c => c.id === p.cardId);
    if (!ec) return false;
    const eEnd = p.slotIndex + ec.time / SLOT_MIN;
    return !(endSlot <= p.slotIndex || slotIndex >= eEnd);
  });
  if (overlaps) { showToast('Time slot overlaps an existing activity'); return false; }

  existing.push({ instanceId: uid(), cardId, slotIndex });
  saveState();
  renderGrid();
  showToast(`Placed: ${card.name} at ${minutesToLabel(slotIndex * SLOT_MIN)}`);
  return true;
}

// ── Card Reorder (within left panel) ───────────────────────
function applyReorder(cardId, insertBeforeId) {
  const fromIdx = state.cards.findIndex(c => c.id === cardId);
  if (fromIdx === -1) return;
  const [card] = state.cards.splice(fromIdx, 1);
  if (insertBeforeId) {
    const toIdx = state.cards.findIndex(c => c.id === insertBeforeId);
    state.cards.splice(toIdx > -1 ? toIdx : state.cards.length, 0, card);
  } else {
    state.cards.push(card);
  }
  saveState();
  renderCards();
}

function findInsertBefore(clientY, cardList, draggingId) {
  const cards = [...cardList.querySelectorAll('.activity-card:not(.dragging)')];
  for (const card of cards) {
    const r = card.getBoundingClientRect();
    if (clientY < r.top + r.height / 2) return card.dataset.cardId;
  }
  return null;
}

function startReorder(e, cardId, srcEl) {
  e.preventDefault();
  const card = state.cards.find(c => c.id === cardId);
  if (!card) return;

  srcEl.classList.add('dragging');
  const rect = srcEl.getBoundingClientRect();
  const grabOffsetY = e.clientY - rect.top;

  const ghost = buildGhost(card, rect.width);
  ghost.style.left = rect.left + 'px';
  ghost.style.top  = (e.clientY - grabOffsetY) + 'px';

  const indicator = document.createElement('div');
  indicator.className = 'reorder-indicator';
  const cardList = document.getElementById('card-list');

  let insertBeforeId = null;

  const move = e => {
    ghost.style.top = (e.clientY - grabOffsetY) + 'px';
    insertBeforeId = findInsertBefore(e.clientY, cardList, cardId);
    const refEl = insertBeforeId
      ? cardList.querySelector(`[data-card-id="${insertBeforeId}"]`)
      : null;
    if (refEl) cardList.insertBefore(indicator, refEl);
    else        cardList.appendChild(indicator);
  };

  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    ghost.remove();
    indicator.remove();
    srcEl.classList.remove('dragging');
    applyReorder(cardId, insertBeforeId);
  };

  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function startReorderTouch(e, cardId, srcEl) {
  e.preventDefault();
  const card = state.cards.find(c => c.id === cardId);
  if (!card) return;

  const touch = e.touches[0];
  srcEl.classList.add('dragging');
  const rect = srcEl.getBoundingClientRect();
  const grabOffsetY = touch.clientY - rect.top;

  const ghost = buildGhost(card, rect.width);
  ghost.style.left = rect.left + 'px';
  ghost.style.top  = (touch.clientY - grabOffsetY) + 'px';

  const indicator = document.createElement('div');
  indicator.className = 'reorder-indicator';
  const cardList = document.getElementById('card-list');

  let insertBeforeId = null;

  const move = e => {
    const t = e.touches[0];
    ghost.style.top = (t.clientY - grabOffsetY) + 'px';
    insertBeforeId = findInsertBefore(t.clientY, cardList, cardId);
    const refEl = insertBeforeId
      ? cardList.querySelector(`[data-card-id="${insertBeforeId}"]`)
      : null;
    if (refEl) cardList.insertBefore(indicator, refEl);
    else        cardList.appendChild(indicator);
  };

  const end = () => {
    document.removeEventListener('touchmove', move);
    document.removeEventListener('touchend', end);
    ghost.remove();
    indicator.remove();
    srcEl.classList.remove('dragging');
    applyReorder(cardId, insertBeforeId);
  };

  document.addEventListener('touchmove', move, { passive: false });
  document.addEventListener('touchend', end);
}

// ── Grid Card Drag (reposition a placed card) ──────────────

/** Silent version of placeCard — no toast, no renderGrid. Returns true on success. */
function placeCardDirect(cardId, person, slotIndex, day) {
  const card = state.cards.find(c => c.id === cardId);
  if (!card) return false;

  const slotsNeeded = card.time / SLOT_MIN;
  slotIndex = Math.max(0, Math.min(TOTAL_SLOTS - slotsNeeded, Math.floor(slotIndex)));
  const endSlot = slotIndex + slotsNeeded;

  if (!state.schedule[day]) state.schedule[day] = {};
  if (!state.schedule[day][person]) state.schedule[day][person] = [];
  const existing = state.schedule[day][person];

  if (totalPlaced(cardId) >= card.count) return false;
  if (existing.some(p => p.cardId === cardId)) return false;
  const overlaps = existing.some(p => {
    const ec = state.cards.find(c => c.id === p.cardId);
    if (!ec) return false;
    const eEnd = p.slotIndex + ec.time / SLOT_MIN;
    return !(endSlot <= p.slotIndex || slotIndex >= eEnd);
  });
  if (overlaps) return false;

  existing.push({ instanceId: uid(), cardId, slotIndex });
  return true;
}

/** Compute the slot index for a drop at clientY into colEl, using current grabOffsetY. */
function computeDropSlot(clientY, colEl) {
  const slotsEl = colEl.querySelector('.col-slots');
  if (!slotsEl) return 0;
  const rect    = slotsEl.getBoundingClientRect();
  const rawY    = clientY - rect.top;
  const offsetY = dragState ? (dragState.grabOffsetY || 0) : 0;
  return Math.floor((rawY - offsetY) / unitPx());
}

function restorePlacement(savedPlacement, savedPerson, day) {
  if (!state.schedule[day]) state.schedule[day] = {};
  if (!state.schedule[day][savedPerson]) state.schedule[day][savedPerson] = [];
  state.schedule[day][savedPerson].push(savedPlacement);
  saveState();
  renderGrid();
}

function startGridCardDrag(e, card, placement, person, srcEl) {
  e.preventDefault();

  // Check if this is a multi-card drag (dragged card is in a selection of 2+)
  const isMulti = selectedPlacedIds.size > 1 && selectedPlacedIds.has(placement.instanceId);

  // Collect and remove all other selected placements from state
  const extras = [];
  if (isMulti) {
    DAYS.forEach(day => {
      if (!state.schedule[day]) return;
      Object.keys(state.schedule[day]).forEach(p => {
        if (!state.schedule[day][p]) return;
        state.schedule[day][p].forEach(s => {
          if (s.instanceId !== placement.instanceId && selectedPlacedIds.has(s.instanceId)) {
            const c = state.cards.find(c => c.id === s.cardId);
            if (c) extras.push({ placement: { ...s }, person: p, day, card: c,
              slotOffset: s.slotIndex - placement.slotIndex });
          }
        });
        // Remove extras from state
        state.schedule[day][p] = state.schedule[day][p].filter(
          s => !selectedPlacedIds.has(s.instanceId) || s.instanceId === placement.instanceId
        );
      });
    });
  }

  // Temporarily remove primary card from state so overlap check ignores it
  const savedPlacement = { ...placement };
  const savedDay = srcEl.dataset.day;
  const daySchedule = state.schedule[savedDay];
  if (daySchedule && daySchedule[person]) {
    daySchedule[person] = daySchedule[person].filter(s => s.instanceId !== placement.instanceId);
  }

  srcEl.classList.add('dragging');
  const rect = srcEl.getBoundingClientRect();
  const grabOffsetY = e.clientY - rect.top;

  const ghost = buildGhost(card, rect.width, isMulti ? extras.length : 0);
  ghost.style.left = (e.clientX - 16) + 'px';
  ghost.style.top  = (e.clientY - grabOffsetY) + 'px';

  dragState = { cardId: card.id, ghost, grabOffsetY };

  const move = e => {
    ghost.style.left = (e.clientX - 16) + 'px';
    ghost.style.top  = (e.clientY - grabOffsetY) + 'px';
    highlightColumn(e.clientX, e.clientY);
    autoScrollUpdate(e.clientY);
  };

  const up = e => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    autoScrollStop();
    clearHighlight();
    ghost.remove();
    srcEl.classList.remove('dragging');

    const target = document.elementFromPoint(e.clientX, e.clientY);
    const col    = target?.closest('.grid-column');

    if (isMulti) {
      // Multi-drag: main card goes to the drop target; each extra shifts by the same
      // time offset but stays in its own original column (avoids slot conflicts when
      // cards from different columns share the same start time).
      let mainOk = false;
      if (col) {
        const targetSlot = computeDropSlot(e.clientY, col);
        mainOk = placeCardDirect(card.id, col.dataset.person, targetSlot, col.dataset.day);
        if (mainOk) {
          const timeShift = targetSlot - savedPlacement.slotIndex;
          const failedExtras = [];
          extras.forEach(ex => {
            const newSlot = ex.placement.slotIndex + timeShift;
            const ok = placeCardDirect(ex.card.id, ex.person, newSlot, ex.day);
            if (!ok) failedExtras.push(ex);
          });
          // Restore any extras that couldn't be placed
          failedExtras.forEach(ex => {
            if (!state.schedule[ex.day]) state.schedule[ex.day] = {};
            if (!state.schedule[ex.day][ex.person]) state.schedule[ex.day][ex.person] = [];
            state.schedule[ex.day][ex.person].push(ex.placement);
          });
          showToast(`Moved ${1 + extras.length - failedExtras.length} cards`);
        }
      }
      if (!mainOk) {
        // Restore everything
        if (!state.schedule[savedDay]) state.schedule[savedDay] = {};
        if (!state.schedule[savedDay][person]) state.schedule[savedDay][person] = [];
        state.schedule[savedDay][person].push(savedPlacement);
        extras.forEach(ex => {
          if (!state.schedule[ex.day]) state.schedule[ex.day] = {};
          if (!state.schedule[ex.day][ex.person]) state.schedule[ex.day][ex.person] = [];
          state.schedule[ex.day][ex.person].push(ex.placement);
        });
      }
      saveState();
      renderGrid();
      clearPlacedSelection();
    } else {
      // Single drag: existing behavior
      const placed = col ? dropOnColumn(e.clientY, col, card.id) : false;
      if (!placed) restorePlacement(savedPlacement, person, savedDay);
    }

    dragState = null;
  };

  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function startGridCardDragTouch(e, card, placement, person, srcEl) {
  e.preventDefault();

  const savedPlacement = { ...placement };
  const savedDay = srcEl.dataset.day;
  const daySchedule = state.schedule[savedDay];
  if (daySchedule && daySchedule[person]) {
    daySchedule[person] = daySchedule[person].filter(s => s.instanceId !== placement.instanceId);
  }

  const touch = e.touches[0];
  srcEl.classList.add('dragging');
  const rect = srcEl.getBoundingClientRect();
  const grabOffsetY = touch.clientY - rect.top;

  const ghost = buildGhost(card, rect.width);
  ghost.style.left = (touch.clientX - 16) + 'px';
  ghost.style.top  = (touch.clientY - grabOffsetY) + 'px';

  dragState = { cardId: card.id, ghost, grabOffsetY };

  const move = e => {
    const t = e.touches[0];
    ghost.style.left = (t.clientX - 16) + 'px';
    ghost.style.top  = (t.clientY - grabOffsetY) + 'px';
    highlightColumn(t.clientX, t.clientY);
    autoScrollUpdate(t.clientY);
  };

  const end = e => {
    document.removeEventListener('touchmove', move);
    document.removeEventListener('touchend', end);
    autoScrollStop();
    clearHighlight();
    ghost.remove();
    srcEl.classList.remove('dragging');

    const t      = e.changedTouches[0];
    const target = document.elementFromPoint(t.clientX, t.clientY);
    const col    = target?.closest('.grid-column');
    const placed = col ? dropOnColumn(t.clientY, col, card.id) : false;
    if (!placed) restorePlacement(savedPlacement, person, savedDay);

    dragState = null;
  };

  document.addEventListener('touchmove', move, { passive: false });
  document.addEventListener('touchend', end);
}

// ── Custom Drag & Drop (to schedule grid) ──────────────────
function buildGhost(card, w, extraCount = 0) {
  const ghost = document.createElement('div');
  ghost.className = 'drag-ghost';
  ghost.style.width = Math.min(220, w) + 'px';

  const n = ghost.appendChild(document.createElement('div'));
  n.className = 'card-name';
  n.textContent = card.name;

  const t = ghost.appendChild(document.createElement('div'));
  t.className = 'card-time-label';
  t.textContent = card.time + ' min';

  if (extraCount > 0) {
    const badge = ghost.appendChild(document.createElement('div'));
    badge.className = 'ghost-extra-badge';
    badge.textContent = '+' + extraCount;
  }

  document.body.appendChild(ghost);
  return ghost;
}

// ── Auto-scroll while dragging near edges of schedule-wrapper ──
let _autoScrollY  = null;
let _autoScrollRAF = null;

function autoScrollUpdate(clientY) {
  _autoScrollY = clientY;
  if (!_autoScrollRAF) _autoScrollRAF = requestAnimationFrame(_autoScrollTick);
}

function autoScrollStop() {
  _autoScrollY = null;
  if (_autoScrollRAF) { cancelAnimationFrame(_autoScrollRAF); _autoScrollRAF = null; }
}

function _autoScrollTick() {
  _autoScrollRAF = null;
  if (_autoScrollY === null) return;
  const wrapper = document.getElementById('schedule-wrapper');
  const r       = wrapper.getBoundingClientRect();
  const ZONE    = 80;
  const MAX     = 12;
  // Clamp zone edges to the visible viewport so the trigger area is never
  // outside what the user can see (handles partial off-screen wrappers).
  const top    = Math.max(r.top,    0);
  const bottom = Math.min(r.bottom, window.innerHeight);
  let speed = 0;
  if (_autoScrollY < top    + ZONE) speed = -MAX * (1 - Math.max(0, _autoScrollY - top)    / ZONE);
  if (_autoScrollY > bottom - ZONE) speed =  MAX * (1 - Math.max(0, bottom - _autoScrollY) / ZONE);
  if (speed !== 0) {
    const before = wrapper.scrollTop;
    wrapper.scrollTop += speed;
    // If the wrapper didn't move, try the content-area (mobile stacked layout),
    // then fall back to the page (non-fullscreen desktop).
    if (wrapper.scrollTop === before) {
      const ca = document.querySelector('.content-area');
      const caBefore = ca ? ca.scrollTop : null;
      if (ca) ca.scrollTop += speed;
      if (!ca || ca.scrollTop === caBefore) window.scrollBy(0, speed);
    }
  }
  _autoScrollRAF = requestAnimationFrame(_autoScrollTick);
}

function dropOnColumn(clientY, colEl, cardId) {
  const slotsEl = colEl.querySelector('.col-slots');
  if (!slotsEl) return false;
  const rect = slotsEl.getBoundingClientRect();
  const rawY  = clientY - rect.top;
  // Adjust for ghost offset so card top aligns correctly
  const offsetY = dragState ? (dragState.grabOffsetY || 0) : 0;
  const slotIndex = Math.floor((rawY - offsetY) / unitPx());
  return placeCard(cardId, colEl.dataset.person, slotIndex, colEl.dataset.day);
}

function highlightColumn(x, y) {
  document.querySelectorAll('.grid-column.drop-over').forEach(c => c.classList.remove('drop-over'));
  const el = document.elementFromPoint(x, y);
  el?.closest('.grid-column')?.classList.add('drop-over');
}

function clearHighlight() {
  document.querySelectorAll('.grid-column.drop-over').forEach(c => c.classList.remove('drop-over'));
}

function startDrag(e, cardId, srcEl) {
  e.preventDefault();
  const card = state.cards.find(c => c.id === cardId);
  if (!card) return;
  if (totalPlaced(cardId) >= card.count) {
    showToast(`All of the ${card.name} cards have been placed.`);
    return;
  }

  srcEl.classList.add('dragging');
  const rect       = srcEl.getBoundingClientRect();
  const grabOffsetY = e.clientY - rect.top;

  const ghost = buildGhost(card, rect.width);

  dragState = { cardId, ghost, grabOffsetY };

  const move = e => {
    ghost.style.left = (e.clientX - 16) + 'px';
    ghost.style.top  = (e.clientY - grabOffsetY) + 'px';
    highlightColumn(e.clientX, e.clientY);
    autoScrollUpdate(e.clientY);
  };

  const up = e => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    autoScrollStop();
    clearHighlight();
    ghost.remove();
    srcEl.classList.remove('dragging');

    const target = document.elementFromPoint(e.clientX, e.clientY);
    const col    = target?.closest('.grid-column');
    if (col) dropOnColumn(e.clientY, col, cardId);

    dragState = null;
  };

  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function startDragTouch(e, cardId, srcEl) {
  const card = state.cards.find(c => c.id === cardId);
  if (!card) return;
  if (totalPlaced(cardId) >= card.count) {
    showToast(`All of the ${card.name} cards have been placed.`);
    return;
  }

  // Don't prevent default at touchstart — allows native scroll until drag is committed
  const touch0      = e.touches[0];
  const startX      = touch0.clientX;
  const startY      = touch0.clientY;
  const rect        = srcEl.getBoundingClientRect();
  const grabOffsetY = touch0.clientY - rect.top;

  let ghost     = null;
  let committed = false;

  const move = e => {
    const t = e.touches[0];
    if (!committed) {
      const dx = t.clientX - startX;
      const dy = t.clientY - startY;
      if (Math.hypot(dx, dy) < 8) return; // wait for clear drag intent
      committed = true;
      srcEl.classList.add('dragging');
      ghost = buildGhost(card, rect.width);
      dragState = { cardId, ghost, grabOffsetY };
    }
    e.preventDefault(); // stop scroll once drag is committed
    ghost.style.left = (t.clientX - 16) + 'px';
    ghost.style.top  = (t.clientY - grabOffsetY) + 'px';
    highlightColumn(t.clientX, t.clientY);
    autoScrollUpdate(t.clientY);
  };

  const end = e => {
    document.removeEventListener('touchmove', move);
    document.removeEventListener('touchend', end);

    if (!committed) { autoScrollStop(); return; }

    autoScrollStop();
    clearHighlight();
    ghost.remove();
    srcEl.classList.remove('dragging');

    const t      = e.changedTouches[0];
    const target = document.elementFromPoint(t.clientX, t.clientY);
    const col    = target?.closest('.grid-column');
    if (col) dropOnColumn(t.clientY, col, cardId);

    dragState = null;
  };

  document.addEventListener('touchmove', move, { passive: false });
  document.addEventListener('touchend', end);
}

// ── UI: Day buttons ─────────────────────────────────────────
document.querySelectorAll('.day-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const day = btn.dataset.day;
    const idx = state.selectedDays.indexOf(day);
    if (idx === -1) {
      state.selectedDays.push(day);
    } else if (state.selectedDays.length > 1) {
      state.selectedDays.splice(idx, 1);
    }
    state.selectedDays.sort((a, b) => DAYS.indexOf(a) - DAYS.indexOf(b));
    saveState();
    renderDayButtons();
    updateLayoutClass();
    renderGrid();
  });
});

// ── UI: Grid sort ───────────────────────────────────────────
document.getElementById('grid-sort-select').addEventListener('change', function () {
  gridSort = this.value;
  renderGrid();
});

// ── UI: Add Person ──────────────────────────────────────────
const addPersonBtn = document.getElementById('add-person-btn');

addPersonBtn.addEventListener('click', startAddPerson);

function startAddPerson() {
  addPersonBtn.hidden = true;

  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = 'Name…';
  input.maxLength = 40;
  input.autocomplete = 'off';
  input.className = 'add-btn add-btn-input';
  input.setAttribute('aria-label', 'Enter person name, press Enter to add');

  addPersonBtn.after(input);
  input.focus();

  const commit = () => {
    const name = clean(input.value);
    input.remove();
    addPersonBtn.hidden = false;
    if (name && !state.people.includes(name)) {
      state.people.push(name);
      state.selectedPeople.push(name);
      state.personColors[name] = PERSON_COLORS[(state.people.length - 1) % PERSON_COLORS.length];
      saveState();
      renderNameButtons();
      renderGrid();
    }
  };

  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') input.blur();
    if (e.key === 'Escape') { input.value = ''; input.blur(); }
  });
}

// ── UI: Add Card ────────────────────────────────────────────
const addCardBtn    = document.getElementById('add-card-btn');
const newCardForm   = document.getElementById('new-card-form');
const newCardName   = document.getElementById('new-card-name');
const newCardDesc   = document.getElementById('new-card-desc');
const newCardTime   = document.getElementById('new-card-time');
const newCardCount    = document.getElementById('new-card-count');
const newCardCountSel = document.getElementById('new-card-count-sel');
newCardCount.addEventListener('input', () => {
  const v = parseInt(newCardCount.value);
  if (!isNaN(v)) newCardCount.value = Math.min(5, Math.max(1, v));
});
// Keep the hidden number input in sync so saveNewCard always reads from it
newCardCountSel.addEventListener('change', () => {
  newCardCount.value = newCardCountSel.value;
});
const saveCardBtn   = document.getElementById('save-card-btn');
const cancelCardBtn = document.getElementById('cancel-card-btn');
const newCardPerson  = document.getElementById('new-card-person');

addCardBtn.addEventListener('click', () => {
  const open = newCardForm.hidden;
  newCardForm.hidden = !open;
  addCardBtn.setAttribute('aria-expanded', String(open));
  if (open) {
    const mobile = window.innerWidth < 768;
    newCardCount.style.display    = mobile ? 'none'  : '';
    newCardCountSel.style.display = mobile ? 'block' : 'none';
    newCardName.value  = '';
    newCardDesc.value  = '';
    newCardTime.value  = '';
    newCardCount.value = '1';
    newCardCountSel.value = '1';
    newCardPerson.value = '';
    newCardName.focus();
  }
});

saveCardBtn.addEventListener('click', saveNewCard);
cancelCardBtn.addEventListener('click', () => {
  newCardForm.hidden = true;
  addCardBtn.setAttribute('aria-expanded', 'false');
});

newCardForm.addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.tagName !== 'SELECT' && e.target !== saveCardBtn && e.target !== cancelCardBtn) {
    e.preventDefault();
    saveNewCard();
  }
  if (e.key === 'Escape') {
    newCardForm.hidden = true;
    addCardBtn.setAttribute('aria-expanded', 'false');
  }
});

function saveNewCard() {
  const name  = clean(newCardName.value);
  const desc  = clean(newCardDesc.value);
  const time  = parseInt(newCardTime.value);
  const count = Math.min(5, Math.max(1, parseInt(newCardCount.value) || 1));

  if (!name)                    { newCardName.focus();  showToast('Please enter a name'); return; }
  if (isNaN(time) || time < 5)  { newCardTime.focus();  showToast('Enter duration ≥ 5 min'); return; }

  const assignedPerson = newCardPerson.value || '';
  const card = { id: uid(), name, description: desc, time, count: Math.max(1, count), assignedPerson };
  state.cards.push(card);
  saveState();
  renderCards();

  newCardForm.hidden = true;
  addCardBtn.setAttribute('aria-expanded', 'false');
  showToast(`Card added: ${name}`);
}

// ── UI: Zoom ────────────────────────────────────────────────
const zoomSlider  = document.getElementById('zoom-slider');
const zoomDisplay = document.getElementById('zoom-display');

zoomSlider.addEventListener('input', () => {
  zoomLevel = parseInt(zoomSlider.value);
  zoomDisplay.textContent = zoomLevel + '%';
  renderGrid();
  updateBadges();
});

const colZoomSlider  = document.getElementById('col-zoom-slider');
const colZoomDisplay = document.getElementById('col-zoom-display');

colZoomSlider.addEventListener('input', () => {
  colZoom = parseInt(colZoomSlider.value);
  colZoomDisplay.textContent = colZoom + '%';
  saveColZoom();
  renderGrid();
});

// ── UI: Fullscreen ──────────────────────────────────────────
function setFullscreen(on) {
  const main   = document.getElementById('main-section');
  const toggle = document.getElementById('fullscreen-toggle');
  main.classList.toggle('fullscreen', on);
  toggle.setAttribute('aria-pressed', String(on));
  toggle.setAttribute('aria-label', on ? 'Exit fullscreen' : 'Toggle fullscreen');
  document.getElementById('fullscreen-icon').textContent = on ? '✕' : '⛶';
}

document.getElementById('fullscreen-toggle').addEventListener('click', function () {
  const on = !document.getElementById('main-section').classList.contains('fullscreen');
  setFullscreen(on);
});

document.getElementById('exit-fullscreen-btn').addEventListener('click', function () {
  setFullscreen(false);
});

// ── UI: Export ──────────────────────────────────────────────
document.getElementById('export-btn').addEventListener('click', () => {
  const data = {
    version:        '1.1',
    exported:       new Date().toISOString(),
    people:         state.people.map(String),
    selectedDays:   state.selectedDays.slice(),
    selectedPeople: state.selectedPeople.map(String),
    personColors:   Object.fromEntries(
      Object.entries(state.personColors).map(([k, v]) => [String(k), String(v)])
    ),
    cards: state.cards.map(c => ({
      id:             String(c.id),
      name:           String(c.name),
      description:    String(c.description),
      time:           Number(c.time),
      count:          Number(c.count),
      assignedPerson: String(c.assignedPerson || ''),
    })),
    schedule: JSON.parse(JSON.stringify(state.schedule)),
  };

  const json = JSON.stringify(data, null, 2)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  const blob = new Blob([json], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `my_schedule_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  showToast('Schedule exported');
});

// ── UI: Import ──────────────────────────────────────────────
document.getElementById('import-btn').addEventListener('click', () => {
  document.getElementById('import-file').click();
});

document.getElementById('import-file').addEventListener('change', function () {
  const file = this.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    try {
      const parsed = JSON.parse(evt.target.result);
      // Re-use loadState's validation by writing to localStorage then reloading
      localStorage.setItem(STORAGE_KEY, JSON.stringify(parsed));
      loadState();
      renderAll();
      showToast('Schedule imported');
    } catch {
      showToast('Import failed — file may be invalid or corrupted');
      saveState(); // restore previous good state
    }
  };
  reader.readAsText(file);
  this.value = ''; // allow re-importing the same file
});

// ── UI: Clear Day ───────────────────────────────────────────
document.getElementById('clear-btn').addEventListener('click', () => {
  const days = state.selectedDays;
  const label = days.length === 1 ? days[0] : `${days.length} selected days`;
  showConfirm(`Clear all placements for ${label}?`, () => {
    days.forEach(d => { state.schedule[d] = {}; });
    saveState();
    renderGrid();
    showToast(`${label} cleared`);
  }, 'Clear');
});

// ── UI: Print ───────────────────────────────────────────────
document.getElementById('print-btn').addEventListener('click', printSchedule);
document.getElementById('undo-btn').addEventListener('click', undo);
document.getElementById('redo-btn').addEventListener('click', redo);

function printSchedule() {
  if (state.selectedPeople.length === 0) { showToast('Select at least one person to print'); return; }
  if (state.selectedDays.length   === 0) { showToast('Select at least one day to print');    return; }

  // Build column pairs in the same order as the grid
  const isMultiDay = state.selectedDays.length > 1;
  const colPairs = [];
  if (gridSort === 'person') {
    state.selectedPeople.forEach(person =>
      state.selectedDays.forEach(day => colPairs.push({ day, person }))
    );
  } else {
    state.selectedDays.forEach(day =>
      state.selectedPeople.forEach(person => colPairs.push({ day, person }))
    );
  }

  const PX_PER_MIN    = 50 / 60;                        // 50 px per hour
  const TOTAL_H       = (END_HOUR - START_HOUR) * 50;   // px height of slot area
  const landscape     = colPairs.length >= 4;
  const COLS_PER_PAGE = 5;                               // columns that fit per printed page
  const CH_HEIGHT     = 24;                              // fixed .ch header height (px)

  // Simple HTML-escape for print output
  const esc = s => String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // ── Time column (shared across all pages) ──
  // top values are offset by CH_HEIGHT so 6 AM aligns with the top of .cs (below .ch header)
  let timeHTML = '';
  for (let h = START_HOUR; h <= END_HOUR; h++) {
    const top   = CH_HEIGHT + (h - START_HOUR) * 50;
    const label = minutesToLabel((h - START_HOUR) * 60);
    timeHTML += `<div class="tl" style="top:${top}px">${esc(label)}</div>`;
  }

  // ── Build one grid column's HTML ──
  const buildColHTML = ({ day, person }) => {
    const headerLabel = isMultiDay ? `${day.slice(0, 3)} · ${esc(person)}` : esc(person);
    const placements  = (state.schedule[day] || {})[person] || [];

    let linesHTML = '';
    for (let h = START_HOUR; h < END_HOUR; h++) {
      linesHTML += `<div class="hl" style="top:${(h - START_HOUR) * 50}px"></div>`;
      linesHTML += `<div class="hl half" style="top:${(h - START_HOUR) * 50 + 25}px"></div>`;
    }

    let cardsHTML = '';
    placements.forEach(pl => {
      const card = state.cards.find(c => c.id === pl.cardId);
      if (!card) return;
      const top    = pl.slotIndex * SLOT_MIN * PX_PER_MIN;
      const height = Math.max(18, card.time * PX_PER_MIN);
      const color  = esc(state.personColors[card.assignedPerson] || '#e0e8e4');
      const time   = esc(minutesToLabel(pl.slotIndex * SLOT_MIN));
      cardsHTML += `<div class="pc" style="top:${top.toFixed(1)}px;height:${height.toFixed(1)}px;background:${color}">
        <div class="pn">${esc(card.name)}</div>
        <div class="pt">${time} · ${card.time} min</div>
      </div>`;
    });

    return `<div class="gc">
      <div class="ch">${headerLabel}</div>
      <div class="cs" style="height:${TOTAL_H}px">${linesHTML}${cardsHTML}</div>
    </div>`;
  };

  const dayLabel  = state.selectedDays.join(', ');
  const sortLabel = gridSort === 'person' ? 'Sorted by Person' : 'Sorted by Day';
  const dateStr   = new Date().toLocaleDateString(undefined, { year:'numeric', month:'long', day:'numeric' });

  const pageHeader = `<div class="ph">
  <h1>My Scheduler</h1>
  <p>${esc(dayLabel)} &nbsp;·&nbsp; ${sortLabel} &nbsp;·&nbsp; ${esc(dateStr)}</p>
</div>`;

  // ── Chunk columns into pages and build page sections ──
  let pagesHTML = '';
  for (let i = 0; i < colPairs.length; i += COLS_PER_PAGE) {
    const chunk   = colPairs.slice(i, i + COLS_PER_PAGE);
    const isLast  = i + COLS_PER_PAGE >= colPairs.length;
    const colsHTML = chunk.map(buildColHTML).join('');
    pagesHTML += `<div class="${isLast ? 'page' : 'page page-break'}">
${pageHeader}<div class="pg">
  <div class="tc">${timeHTML}</div>
  ${colsHTML}
</div>
</div>`;
  }

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>My Scheduler — ${esc(dayLabel)}</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Georgia, 'Times New Roman', serif; font-size: 11px; color: #234141; background: #fff; }
  .ph  { padding: 10px 14px 8px; border-bottom: 2px solid #234141; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: baseline; }
  .ph h1 { font-size: 16px; color: #234141; }
  .ph p  { font-size: 9px; color: #666A59; }
  .pg  { display: flex; align-items: stretch; padding: 0 14px; gap: 0; }
  .tc  { flex: 0 0 48px; position: relative; }
  .tl  { position: absolute; left: 0; right: 4px; font-size: 8px; color: #666A59; text-align: right; line-height: 1; transform: translateY(-50%); }
  .gc  { flex: 1; min-width: 80px; border-left: 1px solid #8C9289; border-bottom: 2px solid #98ACAB; display: flex; flex-direction: column; }
  .gc:last-child { border-right: 1px solid #8C9289; }
  .ch  { height: ${CH_HEIGHT}px; padding: 3px 5px; font-size: 10px; font-weight: 700; background: #234141; border-bottom: 2px solid #98ACAB; text-align: center; color: #FFFFFF; overflow: hidden; box-sizing: border-box; }
  .cs  { position: relative; }
  .hl  { position: absolute; left: 0; right: 0; border-top: 1px solid #D1C6B4; }
  .hl.half { border-top-style: dashed; border-top-color: #E7E0D0; }
  .pc  { position: absolute; left: 2px; right: 2px; border-radius: 3px; padding: 2px 4px; overflow: hidden; border: 1px solid rgba(0,0,0,0.12); }
  .pn  { font-weight: 700; font-size: 9px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .pt  { font-size: 8px; color: #666A59; margin-top: 1px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .page-break { break-after: page; page-break-after: always; }
  @media print {
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    @page { margin: 0.4in; ${landscape ? 'size: landscape;' : ''} }
  }
</style>
</head>
<body>
${pagesHTML}
</body>
</html>`;

  const win = window.open('', '_blank', 'width=900,height=700');
  if (!win) { showToast('Pop-up blocked — please allow pop-ups and try again'); return; }
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 400);
}

// ── Card rubber-band selection ───────────────────────────────
function startRubberBand(startX, startY, additive) {
  // Snapshot pre-drag selection so additive mode works correctly during move
  const preSelection = additive ? new Set(selectedCardIds) : new Set();

  const band = document.createElement('div');
  band.className = 'rubber-band-rect';
  band.style.display = 'none';
  document.body.appendChild(band);

  let started = false;

  const move = e => {
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    if (!started && Math.hypot(dx, dy) < 6) return;
    started = true;

    const left   = Math.min(startX, e.clientX);
    const top    = Math.min(startY, e.clientY);
    const right  = Math.max(startX, e.clientX);
    const bottom = Math.max(startY, e.clientY);

    band.style.display = 'block';
    band.style.left    = left   + 'px';
    band.style.top     = top    + 'px';
    band.style.width   = (right - left)  + 'px';
    band.style.height  = (bottom - top)  + 'px';

    // Rebuild selection each move: preSelection + any card intersecting the rect
    selectedCardIds.clear();
    preSelection.forEach(id => selectedCardIds.add(id));

    document.querySelectorAll('.activity-card').forEach(cardEl => {
      const cr = cardEl.getBoundingClientRect();
      if (cr.left < right && cr.right > left && cr.top < bottom && cr.bottom > top) {
        selectedCardIds.add(cardEl.dataset.cardId);
      }
    });

    document.querySelectorAll('.activity-card').forEach(cardEl => {
      cardEl.classList.toggle('card-selected', selectedCardIds.has(cardEl.dataset.cardId));
    });
  };

  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup',   up);
    band.remove();
    // Plain click on empty area (no drag) clears selection
    if (!started && !additive) clearCardSelection();
  };

  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup',   up);
}

// Rubber-band starts from empty space in the card list
document.getElementById('card-list').addEventListener('mousedown', e => {
  if (e.button !== 0) return;
  if (e.target !== document.getElementById('card-list')) return; // empty space only
  e.preventDefault();
  startRubberBand(e.clientX, e.clientY, e.ctrlKey || e.metaKey);
});

// ── Placed-card rubber-band selection ───────────────────────
function startPlacedRubberBand(startX, startY, additive) {
  const preSelection = additive ? new Set(selectedPlacedIds) : new Set();

  const band = document.createElement('div');
  band.className = 'rubber-band-rect';
  band.style.display = 'none';
  document.body.appendChild(band);

  let started = false;

  const move = e => {
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    if (!started && Math.hypot(dx, dy) < 6) return;
    started = true;

    const left   = Math.min(startX, e.clientX);
    const top    = Math.min(startY, e.clientY);
    const right  = Math.max(startX, e.clientX);
    const bottom = Math.max(startY, e.clientY);

    band.style.display = 'block';
    band.style.left    = left   + 'px';
    band.style.top     = top    + 'px';
    band.style.width   = (right - left)  + 'px';
    band.style.height  = (bottom - top)  + 'px';

    selectedPlacedIds.clear();
    preSelection.forEach(id => selectedPlacedIds.add(id));

    document.querySelectorAll('.placed-card').forEach(cardEl => {
      const cr = cardEl.getBoundingClientRect();
      if (cr.left < right && cr.right > left && cr.top < bottom && cr.bottom > top) {
        selectedPlacedIds.add(cardEl.dataset.instanceId);
      }
    });

    document.querySelectorAll('.placed-card').forEach(cardEl => {
      cardEl.classList.toggle('placed-selected',
        selectedPlacedIds.has(cardEl.dataset.instanceId));
    });
  };

  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup',   up);
    band.remove();
    if (!started && !additive) clearPlacedSelection();
  };

  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup',   up);
}

// Rubber-band on empty grid space (col-slots but not on a placed card or button)
document.getElementById('schedule-grid').addEventListener('mousedown', e => {
  if (e.button !== 0) return;
  if (e.target.closest('.placed-card') || e.target.closest('button')) return;
  if (!e.target.closest('.col-slots')) return;
  e.preventDefault();
  startPlacedRubberBand(e.clientX, e.clientY, e.ctrlKey || e.metaKey);
});

// Delete selected placed cards
function deleteSelectedPlaced() {
  DAYS.forEach(day => {
    if (!state.schedule[day]) return;
    Object.keys(state.schedule[day]).forEach(person => {
      state.schedule[day][person] = state.schedule[day][person].filter(
        s => !selectedPlacedIds.has(s.instanceId)
      );
    });
  });
  clearPlacedSelection();
  saveState();
  renderGrid();
}

// Escape clears both selections; Delete/Backspace removes selected placed cards
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    clearCardSelection();
    clearPlacedSelection();
  }
  if ((e.key === 'Delete' || e.key === 'Backspace') && selectedPlacedIds.size > 0) {
    // Only delete placed cards when not focused inside an input/select
    const tag = document.activeElement?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
    e.preventDefault();
    deleteSelectedPlaced();
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
    const tag = document.activeElement?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
    e.preventDefault();
    undo();
  }
  if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
    const tag = document.activeElement?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
    e.preventDefault();
    redo();
  }
});

// ── Render All ──────────────────────────────────────────────
function renderAll() {
  renderDayButtons();
  renderNameButtons();
  renderCards();
  renderGrid();
}

// ── Panel Resizer ────────────────────────────────────────────
(function initPanelResizer() {
  const COMMIT_PX   = 300;  // snap threshold evaluated on mouse-up only
  const STORAGE_KEY = 'cm_panel_width';

  const resizer     = document.getElementById('panel-resizer');
  const leftPanel   = document.querySelector('.left-panel');
  const contentArea = document.querySelector('.content-area');

  let panelCollapsed = false;
  _collapsePanel = () => { if (!panelCollapsed) collapse(); };

  function collapse() {
    leftPanel.style.width    = '0';
    leftPanel.style.minWidth = '0';
    leftPanel.style.padding  = '0';
    leftPanel.style.border   = 'none';
    panelCollapsed = true;
    resizer.classList.add('panel-collapsed');
  }

  function open(px) {
    leftPanel.style.width    = Math.max(COMMIT_PX, px) + 'px';
    leftPanel.style.minWidth = '';
    leftPanel.style.padding  = '';
    leftPanel.style.border   = '';
    panelCollapsed = false;
    resizer.classList.remove('panel-collapsed');
  }

  function savePanelWidth() {
    localStorage.setItem(STORAGE_KEY, panelCollapsed ? 0 : parseInt(leftPanel.style.width) || 0);
  }

  // Restore saved width on load (desktop only)
  if (window.innerWidth >= 768) {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved !== null) {
      const w = parseFloat(saved);
      if (w < COMMIT_PX) collapse(); else open(w);
    }
  }

  function onDragStart(getClientX) {
    resizer.classList.add('resizing');
    const wasCollapsed = panelCollapsed;

    // Strip padding/border for the whole drag so the panel tracks the cursor cleanly
    leftPanel.style.padding = '0';
    leftPanel.style.border  = 'none';

    const move = e => {
      const maxPx = contentArea.offsetWidth * 0.65;
      const px = Math.max(0, Math.min(getClientX(e) - contentArea.getBoundingClientRect().left, maxPx));
      leftPanel.style.width    = px + 'px';
      leftPanel.style.minWidth = px > 0 ? '' : '0';
      resizer.classList.toggle('panel-collapsed', px < 1);
    };

    const end = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup',   end);
      document.removeEventListener('touchmove', move);
      document.removeEventListener('touchend',  end);
      resizer.classList.remove('resizing');

      const w = parseInt(leftPanel.style.width) || 0;
      if (w < COMMIT_PX) {
        // Below threshold on release:
        //   closing drag → snap shut; opening drag → snap to minimum open
        if (wasCollapsed) open(COMMIT_PX); else collapse();
      } else {
        // Above threshold → leave where it is, restore padding/border
        open(w);
      }

      savePanelWidth();
    };

    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup',   end);
    document.addEventListener('touchmove', move, { passive: false });
    document.addEventListener('touchend',  end);
  }

  resizer.addEventListener('mousedown', e => {
    e.preventDefault();
    onDragStart(e => e.clientX);
  });

  resizer.addEventListener('touchstart', e => {
    e.preventDefault();
    onDragStart(e => e.touches[0].clientX);
  }, { passive: false });
})();

// ── Nav tooltips ────────────────────────────────────────────
(function () {
  const tip = document.createElement('div');
  tip.className = 'nav-tip';
  tip.setAttribute('role', 'tooltip');
  tip.style.display = 'none';
  document.body.appendChild(tip);

  let timer;

  document.querySelectorAll('.nav-btn[data-tip-title]').forEach(btn => {
    btn.addEventListener('mouseenter', () => {
      timer = setTimeout(() => {
        const key = btn.dataset.tipKey
          ? `<span class="nav-tip-key">${btn.dataset.tipKey}</span>` : '';
        tip.innerHTML = `<strong>${btn.dataset.tipTitle}</strong>${btn.dataset.tipBody || ''}${key}`;
        tip.style.display = '';

        const r = btn.getBoundingClientRect();
        const tw = tip.offsetWidth;
        const left = Math.max(8, Math.min(r.left + r.width / 2 - tw / 2, window.innerWidth - tw - 8));
        tip.style.left = left + 'px';
        tip.style.top  = (r.bottom + 10) + 'px';

        // Position arrow relative to button centre
        const arrowLeft = (r.left + r.width / 2) - left;
        tip.style.setProperty('--arrow-left', arrowLeft + 'px');
      }, 280);
    });

    btn.addEventListener('mouseleave', () => {
      clearTimeout(timer);
      tip.style.display = 'none';
    });
  });
})();

// ── Init ────────────────────────────────────────────────────
loadState();
// Sync col-zoom slider to persisted value
colZoomSlider.value = colZoom;
colZoomDisplay.textContent = colZoom + '%';
renderAll();
