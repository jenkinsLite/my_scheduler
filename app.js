'use strict';

/* ============================================================
   My Scheduler — app.js
   All state lives in localStorage. No external dependencies.
   ============================================================ */

// ── Constants ──────────────────────────────────────────────
const DAYS         = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
const START_HOUR   = 6;    // 6:00 AM
const END_HOUR     = 22;   // 10:00 PM
const SLOT_MIN     = 5;    // minutes per grid row
const TOTAL_SLOTS  = ((END_HOUR - START_HOUR) * 60) / SLOT_MIN;  // 192
const STORAGE_KEY  = 'cm_scheduler_v1';
const BASE_UNIT_PX = 12;   // px per slot at 100 % zoom
let   _idCounter   = Date.now();

// ── State ──────────────────────────────────────────────────
let state = {
  people:        [],
  selectedDay:   'Monday',
  selectedPeople:[],
  cards:         [],
  // schedule[day][personName] = [ { instanceId, cardId, slotIndex } ]
  schedule:      {}
};

let zoomLevel  = 100;  // percent
let dragState  = null; // active drag info

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
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
  catch(e) { console.error('Save failed', e); }
}

function parseCard(c) {
  return {
    id:          String(c.id   || uid()),
    name:        clean(c.name  || ''),
    description: clean(c.description || ''),
    time:        Math.max(5, Math.min(480, parseInt(c.time) || 30)),
    count:       Math.max(1, Math.min(99,  parseInt(c.count) || 1)),
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

    state.selectedDay = DAYS.includes(p.selectedDay) ? p.selectedDay : 'Monday';

    state.selectedPeople = Array.isArray(p.selectedPeople)
      ? p.selectedPeople.map(x => clean(String(x))).filter(x => state.people.includes(x))
      : [];

    state.cards = Array.isArray(p.cards) ? p.cards.map(parseCard) : [];

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
    const active = btn.dataset.day === state.selectedDay;
    btn.classList.toggle('active', active);
    btn.setAttribute('aria-pressed', String(active));
  });
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
    btn.setAttribute('aria-label', `${person} — click to toggle, double-click to rename`);

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

    list.appendChild(btn);
  });
}

function inlineEditPerson(btn, oldName) {
  // Wrapper holds the input + delete button side by side
  const wrapper = document.createElement('div');
  wrapper.className = 'name-btn-edit-wrapper';

  const input = document.createElement('input');
  input.type = 'text';
  input.value = oldName;
  input.className = 'name-btn-edit-input';
  input.maxLength = 40;
  input.setAttribute('aria-label', 'Rename person — press Enter to save, Escape to cancel');

  const delBtn = document.createElement('button');
  delBtn.className = 'name-btn-delete-btn';
  delBtn.textContent = '×';
  delBtn.setAttribute('aria-label', `Delete ${oldName}`);

  wrapper.appendChild(input);
  wrapper.appendChild(delBtn);
  btn.replaceWith(wrapper);
  input.focus();
  input.select();

  const commit = () => {
    const newName = clean(input.value);
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
      saveState();
    }
    renderNameButtons();
    renderGrid();
  };

  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter')  input.blur();
    if (e.key === 'Escape') renderNameButtons();
  });

  // Prevent input blur when clicking delete so the click still fires
  delBtn.addEventListener('mousedown', e => e.preventDefault());
  delBtn.addEventListener('click', () => {
    input.removeEventListener('blur', commit); // prevent commit from running on removal
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
  saveState();
  renderNameButtons();
  renderGrid();
}

// ── Render: Activity cards (left panel) ────────────────────
function renderCards() {
  const list = document.getElementById('card-list');
  list.innerHTML = '';
  state.cards.forEach(card => list.appendChild(buildCardEl(card)));
}

function buildCardEl(card) {
  const placed     = totalPlaced(card.id);
  const fullyPlaced = placed >= card.count;
  const heightPx   = Math.max(72, slotToPx(card.time / SLOT_MIN));

  const el = document.createElement('div');
  el.className = 'activity-card' + (fullyPlaced ? ' fully-placed' : '');
  el.setAttribute('role', 'listitem');
  el.dataset.cardId = card.id;
  el.style.minHeight = heightPx + 'px';
  el.setAttribute('aria-label',
    `${card.name}: ${card.description}. ${card.time} min. ${placed}/${card.count} placed. Drag to schedule.`);

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
  badge.textContent = `${placed}/${card.count}`;
  badge.setAttribute('aria-label', `${placed} of ${card.count} placed`);

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
    if (e.target.closest('button, input, .card-drag-handle')) return;
    startDrag(e, card.id, el);
  });
  el.addEventListener('touchstart', e => {
    if (e.target.closest('button, input, .card-drag-handle')) return;
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
    if (field === 'time') renderGrid();
  };

  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter')  input.blur();
    if (e.key === 'Escape') renderCards();
  });
}

function inlineEditCount(badge, card) {
  const input = document.createElement('input');
  input.type = 'number';
  input.value = card.count;
  input.min = '1'; input.max = '99';
  input.className = 'inline-edit-input';
  input.style.width = '50px';
  input.setAttribute('aria-label', 'Edit count');
  badge.replaceWith(input);
  input.focus(); input.select();

  const commit = () => {
    const v = parseInt(input.value);
    if (!isNaN(v) && v >= 1) card.count = v;
    saveState();
    renderCards();
    updateBadges();
  };
  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter')  input.blur();
    if (e.key === 'Escape') renderCards();
  });
}

function deleteCard(cardId) {
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
    const placed = totalPlaced(card.id);
    const badge  = el.querySelector('.card-count-badge');
    if (badge) {
      badge.textContent = `${placed}/${card.count}`;
      badge.setAttribute('aria-label', `${placed} of ${card.count} placed`);
    }
    el.classList.toggle('fully-placed', placed >= card.count);
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
  timeCol.innerHTML = '';
  gridEl.innerHTML  = '';

  // ── Time labels ──
  for (let i = 0; i <= TOTAL_SLOTS; i++) {
    const label   = document.createElement('div');
    label.className = 'time-label';
    label.style.height = unitPx() + 'px';
    const mins    = i * SLOT_MIN;
    const isHour  = mins % 60 === 0;
    const isHalf  = mins % 30 === 0;
    if (isHour) {
      label.classList.add('hour');
      label.textContent = minutesToLabel(mins);
    } else if (isHalf) {
      label.classList.add('show');
      label.textContent = minutesToLabel(mins);
    }
    timeCol.appendChild(label);
  }

  // ── Grid columns ──
  state.selectedPeople.forEach(person => {
    const col = document.createElement('div');
    col.className = 'grid-column';
    col.dataset.person = person;
    col.setAttribute('role', 'gridcell');
    col.setAttribute('aria-label', `${person}'s schedule`);

    // Column header
    const colHeader = col.appendChild(document.createElement('div'));
    colHeader.className = 'col-header';
    colHeader.textContent = person;
    colHeader.setAttribute('aria-hidden', 'true');

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
    const daySchedule = (state.schedule[state.selectedDay] || {})[person] || [];
    daySchedule.forEach(placement => {
      const card = state.cards.find(c => c.id === placement.cardId);
      if (!card) return;
      placedCardEl(card, placement, person, slots);
    });
  });

  updateBadges();
}

function placedCardEl(card, placement, person, slotsContainer) {
  const heightPx = Math.max(unitPx(), slotToPx(card.time / SLOT_MIN));
  const topPx    = slotToPx(placement.slotIndex);

  const el = document.createElement('div');
  el.className = 'placed-card';
  el.style.top    = topPx    + 'px';
  el.style.height = heightPx + 'px';
  el.dataset.instanceId = placement.instanceId;
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
    removePlaced(person, placement.instanceId);
  });

  // Drag to reposition
  el.addEventListener('mousedown', e => {
    if (e.target.closest('button')) return;
    startGridCardDrag(e, card, placement, person, el);
  });
  el.addEventListener('touchstart', e => {
    if (e.target.closest('button')) return;
    startGridCardDragTouch(e, card, placement, person, el);
  }, { passive: false });

  slotsContainer.appendChild(el);
}

function removePlaced(person, instanceId) {
  const day = state.schedule[state.selectedDay];
  if (!day || !day[person]) return;
  day[person] = day[person].filter(s => s.instanceId !== instanceId);
  saveState();
  renderGrid();
}

// ── Place card into schedule ─────────────────────────────────
function placeCard(cardId, person, slotIndex) {
  const card = state.cards.find(c => c.id === cardId);
  if (!card) return;

  slotIndex = Math.max(0, Math.min(TOTAL_SLOTS - 1, Math.floor(slotIndex)));
  const endSlot = slotIndex + card.time / SLOT_MIN;

  if (!state.schedule[state.selectedDay]) state.schedule[state.selectedDay] = {};
  if (!state.schedule[state.selectedDay][person]) state.schedule[state.selectedDay][person] = [];
  const existing = state.schedule[state.selectedDay][person];

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
function restorePlacement(savedPlacement, savedPerson) {
  if (!state.schedule[state.selectedDay]) state.schedule[state.selectedDay] = {};
  if (!state.schedule[state.selectedDay][savedPerson]) state.schedule[state.selectedDay][savedPerson] = [];
  state.schedule[state.selectedDay][savedPerson].push(savedPlacement);
  saveState();
  renderGrid();
}

function startGridCardDrag(e, card, placement, person, srcEl) {
  e.preventDefault();

  // Temporarily remove from state so overlap check ignores it
  const savedPlacement = { ...placement };
  const day = state.schedule[state.selectedDay];
  if (day && day[person]) {
    day[person] = day[person].filter(s => s.instanceId !== placement.instanceId);
  }

  srcEl.classList.add('dragging');
  const rect = srcEl.getBoundingClientRect();
  const grabOffsetY = e.clientY - rect.top;

  const ghost = buildGhost(card, rect.width);
  ghost.style.left = (e.clientX - 16) + 'px';
  ghost.style.top  = (e.clientY - grabOffsetY) + 'px';

  dragState = { cardId: card.id, ghost, grabOffsetY };

  const move = e => {
    ghost.style.left = (e.clientX - 16) + 'px';
    ghost.style.top  = (e.clientY - grabOffsetY) + 'px';
    highlightColumn(e.clientX, e.clientY);
  };

  const up = e => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    clearHighlight();
    ghost.remove();
    srcEl.classList.remove('dragging');

    const target = document.elementFromPoint(e.clientX, e.clientY);
    const col    = target?.closest('.grid-column');
    const placed = col ? dropOnColumn(e.clientY, col, card.id) : false;
    if (!placed) restorePlacement(savedPlacement, person);

    dragState = null;
  };

  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function startGridCardDragTouch(e, card, placement, person, srcEl) {
  e.preventDefault();

  const savedPlacement = { ...placement };
  const day = state.schedule[state.selectedDay];
  if (day && day[person]) {
    day[person] = day[person].filter(s => s.instanceId !== placement.instanceId);
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
  };

  const end = e => {
    document.removeEventListener('touchmove', move);
    document.removeEventListener('touchend', end);
    clearHighlight();
    ghost.remove();
    srcEl.classList.remove('dragging');

    const t      = e.changedTouches[0];
    const target = document.elementFromPoint(t.clientX, t.clientY);
    const col    = target?.closest('.grid-column');
    const placed = col ? dropOnColumn(t.clientY, col, card.id) : false;
    if (!placed) restorePlacement(savedPlacement, person);

    dragState = null;
  };

  document.addEventListener('touchmove', move, { passive: false });
  document.addEventListener('touchend', end);
}

// ── Custom Drag & Drop (to schedule grid) ──────────────────
function buildGhost(card, w) {
  const ghost = document.createElement('div');
  ghost.className = 'drag-ghost';
  ghost.style.width = Math.min(220, w) + 'px';

  const n = ghost.appendChild(document.createElement('div'));
  n.className = 'card-name';
  n.textContent = card.name;

  const t = ghost.appendChild(document.createElement('div'));
  t.className = 'card-time-label';
  t.textContent = card.time + ' min';

  document.body.appendChild(ghost);
  return ghost;
}

function dropOnColumn(clientY, colEl, cardId) {
  const slotsEl = colEl.querySelector('.col-slots');
  if (!slotsEl) return false;
  const rect = slotsEl.getBoundingClientRect();
  const rawY  = clientY - rect.top;
  // Adjust for ghost offset so card top aligns correctly
  const offsetY = dragState ? (dragState.grabOffsetY || 0) : 0;
  const slotIndex = Math.floor((rawY - offsetY) / unitPx());
  return placeCard(cardId, colEl.dataset.person, slotIndex);
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

  srcEl.classList.add('dragging');
  const rect       = srcEl.getBoundingClientRect();
  const grabOffsetY = e.clientY - rect.top;

  const ghost = buildGhost(card, rect.width);

  dragState = { cardId, ghost, grabOffsetY };

  const move = e => {
    ghost.style.left = (e.clientX - 16) + 'px';
    ghost.style.top  = (e.clientY - grabOffsetY) + 'px';
    highlightColumn(e.clientX, e.clientY);
  };

  const up = e => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
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
  };

  const end = e => {
    document.removeEventListener('touchmove', move);
    document.removeEventListener('touchend', end);

    if (!committed) return; // was a tap or scroll, not a drag

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
    state.selectedDay = btn.dataset.day;
    saveState();
    renderDayButtons();
    renderGrid();
  });
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
const newCardCount  = document.getElementById('new-card-count');
const saveCardBtn   = document.getElementById('save-card-btn');
const cancelCardBtn = document.getElementById('cancel-card-btn');

addCardBtn.addEventListener('click', () => {
  const open = newCardForm.hidden;
  newCardForm.hidden = !open;
  addCardBtn.setAttribute('aria-expanded', String(open));
  if (open) {
    newCardName.value  = '';
    newCardDesc.value  = '';
    newCardTime.value  = '';
    newCardCount.value = '1';
    newCardName.focus();
  }
});

saveCardBtn.addEventListener('click', saveNewCard);
cancelCardBtn.addEventListener('click', () => {
  newCardForm.hidden = true;
  addCardBtn.setAttribute('aria-expanded', 'false');
});

newCardForm.addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target !== saveCardBtn && e.target !== cancelCardBtn) {
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
  const count = parseInt(newCardCount.value) || 1;

  if (!name)                    { newCardName.focus();  showToast('Please enter a name'); return; }
  if (isNaN(time) || time < 5)  { newCardTime.focus();  showToast('Enter duration ≥ 5 min'); return; }

  const card = { id: uid(), name, description: desc, time, count: Math.max(1, count) };
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

// ── UI: Clear Day ───────────────────────────────────────────
document.getElementById('clear-btn').addEventListener('click', () => {
  if (!confirm(`Clear all placements for ${state.selectedDay}?`)) return;
  state.schedule[state.selectedDay] = {};
  saveState();
  renderGrid();
  showToast(`${state.selectedDay} cleared`);
});

// ── Render All ──────────────────────────────────────────────
function renderAll() {
  renderDayButtons();
  renderNameButtons();
  renderCards();
  renderGrid();
}

// ── Init ────────────────────────────────────────────────────
loadState();
renderAll();
