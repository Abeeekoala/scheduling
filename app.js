const state = {
  staff: [],
  closures: [],
  holidays: [],
  results: null,
  settings: {
    month: null,
    newJoinerLimit: 1,
  },
};

const STORAGE_KEY = 'schedulerStateV1';

const WEEKDAY_LABELS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const CALENDAR_WEEKDAY_LABELS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

const SHIFT_BLUEPRINTS = {
  L: {
    default: [
      { key: 'morning', label: 'Morning (4h)', hours: 4, required: 4 },
      { key: 'afternoon', label: 'Afternoon (3h)', hours: 3, required: 3 },
      { key: 'night', label: 'Night (3h)', hours: 3, required: 2 },
    ],
  },
  C: {
    weekday: [
      { key: 'morning', label: 'Morning (4h)', hours: 4, required: 2 },
      { key: 'afternoon', label: 'Afternoon (3h)', hours: 3, required: 2 },
      { key: 'night', label: 'Night (3h)', hours: 3, required: 1 },
    ],
    saturday: [
      { key: 'morning', label: 'Sat Morning (3.5h)', hours: 3.5, required: 2 },
      { key: 'afternoon', label: 'Sat Afternoon (3h)', hours: 3, required: 2 },
    ],
  },
};

const SHIFT_ORDER = ['morning', 'afternoon', 'night'];
const SHIFT_LABELS = {
  morning: '早',
  afternoon: '午',
  night: '晚',
};

function closeAllMultiDropdowns() {
  document.querySelectorAll('.multi-dropdown.open').forEach(dropdown => dropdown.classList.remove('open'));
}

function initializeMultiSelects(root = document) {
  if (!root) return;
  const selects = root.querySelectorAll('select[data-multi-select]:not([data-multi-init])');
  selects.forEach(select => {
    select.dataset.multiInit = 'true';
    select.classList.add('multi-select-hidden');
    const placeholder = select.dataset.placeholder || 'Select';
    const wrapper = document.createElement('div');
    wrapper.className = 'multi-dropdown';
    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'multi-dropdown-toggle';
    const labelSpan = document.createElement('span');
    toggle.appendChild(labelSpan);
    const menu = document.createElement('div');
    menu.className = 'multi-dropdown-menu';

    const checkboxes = [];
    Array.from(select.options).forEach(option => {
      const optionLabel = document.createElement('label');
      optionLabel.className = 'multi-dropdown-option';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = option.value;
      checkbox.checked = option.selected;
      const text = document.createTextNode(option.textContent.trim());
      checkbox.addEventListener('change', () => {
        option.selected = checkbox.checked;
        updateSummary();
        select.dispatchEvent(new Event('change', { bubbles: true }));
      });
      optionLabel.appendChild(checkbox);
      optionLabel.appendChild(text);
      menu.appendChild(optionLabel);
      checkboxes.push({ checkbox, option });
    });

    toggle.addEventListener('click', event => {
      event.stopPropagation();
      const isOpen = wrapper.classList.contains('open');
      closeAllMultiDropdowns();
      if (!isOpen) {
        wrapper.classList.add('open');
      }
    });

    menu.addEventListener('click', event => event.stopPropagation());

    const updateSummary = () => {
      const selectedTexts = Array.from(select.selectedOptions).map(opt => opt.text.trim());
      labelSpan.textContent = selectedTexts.length ? selectedTexts.join(', ') : placeholder;
      checkboxes.forEach(({ checkbox, option }) => {
        if (checkbox.checked !== option.selected) {
          checkbox.checked = option.selected;
        }
      });
    };

    select.addEventListener('change', () => updateSummary());

    const parent = select.parentNode;
    parent.insertBefore(wrapper, select.nextSibling);
    wrapper.appendChild(toggle);
    wrapper.appendChild(menu);
    updateSummary();
  });
}

document.addEventListener('click', event => {
  if (!event.target.closest('.multi-dropdown')) {
    closeAllMultiDropdowns();
  }
});

const uid = () =>
  (typeof crypto !== 'undefined' && crypto.randomUUID)
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;

const HIGH_SKILL_THRESHOLD = 4;

function loadPersistedState() {
  if (typeof window === 'undefined' || !window.localStorage) return;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    state.staff = (parsed.staff || []).map(staff => {
      const daysOff = Array.isArray(staff.daysOff)
        ? staff.daysOff
            .map(Number)
            .filter(day => Number.isInteger(day) && day >= 1 && day <= 31)
        : [];
      const preferences = (staff.preferences || []).map(pref => ({
        ...pref,
        location: pref.location || 'any',
        shifts: Array.isArray(pref.shifts) && pref.shifts.length ? pref.shifts : ['all'],
      }));
      return {
        ...staff,
        preferences,
        collapsed: Boolean(staff.collapsed),
        preferredLocation: staff.preferredLocation || 'any',
        locations: Array.isArray(staff.locations) && staff.locations.length ? staff.locations : ['L', 'C'],
        daysOff,
        avoidWeekdayShifts: staff.avoidWeekdayShifts || [],
        avoidWeekendShifts: staff.avoidWeekendShifts || [],
        skillScore: typeof staff.skillScore === 'number' ? staff.skillScore : 3,
      };
    });
    state.closures = parsed.closures || [];
    state.holidays = parsed.holidays || [];
    state.settings = {
      month: parsed.settings?.month || null,
      newJoinerLimit:
        typeof parsed.settings?.newJoinerLimit === 'number'
          ? parsed.settings.newJoinerLimit
          : 1,
    };
  } catch (error) {
    console.warn('Unable to load saved schedule data.', error);
  }
}

function persistState() {
  if (typeof window === 'undefined' || !window.localStorage) return;
  try {
    const payload = {
      staff: state.staff,
      closures: state.closures,
      holidays: state.holidays,
      settings: state.settings,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch (error) {
    console.warn('Unable to persist schedule data.', error);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  initializeMultiSelects(document);
  const monthInput = document.getElementById('month-input');
  const newJoinerInput = document.getElementById('new-joiner-limit');
  const today = new Date();
  const defaultMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;

  loadPersistedState();

  if (!state.settings.month) {
    state.settings.month = defaultMonth;
    persistState();
  }

  monthInput.value = state.settings.month;
  newJoinerInput.value = state.settings.newJoinerLimit ?? 1;

  document.getElementById('staff-form').addEventListener('submit', handleStaffSubmit);
  const staffListElement = document.getElementById('staff-list');
  staffListElement.addEventListener('click', handleStaffListClick);
  staffListElement.addEventListener('submit', handleStaffListSubmit);
  staffListElement.addEventListener('change', handleStaffListChange);

  document.getElementById('closure-form').addEventListener('submit', handleClosureSubmit);
  document.getElementById('closure-list').addEventListener('click', handleClosureListClick);

  document.getElementById('holiday-form').addEventListener('submit', handleHolidaySubmit);
  document.getElementById('holiday-list').addEventListener('click', handleHolidayListClick);

  document.getElementById('generate-btn').addEventListener('click', handleGenerate);
  document.getElementById('clear-results').addEventListener('click', () => {
    state.results = null;
    renderResults();
    setStatus('Results cleared.');
  });

  monthInput.addEventListener('change', event => {
    state.settings.month = event.target.value || state.settings.month;
    persistState();
  });

  newJoinerInput.addEventListener('change', event => {
    const numeric = Math.max(0, Number(event.target.value) || 0);
    state.settings.newJoinerLimit = numeric;
    event.target.value = numeric;
    persistState();
  });

  renderStaffList();
  renderClosures();
  renderHolidays();
  renderResults();
});

function handleStaffSubmit(event) {
  event.preventDefault();
  const name = document.getElementById('staff-name').value.trim();
  const weeklyCap = Number(document.getElementById('staff-weekly').value) || 40;
  const monthlyCapInput = document.getElementById('staff-monthly').value;
  const monthlyCap = monthlyCapInput ? Number(monthlyCapInput) : null;
  const isNewJoiner = document.getElementById('staff-new-joiner').checked;
  const locations = Array.from(document.getElementById('staff-locations').selectedOptions).map(opt => opt.value);
  const preferredLocation = document.getElementById('staff-preferred-location').value || 'any';
  const avoidWeekdayShifts = Array.from(document.getElementById('staff-weekday-preferences').selectedOptions).map(opt => opt.value);
  const avoidWeekendShifts = Array.from(document.getElementById('staff-weekend-preferences').selectedOptions).map(opt => opt.value);
  const skillScore = Number(document.getElementById('staff-skill').value) || 3;

  if (!name) {
    setStatus('Please provide a staff name.', true);
    return;
  }
  if (!locations.length) {
    setStatus('Select at least one allowed location.', true);
    return;
  }

  state.staff.push({
    id: uid(),
    name,
    weeklyCap,
    monthlyCap,
    isNewJoiner,
    locations,
    preferences: [],
    collapsed: false,
    preferredLocation,
    daysOff: [],
    avoidWeekdayShifts,
    avoidWeekendShifts,
    skillScore,
  });
  event.target.reset();
  document.getElementById('staff-weekly').value = 40;
  document.getElementById('staff-preferred-location').value = 'any';
  const avoidWeekdaySelect = document.getElementById('staff-weekday-preferences');
  const avoidWeekendSelect = document.getElementById('staff-weekend-preferences');
  [avoidWeekdaySelect, avoidWeekendSelect].forEach(select => {
    Array.from(select.options).forEach(opt => (opt.selected = false));
    select.dispatchEvent(new Event('change'));
  });
  const locationsSelect = document.getElementById('staff-locations');
  Array.from(locationsSelect.options).forEach(opt => (opt.selected = true));
  locationsSelect.dispatchEvent(new Event('change'));
  persistState();
  renderStaffList();
  setStatus(`Added ${name} to the pool.`);
}

function handleStaffListClick(event) {
  const card = event.target.closest('.staff-card');
  if (!card) return;
  const staffId = card.dataset.staffId;

  if (event.target.matches('[data-action="toggle-card"]') || event.target.closest('[data-action="toggle-card"]')) {
    const staff = state.staff.find(s => s.id === staffId);
    if (!staff) return;
    staff.collapsed = !staff.collapsed;
    persistState();
    renderStaffList();
    return;
  }

  if (event.target.matches('[data-action="remove-staff"]')) {
    state.staff = state.staff.filter(staff => staff.id !== staffId);
    persistState();
    renderStaffList();
    setStatus('Removed staff member.');
    return;
  }

  if (event.target.matches('[data-action="remove-pref"]')) {
    const prefId = event.target.dataset.prefId;
    const staff = state.staff.find(s => s.id === staffId);
    if (!staff) return;
    staff.preferences = (staff.preferences || []).filter(pref => pref.id !== prefId);
    persistState();
    renderStaffList();
    setStatus('Removed shift-specific day off.');
  }
}

function handleStaffListSubmit(event) {
  if (event.target.matches('.days-off-form')) {
    event.preventDefault();
    const staffId = event.target.dataset.staffId;
    const staff = state.staff.find(s => s.id === staffId);
    if (!staff) return;
    const rawInput = event.target.daysOff.value.trim();
    const parsed = parseDaysOffInput(rawInput);
    if (!parsed.success) {
      setStatus(parsed.error, true);
      return;
    }
    staff.daysOff = parsed.days;
    persistState();
    renderStaffList();
    setStatus(`Updated days off for ${staff.name}.`);
    return;
  }

  if (event.target.matches('.shift-off-form')) {
    event.preventDefault();
    const staffId = event.target.dataset.staffId;
    const staff = state.staff.find(s => s.id === staffId);
    if (!staff) return;
    const type = event.target.ruleType.value;
    const shiftValue = event.target.shiftKey.value;
    const location = event.target.shiftLocation.value;
    let weekday = null;
    let date = null;
    if (type === 'weekday') {
      weekday = Number(event.target.weekday.value);
      if (Number.isNaN(weekday)) {
        setStatus('Select a weekday.', true);
        return;
      }
    } else {
      date = event.target.prefDate.value;
      if (!date) {
        setStatus('Select a date.', true);
        return;
      }
    }

    const preference = {
      id: uid(),
      type,
      weekday,
      date,
      location,
      shifts: shiftValue === 'all' ? ['all'] : [shiftValue],
    };
    staff.preferences = staff.preferences || [];
    staff.preferences.push(preference);
    persistState();
    renderStaffList();
    setStatus('Shift-specific day off added.');
    event.target.reset();
    syncShiftOffFields(event.target, 'weekday');
  }
}

function handleStaffListChange(event) {
  const card = event.target.closest('.staff-card');
  if (!card) return;
  const staffId = card.dataset.staffId;
  const staff = state.staff.find(s => s.id === staffId);
  if (!staff) return;

  if (event.target.matches('[data-action="new-joiner"]')) {
    staff.isNewJoiner = event.target.checked;
    persistState();
    setStatus(`New joiner status updated for ${staff.name}.`);
  }

  if (event.target.matches('[data-action="preferred-location"]')) {
    staff.preferredLocation = event.target.value;
    persistState();
    renderStaffList();
    setStatus(`Preferred location updated for ${staff.name}.`);
  }

  if (event.target.matches('[data-action="monthly-cap"]')) {
    const value = event.target.value.trim();
    staff.monthlyCap = value ? Number(value) : null;
    persistState();
    renderStaffList();
    setStatus(`Monthly cap updated for ${staff.name}.`);
  }

  if (event.target.matches('[data-action="allowed-locations"]')) {
    const selected = Array.from(event.target.selectedOptions).map(opt => opt.value);
    if (!selected.length) {
      setStatus('Select at least one allowed location.', true);
      renderStaffList();
      return;
    }
    staff.locations = selected;
    persistState();
    renderStaffList();
    setStatus(`Allowed locations updated for ${staff.name}.`);
  }

  if (event.target.matches('[data-action="skill-score"]')) {
    let value = Number(event.target.value);
    if (!Number.isFinite(value)) {
      value = staff.skillScore ?? 3;
    }
    value = Math.min(5, Math.max(1, value));
    staff.skillScore = value;
    event.target.value = value;
    persistState();
    setStatus(`Skill score updated for ${staff.name}.`);
  }

  if (event.target.matches('[data-action="weekday-preferences"]')) {
    staff.avoidWeekdayShifts = Array.from(event.target.selectedOptions).map(opt => opt.value);
    persistState();
    renderStaffList();
    setStatus(`Weekday shift preferences updated for ${staff.name}.`);
  }

  if (event.target.matches('[data-action="weekend-preferences"]')) {
    staff.avoidWeekendShifts = Array.from(event.target.selectedOptions).map(opt => opt.value);
    persistState();
    renderStaffList();
    setStatus(`Weekend shift preferences updated for ${staff.name}.`);
  }

  if (event.target.matches('.shift-off-type')) {
    const form = event.target.closest('.shift-off-form');
    if (form) {
      syncShiftOffFields(form, event.target.value);
    }
  }
}

function parseDaysOffInput(raw) {
  if (!raw) {
    return { success: true, days: [] };
  }
  const tokens = raw.split(',').map(token => token.trim()).filter(Boolean);
  const days = [];
  const seen = new Set();
  for (const token of tokens) {
    const value = Number(token);
    if (!Number.isInteger(value) || value < 1 || value > 31) {
      return { success: false, error: `Invalid day "${token}". Use numbers between 1 and 31.` };
    }
    if (!seen.has(value)) {
      seen.add(value);
      days.push(value);
    }
  }
  days.sort((a, b) => a - b);
  return { success: true, days };
}

function syncShiftOffFields(form, type) {
  if (!form) return;
  const weekdayField = form.querySelector('.shift-off-weekday');
  const dateField = form.querySelector('.shift-off-date');
  if (!weekdayField || !dateField) return;
  const showWeekday = type === 'weekday';
  weekdayField.style.display = showWeekday ? 'flex' : 'none';
  dateField.style.display = showWeekday ? 'none' : 'flex';
}

function handleClosureSubmit(event) {
  event.preventDefault();
  const date = document.getElementById('closure-date').value;
  const location = document.getElementById('closure-location').value;
  if (!date) return;
  if (state.closures.some(item => item.date === date && item.location === location)) {
    setStatus('Closure already recorded.', true);
    return;
  }
  state.closures.push({ id: uid(), date, location });
  persistState();
  renderClosures();
  event.target.reset();
}

function handleClosureListClick(event) {
  if (!event.target.matches('button')) return;
  const id = event.target.dataset.id;
  state.closures = state.closures.filter(item => item.id !== id);
  persistState();
  renderClosures();
}

function handleHolidaySubmit(event) {
  event.preventDefault();
  const date = document.getElementById('holiday-date').value;
  if (!date) return;
  if (state.holidays.some(item => item.date === date)) {
    setStatus('Holiday already listed.', true);
    return;
  }
  state.holidays.push({ id: uid(), date });
  persistState();
  renderHolidays();
  event.target.reset();
}

function handleHolidayListClick(event) {
  if (!event.target.matches('button')) return;
  const id = event.target.dataset.id;
  state.holidays = state.holidays.filter(item => item.id !== id);
  persistState();
  renderHolidays();
}

function handleGenerate() {
  if (!state.staff.length) {
    setStatus('Add staff before generating a schedule.', true);
    return;
  }
  const monthValue = document.getElementById('month-input').value;
  if (!monthValue) {
    setStatus('Select a target month.', true);
    return;
  }
  if (state.settings.month !== monthValue) {
    state.settings.month = monthValue;
    persistState();
  }
  const [yearStr, monthStr] = monthValue.split('-');
  const year = Number(yearStr);
  const monthIndex = Number(monthStr) - 1;
  if (isNaN(year) || isNaN(monthIndex)) {
    setStatus('Invalid month input.', true);
    return;
  }

  const result = buildSchedule({ year, monthIndex });
  state.results = result;
  renderResults();
  if (result) {
    setStatus('Schedule generated.');
  }
}

function renderStaffList() {
  const list = document.getElementById('staff-list');
  if (!state.staff.length) {
    list.innerHTML = '<p>No staff added yet.</p>';
    return;
  }
  const monthLabel = getActiveMonthLabel();
  list.innerHTML = state.staff
    .map(staff => {
      const detailsId = `staff-details-${staff.id}`;
      const chevron = staff.collapsed ? '▸' : '▾';
      const preferredLabel = staff.preferredLocation === 'any'
        ? 'No location preference'
        : `Prefers Location ${staff.preferredLocation}`;
      const daysOffValue = (staff.daysOff || []).join(', ');
      const daysOffSummary = staff.daysOff?.length
        ? staff.daysOff.join(', ')
        : 'None set';
      const shiftPrefList = (staff.preferences && staff.preferences.length)
        ? staff.preferences
            .map(pref => `
              <div class="preference-item">
                <span>${describePreference(pref)}</span>
                <button type="button" class="ghost" data-action="remove-pref" data-pref-id="${pref.id}">✕</button>
              </div>
            `)
            .join('')
        : '<p class="muted">No shift-specific offs.</p>';

      return `
        <article class="staff-card ${staff.collapsed ? 'collapsed' : ''}" data-staff-id="${staff.id}">
          <header>
            <div class="staff-header-main">
              <button type="button" class="staff-toggle" aria-controls="${detailsId}" aria-expanded="${!staff.collapsed}" data-action="toggle-card">
                <span class="chevron">${chevron}</span>
              </button>
              <div>
                <strong>${staff.name}</strong>
                ${staff.isNewJoiner ? '<span class="tag">New joiner</span>' : ''}
              </div>
            </div>
            <div class="staff-header-actions">
              <span class="staff-pref-count">${preferredLabel}</span>
              <button type="button" data-action="remove-staff" class="danger">Remove</button>
            </div>
          </header>
          <div class="staff-details" id="${detailsId}">
            <div class="staff-meta">
              <span>Weekly ≤ ${staff.weeklyCap}h</span>
              <span>Monthly ≤ ${staff.monthlyCap ? staff.monthlyCap + 'h' : 'auto cap'}</span>
            </div>
            <div class="staff-controls">
              <label class="checkbox">
                <input type="checkbox" data-action="new-joiner" ${staff.isNewJoiner ? 'checked' : ''} />
                New joiner
              </label>
              <label>
                Allowed locations
                <select data-action="allowed-locations" multiple data-multi-select data-placeholder="Select locations">
                  <option value="L" ${staff.locations?.includes('L') ? 'selected' : ''}>Location L</option>
                  <option value="C" ${staff.locations?.includes('C') ? 'selected' : ''}>Location C</option>
                </select>
              </label>
              <label>
                Monthly cap
                <input type="number" data-action="monthly-cap" min="1" value="${staff.monthlyCap ?? ''}" placeholder="Default" />
              </label>
              <label>
                Skill score
                <input type="number" data-action="skill-score" min="1" max="5" value="${staff.skillScore ?? 3}" />
              </label>
              <label>
                Preferred location
                <select data-action="preferred-location">
                  <option value="any" ${staff.preferredLocation === 'any' ? 'selected' : ''}>No preference</option>
                  <option value="L" ${staff.preferredLocation === 'L' ? 'selected' : ''}>Prefer Location L</option>
                  <option value="C" ${staff.preferredLocation === 'C' ? 'selected' : ''}>Prefer Location C</option>
                </select>
              </label>
              <label>
                Weekday preferences
                <select data-action="weekday-preferences" multiple data-multi-select data-placeholder="No weekday preference">
                  ${SHIFT_ORDER.map(shift => `
                    <option value="${shift}" ${staff.avoidWeekdayShifts?.includes(shift) ? 'selected' : ''}>
                      Avoid ${shift}
                    </option>
                  `).join('')}
                </select>
              </label>
              <label>
                Weekend preferences
                <select data-action="weekend-preferences" multiple data-multi-select data-placeholder="No weekend preference">
                  ${SHIFT_ORDER.map(shift => `
                    <option value="${shift}" ${staff.avoidWeekendShifts?.includes(shift) ? 'selected' : ''}>
                      Avoid ${shift}
                    </option>
                  `).join('')}
                </select>
              </label>
            </div>
            <div class="days-off-section">
              <form class="days-off-form" data-staff-id="${staff.id}">
                <label>
                  Days off (${monthLabel})
                  <input type="text" name="daysOff" value="${daysOffValue}" placeholder="1, 5, 12" />
                </label>
                <button type="submit">Save</button>
              </form>
              <p class="note">Comma-separated day numbers (1-31). Entire day will be off across both locations.</p>
              <div class="days-off-summary">Current days off: ${daysOffSummary}</div>
            </div>
            <div class="shift-off-section">
              <h4>Shift-specific offs</h4>
              <form class="shift-off-form" data-staff-id="${staff.id}">
                <label>
                  Rule type
                  <select name="ruleType" class="shift-off-type">
                    <option value="weekday">Recurring weekday</option>
                    <option value="date">Specific date</option>
                  </select>
                </label>
                <label class="shift-off-weekday">
                  Weekday
                  <select name="weekday">
                    ${WEEKDAY_LABELS.map((day, idx) => `<option value="${idx}">${day}</option>`).join('')}
                  </select>
                </label>
                <label class="shift-off-date" style="display:none">
                  Date
                  <input type="date" name="prefDate" />
                </label>
                <label>
                  Shift
                  <select name="shiftKey">
                    <option value="all">Whole day</option>
                    <option value="morning">Morning</option>
                    <option value="afternoon">Afternoon</option>
                    <option value="night">Night</option>
                  </select>
                </label>
                <label>
                  Location
                  <select name="shiftLocation">
                    <option value="any">Any</option>
                    <option value="L">Location L</option>
                    <option value="C">Location C</option>
                  </select>
                </label>
                <button type="submit">Add</button>
              </form>
              <div class="preference-list">${shiftPrefList}</div>
            </div>
          </div>
        </article>
      `;
    })
    .join('');
  initializeMultiSelects(list);
}

function getActiveMonthLabel() {
  if (!state.settings.month) return 'selected month';
  const [yearStr, monthStr] = state.settings.month.split('-');
  const year = Number(yearStr);
  const monthIndex = Number(monthStr) - 1;
  if (Number.isNaN(year) || Number.isNaN(monthIndex)) {
    return 'selected month';
  }
  const date = new Date(year, monthIndex, 1);
  return date.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
}

function renderClosures() {
  const list = document.getElementById('closure-list');
  if (!state.closures.length) {
    list.innerHTML = '<li>No closures.</li>';
    return;
  }
  list.innerHTML = state.closures
    .map(entry => `
      <li>
        ${entry.date} • ${entry.location === 'all' ? 'All' : 'Location ' + entry.location}
        <button type="button" data-id="${entry.id}">✕</button>
      </li>
    `)
    .join('');
}

function renderHolidays() {
  const list = document.getElementById('holiday-list');
  if (!state.holidays.length) {
    list.innerHTML = '<li>No legal holidays set.</li>';
    return;
  }
  list.innerHTML = state.holidays
    .map(entry => `
      <li>
        ${entry.date}
        <button type="button" data-id="${entry.id}">✕</button>
      </li>
    `)
    .join('');
}

function renderResults() {
  const container = document.getElementById('results');
  if (!state.results) {
    container.innerHTML = '<p>No schedule generated yet.</p>';
    return;
  }
  const { schedule, warnings, stats, coverage, monthMeta, defaultMonthlyCap } = state.results;
  const monthContext = getMonthContext(monthMeta);

  const warningBlock = warnings.length
    ? `<div class="warn">${warnings.map(w => `<div>⚠️ ${w}</div>`).join('')}</div>`
    : '<div class="success">All shifts satisfied.</div>';

  const calendars = Object.entries(schedule)
    .map(([location, days]) => renderLocationCalendar(location, days, monthContext.year, monthContext.monthIndex))
    .join('');

  const statsList = stats
    .map(stat => `
      <li>
        <strong>${stat.name}</strong>: ${stat.hours.toFixed(1)}h • ${stat.shiftCount} shift(s) • max 7-day hours: ${stat.max7DayHours?.toFixed(1) || '0.0'}h
        ${stat.flags.length ? `<span class="warn">(${stat.flags.join(', ')})</span>` : ''}
        <div class="muted small-text">Off days: ${stat.offDayCount}</div>
      </li>
    `)
    .join('');

  container.innerHTML = `
    <div class="summary">
      <p><strong>${monthContext.label}</strong></p>
      <p><strong>Coverage:</strong> ${coverage.filled} / ${coverage.total} slots (${coverage.percent}%).</p>
      ${warningBlock}
    </div>
    <div class="calendar-wrapper">${calendars}</div>
    <div class="summary">
      <h3>Staff load</h3>
      <p class="muted">Default monthly cap: ${(defaultMonthlyCap ?? 0).toFixed(1)}h</p>
      <ul>${statsList}</ul>
    </div>
  `;
  bindDownloadButtons(schedule, monthContext);
}

function getMonthContext(meta) {
  if (meta && typeof meta.year === 'number' && typeof meta.monthIndex === 'number') {
    return meta;
  }
  if (state.settings.month) {
    const [yearStr, monthStr] = state.settings.month.split('-');
    const fallbackYear = Number(yearStr);
    const fallbackMonth = Number(monthStr) - 1;
    if (!Number.isNaN(fallbackYear) && !Number.isNaN(fallbackMonth)) {
      const labelDate = new Date(fallbackYear, fallbackMonth, 1);
      return {
        year: fallbackYear,
        monthIndex: fallbackMonth,
        label: labelDate.toLocaleDateString(undefined, { month: 'long', year: 'numeric' }),
      };
    }
  }
  const today = new Date();
  return {
    year: today.getFullYear(),
    monthIndex: today.getMonth(),
    label: today.toLocaleDateString(undefined, { month: 'long', year: 'numeric' }),
  };
}

function renderLocationCalendar(location, days, year, monthIndex) {
  const cells = buildCalendarMatrix(year, monthIndex);
  const locationSchedule = days || {};
  const weekdayHeaders = CALENDAR_WEEKDAY_LABELS.map(day => `<div class="calendar-header">${day}</div>`).join('');
  const cellMarkup = cells
    .map((dayObj, index) => renderDayCell(dayObj, locationSchedule, index % 7 === 0))
    .join('');

  return `
    <section class="calendar-card">
      <div class="calendar-card-header">
        <h3>Location ${location}</h3>
        <button type="button" class="ghost download-csv" data-download-location="${location}">Download CSV</button>
      </div>
      <div class="calendar-grid">
        ${weekdayHeaders}
        ${cellMarkup}
      </div>
    </section>
  `;
}

function bindDownloadButtons(schedule, monthContext) {
  const container = document.getElementById('results');
  if (!container) return;
  container.querySelectorAll('[data-download-location]').forEach(button => {
    button.addEventListener('click', () => {
      const location = button.dataset.downloadLocation;
      const days = schedule[location] || {};
      const csv = buildLocationCsvTable({ location, days, monthContext });
      const year = monthContext.year ?? new Date().getFullYear();
      const month = typeof monthContext.monthIndex === 'number'
        ? String(monthContext.monthIndex + 1).padStart(2, '0')
        : String(new Date().getMonth() + 1).padStart(2, '0');
      const fileName = `schedule-${location}-${year}-${month}.csv`;
      triggerCsvDownload(csv, fileName);
    });
  });
}

function buildCalendarMatrix(year, monthIndex) {
  const rawFirstDay = new Date(year, monthIndex, 1).getDay();
  const firstDay = (rawFirstDay + 6) % 7; // shift so Monday is column 0
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const totalCells = Math.ceil((firstDay + daysInMonth) / 7) * 7;
  const cells = [];
  for (let index = 0; index < totalCells; index += 1) {
    const dayNumber = index - firstDay + 1;
    if (dayNumber < 1 || dayNumber > daysInMonth) {
      cells.push(null);
    } else {
      const dateObj = new Date(year, monthIndex, dayNumber);
      cells.push({
        iso: formatISO(dateObj),
        dayNumber,
      });
    }
  }
  return cells;
}

function renderDayCell(dayObj, locationSchedule, showLabels) {
  if (!dayObj) {
    const placeholderSchedule = {
      morning: { names: [], required: 0 },
      afternoon: { names: [], required: 0 },
      night: { names: [], required: 0 },
    };
    const placeholderRows = showLabels
      ? SHIFT_ORDER.map(shift => renderShiftRow(shift, placeholderSchedule, true)).join('')
      : '';
    return `
      <div class="calendar-cell calendar-cell--empty">
        <div class="calendar-date">&nbsp;</div>
        ${placeholderRows}
      </div>
    `;
  }
  const scheduleDay = locationSchedule[dayObj.iso];
  if (!scheduleDay) {
    const placeholderSchedule = {
      morning: { names: [], required: 0 },
      afternoon: { names: [], required: 0 },
      night: { names: [], required: 0 },
    };
    const rows = showLabels
      ? SHIFT_ORDER.map(shift => renderShiftRow(shift, placeholderSchedule, true)).join('')
      : '';
    return `
      <div class="calendar-cell calendar-cell--closed">
        <div class="calendar-date">${dayObj.dayNumber}</div>
        ${rows}
        <div class="calendar-closed">Closed</div>
      </div>
    `;
  }

  const shiftRows = SHIFT_ORDER.map(shift => renderShiftRow(shift, scheduleDay, showLabels)).join('');

  return `
    <div class="calendar-cell">
      <div class="calendar-date">${dayObj.dayNumber}</div>
      ${shiftRows}
    </div>
  `;
}

function renderShiftRow(shiftKey, scheduleDay, showLabel) {
  const shiftInfo = scheduleDay[shiftKey];
  const hasShift = Boolean(shiftInfo);
  const assignments = hasShift ? shiftInfo.names : [];
  const required = hasShift ? shiftInfo.required : 0;
  const slots = renderShiftSlots(assignments, hasShift, required);
  const shortage = hasShift ? Math.max(0, required - assignments.length) : 0;
  const rowClasses = ['shift-row'];
  if (!hasShift) rowClasses.push('shift-row--disabled');
  return `
    <div class="${rowClasses.join(' ')}">
      <div class="shift-label">${showLabel ? SHIFT_LABELS[shiftKey] : ''}</div>
      <div class="shift-grid">${slots}</div>
      ${shortage > 0 ? '<div class="shift-note warn">⚠️</div>' : ''}
    </div>
  `;
}

function renderShiftSlots(assignments, isActive, required) {
  if (!isActive) {
    return Array.from({ length: 4 })
      .map(() => '<div class="shift-slot placeholder">-</div>')
      .join('');
  }
  if (!assignments.length) {
    return Array.from({ length: 4 })
      .map(() => '<div class="shift-slot empty"></div>')
      .join('');
  }
  if (assignments.length <= 4) {
    const slots = Array.from({ length: 4 }, (_, idx) => assignments[idx] || '');
    return slots
      .map(name => `<div class="shift-slot">${name ? name : ''}</div>`)
      .join('');
  }
  const slots = [
    assignments[0],
    assignments[1] || '',
    assignments[2] || '',
    `+${assignments.length - 3} more`,
  ];
  return slots
    .map(name => `<div class="shift-slot">${name || ''}</div>`)
    .join('');
}

function buildLocationCsvTable({ location, days, monthContext }) {
  const year = monthContext.year ?? new Date().getFullYear();
  const monthIndex = typeof monthContext.monthIndex === 'number'
    ? monthContext.monthIndex
    : new Date().getMonth();
  const cells = buildCalendarMatrix(year, monthIndex);
  const rows = [];
  rows.push([`Location ${location}`]);
  rows.push([]);

  const STAFF_COLUMNS = 4;

  for (let index = 0; index < cells.length; index += 7) {
    const weekCells = cells.slice(index, index + 7);
    if (!weekCells.length) continue;

    const weekdayRow = ['Shift'];
    const dayNumberRow = ['Day'];
    for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
      const cell = weekCells[dayIndex] || null;
      const dayLabel = CALENDAR_WEEKDAY_LABELS[dayIndex];
      weekdayRow.push(dayLabel);
      for (let s = 1; s < STAFF_COLUMNS; s += 1) {
        weekdayRow.push('');
      }
      const dayNumber = cell ? cell.dayNumber : '';
      dayNumberRow.push(dayNumber || '');
      for (let s = 1; s < STAFF_COLUMNS; s += 1) {
        dayNumberRow.push('');
      }
    }
    rows.push(weekdayRow);
    rows.push(dayNumberRow);

    SHIFT_ORDER.forEach(shiftKey => {
      const shiftRow = [SHIFT_LABELS[shiftKey]];
      for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
        const cell = weekCells[dayIndex] || null;
        const slots = Array(STAFF_COLUMNS).fill('');
        if (!cell) {
          shiftRow.push(...slots);
          continue;
        }
        const scheduleDay = (days || {})[cell.iso];
        if (!scheduleDay) {
          slots[0] = 'Closed';
          shiftRow.push(...slots);
          continue;
        }
        const shiftInfo = scheduleDay[shiftKey];
        if (!shiftInfo) {
          shiftRow.push(...slots);
          continue;
        }
        shiftInfo.names.slice(0, STAFF_COLUMNS).forEach((name, idx) => {
          slots[idx] = name || '';
        });
        if (shiftInfo.names.length > STAFF_COLUMNS) {
          const extra = shiftInfo.names.length - STAFF_COLUMNS;
          slots[STAFF_COLUMNS - 1] = `${slots[STAFF_COLUMNS - 1]} +${extra} more`.trim();
        }
        if (shiftInfo.names.length < shiftInfo.required) {
          const warnIndex = Math.min(shiftInfo.names.length, STAFF_COLUMNS - 1);
          slots[warnIndex] = slots[warnIndex] ? `${slots[warnIndex]} ⚠️` : '⚠️';
        }
        shiftRow.push(...slots);
      }
      rows.push(shiftRow);
    });

    rows.push([]);
  }

  if (rows.length && rows[rows.length - 1].length === 0) {
    rows.pop();
  }

  return rows.map(row => row.map(toCsvCell).join(',')).join('\n');
}

function toCsvCell(value) {
  const str = value == null ? '' : String(value);
  if (/[",\n]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

function triggerCsvDownload(csv, fileName) {
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(link.href), 0);
}

function describePreference(pref) {
  const shiftText = pref.shifts.includes('all') ? 'all shifts' : pref.shifts.join(', ');
  const locationText = pref.location === 'any' ? 'any location' : `location ${pref.location}`;
  if (pref.type === 'weekday') {
    return `${WEEKDAY_LABELS[pref.weekday]} • ${shiftText} @ ${locationText}`;
  }
  return `${formatDisplayDate(pref.date)} • ${shiftText} @ ${locationText}`;
}

function buildSchedule({ year, monthIndex }) {
  const monthMeta = deriveMonthMeta(year, monthIndex);
  const holidaySet = new Set(state.holidays.map(item => item.date));
  const closures = buildClosureSets();
  const shiftEntries = buildShiftEntries({ year, monthIndex, closures });
  if (!shiftEntries.length) {
    setStatus('No open shifts in the selected month.', true);
    return null;
  }
  const activeDates = Array.from(new Set(shiftEntries.map(entry => entry.date))).sort();

  const workableDays = countWorkableDays({ year, monthIndex, holidaySet });
  const defaultMonthlyCap = workableDays * 8;
  const maxWorkingDaysAllowed = activeDates.length > 10 ? activeDates.length - 10 : activeDates.length;
  const staffContext = state.staff.map(staff => ({
    ...staff,
    skillScore: typeof staff.skillScore === 'number' ? staff.skillScore : 3,
    weeklyHours: new Map(),
    monthlyHours: 0,
    monthlyCap: Math.min(staff.monthlyCap ?? Infinity, defaultMonthlyCap),
    dayAssignments: {},
    lastWorkedDay: null,
    streak: 0,
    totalAssignments: 0,
    dailyHours: [],
    max7DayHours: 0,
    lastLocation: null,
    workedDates: new Set(),
    maxWorkingDays: maxWorkingDaysAllowed,
    newJoinerAssignments: 0,
  }));

  const schedule = { L: {}, C: {} };
  const warnings = [];
  const coverage = { filled: 0, total: 0 };
  const shiftNewJoiners = {};
  const newJoinerLimit = Number(state.settings.newJoinerLimit) || 1;

  shiftEntries.forEach(entry => {
    if (!schedule[entry.location][entry.date]) {
      schedule[entry.location][entry.date] = {};
    }
    schedule[entry.location][entry.date][entry.key] = { names: [], required: entry.required };
    const shiftBlock = schedule[entry.location][entry.date][entry.key];
    const assignedForShift = [];

    for (let slot = 0; slot < entry.required; slot += 1) {
      coverage.total += 1;
      const candidate = pickStaffForShift({
        staffContext,
        entry,
        shiftNewJoiners,
        newJoinerLimit,
        currentShiftStaff: assignedForShift,
      });
      if (!candidate) {
        warnings.push(
          `Unable to fill ${entry.location} ${entry.key} on ${formatDisplayDate(entry.date)} (slot ${slot + 1}).`
        );
        continue;
      }
      coverage.filled += 1;
      assignedForShift.push(candidate);
      shiftBlock.names.push(candidate.name);
    }
  });

  const stats = staffContext.map(staff => {
    const shiftCount = Object.values(staff.dayAssignments).reduce((acc, shifts) => acc + shifts.length, 0);
    const offDayCount = activeDates.length - staff.workedDates.size;
    return {
      name: staff.name,
      hours: staff.monthlyHours,
      shiftCount,
      max7DayHours: staff.max7DayHours || 0,
      offDayCount,
      flags: buildStaffFlags(staff, workableDays),
    };
  });

  const percent = coverage.total ? ((coverage.filled / coverage.total) * 100).toFixed(1) : '0.0';

  return {
    schedule,
    warnings,
    stats,
    coverage: { ...coverage, percent },
    monthMeta,
    defaultMonthlyCap,
    workingDates: activeDates,
  };
}

function pickStaffForShift({ staffContext, entry, shiftNewJoiners, newJoinerLimit, currentShiftStaff = [] }) {
  let candidates = staffContext.filter(staff => canWorkShift({ staff, entry, shiftNewJoiners, newJoinerLimit }));
  if (!candidates.length) return null;
  const hasHighSkillAssigned = currentShiftStaff.some(staff => staff.skillScore >= HIGH_SKILL_THRESHOLD);
  if (hasHighSkillAssigned) {
    const nonHigh = candidates.filter(staff => staff.skillScore < HIGH_SKILL_THRESHOLD);
    if (nonHigh.length) {
      candidates = nonHigh;
    }
  }
  candidates.sort((a, b) => {
    const aDayAssignments = a.dayAssignments[entry.date] || [];
    const bDayAssignments = b.dayAssignments[entry.date] || [];
    const sameLocationDiff =
      bDayAssignments.filter(item => item.location === entry.location).length -
      aDayAssignments.filter(item => item.location === entry.location).length;
    if (sameLocationDiff !== 0) return sameLocationDiff;

    const sameDayDiff = bDayAssignments.length - aDayAssignments.length;
    if (sameDayDiff !== 0) return sameDayDiff;

    const flexDiff = locationFlexScore(a, entry.location) - locationFlexScore(b, entry.location);
    if (flexDiff !== 0) return flexDiff;

    if (a.isNewJoiner && b.isNewJoiner && a.newJoinerAssignments !== b.newJoinerAssignments) {
      return a.newJoinerAssignments - b.newJoinerAssignments;
    }

    const offPrefDiff = (b.daysOff?.length || 0) - (a.daysOff?.length || 0);
    if (offPrefDiff !== 0) return offPrefDiff;

    const continuityDiff = locationContinuityScore(a, entry.location) - locationContinuityScore(b, entry.location);
    if (continuityDiff !== 0) return continuityDiff;
    const prefDiff = locationPreferenceScore(a, entry.location) - locationPreferenceScore(b, entry.location);
    if (prefDiff !== 0) return prefDiff;
    if (a.monthlyHours !== b.monthlyHours) return a.monthlyHours - b.monthlyHours;
    if (a.totalAssignments !== b.totalAssignments) return a.totalAssignments - b.totalAssignments;
    const restDiff = daysSinceLastShift(b, entry.date) - daysSinceLastShift(a, entry.date);
    if (restDiff !== 0) return restDiff;
    return 0;
  });
  const chosen = candidates[0];
  commitAssignment({ staff: chosen, entry, shiftNewJoiners });
  return chosen;
}

function canWorkShift({ staff, entry, shiftNewJoiners, newJoinerLimit }) {
  if (!staff.locations.includes(entry.location)) return false;
  if (staff.preferredLocation && staff.preferredLocation !== 'any' && staff.preferredLocation !== entry.location) {
    return false;
  }
  const dayOfWeek = new Date(`${entry.date}T00:00:00`).getDay();
  const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
  const avoidList = isWeekend ? staff.avoidWeekendShifts : staff.avoidWeekdayShifts;
  if (avoidList?.includes(entry.key)) return false;
  if (isStaffUnavailable(staff, entry)) return false;

  const dateAssignments = staff.dayAssignments[entry.date] || [];
  const alreadyWorkingToday = dateAssignments.length > 0;
  if (!alreadyWorkingToday && staff.workedDates && staff.workedDates.size >= staff.maxWorkingDays) {
    return false;
  }
  if (dateAssignments.some(item => item.shift === entry.key)) return false;
  const hasMorning = dateAssignments.some(item => item.shift === 'morning');
  const hasAfternoon = dateAssignments.some(item => item.shift === 'afternoon');
  const hasNight = dateAssignments.some(item => item.shift === 'night');
  const coveredDay = hasMorning && hasAfternoon;
  if (entry.key === 'night' && hasMorning && !coveredDay) return false;
  if (entry.key === 'morning' && hasNight) return false;

  const weekKey = getWeekKey(entry.date);
  const weeklyHours = staff.weeklyHours.get(weekKey) || 0;
  if (weeklyHours + entry.hours > staff.weeklyCap) return false;
  if (staff.monthlyHours + entry.hours > staff.monthlyCap) return false;

  if (wouldHitSevenDayStreak(staff, entry.date)) return false;

  if (staff.isNewJoiner) {
    const shiftKey = getShiftKey(entry);
    const shiftCount = shiftNewJoiners[shiftKey] || 0;
    if (shiftCount >= newJoinerLimit) return false;
  }

  return true;
}

function isStaffUnavailable(staff, entry) {
  if (Array.isArray(staff.daysOff) && staff.daysOff.length) {
    const dayOfMonth = getDayOfMonth(entry.date);
    if (staff.daysOff.includes(dayOfMonth)) {
      return true;
    }
  }
  return (staff.preferences || []).some(pref => {
    if (pref.location !== 'any' && pref.location !== entry.location) return false;
    if (pref.shifts.includes('all') || pref.shifts.includes(entry.key)) {
      if (pref.type === 'weekday' && pref.weekday === entry.dayOfWeek) return true;
      if (pref.type === 'date' && pref.date === entry.date) return true;
    }
    return false;
  });
}

function commitAssignment({ staff, entry, shiftNewJoiners }) {
  const weekKey = getWeekKey(entry.date);
  const weeklyHours = staff.weeklyHours.get(weekKey) || 0;
  staff.weeklyHours.set(weekKey, weeklyHours + entry.hours);
  staff.monthlyHours += entry.hours;

  if (!staff.dayAssignments[entry.date]) {
    staff.dayAssignments[entry.date] = [];
  }
  if (staff.dayAssignments[entry.date].length === 0 && staff.workedDates) {
    staff.workedDates.add(entry.date);
  }
  staff.dayAssignments[entry.date].push({ location: entry.location, shift: entry.key });
  staff.totalAssignments += 1;
  staff.lastLocation = entry.location;

  if (staff.lastWorkedDay === entry.date) {
    // already counted for streak purposes
  } else {
    if (staff.lastWorkedDay) {
      const diff = diffInDays(staff.lastWorkedDay, entry.date);
      staff.streak = diff === 1 ? staff.streak + 1 : 1;
    } else {
      staff.streak = 1;
    }
    staff.lastWorkedDay = entry.date;
  }

  updateRollingHours(staff, entry);

  if (staff.isNewJoiner) {
    staff.newJoinerAssignments = (staff.newJoinerAssignments || 0) + 1;
    const shiftKey = getShiftKey(entry);
    shiftNewJoiners[shiftKey] = (shiftNewJoiners[shiftKey] || 0) + 1;
  }
}

function wouldHitSevenDayStreak(staff, targetDate) {
  if (!staff.lastWorkedDay) return false;
  if (staff.lastWorkedDay === targetDate) return false;
  const diff = diffInDays(staff.lastWorkedDay, targetDate);
  if (diff === 1 && staff.streak >= 6) {
    return true;
  }
  return false;
}

function buildStaffFlags(staff, workableDays) {
  const flags = [];
  if (staff.monthlyHours >= staff.monthlyCap - 0.1) {
    flags.push('monthly limit reached');
  }
  const expected = workableDays ? ((staff.monthlyHours / (workableDays * 8)) * 100) : 0;
  if (expected < 30) {
    flags.push('low utilization');
  }
  return flags;
}

function updateRollingHours(staff, entry) {
  if (!staff.dailyHours) {
    staff.dailyHours = [];
    staff.max7DayHours = 0;
  }
  staff.dailyHours.push({ date: entry.date, hours: entry.hours });
  staff.dailyHours.sort((a, b) => (a.date < b.date ? -1 : 1));
  let windowHours = 0;
  let startIndex = 0;
  for (let endIndex = 0; endIndex < staff.dailyHours.length; endIndex += 1) {
    windowHours += staff.dailyHours[endIndex].hours;
    while (
      diffInDays(staff.dailyHours[startIndex].date, staff.dailyHours[endIndex].date) >= 7
    ) {
      windowHours -= staff.dailyHours[startIndex].hours;
      startIndex += 1;
    }
    if (windowHours > (staff.max7DayHours || 0)) {
      staff.max7DayHours = windowHours;
    }
  }
}

function deriveMonthMeta(year, monthIndex) {
  const date = new Date(year, monthIndex, 1);
  const label = date.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
  return { label, year, monthIndex };
}

function buildShiftEntries({ year, monthIndex, closures }) {
  const entries = [];
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day += 1) {
    const dateObj = new Date(year, monthIndex, day);
    const iso = formatISO(dateObj);
    const dayOfWeek = dateObj.getDay();

    if (!isClosed('C', iso, closures)) {
      const shifts = dayOfWeek === 0
        ? []
        : dayOfWeek === 6
          ? SHIFT_BLUEPRINTS.C.saturday
          : SHIFT_BLUEPRINTS.C.weekday;
      shifts.forEach(shift => entries.push({ ...shift, location: 'C', date: iso, dayOfWeek }));
    }

    if (!isClosed('L', iso, closures)) {
      SHIFT_BLUEPRINTS.L.default.forEach(shift => {
        let required = shift.required;
        if (shift.key === 'morning' && (dayOfWeek === 2 || dayOfWeek === 4)) {
          required = 3;
        }
        if (dayOfWeek === 5) {
          if (shift.key === 'afternoon') required = 2;
          if (shift.key === 'night') required = 1;
        }
        if ((dayOfWeek === 0 || dayOfWeek === 6) && shift.key === 'night') {
          return;
        }
        entries.push({ ...shift, required, location: 'L', date: iso, dayOfWeek });
      });
    }
  }
  return entries;
}

function buildClosureSets() {
  const sets = { all: new Set(), L: new Set(), C: new Set() };
  state.closures.forEach(entry => {
    if (entry.location === 'all') {
      sets.all.add(entry.date);
    } else {
      sets[entry.location].add(entry.date);
    }
  });
  return sets;
}

function isClosed(location, date, closures) {
  return closures.all.has(date) || closures[location].has(date);
}

function countWorkableDays({ year, monthIndex, holidaySet }) {
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  let count = 0;
  for (let day = 1; day <= daysInMonth; day += 1) {
    const dateObj = new Date(year, monthIndex, day);
    const iso = formatISO(dateObj);
    const dow = dateObj.getDay();
    const isWeekend = dow === 0 || dow === 6;
    if (isWeekend) continue;
    if (holidaySet.has(iso)) continue;
    count += 1;
  }
  return count;
}

function getWeekKey(dateStr) {
  const date = new Date(`${dateStr}T00:00:00`);
  const firstDay = new Date(date.getFullYear(), date.getMonth(), 1);
  const diff = Math.floor((date - firstDay) / (24 * 60 * 60 * 1000));
  const offset = firstDay.getDay();
  const weekIndex = Math.floor((diff + offset) / 7);
  return `${date.getFullYear()}-${date.getMonth()}-${weekIndex}`;
}

function diffInDays(a, b) {
  const first = new Date(`${a}T00:00:00`);
  const second = new Date(`${b}T00:00:00`);
  return Math.round((second - first) / (24 * 60 * 60 * 1000));
}

function formatISO(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function formatDisplayDate(iso) {
  const date = new Date(`${iso}T00:00:00`);
  return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric', weekday: 'short' });
}

function getDayOfMonth(isoDate) {
  return Number(isoDate.split('-')[2]);
}

function locationPreferenceScore(staff, location) {
  if (!staff.preferredLocation || staff.preferredLocation === 'any') return 1;
  return staff.preferredLocation === location ? 0 : 2;
}

function locationFlexScore(staff, location) {
  if (staff.preferredLocation && staff.preferredLocation !== 'any') {
    return staff.preferredLocation === location ? 0 : 5;
  }
  if (!Array.isArray(staff.locations) || !staff.locations.length) return 5;
  if (staff.locations.length === 1 && staff.locations[0] === location) return 0;
  return staff.locations.length;
}

function locationContinuityScore(staff, location) {
  if (!staff.lastLocation) return 1;
  return staff.lastLocation === location ? 0 : 2;
}

function daysSinceLastShift(staff, targetDate) {
  if (!staff.lastWorkedDay) return Number.POSITIVE_INFINITY;
  return diffInDays(staff.lastWorkedDay, targetDate);
}

function getShiftKey(entry) {
  return `${entry.date}|${entry.key}`;
}

function setStatus(message, isError = false) {
  const el = document.getElementById('status-message');
  el.textContent = message;
  el.style.color = isError ? '#c0392b' : '#057a55';
}
