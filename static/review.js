(() => {
  const data = document.querySelector('#roster-data');
  if (!data) return;
  const rosters = JSON.parse(data.textContent);
  const tabs = [...document.querySelectorAll('[data-roster-tab]')];
  const panels = [...document.querySelectorAll('[data-roster-panel]')];
  const search = document.querySelector('#roster-search');
  const zoneFilter = document.querySelector('#roster-zone-filter');
  const results = document.querySelector('#roster-results');
  const source = document.querySelector('#review-source');
  const missionary = document.querySelector('#review-missionary');
  const companionList = document.querySelector('#review-companions');
  const companionSearch = document.querySelector('#companion-search');
  const editorFields = document.querySelector('#editor-fields');
  const editorPanel = document.querySelector('#correction-panel');
  const normalize = value => value.toLocaleLowerCase().normalize('NFD').replace(/\p{M}/gu, '');
  const option = (label, value) => new Option(label, value);
  let activeRoster = 'previous';

  document.querySelectorAll('[data-roster-row]').forEach(row => {
    const person = rosters[row.dataset.source][Number(row.dataset.index)];
    row.dataset.search = normalize([person.title, person.full_name, person.zone, person.area, person.position, ...person.companions].join(' '));
  });

  function filterRoster() {
    const query = normalize(search.value.trim());
    let visible = 0;
    const panel = panels.find(item => item.dataset.rosterPanel === activeRoster);
    panel.querySelectorAll('[data-roster-row]').forEach(row => {
      row.hidden = !row.dataset.search.includes(query) || Boolean(zoneFilter.value && row.dataset.zone !== zoneFilter.value);
      if (!row.hidden) visible += 1;
    });
    panel.querySelector('.roster-empty').hidden = visible !== 0;
    results.textContent = `Showing ${visible} of ${rosters[activeRoster].length} missionaries`;
  }

  function showPanel(value) {
    activeRoster = value;
    tabs.forEach(tab => {
      const selected = tab.dataset.rosterTab === value;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;
    });
    panels.forEach(panel => { panel.hidden = panel.dataset.rosterPanel !== value; });
    const previousZone = zoneFilter.value;
    const zones = [...new Set(rosters[value].map(row => row.zone))].sort();
    zoneFilter.replaceChildren(option('All zones', ''), ...zones.map(zone => option(zone, zone)));
    zoneFilter.value = zones.includes(previousZone) ? previousZone : '';
    filterRoster();
  }

  tabs.forEach((tab, index) => {
    tab.addEventListener('click', () => showPanel(tab.dataset.rosterTab));
    tab.addEventListener('keydown', event => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      const nextIndex = event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 :
        (index + (event.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
      tabs[nextIndex].click();
      tabs[nextIndex].focus();
    });
  });
  document.querySelector('.roster-tabs').hidden = false;
  search.addEventListener('input', filterRoster);
  zoneFilter.addEventListener('change', filterRoster);
  document.querySelectorAll('[data-open-fold]').forEach(link => {
    link.addEventListener('click', () => { document.getElementById(link.dataset.openFold).open = true; });
  });
  showPanel(activeRoster);

  if (!source) return; // Completed reviews keep roster navigation available.

  function selectedScope() {
    const count = companionList.querySelectorAll('input:checked').length;
    document.querySelector('#companion-count').textContent = `${count} selected`;
    document.querySelector('#assignment-scope').textContent = count ?
      `Zone and area will update ${count + 1} people: this missionary and ${count} companion${count === 1 ? '' : 's'}.` :
      'Only this missionary will be updated. No companions are selected.';
  }

  function filterCompanions() {
    const query = normalize(companionSearch.value.trim());
    let visible = 0;
    companionList.querySelectorAll('.companion-choice').forEach(choice => {
      choice.hidden = !choice.dataset.search.includes(query);
      if (!choice.hidden) visible += 1;
    });
    document.querySelector('#companion-empty').hidden = visible !== 0;
  }

  function showPerson() {
    const hasSelection = missionary.value !== '';
    editorFields.hidden = !hasSelection;
    editorFields.disabled = !hasSelection;
    document.querySelector('#editor-empty').hidden = hasSelection;
    document.querySelectorAll('[data-roster-row]').forEach(row => {
      row.classList.toggle('is-selected', hasSelection && row.dataset.source === source.value && row.dataset.index === missionary.value);
    });
    if (!hasSelection) return;
    const rows = rosters[source.value];
    const row = rows[Number(missionary.value)];
    document.querySelector('#review-zone').value = row.zone;
    document.querySelector('#review-area').value = row.area;
    document.querySelector('#review-role').value = row.position;
    document.querySelector('#review-reason').value = '';
    document.querySelector('#editor-person-name').textContent = `${row.title} ${row.full_name}`.trim();
    document.querySelector('#editor-avatar').textContent = row.full_name.slice(0, 1).toLocaleUpperCase();
    companionSearch.value = '';
    companionList.replaceChildren();
    rows.map((person, index) => ({person, index})).sort((a, b) => a.person.full_name.localeCompare(b.person.full_name)).forEach(({person, index}) => {
      if (index === Number(missionary.value)) return;
      const label = document.createElement('label');
      label.className = 'companion-choice';
      label.dataset.search = normalize(`${person.full_name} ${person.area} ${person.title}`);
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.name = 'companions';
      checkbox.value = String(index);
      checkbox.checked = row.companions.includes(person.full_name);
      checkbox.addEventListener('change', selectedScope);
      const text = document.createElement('span');
      const name = document.createElement('strong');
      name.textContent = person.full_name;
      const detail = document.createElement('small');
      detail.textContent = `${person.title || 'Missionary'} · ${person.area}`;
      text.append(name, detail);
      label.append(checkbox, text);
      companionList.append(label);
    });
    selectedScope();
    filterCompanions();
  }

  function showEditorRoster(index = '') {
    const rows = rosters[source.value];
    const entries = rows.map((row, i) => ({row, i})).sort((a, b) => a.row.full_name.localeCompare(b.row.full_name));
    missionary.replaceChildren(option('Choose a missionary…', ''), ...entries.map(({row, i}) => option(`${row.title} ${row.full_name} — ${row.area}`, i)));
    missionary.value = String(index);
    const context = document.querySelector('#editor-context');
    context.classList.toggle('is-current', source.value === 'current');
    context.textContent = source.value === 'previous' ?
      'Previous baseline changes where missionaries move FROM. Use it for emergency transfers before this cycle.' :
      'Current roster changes where missionaries move TO in the new transfer.';
    for (const [field, selector] of [['zone', '#review-zones'], ['area', '#review-areas'], ['position', '#review-roles']]) {
      document.querySelector(selector).replaceChildren(...[...new Set(rows.map(row => row[field]).filter(Boolean))].sort().map(value => option(value, value)));
    }
    showPerson();
  }

  source.addEventListener('change', () => {
    showEditorRoster();
    showPanel(source.value);
  });
  missionary.addEventListener('change', showPerson);
  companionSearch.addEventListener('input', filterCompanions);
  document.querySelector('#clear-companions').addEventListener('click', () => {
    companionList.querySelectorAll('input').forEach(checkbox => { checkbox.checked = false; });
    selectedScope();
  });
  document.querySelector('#reset-assignment').addEventListener('click', showPerson);
  document.querySelectorAll('[data-edit-source]').forEach(button => {
    button.addEventListener('click', () => {
      source.value = button.dataset.editSource;
      showEditorRoster(button.dataset.editIndex);
      editorPanel.scrollTop = 0;
      if (window.matchMedia('(max-width: 860px)').matches) {
        const heading = document.querySelector('#editor-heading');
        heading.tabIndex = -1;
        heading.focus({preventScroll: true});
        editorPanel.scrollIntoView({behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth', block: 'start'});
      } else {
        document.querySelector('#review-zone').focus();
      }
    });
  });
  showEditorRoster();
})();
