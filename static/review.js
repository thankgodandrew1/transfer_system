const rosterData = document.querySelector('#roster-data');
if (rosterData) {
  const rosters = JSON.parse(rosterData.textContent);
  const source = document.querySelector('#review-source');
  const missionary = document.querySelector('#review-missionary');
  const companionSelect = document.querySelector('#review-companions');
  const option = (label, value) => new Option(label, value);
  function showPerson() {
    const rows = rosters[source.value];
    const row = rows[Number(missionary.value)];
    if (!row) return;
    document.querySelector('#review-zone').value = row.zone;
    document.querySelector('#review-area').value = row.area;
    document.querySelector('#review-role').value = row.position;
    companionSelect.replaceChildren();
    rows.forEach((candidate, index) => {
      if (index === Number(missionary.value)) return;
      const item = option(`${candidate.title} ${candidate.full_name} — ${candidate.area}`, index);
      item.selected = row.companions.includes(candidate.full_name);
      companionSelect.append(item);
    });
  }
  function showRoster() {
    const rows = rosters[source.value];
    missionary.replaceChildren(...rows.map((row, index) => option(`${row.title} ${row.full_name} — ${row.area}`, index)));
    for (const [field, selector] of [['zone', '#review-zones'], ['area', '#review-areas']]) {
      document.querySelector(selector).replaceChildren(...[...new Set(rows.map(row => row[field]))].sort().map(value => option(value, value)));
    }
    showPerson();
  }
  if (source) {
    source.addEventListener('change', showRoster);
    missionary.addEventListener('change', showPerson);
    showRoster();
  }
  document.querySelector('#roster-search').addEventListener('input', event => {
    const search = event.target.value.toLocaleLowerCase();
    document.querySelectorAll('.roster-table tbody tr').forEach(row => {
      row.hidden = !row.textContent.toLocaleLowerCase().includes(search);
    });
  });
}
