let tcApi;
let groups = [];
let selectedGroups = new Set();
let groupColors = {};
let groupColorReasons = {};
let sortColumns = [];
let expandedGroups = new Set();
let coloringActive = false;
let currentViewMode = 'selected';
let lastSelectedRow = null;
let colorMode = 'random';

const servicesPalette = [
  [0, 0, 255], [0, 255, 0], [0, 255, 255], [255, 0, 0], [255, 0, 255],
  [191, 127, 255], [127, 191, 255], [127, 255, 191], [191, 255, 127], [255, 127, 191],
  [255, 191, 127], [0, 127, 255], [127, 0, 255], [127, 255, 0], [0, 255, 127],
  [255, 127, 0], [255, 0, 127], [159, 127, 255], [127, 159, 255], [127, 255, 159],
  [159, 255, 127], [255, 127, 159], [255, 159, 127], [0, 191, 223], [27, 118, 97],
  [191, 223, 223], [223, 0, 191], [223, 191, 223], [0, 223, 255], [223, 0, 255],
  [200, 150, 80], [0, 255, 223], [205, 182, 188], [255, 0, 223], [223, 191, 95],
  [191, 223, 95], [191, 95, 223], [223, 95, 191], [95, 191, 223], [95, 223, 191]
];
let servicesPaletteIndex = 0;

const debugConsole = document.getElementById('debugConsole');

function debug(msg) {
  const time = new Date().toLocaleTimeString();
  debugConsole.textContent += `[${time}] ${msg}\n`;
  debugConsole.scrollTop = debugConsole.scrollHeight;
}

const columnConfig = [
  { key: 'expand', name: '#', visible: true },
  { key: 'mark', name: 'Mark', visible: true },
  { key: 'h1', name: 'H1', visible: true },
  { key: 'v1', name: 'V1', visible: true },
  { key: 'z1', name: 'Z1', visible: true },
  { key: 'h2', name: 'H2', visible: true },
  { key: 'v2', name: 'V2', visible: true },
  { key: 'z2', name: 'Z2', visible: true },
  { key: 'subassembly', name: 'Subassembly -1', visible: true },
  { key: 'diameter', name: 'Diam.', visible: true },
  { key: 'shape', name: 'Shape code', visible: true },
  { key: 'length', name: 'Length', visible: true },
  { key: 'quantity', name: 'Count', visible: true },
  { key: 'reason', name: 'Color Reason', visible: true }
];

let filters = [];
let groupBy = [];

// Views Dropdown
const viewsDropdown = document.getElementById('viewsDropdown');
const viewsMenu = document.getElementById('viewsMenu');
const viewsText = document.getElementById('viewsText');

viewsDropdown.addEventListener('click', (e) => {
  e.stopPropagation();
  viewsMenu.classList.toggle('show');
});

document.querySelectorAll('#viewsMenu .views-item').forEach(item => {
  item.addEventListener('click', (e) => {
    e.stopPropagation();
    viewsText.textContent = item.textContent.trim();
    currentViewMode = item.dataset.value;

    document.querySelectorAll('.checkmark').forEach(img => img.classList.remove('selected'));
    item.querySelector('.checkmark').classList.add('selected');

    document.querySelectorAll('.views-item').forEach(i => i.classList.remove('selected'));
    item.classList.add('selected');

    viewsMenu.classList.remove('show');
  });
});

// Close popups only on true outside clicks
document.addEventListener('click', (e) => {
  document.querySelectorAll('.popup').forEach(popup => {
    if (popup.style.display === 'block' && !popup.contains(e.target)) {
      popup.style.display = 'none';
    }
  });
});

// Stop propagation on all interactive elements
document.querySelectorAll('.header-actions button, .header-actions button *, .popup, .popup *, .views-dropdown-button, .views-dropdown-button *, .views-dropdown-menu').forEach(el => {
  el.addEventListener('click', (e) => e.stopPropagation());
});

// Settings Popup
const settingsButton = document.getElementById('settingsButton');
const settingsPopup = document.getElementById('settingsPopup');
const closeSettings = document.getElementById('closeSettings');
const colorModeSelect = document.getElementById('colorModeSelect');

settingsButton.onclick = () => {
  const rect = settingsButton.getBoundingClientRect();
  settingsPopup.style.top = rect.bottom + 'px';
  settingsPopup.style.right = (window.innerWidth - rect.right) + 'px';
  settingsPopup.style.display = 'block';
  colorModeSelect.value = colorMode;
};

closeSettings.onclick = () => settingsPopup.style.display = 'none';

colorModeSelect.onchange = () => {
  colorMode = colorModeSelect.value;
  servicesPaletteIndex = 0;
  if (coloringActive) applyGroupColors();
};

// Hide Columns
const hideColumnsButton = document.getElementById('hideColumnsButton');
const hideColumnsPopup = document.getElementById('hideColumnsPopup');
const columnList = document.getElementById('columnList');

hideColumnsButton.onclick = () => {
  columnList.innerHTML = '';
  columnConfig.forEach(col => {
    if (col.key === 'expand' || col.key === 'reason') return;
    const div = document.createElement('div');
    div.className = 'column-item';
    div.innerHTML = `
      <span>${col.name}</span>
      <label class="toggle-switch">
        <input type="checkbox" ${col.visible ? 'checked' : ''} data-key="${col.key}">
        <span class="slider"></span>
      </label>
    `;
    columnList.appendChild(div);
  });
  const rect = hideColumnsButton.getBoundingClientRect();
  hideColumnsPopup.style.top = rect.bottom + 'px';
  hideColumnsPopup.style.right = (window.innerWidth - rect.right) + 'px';
  hideColumnsPopup.style.display = 'block';
};

columnList.addEventListener('change', (e) => {
  if (e.target.type === 'checkbox') {
    const key = e.target.dataset.key;
    const col = columnConfig.find(c => c.key === key);
    if (col) col.visible = e.target.checked;
    renderTable();
  }
});

// Filter
const filterButton = document.getElementById('filterButton');
const filterPopup = document.getElementById('filterPopup');
const filterRules = document.getElementById('filterRules');

filterButton.onclick = () => {
  renderFilterRules();
  const rect = filterButton.getBoundingClientRect();
  filterPopup.style.top = rect.bottom + 'px';
  filterPopup.style.right = (window.innerWidth - rect.right) + 'px';
  filterPopup.style.display = 'block';
};

function renderFilterRules() {
  filterRules.innerHTML = '';
  filters.forEach((f, i) => {
    const div = document.createElement('div');
    div.className = 'filter-rule';
    div.innerHTML = `
      <select class="filter-col">
        ${columnConfig.filter(c => c.key !== 'expand' && c.key !== 'reason').map(c => 
          `<option value="${c.key}" ${f.col===c.key?'selected':''}>${c.name}</option>`
        ).join('')}
      </select>
      <select class="filter-op">
        <option value="contains">contains</option>
        <option value="noneof" ${f.op==='noneof'?'selected':''}>is none of</option>
      </select>
      <input type="text" class="filter-val" placeholder="Enter value" value="${f.val||''}" style="display:${f.op==='noneof'?'none':''}">
      <div class="filter-tags" style="display:${f.op==='noneof'?'flex':'none'}"></div>
      <button class="remove-filter">×</button>
    `;
    if (f.op === 'noneof' && f.tags) {
      const tagsDiv = div.querySelector('.filter-tags');
      f.tags.forEach(tag => {
        const span = document.createElement('span');
        span.className = 'tag';
        span.innerHTML = `${tag} <span class="tag-remove">×</span>`;
        tagsDiv.appendChild(span);
      });
    }
    filterRules.appendChild(div);
  });
  if (filters.length === 0) filterRules.innerHTML = '<i>No filters</i>';
}

document.getElementById('addFilter').onclick = () => {
  filters.push({ col: 'mark', op: 'contains', val: '' });
  renderFilterRules();
  renderTable();
};

filterRules.addEventListener('click', (e) => {
  if (e.target.classList.contains('remove-filter')) {
    const index = Array.from(filterRules.children).indexOf(e.target.closest('.filter-rule'));
    filters.splice(index, 1);
    renderFilterRules();
    renderTable();
  }
  if (e.target.classList.contains('tag-remove')) {
    const tagSpan = e.target.parentElement;
    const tagsDiv = tagSpan.parentElement;
    const rule = tagsDiv.closest('.filter-rule');
    const index = Array.from(filterRules.children).indexOf(rule);
    const tagText = tagSpan.textContent.trim().slice(0, -1);
    filters[index].tags = filters[index].tags.filter(t => t !== tagText);
    tagSpan.remove();
    renderTable();
  }
});

filterRules.addEventListener('change', (e) => {
  const rule = e.target.closest('.filter-rule');
  const index = Array.from(filterRules.children).indexOf(rule);
  if (e.target.classList.contains('filter-col')) filters[index].col = e.target.value;
  if (e.target.classList.contains('filter-op')) {
    filters[index].op = e.target.value;
    rule.querySelector('.filter-val').style.display = e.target.value === 'noneof' ? 'none' : '';
    rule.querySelector('.filter-tags').style.display = e.target.value === 'noneof' ? 'flex' : 'none';
    if (e.target.value === 'noneof') filters[index].tags = [];
  }
  renderTable();
});

filterRules.addEventListener('input', (e) => {
  if (e.target.classList.contains('filter-val')) {
    const rule = e.target.closest('.filter-rule');
    const index = Array.from(filterRules.children).indexOf(rule);
    filters[index].val = e.target.value;
    renderTable();
  }
});

filterRules.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.classList.contains('filter-val')) {
    e.preventDefault();
    const rule = e.target.closest('.filter-rule');
    const index = Array.from(filterRules.children).indexOf(rule);
    const val = e.target.value.trim();
    if (val && filters[index].op === 'noneof') {
      filters[index].tags = filters[index].tags || [];
      if (!filters[index].tags.includes(val)) {
        filters[index].tags.push(val);
        const tagsDiv = rule.querySelector('.filter-tags');
        const span = document.createElement('span');
        span.className = 'tag';
        span.innerHTML = `${val} <span class="tag-remove">×</span>`;
        tagsDiv.appendChild(span);
      }
      e.target.value = '';
      renderTable();
    }
  }
});

// Group by
const groupButton = document.getElementById('groupButton');
const groupPopup = document.getElementById('groupPopup');
const groupRules = document.getElementById('groupRules');
const groupCount = document.getElementById('groupCount');
const groupPlural = document.getElementById('groupPlural');

groupButton.onclick = () => {
  renderGroupRules();
  groupCount.textContent = groupBy.length;
  groupPlural.textContent = groupBy.length === 1 ? '' : 's';
  const rect = groupButton.getBoundingClientRect();
  groupPopup.style.top = rect.bottom + 'px';
  groupPopup.style.right = (window.innerWidth - rect.right) + 'px';
  groupPopup.style.display = 'block';
};

function renderGroupRules() {
  groupRules.innerHTML = '';
  groupBy.forEach((col, i) => {
    const div = document.createElement('div');
    div.className = 'group-rule';
    div.innerHTML = `
      <select>
        ${columnConfig.filter(c => c.key !== 'expand' && c.key !== 'reason').map(c => 
          `<option value="${c.key}" ${col===c.key?'selected':''}>${c.name}</option>`
        ).join('')}
      </select>
      <button class="remove-group">×</button>
    `;
    groupRules.appendChild(div);
  });
  if (groupBy.length === 0) groupRules.innerHTML = '<i>No grouping</i>';
}

document.getElementById('addGroup').onclick = () => {
  groupBy.push('mark');
  renderGroupRules();
  groupCount.textContent = groupBy.length;
  groupPlural.textContent = groupBy.length === 1 ? '' : 's';
  renderTable();
};

groupRules.addEventListener('click', (e) => {
  if (e.target.classList.contains('remove-group')) {
    const index = Array.from(groupRules.children).indexOf(e.target.closest('.group-rule'));
    groupBy.splice(index, 1);
    renderGroupRules();
    groupCount.textContent = groupBy.length;
    groupPlural.textContent = groupBy.length === 1 ? '' : 's';
    renderTable();
  }
});

groupRules.addEventListener('change', (e) => {
  const index = Array.from(groupRules.children).indexOf(e.target.closest('.group-rule'));
  groupBy[index] = e.target.value;
  renderTable();
});

document.getElementById('collapseAll').onclick = () => { expandedGroups.clear(); renderTable(); };
document.getElementById('expandAll').onclick = () => { groups.forEach(g => expandedGroups.add(g.key)); renderTable(); };

async function init() {
  try {
    tcApi = await TrimbleConnectWorkspace.connect(window.parent);
    debug('Connected successfully');
  } catch (err) {
    debug('Connection error: ' + err.message);
  }
}

async function applyGroupColors() {
  if (!tcApi || groups.length === 0) return;

  groupColors = {};
  groupColorReasons = {};
  servicesPaletteIndex = 0;

  for (const group of groups) {
    let color, reason;

    if (colorMode === 'random') {
      color = {
        r: Math.floor(Math.random() * 256),
        g: Math.floor(Math.random() * 256),
        b: Math.floor(Math.random() * 256),
        a: 255
      };
      reason = 'Random';
    } else if (colorMode === 'deliveries') {
      const flat = flattenProps(group.sampleProps);
      const subassembly = flat['Subassembly -1'] || '';
      const diameter = flat['Size'] || 'unknown';

      const isLooseOrInsitu = /LOOSE|SITU/i.test(subassembly);

      if (isLooseOrInsitu) {
        const lightMap = {
          '8':  [128, 0, 255],
          '10': [88, 145, 163],
          '12': [206, 150, 45],
          '16': [0, 85, 255],
          '20': [255, 83, 255],
          '25': [85, 255, 85],
          '32': [255, 0, 0],
          '40': [0, 198, 227]
        };
        const rgb = lightMap[diameter] || [255, 255, 255];
        color = { r: rgb[0], g: rgb[1], b: rgb[2], a: 255 };
        reason = `Rule-based (Deliveries) - Loose/In-Situ (Dia ${diameter})`;
      } else {
        const darkMap = {
          '8':  [0, 0, 127],
          '10': [0, 17, 35],
          '12': [78, 22, 0],
          '16': [0, 0, 127],
          '20': [127, 0, 127],
          '25': [0, 127, 0],
          '32': [127, 0, 0],
          '40': [0, 70, 99]
        };
        const rgb = darkMap[diameter] || [200, 200, 200];
        color = { r: rgb[0], g: rgb[1], b: rgb[2], a: 255 };
        reason = `Rule-based (Deliveries) - Prefab/Assembly (Dia ${diameter})`;
      }
    } else if (colorMode === 'services') {
      const rgb = servicesPalette[servicesPaletteIndex % servicesPalette.length];
      color = { r: rgb[0], g: rgb[1], b: rgb[2], a: 255 };
      reason = 'Rule-based (Services) - Palette color';
      servicesPaletteIndex++;
    }

    groupColors[group.key] = color;
    groupColorReasons[group.key] = reason;
    group.color = color;

    try {
      await tcApi.viewer.setObjectState({
        modelId: group.modelId,
        objectIds: group.objectIds
      }, { color });
    } catch (e) {
      debug('Color apply error: ' + e.message);
    }
  }
  renderTable();
}

async function resetAllColors() {
  if (!tcApi) return;
  try {
    await tcApi.viewer.setObjectState(undefined, { color: "reset" });
    groupColors = {};
    groupColorReasons = {};
    groups.forEach(g => g.color = null);
    renderTable();
  } catch (e) {
    debug('Reset colors error: ' + e.message);
  }
}

async function toggleColorGroups() {
  coloringActive = !coloringActive;
  const button = document.getElementById('colorGroups');
  button.classList.toggle('active', coloringActive);

  if (coloringActive) {
    servicesPaletteIndex = 0;
    await applyGroupColors();
  } else {
    await resetAllColors();
  }
}

async function selectAndZoomGroup(group) {
  const modelObjectIds = [{ modelId: group.modelId, objectRuntimeIds: group.objectIds }];
  try {
    await tcApi.viewer.setSelection({ modelObjectIds }, "set");
    await tcApi.viewer.zoomTo({ modelObjectIds });
  } catch (e) {}
}

function appliesFilters(group) {
  return filters.every(f => {
    const val = group[f.col] || '';
    const str = val.toString().toLowerCase();
    if (f.op === 'contains') return str.includes((f.val||'').toLowerCase());
    if (f.op === 'noneof') return !f.tags.some(tag => str === tag.toLowerCase());
    return true;
  });
}

function renderTable() {
  const tbody = document.querySelector('#rebarTable tbody');
  tbody.innerHTML = '';

  let filteredGroups = groups.filter(appliesFilters);

  let rows = [];
  if (groupBy.length === 0) {
    rows = filteredGroups;
  } else {
    const grouped = {};
    filteredGroups.forEach(g => {
      const key = groupBy.map(col => g[col] ?? '').join(' │ ');
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(g);
    });
    Object.keys(grouped).sort().forEach(key => {
      const groupRow = document.createElement('tr');
      groupRow.className = 'group-header';
      groupRow.innerHTML = `<td colspan="14"><strong>Group: ${key} (${grouped[key].length} items)</strong></td>`;
      tbody.appendChild(groupRow);
      grouped[key].forEach(g => rows.push(g));
    });
  }

  let totalBars = 0;
  rows.forEach(group => {
    totalBars += group.quantity;

    const color = group.color || { r: 255, g: 255, b: 255 };
    const reason = groupColorReasons[group.key] || (coloringActive ? colorMode.charAt(0).toUpperCase() + colorMode.slice(1) : '--');
    const isSelected = selectedGroups.has(group);

    const groupTr = document.createElement('tr');
    groupTr.className = 'group-row' + (isSelected ? ' selected' : '');
    groupTr.dataset.groupKey = group.key;

    let html = `<td class="expand-cell" style="background-color: rgb(${color.r},${color.g},${color.b});">
      <img src="${expandedGroups.has(group.key) ? 'chevron-down.svg' : 'chevron-right.svg'}" class="expand-icon">
    </td>`;
    columnConfig.forEach(col => {
      if (!col.visible) return;
      if (col.key === 'expand') return;
      if (col.key === 'reason') {
        html += `<td class="color-reason">${reason}</td>`;
        return;
      }
      let value = group[col.key] || '--';
      if (col.key === 'quantity') value = group.quantity;
      const cls = ['mark', 'diameter', 'quantity'].includes(col.key) ? 'center' : 'numeric';
      html += `<td class="${cls}">${value}</td>`;
    });
    groupTr.innerHTML = html;
    groupTr.addEventListener('click', (e) => handleRowClick(e, group, groupTr));
    tbody.appendChild(groupTr);

    if (expandedGroups.has(group.key)) {
      group.individualBars.forEach((bar, idx) => {
        const childTr = document.createElement('tr');
        childTr.className = 'child-row' + (isSelected ? ' selected' : '');
        childTr.dataset.groupKey = group.key;
        let childHtml = `<td class="center index-cell">${idx + 1}</td>`;
        columnConfig.forEach(col => {
          if (!col.visible) return;
          if (col.key === 'expand') return;
          if (col.key === 'reason') {
            childHtml += `<td class="color-reason gray-bg">${reason}</td>`;
            return;
          }
          let value = group[col.key] || '--';
          if (['h1','v1','z1','h2','v2','z2'].includes(col.key)) value = bar[col.key] || '--';
          if (col.key === 'quantity') value = group.quantity;
          const cls = ['mark', 'diameter', 'quantity'].includes(col.key) ? 'center gray-bg' : 'numeric';
          childHtml += `<td class="${cls}">${value}</td>`;
        });
        childTr.innerHTML = childHtml;
        childTr.addEventListener('click', (e) => handleRowClick(e, group, childTr));
        tbody.appendChild(childTr);
      });
    }
  });

  if (rows.length > 0) {
    const totalsTr = document.createElement('tr');
    totalsTr.className = 'totals-row';
    let totalsHtml = `<td colspan="2"><strong>Total</strong></td>`;
    columnConfig.slice(2, -1).forEach(col => {
      totalsHtml += col.visible ? '<td></td>' : '';
    });
    totalsHtml += `<td class="center"><strong>${totalBars}</strong></td><td></td>`;
    totalsTr.innerHTML = totalsHtml;
    tbody.appendChild(totalsTr);
  }
}

function handleRowClick(e, group, rowElement) {
  if (e.target.classList.contains('expand-icon') || e.target.closest('.expand-icon')) {
    e.stopPropagation();
    expandedGroups.has(group.key) ? expandedGroups.delete(group.key) : expandedGroups.add(group.key);
    renderTable();
    return;
  }

  const groupRows = Array.from(document.querySelectorAll('tr.group-row'));

  if (e.shiftKey && lastSelectedRow) {
    const currentIndex = groupRows.indexOf(rowElement);
    const lastIndex = groupRows.indexOf(lastSelectedRow);
    const start = Math.min(currentIndex, lastIndex);
    const end = Math.max(currentIndex, lastIndex);

    document.querySelectorAll('#rebarTable tbody tr').forEach(tr => tr.classList.remove('selected'));

    for (let i = start; i <= end; i++) {
      const tr = groupRows[i];
      tr.classList.add('selected');
      const gKey = tr.dataset.groupKey;
      document.querySelectorAll(`tr.child-row[data-group-key="${gKey}"]`).forEach(c => c.classList.add('selected'));
    }
  } else if (e.ctrlKey || e.metaKey) {
    rowElement.classList.toggle('selected');
    const gKey = rowElement.dataset.groupKey;
    const isNowSelected = rowElement.classList.contains('selected');
    document.querySelectorAll(`tr.child-row[data-group-key="${gKey}"]`).forEach(c => {
      if (isNowSelected) c.classList.add('selected');
      else c.classList.remove('selected');
    });
  } else {
    document.querySelectorAll('#rebarTable tbody tr').forEach(tr => tr.classList.remove('selected'));
    rowElement.classList.add('selected');
    const gKey = rowElement.dataset.groupKey;
    document.querySelectorAll(`tr.child-row[data-group-key="${gKey}"]`).forEach(c => c.classList.add('selected'));
  }

  if (rowElement.classList.contains('group-row') && rowElement.classList.contains('selected')) {
    lastSelectedRow = rowElement;
  }

  selectedGroups.clear();
  document.querySelectorAll('tr.group-row.selected').forEach(tr => {
    const gKey = tr.dataset.groupKey;
    const grp = groups.find(g => g.key === gKey);
    if (grp) selectedGroups.add(grp);
  });

  if (selectedGroups.size === 1) selectAndZoomGroup([...selectedGroups][0]);
}

function flattenProps(props) {
  const flat = {};
  if (!props?.properties) return flat;
  props.properties.forEach(pset => {
    if (pset.properties) pset.properties.forEach(prop => {
      if (prop.value !== undefined) flat[prop.name] = prop.value;
    });
  });
  return flat;
}

async function processRebar() {
  debug('=== REFRESH STARTED ===');
  groups = [];
  selectedGroups.clear();
  lastSelectedRow = null;
  expandedGroups = new Set();

  const tbody = document.querySelector('#rebarTable tbody');
  tbody.innerHTML = '<tr><td colspan="14">Processing...</td></tr>';

  try {
    let objects = [];

    if (currentViewMode === 'selected') {
      const selection = await tcApi.viewer.getSelection();
      if (selection.length === 0) throw new Error('No objects selected.');
      for (const sel of selection) {
        const modelId = sel.modelId;
        const runtimeIds = sel.objectRuntimeIds || [];
        if (!modelId || runtimeIds.length === 0) continue;
        const propsArray = await tcApi.viewer.getObjectProperties(modelId, runtimeIds);
        propsArray.forEach((props, i) => {
          if (hasRebarProps(props)) objects.push({ id: runtimeIds[i], props, modelId });
        });
      }
    } else if (currentViewMode === 'visible') {
      const result = await tcApi.viewer.getObjects({ objectState: { visibility: 'visible' } });
      for (const item of result) {
        const modelId = item.modelId;
        const runtimeIds = item.objectRuntimeIds || [];
        if (runtimeIds.length === 0) continue;
        const propsArray = await tcApi.viewer.getObjectProperties(modelId, runtimeIds);
        propsArray.forEach((props, i) => {
          if (hasRebarProps(props)) objects.push({ id: runtimeIds[i], props, modelId });
        });
      }
    }

    if (objects.length === 0) throw new Error('No rebar elements found.');

    const groupMap = {};
    for (const item of objects) {
      const flat = flattenProps(item.props);
      const mark = flat['Serial number'] || flat['Rebar mark'] || flat['Group position number'] || 'Unknown';
      const dia = flat['Size'] || '';
      const shape = flat['Shape'] || flat['Shape code'] || '';
      const length = Math.round(Number(flat['Length'] || 0));
      const key = `${mark}|${dia}|${length}|${shape}|${item.modelId}`;

      if (!groupMap[key]) {
        groupMap[key] = {
          key, mark, diameter: dia, shape, length,
          subassembly: flat['Subassembly -1'] || '',
          objectIds: [], sampleProps: item.props, modelId: item.modelId,
          individualBars: []
        };
      }
      groupMap[key].objectIds.push(item.id);
      groupMap[key].individualBars.push({
        h1: flat['Start Point X'] || '',
        v1: flat['Start Point Y'] || '',
        z1: flat['Start Point Z'] || '',
        h2: flat['End Point X'] || '',
        v2: flat['End Point Y'] || '',
        z2: flat['End Point Z'] || ''
      });
    }

    groups = Object.values(groupMap);
    groups.forEach(g => {
      g.quantity = g.objectIds.length;

      const startsH = new Set(g.individualBars.map(b => b.h1));
      g.h1 = startsH.size === 1 ? g.individualBars[0].h1 : '';
      const startsV = new Set(g.individualBars.map(b => b.v1));
      g.v1 = startsV.size === 1 ? g.individualBars[0].v1 : '';
      const startsZ = new Set(g.individualBars.map(b => b.z1));
      g.z1 = startsZ.size === 1 ? g.individualBars[0].z1 : '';

      const endsH = new Set(g.individualBars.map(b => b.h2));
      g.h2 = endsH.size === 1 ? g.individualBars[0].h2 : '';
      const endsV = new Set(g.individualBars.map(b => b.v2));
      g.v2 = endsV.size === 1 ? g.individualBars[0].v2 : '';
      const endsZ = new Set(g.individualBars.map(b => b.z2));
      g.z2 = endsZ.size === 1 ? g.individualBars[0].z2 : '';
    });

    sortColumns = [
      { col: 'diameter', dir: 'desc' },
      { col: 'subassembly', dir: 'asc' }
    ];
    sortTable();

    document.getElementById('status').textContent = `Success! Found ${groups.length} groups from ${objects.length} bars.`;
    document.getElementById('status').classList.add('good');

    if (coloringActive) await applyGroupColors();
  } catch (err) {
    debug('ERROR: ' + err.message);
    tbody.innerHTML = `<tr><td colspan="14">Error: ${err.message}</td></tr>`;
    document.getElementById('status').textContent = 'Failed: ' + err.message;
    document.getElementById('status').classList.add('error');
  }
}

function hasRebarProps(props) {
  return props.properties?.some(p => p.name === 'Tekla Reinforcement - Bending List' || p.name === 'ICOS Rebar' || p.name === 'ICOS BBS');
}

// Sort popup
const sortButton = document.getElementById('sortButton');
const sortPopup = document.getElementById('sortPopup');
const sortLevels = document.getElementById('sortLevels');
const addSortBtn = document.getElementById('addSort');
const sortCount = document.querySelector('.sort-count');

const columns = ['mark', 'subassembly', 'diameter', 'shape', 'h1', 'v1', 'z1', 'h2', 'v2', 'z2', 'length', 'quantity'];
const columnNames = { 
  mark: 'Mark', subassembly: 'Subassembly -1', diameter: 'Diam.', shape: 'Shape code',
  h1: 'H1', v1: 'V1', z1: 'Z1', h2: 'H2', v2: 'V2', z2: 'Z2',
  length: 'Length', quantity: 'Count'
};

function createSortLevel(levelIndex, col = 'diameter', dir = 'desc') {
  const level = document.createElement('div');
  level.className = 'sort-level';
  level.innerHTML = `
    <span class="grip">⋮⋮</span>
    <button class="remove-sort">×</button>
    <select class="sort-column">
      ${columns.map(c => `<option value="${c}" ${c===col?'selected':''}>${columnNames[c]}</option>`).join('')}
    </select>
    <select class="sort-direction">
      <option value="asc" ${dir==='asc'?'selected':''}>Up</option>
      <option value="desc" ${dir==='desc'?'selected':''}>Down</option>
    </select>
  `;
  level.querySelector('.remove-sort').onclick = (e) => { e.stopPropagation(); level.remove(); updateSortColumns(); };
  level.querySelectorAll('select').forEach(sel => sel.onchange = (e) => { e.stopPropagation(); updateSortColumns(); });
  return level;
}

function updateSortColumns() {
  sortColumns = [];
  document.querySelectorAll('.sort-level').forEach(level => {
    sortColumns.push({
      col: level.querySelector('.sort-column').value,
      dir: level.querySelector('.sort-direction').value
    });
  });
  sortCount.textContent = `${sortColumns.length} Sort${sortColumns.length !== 1 ? 's' : ''}`;
  sortTable();
}

addSortBtn.onclick = (e) => {
  e.stopPropagation();
  sortLevels.appendChild(createSortLevel(sortLevels.children.length));
  updateSortColumns();
};

new Sortable(sortLevels, { handle: '.grip', animation: 150, onEnd: updateSortColumns });

sortButton.onclick = () => {
  sortLevels.innerHTML = '';
  sortColumns.forEach((sc, i) => sortLevels.appendChild(createSortLevel(i, sc.col, sc.dir)));
  if (sortColumns.length === 0) sortLevels.appendChild(createSortLevel(0, 'diameter', 'desc'));
  updateSortColumns();
  const rect = sortButton.getBoundingClientRect();
  sortPopup.style.top = rect.bottom + 'px';
  sortPopup.style.right = (window.innerWidth - rect.right) + 'px';
  sortPopup.style.display = 'block';
};

document.getElementById('refresh').onclick = processRebar;
document.getElementById('colorGroups').onclick = toggleColorGroups;

init();