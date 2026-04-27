// ── CONFIG ────────────────────────────────────────────────────────────────────

const NODE_FIELDS = [
  { key: 'voltageLevel', label: 'Voltage Level', type: 'text',   placeholder: 'e.g. 415V' },
  { key: 'load',         label: 'Load (kW)',      type: 'number', placeholder: 'e.g. 45'   },
  { key: 'circuitId',    label: 'Circuit ID',     type: 'text',   placeholder: 'e.g. C-01' },
  { key: 'notes',        label: 'Notes',          type: 'text',   placeholder: ''           },
];

const NODE_STYLES = {
  Source: { fill: '#2e7d32', shape: 'rect',   width: 80, height: 40 },
  DB:     { fill: '#1565c0', shape: 'rect',   width: 80, height: 40 },
  DC:     { fill: '#0277bd', shape: 'rect',   width: 80, height: 40 },
  TX:     { fill: '#555555', shape: 'circle', r: 24 },
  PB:     { fill: '#e65100', shape: 'rect',   width: 80, height: 40 },
};

// ── STATE ─────────────────────────────────────────────────────────────────────

let nodes               = [];
let connections         = [];
let selectedId          = null;
let selectedConnection  = null;
let dragState           = null;
let isPanning           = false;
let panX                = 0;
let panY                = 0;
let zoomLevel           = 1.0;
let currentTool         = 'select';
let nodeCounters        = { Source: 0, DB: 0, DC: 0, TX: 0, PB: 0 };
let globalIdCounter     = 0;
let propertiesCollapsed = false;
let connectDragState    = null;

let _panStartX = 0;
let _panStartY = 0;

// ── OFFICE INIT ───────────────────────────────────────────────────────────────

Office.onReady(function (info) {
  console.log('Office.onReady fired. Host:', info.host, 'Platform:', info.platform);

  // Show app regardless of host check — works for all Excel versions
  var sideload = document.getElementById('sideload-msg');
  var appBody  = document.getElementById('app-body');

  if (sideload) sideload.style.display = 'none';
  if (appBody)  appBody.style.display  = 'flex';

  try {
    initCanvas();
    buildPropertiesPanel();
    console.log('App initialized successfully');
  } catch(err) {
    console.error('Init error:', err);
    if (appBody) appBody.innerHTML = '<div style="padding:20px;color:red;font-size:13px;">Error: ' + err.message + '</div>';
  }
});

// ── VIEWPORT TRANSFORM ────────────────────────────────────────────────────────

function applyViewportTransform() {
  const vp = document.getElementById('viewport');
  if (vp) vp.setAttribute('transform', `translate(${panX}, ${panY}) scale(${zoomLevel})`);
  const label = document.getElementById('zoom-label');
  if (label) label.textContent = Math.round(zoomLevel * 100) + '%';
}

function getCanvasPoint(e) {
  const svg  = document.getElementById('canvas');
  const rect = svg.getBoundingClientRect();
  return {
    x: (e.clientX - rect.left - panX) / zoomLevel,
    y: (e.clientY - rect.top  - panY) / zoomLevel,
  };
}

// ── CANVAS INIT ───────────────────────────────────────────────────────────────

function initCanvas() {
  const svg  = document.getElementById('canvas');
  const defs = svg.querySelector('defs');

  const marker = document.createElementNS('http://www.w3.org/2000/svg', 'marker');
  marker.setAttribute('id',           'arrowhead');
  marker.setAttribute('markerWidth',  '8');
  marker.setAttribute('markerHeight', '8');
  marker.setAttribute('refX',         '8');
  marker.setAttribute('refY',         '4');
  marker.setAttribute('orient',       'auto');
  const poly = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
  poly.setAttribute('points', '0 0, 8 4, 0 8');
  poly.setAttribute('fill',   '#455a64');
  marker.appendChild(poly);
  defs.appendChild(marker);

  svg.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    const isBackground = e.target === svg
      || ['viewport', 'connections', 'nodes', 'temp-connections'].includes(e.target.id);
    if (!isBackground) return;
    e.preventDefault();
    isPanning  = true;
    _panStartX = e.clientX - panX;
    _panStartY = e.clientY - panY;
    svg.style.cursor = 'grabbing';
  });

  document.addEventListener('mousemove', (e) => {
    if (!isPanning) return;
    panX = e.clientX - _panStartX;
    panY = e.clientY - _panStartY;
    applyViewportTransform();
    if (selectedConnection) repositionConnDeletePopup();
  });

  document.addEventListener('mouseup', () => {
    if (!isPanning) return;
    isPanning = false;
    const c = document.getElementById('canvas');
    if (c) c.style.cursor = '';
  });

  svg.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    const delta = e.deltaY < 0 ? 0.25 : -0.25;
    zoomLevel = clampZoom(zoomLevel + delta);
    applyViewportTransform();
    if (selectedConnection) repositionConnDeletePopup();
  }, { passive: false });

  const handle          = document.getElementById('resize-handle');
  const canvasContainer = document.getElementById('canvas-container');
  canvasContainer.style.height = '300px';

  handle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    const startY      = e.clientY;
    const startHeight = canvasContainer.offsetHeight;
    const onMove = (mv) => {
      canvasContainer.style.height = Math.max(150, startHeight + mv.clientY - startY) + 'px';
      if (selectedConnection) repositionConnDeletePopup();
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup',   onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',   onUp);
  });
}

function clampZoom(z) {
  return Math.min(3.0, Math.max(0.25, Math.round(z * 100) / 100));
}

// ── TOOL SWITCH ───────────────────────────────────────────────────────────────

function switchTool(tool) {
  currentTool = tool;
  document.body.classList.toggle('tool-connect', tool === 'connect');
  document.getElementById('btn-select') .classList.toggle('active', tool === 'select');
  document.getElementById('btn-connect').classList.toggle('active', tool === 'connect');
}

// ── ADD NODE ──────────────────────────────────────────────────────────────────

function addNode(type) {
  const svg  = document.getElementById('canvas');
  const bbox = svg.getBoundingClientRect();

  nodeCounters[type]++;
  const id   = `node-${++globalIdCounter}`;
  const name = `${type}${nodeCounters[type]}`;

  const fields = {};
  NODE_FIELDS.forEach(f => { fields[f.key] = ''; });

  const node = {
    id,
    name,
    type,
    x: (bbox.width  / 2 - panX) / zoomLevel,
    y: (bbox.height / 2 - panY) / zoomLevel,
    fields,
  };

  nodes.push(node);
  renderNode(node);
  selectNode(id);
}

// ── RENDER NODE ───────────────────────────────────────────────────────────────

function renderNode(node) {
  const nodesG = document.getElementById('nodes');
  const style  = NODE_STYLES[node.type];
  const ns     = 'http://www.w3.org/2000/svg';
  const g      = document.createElementNS(ns, 'g');

  g.setAttribute('id',        node.id);
  g.setAttribute('class',     'node-group');
  g.setAttribute('transform', `translate(${node.x}, ${node.y})`);

  if (style.shape === 'circle') {
    const circle = document.createElementNS(ns, 'circle');
    circle.setAttribute('class',        'node-circle');
    circle.setAttribute('r',            style.r);
    circle.setAttribute('fill',         style.fill);
    circle.setAttribute('stroke',       style.fill);
    circle.setAttribute('stroke-width', '2');
    g.appendChild(circle);

    const ch = document.createElementNS(ns, 'circle');
    ch.setAttribute('class', 'connect-handle');
    ch.setAttribute('cx',    '0');
    ch.setAttribute('cy',    style.r);
    ch.setAttribute('r',     '6');
    g.appendChild(ch);
  } else {
    const hw   = style.width  / 2;
    const hh   = style.height / 2;
    const rect = document.createElementNS(ns, 'rect');
    rect.setAttribute('class',        'node-rect');
    rect.setAttribute('x',            -hw);
    rect.setAttribute('y',            -hh);
    rect.setAttribute('width',        style.width);
    rect.setAttribute('height',       style.height);
    rect.setAttribute('rx',           '6');
    rect.setAttribute('fill',         style.fill);
    rect.setAttribute('stroke',       style.fill);
    rect.setAttribute('stroke-width', '2');
    g.appendChild(rect);

    const ch = document.createElementNS(ns, 'circle');
    ch.setAttribute('class', 'connect-handle');
    ch.setAttribute('cx',    '0');
    ch.setAttribute('cy',    hh);
    ch.setAttribute('r',     '6');
    g.appendChild(ch);
  }

  const text = document.createElementNS(ns, 'text');
  text.setAttribute('class',             'node-label');
  text.setAttribute('x',                 '0');
  text.setAttribute('y',                 '0');
  text.setAttribute('dominant-baseline', 'central');
  text.setAttribute('text-anchor',       'middle');
  text.textContent = node.name;
  g.appendChild(text);

  g.addEventListener('click', (e) => {
    e.stopPropagation();
    dismissSelectedConnection();
    if (currentTool === 'select') selectNode(node.id);
  });

  g.addEventListener('mousedown', (e) => {
    e.stopPropagation();
    if (currentTool === 'select' && e.button === 0) startDrag(e, node);
  });

  const handle = g.querySelector('.connect-handle');
  if (handle) {
    handle.addEventListener('mousedown', (e) => {
      e.stopPropagation();
      e.preventDefault();
      if (currentTool === 'connect') startConnectDrag(e, node);
    });
  }

  nodesG.appendChild(g);
}

// ── DRAG NODE ─────────────────────────────────────────────────────────────────

function startDrag(e, node) {
  e.preventDefault();
  const origin = getCanvasPoint(e);
  dragState = { node, offsetX: origin.x - node.x, offsetY: origin.y - node.y };

  const onMove = (ev) => {
    if (!dragState) return;
    const p  = getCanvasPoint(ev);
    node.x   = p.x - dragState.offsetX;
    node.y   = p.y - dragState.offsetY;
    const el = document.getElementById(node.id);
    if (el) el.setAttribute('transform', `translate(${node.x}, ${node.y})`);
    redrawConnections();
    if (selectedConnection) repositionConnDeletePopup();
  };

  const onUp = () => {
    dragState = null;
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup',   onUp);
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup',   onUp);
}

// ── CONNECT DRAG ──────────────────────────────────────────────────────────────

function startConnectDrag(e, fromNode) {
  const tempG    = document.getElementById('temp-connections');
  const ns       = 'http://www.w3.org/2000/svg';
  const tempLine = document.createElementNS(ns, 'line');
  tempLine.setAttribute('class', 'connector-temp');
  tempLine.setAttribute('x1',    fromNode.x);
  tempLine.setAttribute('y1',    fromNode.y);
  tempLine.setAttribute('x2',    fromNode.x);
  tempLine.setAttribute('y2',    fromNode.y);
  tempG.appendChild(tempLine);
  connectDragState = tempLine;

  const onMove = (ev) => {
    const p = getCanvasPoint(ev);
    tempLine.setAttribute('x2', p.x);
    tempLine.setAttribute('y2', p.y);
    nodes.forEach(n => {
      if (n.id === fromNode.id) return;
      const el    = document.getElementById(n.id);
      if (!el) return;
      const s  = NODE_STYLES[n.type];
      const dx = p.x - n.x;
      const dy = p.y - n.y;
      const over = s.shape === 'circle'
        ? Math.hypot(dx, dy) < s.r + 8
        : Math.abs(dx) < s.width / 2 + 8 && Math.abs(dy) < s.height / 2 + 8;
      el.classList.toggle('node-connect-target', over);
    });
  };

  const onUp = (ev) => {
    if (tempLine.parentNode) tempLine.remove();
    connectDragState = null;
    nodes.forEach(n => document.getElementById(n.id)?.classList.remove('node-connect-target'));

    const p = getCanvasPoint(ev);
    let target = null;
    nodes.forEach(n => {
      if (n.id === fromNode.id) return;
      const s  = NODE_STYLES[n.type];
      const dx = p.x - n.x;
      const dy = p.y - n.y;
      const over = s.shape === 'circle'
        ? Math.hypot(dx, dy) < s.r + 10
        : Math.abs(dx) < s.width / 2 + 10 && Math.abs(dy) < s.height / 2 + 10;
      if (over) target = n;
    });

    if (target && !connections.find(c => c.fromId === fromNode.id && c.toId === target.id)) {
      connections.push({ fromId: fromNode.id, toId: target.id });
      redrawConnections();
      updateConnectionCount();
      showToast('Connected');
    }

    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup',   onUp);
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup',   onUp);
}

// ── NODE EDGE POINT ───────────────────────────────────────────────────────────

function nodeEdgePoint(node, fromX, fromY) {
  const style = NODE_STYLES[node.type];
  const dx    = fromX - node.x;
  const dy    = fromY - node.y;
  const dist  = Math.hypot(dx, dy);
  if (dist < 1) return { x: node.x, y: node.y };
  const ndx = dx / dist;
  const ndy = dy / dist;

  if (style.shape === 'circle') {
    return { x: node.x + ndx * style.r, y: node.y + ndy * style.r };
  }

  const hw  = style.width  / 2;
  const hh  = style.height / 2;
  const adx = Math.abs(ndx);
  const ady = Math.abs(ndy);
  const t   = adx < 1e-9 ? hh / ady
            : ady < 1e-9 ? hw / adx
            : Math.min(hw / adx, hh / ady);
  return { x: node.x + ndx * t, y: node.y + ndy * t };
}

// ── REDRAW CONNECTIONS ────────────────────────────────────────────────────────

function redrawConnections() {
  const g  = document.getElementById('connections');
  g.innerHTML = '';
  const ns = 'http://www.w3.org/2000/svg';

  connections.forEach(({ fromId, toId }) => {
    const from = nodes.find(n => n.id === fromId);
    const to   = nodes.find(n => n.id === toId);
    if (!from || !to) return;

    const isSelected = selectedConnection
      && selectedConnection.fromId === fromId
      && selectedConnection.toId   === toId;

    const ep   = nodeEdgePoint(to, from.x, from.y);
    const line = document.createElementNS(ns, 'line');
    line.setAttribute('x1', from.x);
    line.setAttribute('y1', from.y);
    line.setAttribute('x2', ep.x);
    line.setAttribute('y2', ep.y);
    line.setAttribute('class', isSelected ? 'connector connector-selected' : 'connector');
    line.setAttribute('marker-end', 'url(#arrowhead)');
    g.appendChild(line);

    const hit = document.createElementNS(ns, 'line');
    hit.setAttribute('x1',           from.x);
    hit.setAttribute('y1',           from.y);
    hit.setAttribute('x2',           ep.x);
    hit.setAttribute('y2',           ep.y);
    hit.setAttribute('class',        'connector-hit');
    hit.setAttribute('data-from',    fromId);
    hit.setAttribute('data-to',      toId);
    hit.setAttribute('stroke-width', '12');
    hit.setAttribute('stroke',       'transparent');
    hit.setAttribute('opacity',      '0');
    hit.addEventListener('click', (e) => {
      e.stopPropagation();
      selectedConnection = { fromId, toId };
      redrawConnections();
      repositionConnDeletePopup();
    });
    g.appendChild(hit);
  });
}

// ── CONNECTION POPUP ──────────────────────────────────────────────────────────

function repositionConnDeletePopup() {
  const from = nodes.find(n => n.id === selectedConnection?.fromId);
  const to   = nodes.find(n => n.id === selectedConnection?.toId);
  if (!from || !to) { hideConnDeletePopup(); return; }

  const mx      = (from.x + to.x) / 2;
  const my      = (from.y + to.y) / 2;
  const screenX = mx * zoomLevel + panX;
  const screenY = my * zoomLevel + panY;

  const popup = document.getElementById('conn-delete-popup');
  popup.style.left    = screenX + 'px';
  popup.style.top     = screenY + 'px';
  popup.style.display = 'flex';
}

function hideConnDeletePopup() {
  const popup = document.getElementById('conn-delete-popup');
  if (popup) popup.style.display = 'none';
}

function dismissSelectedConnection() {
  if (!selectedConnection) return;
  selectedConnection = null;
  redrawConnections();
  hideConnDeletePopup();
}

function deleteSelectedConnection() {
  if (!selectedConnection) return;
  const { fromId, toId } = selectedConnection;
  connections = connections.filter(c => !(c.fromId === fromId && c.toId === toId));
  selectedConnection = null;
  redrawConnections();
  updateConnectionCount();
  hideConnDeletePopup();
}

// ── CONNECTION COUNT ──────────────────────────────────────────────────────────

function updateConnectionCount() {
  const el = document.getElementById('conn-counter');
  if (el) el.textContent = `Connections: ${connections.length}`;
}

// ── SELECT NODE ───────────────────────────────────────────────────────────────

function selectNode(id) {
  if (selectedId) document.getElementById(selectedId)?.classList.remove('node-selected');
  selectedId = id;
  if (id) document.getElementById(id)?.classList.add('node-selected');
  updatePropertiesPanel(id);
}

// ── PROPERTIES PANEL ──────────────────────────────────────────────────────────

function buildPropertiesPanel() {
  const panel = document.getElementById('properties-panel');
  panel.innerHTML = `
    <div class="panel-header">
      <span class="panel-title-text">Node Properties</span>
      <button class="collapse-btn" onclick="togglePropertiesCollapse()">&#9650;</button>
    </div>
    <div id="props-body">
      <div class="prop-row">
        <label>Name</label>
        <input id="prop-name" type="text" onchange="updateNodeProp('name', this.value)" />
      </div>
      ${NODE_FIELDS.map(f => `
      <div class="prop-row">
        <label>${f.label}</label>
        <input id="prop-${f.key}" type="${f.type}" placeholder="${f.placeholder}"
               onchange="updateNodeProp('${f.key}', this.value)" />
      </div>`).join('')}
      <button id="delete-node-btn" onclick="deleteNode()">Delete Node</button>
    </div>`;
}

function updatePropertiesPanel(id) {
  const panel = document.getElementById('properties-panel');
  if (!panel) return;
  if (!id) { panel.classList.remove('visible'); return; }
  const node = nodes.find(n => n.id === id);
  if (!node) return;
  panel.classList.add('visible');
  const nameInput = document.getElementById('prop-name');
  if (nameInput) nameInput.value = node.name;
  NODE_FIELDS.forEach(f => {
    const input = document.getElementById(`prop-${f.key}`);
    if (input) input.value = node.fields[f.key] || '';
  });
  const propsBody = document.getElementById('props-body');
  if (propsBody) propsBody.style.display = propertiesCollapsed ? 'none' : 'flex';
}

function togglePropertiesCollapse() {
  propertiesCollapsed = !propertiesCollapsed;
  const body = document.getElementById('props-body');
  const btn  = document.querySelector('.collapse-btn');
  if (body) body.style.display = propertiesCollapsed ? 'none' : 'flex';
  if (btn)  btn.innerHTML      = propertiesCollapsed ? '&#9660;' : '&#9650;';
}

function updateNodeProp(key, value) {
  const node = nodes.find(n => n.id === selectedId);
  if (!node) return;
  if (key === 'name') {
    node.name = value;
    const label = document.querySelector(`#${node.id} .node-label`);
    if (label) label.textContent = value;
  } else {
    node.fields[key] = value;
  }
}

// ── DELETE NODE ───────────────────────────────────────────────────────────────

function deleteNode() {
  if (!selectedId) return;
  if (selectedConnection &&
      (selectedConnection.fromId === selectedId || selectedConnection.toId === selectedId)) {
    dismissSelectedConnection();
  }
  nodes       = nodes.filter(n => n.id !== selectedId);
  connections = connections.filter(c => c.fromId !== selectedId && c.toId !== selectedId);
  document.getElementById(selectedId)?.remove();
  redrawConnections();
  updateConnectionCount();
  selectNode(null);
}

// ── CLEAR CANVAS ──────────────────────────────────────────────────────────────

function clearCanvas() {
  dragState = null;
  isPanning = false;

  if (connectDragState) {
    try { if (connectDragState.parentNode) connectDragState.remove(); } catch (_) {}
    connectDragState = null;
  }

  nodes              = [];
  connections        = [];
  selectedId         = null;
  selectedConnection = null;
  nodeCounters       = { Source: 0, DB: 0, DC: 0, TX: 0, PB: 0 };
  globalIdCounter    = 0;

  const nodesGroup = document.getElementById('nodes');
  const connsGroup = document.getElementById('connections');
  const tempGroup  = document.getElementById('temp-connections');
  if (nodesGroup) nodesGroup.innerHTML = '';
  if (connsGroup) connsGroup.innerHTML = '';
  if (tempGroup)  tempGroup.innerHTML  = '';

  hideConnDeletePopup();
  buildPropertiesPanel();
  updatePropertiesPanel(null);

  panX = 0; panY = 0; zoomLevel = 1.0;
  applyViewportTransform();

  const svg = document.getElementById('canvas');
  if (svg) svg.style.cursor = '';

  updateConnectionCount();
  showToast('Canvas cleared');
}

// ── ZOOM ──────────────────────────────────────────────────────────────────────

function zoomIn() {
  zoomLevel = clampZoom(zoomLevel + 0.25);
  applyViewportTransform();
  if (selectedConnection) repositionConnDeletePopup();
}

function zoomOut() {
  zoomLevel = clampZoom(zoomLevel - 0.25);
  applyViewportTransform();
  if (selectedConnection) repositionConnDeletePopup();
}

function zoomReset() {
  zoomLevel = 1.0;
  applyViewportTransform();
  if (selectedConnection) repositionConnDeletePopup();
}

function resetView() {
  panX = 0; panY = 0; zoomLevel = 1.0;
  applyViewportTransform();
  hideConnDeletePopup();
}

// ── HIERARCHY BUILDER ─────────────────────────────────────────────────────────

function buildHierarchy() {
  const childIdSet = new Set(connections.map(c => c.toId));
  const roots      = nodes.filter(n => !childIdSet.has(n.id));
  const rows       = [];
  const visited    = new Set();

  const visit = (node, parent, level, rootSource, chain) => {
    if (visited.has(node.id)) return;
    visited.add(node.id);

    const newChain = chain ? `${chain} → ${node.name}` : node.name;
    const row      = {
      nodeId:         String(rows.length + 1).padStart(3, '0'),
      name:           node.name,
      type:           node.type,
      parent:         parent ? parent.name : '-',
      distLevel:      `Level ${level}`,
      rootSource:     rootSource,
      hierarchyChain: newChain,
    };
    NODE_FIELDS.forEach(f => { row[f.key] = node.fields[f.key] || ''; });
    rows.push(row);

    connections
      .filter(c => c.fromId === node.id)
      .forEach(c => {
        const child = nodes.find(n => n.id === c.toId);
        if (child) visit(child, node, level + 1, rootSource, newChain);
      });

    visited.delete(node.id);
  };

  roots.forEach(root => visit(root, null, 1, root.name, ''));
  return rows;
}

// ── COLUMN LETTER ──────────────────────────────────────────────────────────────

function colLetter(idx) {
  let result = '', n = idx + 1;
  while (n > 0) {
    const mod = (n - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

// ── EXPORT TO EXCEL ───────────────────────────────────────────────────────────

async function exportToExcel() {
  document.getElementById('export-panel').classList.add('visible');
  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    const select = document.getElementById('export-sheet-select');
    select.innerHTML = sheets.items
      .map(s => `<option value="${s.name}">${s.name}</option>`)
      .join('');
  });
}

function cancelExport() {
  document.getElementById('export-panel').classList.remove('visible');
}

async function confirmExport() {
  const sheetName = document.getElementById('export-sheet-select').value;
  const tableName = (document.getElementById('export-table-name').value || 'NodeHierarchy').trim();
  const rows      = buildHierarchy();

  if (rows.length === 0) { showToast('No nodes to export.'); return; }

  const fixedHeaders = ['Node ID', 'Node Name', 'Type', 'Parent', 'Dist Level', 'Root Source', 'Hierarchy Chain'];
  const fieldHeaders = NODE_FIELDS.map(f => f.label);
  const spareHeaders = ['Spare 1', 'Spare 2', 'Spare 3'];
  const headers      = [...fixedHeaders, ...fieldHeaders, ...spareHeaders];
  const colCount     = headers.length;

  const levelColors = {
    'Level 1': '#c8e6c9',
    'Level 2': '#bbdefb',
    'Level 3': '#ffe0b2',
    'Level 4': '#f8bbd0',
  };

  await Excel.run(async (ctx) => {
    const sheet     = ctx.workbook.worksheets.getItem(sheetName);
    let wasUpdate   = false;

    try {
      const existing = ctx.workbook.tables.getItem(tableName);
      existing.load('name');
      await ctx.sync();
      existing.delete();
      await ctx.sync();
      wasUpdate = true;
    } catch (_) {}

    sheet.getRangeByIndexes(0, 0, 1, colCount).values = [headers];

    const dataValues = rows.map(row => {
      const level  = parseInt(row.distLevel.split(' ')[1], 10);
      const indent = '  '.repeat(level - 1);
      return [
        row.nodeId,
        indent + row.name,
        row.type,
        row.parent,
        row.distLevel,
        row.rootSource,
        row.hierarchyChain,
        ...NODE_FIELDS.map(f => row[f.key] || ''),
        '', '', '',
      ];
    });

    if (dataValues.length > 0) {
      sheet.getRangeByIndexes(1, 0, dataValues.length, colCount).values = dataValues;
    }

    rows.forEach((row, i) => {
      const color = levelColors[row.distLevel] || '#ffffff';
      sheet.getRangeByIndexes(1 + i, 0, 1, colCount).format.fill.color = color;
    });

    const totalRows = 1 + rows.length;
    const tableAddr = `A1:${colLetter(colCount - 1)}${totalRows}`;
    const tbl       = sheet.tables.add(tableAddr, true);
    tbl.name  = tableName;
    tbl.style = 'TableStyleMedium2';
    await ctx.sync();

    const headerRow = tbl.getHeaderRowRange();
    headerRow.format.fill.color = '#1565c0';
    headerRow.format.font.color = '#ffffff';
    headerRow.format.font.bold  = true;

    sheet.getRangeByIndexes(0, 0, totalRows, colCount).format.autofitColumns();
    await ctx.sync();

    cancelExport();
    showToast(wasUpdate ? `Table '${tableName}' updated` : `Table '${tableName}' created`);
  });
}

// ── LOAD FROM EXCEL ───────────────────────────────────────────────────────────

function loadFromExcel() {
  document.getElementById('load-panel').classList.add('visible');
}

function cancelLoad() {
  document.getElementById('load-panel').classList.remove('visible');
}

async function confirmLoad() {
  const tableName = (document.getElementById('load-table-name').value || 'NodeHierarchy').trim();

  await Excel.run(async (ctx) => {
    let table;
    try {
      table = ctx.workbook.tables.getItem(tableName);
      table.load('name');
      await ctx.sync();
    } catch (_) {
      showToast(`Table '${tableName}' not found.`);
      return;
    }

    const headerRange = table.getHeaderRowRange();
    headerRange.load('values');
    const bodyRange = table.getDataBodyRange();
    bodyRange.load('values');
    await ctx.sync();

    const rawHeaders = headerRange.values[0] || [];
    const rawData    = bodyRange.values       || [];

    const colIdx = {};
    rawHeaders.forEach((h, i) => { colIdx[String(h)] = i; });

    const validRows = rawData.filter(row =>
      row.some(cell => cell !== null && cell !== '' && cell !== undefined)
    );

    if (validRows.length === 0) { showToast('Table has no data rows.'); return; }

    dragState = null;
    isPanning = false;
    if (connectDragState) {
      try { if (connectDragState.parentNode) connectDragState.remove(); } catch (_) {}
      connectDragState = null;
    }

    nodes              = [];
    connections        = [];
    selectedId         = null;
    selectedConnection = null;
    nodeCounters       = { Source: 0, DB: 0, DC: 0, TX: 0, PB: 0 };
    globalIdCounter    = 0;

    const nodesGroup = document.getElementById('nodes');
    const connsGroup = document.getElementById('connections');
    const tempGroup  = document.getElementById('temp-connections');
    if (nodesGroup) nodesGroup.innerHTML = '';
    if (connsGroup) connsGroup.innerHTML = '';
    if (tempGroup)  tempGroup.innerHTML  = '';
    hideConnDeletePopup();
    updatePropertiesPanel(null);

    const nameToNode = {};

    validRows.forEach(row => {
      const rawName = String(row[colIdx['Node Name']] ?? '').replace(/^\s+/, '');
      const type    = String(row[colIdx['Type']]      ?? '').trim();
      if (!rawName || !NODE_STYLES[type]) return;

      const id = `node-${++globalIdCounter}`;
      const m  = rawName.match(/^([A-Za-z]+)(\d+)$/);
      if (m && nodeCounters[m[1]] !== undefined) {
        nodeCounters[m[1]] = Math.max(nodeCounters[m[1]], parseInt(m[2], 10));
      }

      const fields = {};
      NODE_FIELDS.forEach(f => { fields[f.key] = String(row[colIdx[f.label]] ?? ''); });

      const node = {
        id,
        name:        rawName,
        type,
        x: 0, y: 0,
        fields,
        _parentName: String(row[colIdx['Parent']] ?? '').trim(),
      };

      nodes.push(node);
      nameToNode[rawName] = node;
    });

    nodes.forEach(node => {
      const parentName = node._parentName;
      delete node._parentName;
      if (parentName && parentName !== '-' && nameToNode[parentName]) {
        const fromNode = nameToNode[parentName];
        if (!connections.find(c => c.fromId === fromNode.id && c.toId === node.id)) {
          connections.push({ fromId: fromNode.id, toId: node.id });
        }
      }
    });

    autoLayoutNodes();
    nodes.forEach(node => renderNode(node));
    redrawConnections();
    updateConnectionCount();

    panX = 0; panY = 0; zoomLevel = 1.0;
    applyViewportTransform();

    cancelLoad();
    showToast(`Loaded ${nodes.length} nodes`);
  });
}

// ── AUTO LAYOUT ───────────────────────────────────────────────────────────────

function autoLayoutNodes() {
  const H  = 120;
  const V  = 100;
  const MX = 100;
  const MY = 80;

  const childIdSet = new Set(connections.map(c => c.toId));
  const roots      = nodes.filter(n => !childIdSet.has(n.id));

  function subtreeWidth(node, seen) {
    seen = seen || new Set();
    if (seen.has(node.id)) return H;
    seen.add(node.id);
    const children = connections
      .filter(c => c.fromId === node.id)
      .map(c => nodes.find(n => n.id === c.toId))
      .filter(Boolean);
    return children.length === 0
      ? H
      : children.reduce((sum, ch) => sum + subtreeWidth(ch, new Set(seen)), 0);
  }

  const placed = new Set();

  function place(node, xLeft, y) {
    if (placed.has(node.id)) return;
    placed.add(node.id);
    const children = connections
      .filter(c => c.fromId === node.id)
      .map(c => nodes.find(n => n.id === c.toId))
      .filter(Boolean);

    if (children.length === 0) {
      node.x = xLeft + H / 2;
      node.y = y;
      return;
    }

    let cx = xLeft;
    children.forEach(child => {
      const w = subtreeWidth(child);
      place(child, cx, y + V);
      cx += w;
    });
    node.x = (children[0].x + children[children.length - 1].x) / 2;
    node.y = y;
  }

  let xOffset = MX;
  roots.forEach(root => {
    const w = subtreeWidth(root);
    place(root, xOffset, MY);
    xOffset += w + H;
  });
}

// ── TOAST ─────────────────────────────────────────────────────────────────────

function showToast(message) {
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), 2000);
}

// ── EXPOSE FUNCTIONS TO GLOBAL SCOPE ─────────────────────────────────────────
// Required because webpack bundles code in a module scope.
// HTML onclick attributes need functions on the window object.
window.addNode                  = addNode;
window.switchTool               = switchTool;
window.zoomIn                   = zoomIn;
window.zoomOut                  = zoomOut;
window.zoomReset                = zoomReset;
window.resetView                = resetView;
window.togglePropertiesCollapse = togglePropertiesCollapse;
window.updateNodeProp           = updateNodeProp;
window.deleteNode               = deleteNode;
window.clearCanvas              = clearCanvas;
window.exportToExcel            = exportToExcel;
window.cancelExport             = cancelExport;
window.confirmExport            = confirmExport;
window.loadFromExcel            = loadFromExcel;
window.cancelLoad               = cancelLoad;
window.confirmLoad              = confirmLoad;
window.deleteSelectedConnection = deleteSelectedConnection;
window.hideConnDeletePopup      = hideConnDeletePopup;
