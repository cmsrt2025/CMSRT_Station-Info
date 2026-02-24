/* ================= CONFIG ================= */
const SHEET_ID = '1wsOOFGM0eVUrpozcOWJFQSpJEBpwH0l7iC0gpWxbx6M';
const SHEET_NAME = 'Sheet1';
const SECTION_SHEET_NAME = 'Sheet2';
const EXCEL_ONLINE_URL = 'https://ccivproject.sharepoint.com/:x:/r/sites/SRTCommu/Shared%20Documents/SRT/CCIV/CMSRT%20Station%20Info.xlsx?d=w5d2ae345211b469d9b81824fcec8cb36&csf=1&web=1&e=Gn2eK9';

// กำหนดค่าคอลัมน์ต่างๆ ตามโครงสร้างใหม่
const CONFIG = {
  noColumn: 'No',
  nameColumn: 'Station Name',
  latColumn: 'Latitude',
  lngColumn: 'Longtitude',
  statusColumn: 'Status',
  sectionFromColumn: 'Hop OFC link A',
  sectionToColumn: 'Hop OFC link B',
  sectionStatusColumn: 'Status',
  sectionInstallColumn: 'Type Cable Install',
  sectionFromColumnIndex: 2,
  sectionToColumnIndex: 3,
  sectionInstallColumnIndex: 4,
  sectionStatusColumnIndex: 10,
  sectionTypeColumn: 'type',
  sectionDistanceColumn: 'Distance',
  sectionLinkColumn: 'Link',
  regionColumn: 'Region',
  dwdmColumn: 'DWDM Site type',
  dwdmColumnIndex: 8,
  mplsColumn: 'MPLS Site Type',
  dwgUrlColumn: 'Ins DWG',
  defaultFilterColumns: ['Consultant', 'Region', 'Province', 'Type of station', 'MPLS Site Type'],
  // คอลัมน์ที่ต้องการให้มี filter (checkbox)
  filterColumns: ['Consultant', 'Region', 'Province', 'Type of station', 'MPLS Site Type', 'Status', 'DWDM Site type'],
  // คอลัมน์ที่ไม่ต้องการแสดงใน popup
  excludeFromPopup: ['Ins DWG']
};
/* ========================================== */

const map = L.map('map').setView([13.7, 100.5], 6);

const layers = {
  street: L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '&copy; OpenStreetMap contributors',
    maxZoom: 19
  }),
  satellite: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
    attribution: '&copy; Esri',
    maxZoom: 19
  })
};

let currentLayer = layers.street.addTo(map);

function changeLayer(layerName) {
  map.removeLayer(currentLayer);
  currentLayer = layers[layerName];
  currentLayer.addTo(map);
  
  document.querySelectorAll('.map-btn').forEach(btn => {
    btn.classList.remove('active');
  });
  event.target.classList.add('active');
}

const menuToggle = document.getElementById('menuToggle');

function setMenuOpen(isOpen) {
  document.body.classList.toggle('menu-open', isOpen);
}

if (menuToggle) {
  menuToggle.addEventListener('click', (event) => {
    event.stopPropagation();
    setMenuOpen(!document.body.classList.contains('menu-open'));
  });
}

let allMarkers = [];
let allStationLabelMarkers = [];
let allSectionLines = [];
let allData = [];
let columnHeaders = [];
let allSectionData = [];
let sectionColumnHeaders = [];
let filterCheckboxes = {};
let currentStyle = 'base';
let showLabels = false;
let showInstallationStations = true;
let showInstallationLines = true;
let allFilterColumns = [];
let selectedFilterColumns = new Set();
let activeFilterColumns = [];

const DWDM_ICON_URLS = {
  'OTM': 'assets/icons/OLM-M24.png',
  'OLA': 'assets/icons/OLA-M12.png',
  'AMP': 'assets/icons/BOOTSER-M5.png',
  'OTM+AMP': 'assets/icons/OLM+BOOTSER.png'
};

const MPLS_ICON_URLS = {
  'Core': 'assets/icons/CORE-M8.png',
  'Agg': 'assets/icons/AGG-M6.png',
  'M-Core1': 'assets/icons/DC.png',
  'M-Core2': 'assets/icons/DR.png'
};

const iconSize = [24, 24];
const iconAnchor = [12, 24];
const popupAnchor = [0, -24];

function getIconForValue(value, urlMap) {
  if (!urlMap) return null;
  const key = value == null ? '' : String(value).trim();
  return urlMap[key] || urlMap.default || null;
}

function createImageIcon(url) {
  if (!url) return null;
  return L.icon({
    iconUrl: url,
    iconSize,
    iconAnchor,
    popupAnchor
  });
}

function getColumnIndexByHeaders(headers, name, fallbackIndex) {
  if (typeof fallbackIndex === 'number' && fallbackIndex >= 0) {
    return fallbackIndex;
  }
  if (!name) return -1;
  const exact = headers.indexOf(name);
  if (exact != -1) return exact;
  const target = String(name).trim().toLowerCase();
  return headers.findIndex(col => String(col).trim().toLowerCase() === target);
}

function getColumnIndex(name, fallbackIndex) {
  return getColumnIndexByHeaders(columnHeaders, name, fallbackIndex);
}

function getSectionColumnIndex(name, fallbackIndex) {
  return getColumnIndexByHeaders(sectionColumnHeaders, name, fallbackIndex);
}

function getSectionColumnIndexAny(names, fallbackIndex) {
  const list = Array.isArray(names) ? names : [names];
  for (const name of list) {
    const idx = getSectionColumnIndex(name);
    if (idx !== -1) return idx;
  }
  if (typeof fallbackIndex === 'number' && fallbackIndex >= 0) {
    return fallbackIndex;
  }
  return -1;
}

function normalizeSectionStatusValue(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  const key = getStatusKey(text);
  if (key === 'status') return '';
  return text;
}

function normalizeColumnName(name) {
  const value = String(name || '').trim().toLowerCase();
  const compact = value.replace(/[^a-z0-9]+/g, '');
  if (compact.startsWith('consult')) {
    return 'consultant';
  }
  return compact;
}

const regionColorMap = new Map();
const regionColorPalette = [
  '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728',
  '#9467bd', '#8c564b', '#e377c2', '#7f7f7f',
  '#bcbd22', '#17becf'
];

const dwdmColorMap = new Map();
const mplsColorMap = new Map();
const statusColorMap = new Map();
const customStatusColorMap = new Map();
const typeColorPalette = [
  '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728',
  '#9467bd', '#8c564b', '#e377c2', '#7f7f7f',
  '#bcbd22', '#17becf'
];
const statusColorPalette = [
  '#22c55e', '#f59e0b', '#ef4444', '#3b82f6', '#6b7280', '#14b8a6'
];
const STATUS_COLOR_STORAGE_KEY = 'cmsrt_status_colors_v1';
const SIMPLE_STATUS_COLOR_OPTIONS = [
  { label: 'Default', value: '' },
  { label: 'Blue', value: '#3b82f6' },
  { label: 'Indigo', value: '#6366f1' },
  { label: 'Violet', value: '#8b5cf6' },
  { label: 'Cyan', value: '#06b6d4' },
  { label: 'Teal', value: '#14b8a6' },
  { label: 'Green', value: '#22c55e' },
  { label: 'Lime', value: '#84cc16' },
  { label: 'Yellow', value: '#eab308' },
  { label: 'Amber', value: '#f59e0b' },
  { label: 'Orange', value: '#f97316' },
  { label: 'Red', value: '#ef4444' },
  { label: 'Rose', value: '#f43f5e' },
  { label: 'Slate', value: '#64748b' },
  { label: 'Gray', value: '#6b7280' },
  { label: 'Black', value: '#111827' }
];

function getStatusKey(status) {
  return String(status || '').trim().toLowerCase();
}

function loadCustomStatusColors() {
  try {
    const raw = localStorage.getItem(STATUS_COLOR_STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    Object.entries(parsed).forEach(([key, color]) => {
      if (key && color) {
        customStatusColorMap.set(key, color);
      }
    });
  } catch (err) {
    console.warn('Unable to load status colors from storage', err);
  }
}

function saveCustomStatusColors() {
  try {
    const asObject = Object.fromEntries(customStatusColorMap.entries());
    localStorage.setItem(STATUS_COLOR_STORAGE_KEY, JSON.stringify(asObject));
  } catch (err) {
    console.warn('Unable to save status colors to storage', err);
  }
}

function setCustomStatusColor(status, color) {
  const key = getStatusKey(status);
  if (!key) return;
  if (!color) {
    customStatusColorMap.delete(key);
    saveCustomStatusColors();
    return;
  }
  customStatusColorMap.set(key, color);
  saveCustomStatusColors();
}

function buildStatusColorPaletteHtml(selectedColor, statusValue) {
  const statusEncoded = encodeURIComponent(String(statusValue || ''));
  return SIMPLE_STATUS_COLOR_OPTIONS.map(item => {
    const isSelected = item.value === selectedColor ? ' is-selected' : '';
    const isDefault = item.value === '' ? ' is-default' : '';
    const swatchStyle = item.value ? ` style="--dot-color:${item.value}"` : '';
    return `<button type="button" class="legend-color-dot${isSelected}${isDefault}" data-status="${statusEncoded}" data-color="${item.value}" title="${item.label}"${swatchStyle}></button>`;
  }).join('');
}

function getCustomStatusColor(statusValue) {
  const key = getStatusKey(statusValue);
  return customStatusColorMap.get(key) || '';
}

function openLegendColorPlate(legend, targetEl, statusValue) {
  legend.querySelectorAll('.legend-color-plate').forEach(node => node.remove());

  const selectedColor = getCustomStatusColor(statusValue);
  const plate = document.createElement('div');
  plate.className = 'legend-color-plate';
  plate.innerHTML = buildStatusColorPaletteHtml(selectedColor, statusValue);
  legend.appendChild(plate);

  const legendRect = legend.getBoundingClientRect();
  const targetRect = targetEl.getBoundingClientRect();
  const left = targetRect.left - legendRect.left + targetRect.width + 8;
  const top = targetRect.top - legendRect.top - 4;
  plate.style.left = `${Math.max(8, left)}px`;
  plate.style.top = `${Math.max(8, top)}px`;

  plate.querySelectorAll('.legend-color-dot').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const status = decodeURIComponent(btn.getAttribute('data-status') || '');
      const color = btn.getAttribute('data-color') || '';
      setCustomStatusColor(status, color);
      applyStyleToMarkers();
      updateLegend();
    });
  });

  const closePlate = (event) => {
    if (!plate.contains(event.target) && event.target !== targetEl) {
      plate.remove();
      document.removeEventListener('click', closePlate, true);
    }
  };
  setTimeout(() => document.addEventListener('click', closePlate, true), 0);
}

function getTypeColor(value, colorMap) {
  if (!value) return '#6b7280';
  const key = String(value).trim().toLowerCase();
  if (!key) return '#6b7280';
  if (!colorMap.has(key)) {
    const color = typeColorPalette[colorMap.size % typeColorPalette.length];
    colorMap.set(key, color);
  }
  return colorMap.get(key);
}

function getDwdmColor(value) {
  return getTypeColor(value, dwdmColorMap);
}

function getMplsColor(value) {
  return getTypeColor(value, mplsColorMap);
}

function getRegionColor(region) {
  if (!region) return '#6b7280';
  const key = String(region).trim().toLowerCase();
  if (!key) return '#6b7280';
  if (!regionColorMap.has(key)) {
    const color = regionColorPalette[regionColorMap.size % regionColorPalette.length];
    regionColorMap.set(key, color);
  }
  return regionColorMap.get(key);
}

function statusColor(status) {
  if (!status) return '#6b7280';
  const s = getStatusKey(status);
  if (!s) return '#6b7280';
  if (s === '-') return '#9ca3af';
  if (s.includes('ดึงสาย')) return '#ef4444';
  if (s.includes('จบงาน')) return '#22c55e';
  if (customStatusColorMap.has(s)) {
    return customStatusColorMap.get(s);
  }
  if (
    s.includes('ยังไม่ได้เริ่ม') ||
    s.includes('not started') ||
    s.includes('notstarted') ||
    s.includes('pending')
  ) {
    return '#9ca3af';
  }
  if (!statusColorMap.has(s)) {
    const color = statusColorPalette[statusColorMap.size % statusColorPalette.length];
    statusColorMap.set(s, color);
  }
  return statusColorMap.get(s);
}

function normalizeKey(value) {
  return String(value || '').trim().toLowerCase();
}

function parseSectionInstall(installValue) {
  const raw = String(installValue || '').trim();
  const method = raw.toLowerCase();
  let methodLabel = 'Unknown';
  let dashArray = '4 6';
  let lineCap = 'butt';
  let lineJoin = 'round';

  if (method.includes('existing')) {
    methodLabel = 'Existing';
    // small spaced dots
    dashArray = '1 10';
    lineCap = 'round';
  } else if (method.includes('underground')) {
    methodLabel = 'Underground';
    // long-short dashed (dimension-line style)
    dashArray = '16 6 4 6';
    lineCap = 'butt';
    lineJoin = 'round';
  } else if (
    method.includes('pole') ||
    method.includes('overhead') ||
    method.includes('aerial')
  ) {
    methodLabel = 'Aerial';
    dashArray = null;
  }

  const status = raw
    .replace(/existing|underground|overhead|aerial|pole/gi, '')
    .replace(/\s+/g, ' ')
    .trim() || raw || '-';

  return {
    status,
    methodLabel,
    dashArray,
    lineCap,
    lineJoin
  };
}

function createSectionLines(sectionRowsData) {
  allSectionLines.forEach(item => map.removeLayer(item.line));
  allSectionLines = [];

  const latIndex = getColumnIndex(CONFIG.latColumn);
  const lngIndex = getColumnIndex(CONFIG.lngColumn);
  const nameIndex = getColumnIndex(CONFIG.nameColumn);
  const sectionNoIndex = 0;
  const sectionLinkIndex = 1;
  const sectionFromIndex = 2;
  const sectionToIndex = 3;
  const sectionInstallIndex = 4;
  const sectionTypeIndex = 5;
  const sectionDistanceIndex = 6;
  const sectionStatusIndex = 10;

  if (latIndex === -1 || lngIndex === -1 || nameIndex === -1) return;

  const stationByName = new Map();
  allData.forEach(row => {
    stationByName.set(normalizeKey(row[nameIndex]), row);
  });

  let createdLineCount = 0;
  sectionRowsData.forEach((sectionRow) => {
    const fromName = sectionRow[sectionFromIndex] || '-';
    const toName = sectionRow[sectionToIndex] || '-';
    if (!fromName || !toName) return;
    if (getStatusKey(fromName) === getStatusKey(CONFIG.sectionFromColumn)) return;
    if (getStatusKey(toName) === getStatusKey(CONFIG.sectionToColumn)) return;
    const fromStationRow = stationByName.get(normalizeKey(fromName));
    const toStationRow = stationByName.get(normalizeKey(toName));
    if (!fromStationRow || !toStationRow) return;

    const fromLat = parseFloat(fromStationRow[latIndex]);
    const fromLng = parseFloat(fromStationRow[lngIndex]);
    const toLat = parseFloat(toStationRow[latIndex]);
    const toLng = parseFloat(toStationRow[lngIndex]);
    if (!Number.isFinite(fromLat) || !Number.isFinite(fromLng)) return;
    if (!Number.isFinite(toLat) || !Number.isFinite(toLng)) return;

    const installText = sectionRow[sectionInstallIndex] || '';
    const parsedInstall = parseSectionInstall(installText);
    const sectionStatusRaw = sectionRow[sectionStatusIndex];
    const sectionStatus = normalizeSectionStatusValue(sectionStatusRaw) || parsedInstall.status;
    const sectionColor = String(sectionStatus || '').trim() === '-'
      ? '#9ca3af'
      : statusColor(sectionStatus);

    const fromLatLng = L.latLng(fromLat, fromLng);
    const toLatLng = L.latLng(toLat, toLng);
    const linePoints = [fromLatLng, toLatLng];

    const line = L.polyline(linePoints, {
      color: sectionColor,
      weight: 4,
      opacity: 0.85,
      dashArray: parsedInstall.dashArray,
      lineCap: parsedInstall.lineCap,
      lineJoin: parsedInstall.lineJoin
    });
    const noValue = sectionRow[sectionNoIndex] || '-';
    const linkValue = sectionRow[sectionLinkIndex] || '-';
    const typeValue = sectionRow[sectionTypeIndex] || '-';
    const distanceValue = sectionRow[sectionDistanceIndex] || '-';

    line.bindPopup(`
      <b>${fromName} -> ${toName}</b><br>
      Status: ${sectionStatus}<br>
      Method: ${parsedInstall.methodLabel}<br>
      No: ${noValue}<br>
      Link: ${linkValue}<br>
      Type: ${typeValue}<br>
      Distance: ${distanceValue}
    `, {
      maxWidth: 320,
      className: 'custom-popup'
    });

    allSectionLines.push({
      line,
      sectionStatus,
      methodLabel: parsedInstall.methodLabel
    });
    createdLineCount++;
  });
  console.log(`Section lines created: ${createdLineCount}`);
}

function updateSectionLinesVisibility() {
  const showLines = currentStyle === 'installation' && showInstallationLines;
  allSectionLines.forEach(item => {
    if (showLines) {
      item.line.addTo(map);
    } else {
      map.removeLayer(item.line);
    }
  });
}

function applyStyleToSectionLines() {
  allSectionLines.forEach(item => {
    const statusText = item.sectionStatus;
    const color = String(statusText || '').trim() === '-'
      ? '#9ca3af'
      : statusColor(statusText);
    if (item.line && item.line.setStyle) {
      item.line.setStyle({ color });
    }
  });
}

function updateInstallationModeOptionsVisibility() {
  const panel = document.getElementById('installationModeOptions');
  if (!panel) return;
  panel.classList.toggle('show', currentStyle === 'installation');
}

function syncInstallationTogglesToUI() {
  const stationsCheckbox = document.getElementById('showInstallationStations');
  const linesCheckbox = document.getElementById('showInstallationLines');
  if (stationsCheckbox) stationsCheckbox.checked = showInstallationStations;
  if (linesCheckbox) linesCheckbox.checked = showInstallationLines;
}

function getOpticalIcon(dwdmType) {
  if (!dwdmType) return null;
  if (String(dwdmType).trim() === '-') return null;
  const url = getIconForValue(dwdmType, DWDM_ICON_URLS);
  return createImageIcon(url);
}

function getMplsIcon(mplsType) {
  const url = getIconForValue(mplsType, MPLS_ICON_URLS);
  return createImageIcon(url);
}

function getMarkerStyle(status, styleMode, region, dwdmType, mplsType) {
  switch (styleMode) {
    case 'optical':
      if (dwdmType && String(dwdmType).trim() === '-') {
        return { radius: 6, fillColor: '#d1d5db' };
      }
      return { radius: 10, fillColor: getDwdmColor(dwdmType) };
    case 'mpls':
      if (mplsType && String(mplsType).trim().toLowerCase() === 'access') {
        return { radius: 6, fillColor: '#d1d5db' };
      }
      return { radius: 10, fillColor: getMplsColor(mplsType) };
    case 'installation':
      if (status && String(status).trim() === '-') {
        return { radius: 6, fillColor: '#d1d5db' };
      }
      return { radius: 10, fillColor: statusColor(status) };
    case 'base':
    default:
      return { radius: 10, fillColor: getRegionColor(region) };
  }
}

function wantsImageMarker(styleMode, dwdmType, mplsType) {
  if (styleMode === 'optical') {
    return dwdmType && String(dwdmType).trim() !== '-';
  }
  if (styleMode === 'mpls') {
    return !(mplsType && String(mplsType).trim().toLowerCase() === 'access');
  }
  return styleMode === 'mpls';
}

function createCircleMarker(latlng, style) {
  return L.circleMarker(latlng, {
    radius: style.radius,
    fillColor: style.fillColor,
    color: '#fff',
    weight: 2,
    fillOpacity: 0.9
  });
}

function updateMarkerLabel(markerObj) {
  // Labels are handled by dedicated global station label markers only.
  // Ensure no tooltip is bound directly on station markers.
  if (!markerObj || !markerObj.marker) return;
  if (markerObj.marker.unbindTooltip) {
    markerObj.marker.unbindTooltip();
  }
}

function applyLabelsToMarkers() {
  updateGlobalStationLabels();
}

function updateGlobalStationLabels() {
  allStationLabelMarkers.forEach(item => {
    if (!item || !item.marker) return;
    if (showLabels) {
      item.marker.addTo(map);
    } else {
      map.removeLayer(item.marker);
    }
  });
}

function updateLegend() {
  const legend = document.getElementById('legend');
  if (!legend) return;

  const items = [];
  let title = 'Legend';

  if (currentStyle === 'base') {
    title = 'Region';
    const regionIndex = getColumnIndex(CONFIG.regionColumn);
    if (regionIndex !== -1) {
      const values = allData.map(r => r[regionIndex]).filter(v => v);
      const unique = [...new Set(values)].sort();
      unique.forEach(value => {
        items.push({ label: value, color: getRegionColor(value), isImage: false });
      });
    }
  } else if (currentStyle === 'optical') {
    title = 'DWDM Site Type';
    const dwdmIndex = getColumnIndex(CONFIG.dwdmColumn, CONFIG.dwdmColumnIndex);
    if (dwdmIndex !== -1) {
      const values = allData.map(r => r[dwdmIndex]).filter(v => v);
      const unique = [...new Set(values)].sort();
      unique.forEach(value => {
        const labelValue = String(value);
        if (labelValue.trim() === '-') {
          items.push({ label: value, color: '#d1d5db', isImage: false });
        } else {
          const iconUrl = getIconForValue(value, DWDM_ICON_URLS);
          items.push({ label: value, iconUrl, isImage: true });
        }
      });
    }
  } else if (currentStyle === 'mpls') {
    title = 'MPLS Site Type';
    const mplsIndex = getColumnIndex(CONFIG.mplsColumn);
    if (mplsIndex !== -1) {
      const values = allData.map(r => r[mplsIndex]).filter(v => v);
      const unique = [...new Set(values)].sort();
      unique.forEach(value => {
        const labelValue = String(value);
        if (labelValue.trim().toLowerCase() === 'access') {
          items.push({ label: value, color: '#d1d5db', isImage: false });
        } else {
          const iconUrl = getIconForValue(value, MPLS_ICON_URLS);
          items.push({ label: value, iconUrl, isImage: true });
        }
      });
    }
  } else if (currentStyle === 'installation') {
    const stationOn = showInstallationStations;
    const lineOn = showInstallationLines;
    title = stationOn && lineOn
      ? 'Status (Indoor/Outdoor)'
      : stationOn
        ? 'Status (Indoor)'
        : lineOn
          ? 'Status (Outdoor)'
          : 'Status';

    if (stationOn) {
      const statusIndex = getColumnIndex(CONFIG.statusColumn);
      if (statusIndex !== -1) {
        const values = allData.map(r => r[statusIndex]).filter(v => v);
        const unique = [...new Set(values)];
        const stationOrder = [
          'เริ่มปรับพื้นที่/ลงเข็ม',
          'เทปูนคอลัมน์',
          'เทปูน Slab',
          'ติดตั้ง E-Stand'
        ];
        const keyToValue = new Map(unique.map(v => [getStatusKey(v), v]));
        const ordered = stationOrder
          .map(v => keyToValue.get(getStatusKey(v)))
          .filter(Boolean);
        const orderedKeys = new Set(ordered.map(v => getStatusKey(v)));
        const extras = unique
          .filter(v => !orderedKeys.has(getStatusKey(v)))
          .sort((a, b) => String(a).localeCompare(String(b)));

        items.push({ kind: 'heading', label: 'Indoor' });
        [...ordered, ...extras].forEach(value => {
          const labelValue = String(value);
          if (labelValue.trim() === '-') {
            items.push({ label: value, color: '#d1d5db', isImage: false });
          } else {
            items.push({
              label: value,
              color: statusColor(value),
              isImage: false,
              statusValue: value,
              editableColor: true
            });
          }
        });
      }
    }

    if (lineOn) {
      if (stationOn) {
        items.push({ kind: 'divider' });
      }
      items.push({ kind: 'heading', label: 'Outdoor' });
      items.push({ label: 'จบงาน', color: '#22c55e', isImage: false });
      items.push({ label: 'ดึงสาย', color: '#ef4444', isImage: false });
      items.push({ label: '-', color: '#9ca3af', isImage: false });

      items.push({ kind: 'divider' });
      items.push({ kind: 'heading', label: 'Line Type' });
      items.push({ label: 'Aerial', lineSample: 'solid' });
      items.push({ label: 'Underground', lineSample: 'dashmix' });
      items.push({ label: 'Existing', lineSample: 'dotted', lineColor: '#9ca3af' });
    }
  }

  let html = `<div class="legend-header"><h4>${title}</h4></div>`;
  if (items.length === 0) {
    html += '<div class="legend-empty">ไม่มีข้อมูล</div>';
  } else {
    items.forEach(item => {
      if (item.kind === 'divider') {
        html += '<div class="legend-divider"></div>';
        return;
      }
      if (item.kind === 'heading') {
        html += `<div class="legend-group-title">${item.label}</div>`;
        return;
      }
      const statusAttr = item.editableColor && item.statusValue
        ? ` data-status="${encodeURIComponent(String(item.statusValue))}" role="button" tabindex="0" title="Pick color"`
        : '';
      const swatchClass = `legend-swatch${item.editableColor ? ' legend-swatch-editable' : ''}`;
      const markerHtml = item.isImage && item.iconUrl
        ? `<img class="legend-icon" src="${item.iconUrl}" alt="">`
        : item.lineSample
          ? `<span class="legend-line ${item.lineSample === 'dashed' ? 'is-dashed' : ''} ${item.lineSample === 'dotted' ? 'is-dotted' : ''} ${item.lineSample === 'dashmix' ? 'is-dashmix' : ''}"${item.lineColor ? ` style="border-top-color:${item.lineColor}"` : ''}></span>`
          : `<span class="${swatchClass}" style="background:${item.color || '#d1d5db'}"${statusAttr}></span>`;
      html += `
        <div class="legend-item">
          ${markerHtml}
          <span class="legend-label">${item.label}</span>
        </div>`;
    });
  }
  legend.innerHTML = html;

  legend.querySelectorAll('.legend-swatch-editable').forEach(swatch => {
    swatch.addEventListener('click', (e) => {
      e.stopPropagation();
      const status = decodeURIComponent(swatch.getAttribute('data-status') || '');
      openLegendColorPlate(legend, swatch, status);
    });
    swatch.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        const status = decodeURIComponent(swatch.getAttribute('data-status') || '');
        openLegendColorPlate(legend, swatch, status);
      }
    });
  });
}

function applyStyleToMarkers() {
  allMarkers.forEach(m => {
    const statusIndex = getColumnIndex(CONFIG.statusColumn);
    const regionIndex = getColumnIndex(CONFIG.regionColumn);
    const dwdmIndex = getColumnIndex(CONFIG.dwdmColumn, CONFIG.dwdmColumnIndex);
    const mplsIndex = getColumnIndex(CONFIG.mplsColumn);
    const status = m.data[statusIndex];
    const region = regionIndex === -1 ? null : m.data[regionIndex];
    const dwdmType = dwdmIndex === -1 ? null : m.data[dwdmIndex];
    const mplsType = mplsIndex === -1 ? null : m.data[mplsIndex];
    const style = getMarkerStyle(status, currentStyle, region, dwdmType, mplsType);
    const useImage = wantsImageMarker(currentStyle, dwdmType, mplsType);
    if (useImage) {
      const icon = currentStyle === 'optical' ? getOpticalIcon(dwdmType) : getMplsIcon(mplsType);
      if (!icon) {
        const circleMarker = createCircleMarker(m.latlng, style);
        if (m.markerType !== 'circle') {
          map.removeLayer(m.marker);
          circleMarker.addTo(map);
          circleMarker.bindPopup(m.popupContent, {
            maxWidth: 420,
            className: 'custom-popup'
          });
          m.marker = circleMarker;
          m.markerType = 'circle';
          updateMarkerLabel(m);
        } else {
          m.marker.setStyle({
            radius: style.radius,
            fillColor: style.fillColor
          });
          updateMarkerLabel(m);
        }
        return;
      }

      if (m.markerType !== 'image') {
        map.removeLayer(m.marker);
        const imageMarker = L.marker(m.latlng, { icon }).addTo(map);
        imageMarker.bindPopup(m.popupContent, {
          maxWidth: 420,
          className: 'custom-popup'
        });
        m.marker = imageMarker;
        m.markerType = 'image';
        updateMarkerLabel(m);
      } else if (m.marker.setIcon) {
        m.marker.setIcon(icon);
        updateMarkerLabel(m);
      }
      return;
    }

    if (m.markerType !== 'circle') {
      map.removeLayer(m.marker);
      const circleMarker = createCircleMarker(m.latlng, style);
      circleMarker.addTo(map);
      circleMarker.bindPopup(m.popupContent, {
        maxWidth: 420,
        className: 'custom-popup'
      });
      m.marker = circleMarker;
      m.markerType = 'circle';
      updateMarkerLabel(m);
    } else {
      m.marker.setStyle({
        radius: style.radius,
        fillColor: style.fillColor
      });
      updateMarkerLabel(m);
    }
  });
  applyStyleToSectionLines();
  updateSectionLinesVisibility();
}

function toggleAllCheckboxes(colName, checked) {
  const checkboxes = filterCheckboxes[colName];
  checkboxes.forEach(cb => {
    cb.checked = checked;
  });
  applyFilters();
}

function toggleDropdown(colName) {
  const checkboxGroup = document.querySelector(`#filter_${colName.replace(/\s+/g, '_')} .checkbox-group`);
  const arrow = document.querySelector(`#filter_${colName.replace(/\s+/g, '_')} .dropdown-arrow`);
  const header = document.querySelector(`#filter_${colName.replace(/\s+/g, '_')} .filter-group-header`);
  
  const isOpen = checkboxGroup.classList.contains('show');
  
  if (isOpen) {
    checkboxGroup.classList.remove('show');
    arrow.classList.remove('open');
    header.classList.remove('active');
  } else {
    checkboxGroup.classList.add('show');
    arrow.classList.add('open');
    header.classList.add('active');
  }
}

function updateActiveFilterColumns() {
  const orderMap = new Map(
    CONFIG.filterColumns.map((col, index) => [normalizeColumnName(col), index])
  );
  const fallbackIndex = CONFIG.filterColumns.length + 1;

  activeFilterColumns = allFilterColumns
    .filter(col => selectedFilterColumns.has(col))
    .sort((a, b) => {
      const aIndex = orderMap.has(normalizeColumnName(a)) ? orderMap.get(normalizeColumnName(a)) : fallbackIndex;
      const bIndex = orderMap.has(normalizeColumnName(b)) ? orderMap.get(normalizeColumnName(b)) : fallbackIndex;
      if (aIndex !== bIndex) return aIndex - bIndex;
      return a.localeCompare(b);
    });
  createFilters(columnHeaders, allData);
  applyFilters();
}

function buildFilterSelector() {
  const list = document.getElementById('filterSelectorList');
  if (!list) return;
  list.innerHTML = '';

  allFilterColumns.forEach(colName => {
    const checkboxItem = document.createElement('div');
    checkboxItem.className = 'checkbox-item';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = `filter_option_${colName}`.replace(/\s+/g, '_');
    checkbox.checked = selectedFilterColumns.has(colName);
    checkbox.onchange = () => {
      if (checkbox.checked) {
        selectedFilterColumns.add(colName);
      } else {
        selectedFilterColumns.delete(colName);
      }
      updateActiveFilterColumns();
    };

    const cbLabel = document.createElement('label');
    cbLabel.htmlFor = checkbox.id;
    cbLabel.textContent = colName;

    checkboxItem.appendChild(checkbox);
    checkboxItem.appendChild(cbLabel);

    checkboxItem.onclick = (e) => {
      if (e.target !== checkbox) {
        checkbox.checked = !checkbox.checked;
        checkbox.onchange();
      }
    };

    list.appendChild(checkboxItem);
  });
}

function setupFilterSelector() {
  const excluded = new Set([CONFIG.latColumn, CONFIG.lngColumn]);
  allFilterColumns = columnHeaders.filter(col => col && !excluded.has(col));

  const configSet = new Set(CONFIG.filterColumns.map(normalizeColumnName));
  selectedFilterColumns = new Set(
    allFilterColumns.filter(col => configSet.has(normalizeColumnName(col)))
  );

  if (CONFIG.defaultFilterColumns && CONFIG.defaultFilterColumns.length > 0) {
    const defaultSet = new Set(CONFIG.defaultFilterColumns.map(normalizeColumnName));
    selectedFilterColumns = new Set(
      allFilterColumns.filter(col => defaultSet.has(normalizeColumnName(col)))
    );
  }

  if (selectedFilterColumns.size === 0) {
    selectedFilterColumns = new Set(allFilterColumns);
  }

  buildFilterSelector();
  updateActiveFilterColumns();
}

function createFilters(columns, data) {
  const dynamicFilters = document.getElementById('dynamicFilters');
  dynamicFilters.innerHTML = '';

  filterCheckboxes = {};
  activeFilterColumns.forEach(colName => {
    const colIndex = getColumnIndex(colName);
    if (colIndex === -1) return;

    const values = data.map(row => row[colIndex]).filter(v => v);
    const uniqueValues = [...new Set(values)].sort();

    if (uniqueValues.length === 0) return;

    const filterGroup = document.createElement('div');
    filterGroup.className = 'filter-group';
    filterGroup.id = `filter_${colName.replace(/\s+/g, '_')}`;

    const headerDiv = document.createElement('div');
    headerDiv.className = 'filter-group-header';

    const titleDiv = document.createElement('div');
    titleDiv.className = 'filter-group-title';

    const arrow = document.createElement('span');
    arrow.className = 'dropdown-arrow';
    arrow.textContent = '▼';

    const label = document.createElement('label');
    label.textContent = colName;

    titleDiv.appendChild(arrow);
    titleDiv.appendChild(label);

    const btnContainer = document.createElement('div');
    btnContainer.className = 'filter-buttons';
    
    const selectAllBtn = document.createElement('button');
    selectAllBtn.className = 'select-all-btn';
    selectAllBtn.textContent = 'ทั้งหมด';
    selectAllBtn.onclick = (e) => {
      e.stopPropagation();
      toggleAllCheckboxes(colName, true);
    };

    const deselectAllBtn = document.createElement('button');
    deselectAllBtn.className = 'select-all-btn';
    deselectAllBtn.textContent = 'ไม่เลือก';
    deselectAllBtn.onclick = (e) => {
      e.stopPropagation();
      toggleAllCheckboxes(colName, false);
    };

    btnContainer.appendChild(selectAllBtn);
    btnContainer.appendChild(deselectAllBtn);

    headerDiv.appendChild(titleDiv);
    headerDiv.appendChild(btnContainer);

    // Add click event to toggle dropdown
    headerDiv.onclick = (e) => {
      if (e.target.tagName !== 'BUTTON') {
        toggleDropdown(colName);
      }
    };

    const checkboxGroup = document.createElement('div');
    checkboxGroup.className = 'checkbox-group';

    filterCheckboxes[colName] = [];

    uniqueValues.forEach(val => {
      const checkboxItem = document.createElement('div');
      checkboxItem.className = 'checkbox-item';

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.id = `cb_${colName}_${val}`.replace(/\s+/g, '_');
      checkbox.value = val;
      checkbox.checked = true;
      checkbox.onchange = applyFilters;

      const cbLabel = document.createElement('label');
      cbLabel.htmlFor = checkbox.id;
      cbLabel.textContent = val;

      checkboxItem.appendChild(checkbox);
      checkboxItem.appendChild(cbLabel);
      checkboxGroup.appendChild(checkboxItem);

      filterCheckboxes[colName].push(checkbox);

      // Click on the entire item to toggle checkbox
      checkboxItem.onclick = (e) => {
        if (e.target !== checkbox) {
          checkbox.checked = !checkbox.checked;
          applyFilters();
        }
      };
    });

    filterGroup.appendChild(headerDiv);
    filterGroup.appendChild(checkboxGroup);
    dynamicFilters.appendChild(filterGroup);
  });
}

function applyFilters() {
  const searchText = document.getElementById('searchBox').value.toLowerCase();
  let visibleCount = 0;

  allMarkers.forEach(m => {
    const noIndex = columnHeaders.indexOf(CONFIG.noColumn);
    const nameIndex = getColumnIndex(CONFIG.nameColumn);
    const stationNo = m.data[noIndex] ? String(m.data[noIndex]) : '';
    const stationName = m.data[nameIndex] || '';
    
    const matchSearch = !searchText || 
                       stationName.toLowerCase().includes(searchText) ||
                       stationNo.toLowerCase().includes(searchText);

    let matchFilters = true;
    for (let colName in filterCheckboxes) {
      const colIndex = getColumnIndex(colName);
      const rowValue = m.data[colIndex];
      
      // Check if at least one checkbox is checked for this column
      const checkedValues = filterCheckboxes[colName]
        .filter(cb => cb.checked)
        .map(cb => cb.value);
      
      // If no checkboxes are checked, show nothing for this filter
      if (checkedValues.length === 0) {
        matchFilters = false;
        break;
      }
      
      // Check if row value is in the checked values
      if (!checkedValues.includes(rowValue)) {
        matchFilters = false;
        break;
      }
    }

    const allowStation = !(currentStyle === 'installation' && !showInstallationStations);
    const show = matchSearch && matchFilters && allowStation;

    if (show) {
      m.marker.addTo(map);
      visibleCount++;
    } else {
      map.removeLayer(m.marker);
    }
  });

  document.getElementById('stationCount').textContent = `แสดง: ${visibleCount} สถานี`;
  updateSectionLinesVisibility();
}

function resetFilters() {
  document.getElementById('searchBox').value = '';
  for (let colName in filterCheckboxes) {
    filterCheckboxes[colName].forEach(cb => {
      cb.checked = true;
    });
  }
  applyFilters();
}

function createPopupContent(rowData) {
  const noIndex = columnHeaders.indexOf(CONFIG.noColumn);
  const nameIndex = columnHeaders.indexOf(CONFIG.nameColumn);
  const latIndex = columnHeaders.indexOf(CONFIG.latColumn);
  const lngIndex = columnHeaders.indexOf(CONFIG.lngColumn);
  const dwgIndex = getColumnIndex(CONFIG.dwgUrlColumn);
  
  const stationNo = rowData[noIndex] || '-';
  const stationName = rowData[nameIndex] || 'ไม่ระบุชื่อ';
  const lat = parseFloat(rowData[latIndex]);
  const lng = parseFloat(rowData[lngIndex]);
  const dwgUrl = dwgIndex === -1 ? '' : rowData[dwgIndex];

  const streetViewUrl = `https://www.google.com/maps?q=&layer=c&cbll=${lat},${lng}`;
  let popupHTML = `
    <div class="popup-title-row">
      <div class="popup-title">🚉 ${stationName}</div>
      <a href="${streetViewUrl}" target="_blank" class="street-view-link">👤 Street View</a>
    </div>
  `;
  popupHTML += `<div class="popup-row">
    <span class="popup-label">หมายเลข:</span>
    <span class="popup-value">${stationNo}</span>
  </div>`;
  
  popupHTML += `<div class="popup-section">`;
  
  columnHeaders.forEach((header, index) => {
    if (index === noIndex || index === nameIndex || index === latIndex || index === lngIndex) return;
    if (CONFIG.excludeFromPopup.includes(header)) return;
    
    const value = rowData[index] || 'ไม่ระบุ';
    popupHTML += `
      <div class="popup-row">
        <span class="popup-label">${header}:</span>
        <span class="popup-value">${value}</span>
      </div>
    `;
  });
  
  popupHTML += `</div>`;

  popupHTML += `
    <div class="popup-row">
      <span class="popup-label">พิกัด:</span>
      <span class="popup-value">${lat.toFixed(6)}, ${lng.toFixed(6)}</span>
    </div>
  `;

  if (dwgUrl && dwgUrl !== '-') {
    popupHTML += `
      <a href="${dwgUrl}" target="_blank" rel="noopener" class="dwg-link">
        Installation DWG.
      </a>
    `;
  }

  return popupHTML;
}

function fetchSheetData(sheetName) {
  return fetch(`https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${sheetName}`)
    .then(res => res.text())
    .then(text => {
      const json = JSON.parse(text.substring(47).slice(0, -2));
      const headers = json.table.cols.map(col => col.label || col.id);
      const rows = json.table.rows.map(r => (r.c || []).map(cell => (cell ? cell.v : null)));
      return { headers, rows };
    });
}

loadCustomStatusColors();

Promise.all([fetchSheetData(SHEET_NAME), fetchSheetData(SECTION_SHEET_NAME)])
  .then(([stationSheet, sectionSheet]) => {
    columnHeaders = stationSheet.headers;
    sectionColumnHeaders = sectionSheet.headers;
    allSectionData = sectionSheet.rows;

    const nameIndex = getColumnIndex(CONFIG.nameColumn);
    const latIndex = getColumnIndex(CONFIG.latColumn);
    const lngIndex = getColumnIndex(CONFIG.lngColumn);
    const statusIndex = getColumnIndex(CONFIG.statusColumn);
    const regionIndex = getColumnIndex(CONFIG.regionColumn);
    const dwdmIndex = getColumnIndex(CONFIG.dwdmColumn, CONFIG.dwdmColumnIndex);
    const mplsIndex = getColumnIndex(CONFIG.mplsColumn);

    if (latIndex === -1 || lngIndex === -1) {
      alert('ไม่พบคอลัมน์พิกัด (Latitude/Longtitude) ใน Google Sheet');
      document.getElementById('loading').textContent = 'ข้อผิดพลาด: ไม่พบคอลัมน์พิกัด';
      return;
    }

    allData = [];
    allMarkers = [];
    allStationLabelMarkers.forEach(item => {
      if (item && item.marker) map.removeLayer(item.marker);
    });
    allStationLabelMarkers = [];

    stationSheet.rows.forEach(rowData => {
      const lat = parseFloat(rowData[latIndex]);
      const lng = parseFloat(rowData[lngIndex]);
      if (!lat || !lng || isNaN(lat) || isNaN(lng)) return;

      allData.push(rowData);

      const status = rowData[statusIndex];
      const region = regionIndex === -1 ? null : rowData[regionIndex];
      const dwdmType = dwdmIndex === -1 ? null : rowData[dwdmIndex];
      const mplsType = mplsIndex === -1 ? null : rowData[mplsIndex];
      const markerStyle = getMarkerStyle(status, currentStyle, region, dwdmType, mplsType);
      const latlng = [lat, lng];
      const popupContent = createPopupContent(rowData);
      const marker = createCircleMarker(latlng, markerStyle).addTo(map);

      marker.bindPopup(popupContent, {
        maxWidth: 420,
        className: 'custom-popup'
      });

      allMarkers.push({
        marker,
        markerType: 'circle',
        latlng,
        popupContent,
        label: nameIndex === -1 ? '' : String(rowData[nameIndex]),
        data: rowData
      });

      if (nameIndex !== -1) {
        const labelText = String(rowData[nameIndex] || '').trim();
        if (labelText) {
          const anchorIcon = L.divIcon({
            className: 'station-label-anchor',
            html: '',
            iconSize: [0, 0],
            iconAnchor: [0, 0]
          });
          const labelMarker = L.marker(latlng, {
            icon: anchorIcon,
            interactive: false,
            keyboard: false
          });
          labelMarker.bindTooltip(labelText, {
            permanent: true,
            direction: 'top',
            offset: [0, -10],
            className: 'marker-label'
          });
          allStationLabelMarkers.push({ marker: labelMarker });
        }
      }
    });

    createSectionLines(allSectionData);
    setupFilterSelector();
    updateLegend();
    applyLabelsToMarkers();
    updateGlobalStationLabels();

    document.getElementById('searchBox').oninput = applyFilters;
    const showLabelsCheckbox = document.getElementById('showLabels');
    if (showLabelsCheckbox) {
      showLabelsCheckbox.addEventListener('change', (e) => {
        showLabels = e.target.checked;
        applyLabelsToMarkers();
        updateGlobalStationLabels();
      });
    }

    document.querySelectorAll('input[name="styleMode"]').forEach(input => {
      input.addEventListener('change', (e) => {
        currentStyle = e.target.value;
        if (currentStyle === 'installation') {
          showInstallationStations = true;
          showInstallationLines = true;
          syncInstallationTogglesToUI();
        }
        applyStyleToMarkers();
        updateLegend();
        updateInstallationModeOptionsVisibility();
        applyFilters();
        updateSectionLinesVisibility();
      });
    });

    updateInstallationModeOptionsVisibility();
    syncInstallationTogglesToUI();
    applyFilters();
    document.getElementById('loading').style.display = 'none';
  })
  .catch(err => {
    console.error('Error loading data:', err);
    document.getElementById('loading').textContent = 'เกิดข้อผิดพลาดในการโหลดข้อมูล';
  });

const mapContainer = document.getElementById('map');
if (mapContainer) {
  mapContainer.addEventListener('click', () => setMenuOpen(false));
}
map.on('click', () => setMenuOpen(false));
window.addEventListener('resize', () => {
  if (window.innerWidth > 576) {
    setMenuOpen(false);
  }
});

const filterSelectAllBtn = document.getElementById('filterSelectAllBtn');
const filterSelectNoneBtn = document.getElementById('filterSelectNoneBtn');
const filterModalOverlay = document.getElementById('filterModalOverlay');
const filterModalClose = document.getElementById('filterModalClose');
const openFilterModal = document.getElementById('openFilterModal');
const updateDataBtn = document.getElementById('updateDataBtn');
const showInstallationStationsCheckbox = document.getElementById('showInstallationStations');
const showInstallationLinesCheckbox = document.getElementById('showInstallationLines');

if (filterSelectAllBtn && filterSelectNoneBtn) {
  filterSelectAllBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    selectedFilterColumns = new Set(allFilterColumns);
    buildFilterSelector();
    updateActiveFilterColumns();
  });

  filterSelectNoneBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    selectedFilterColumns = new Set();
    buildFilterSelector();
    updateActiveFilterColumns();
  });
}

if (openFilterModal && filterModalOverlay) {
  openFilterModal.addEventListener('click', () => {
    filterModalOverlay.classList.add('show');
    filterModalOverlay.setAttribute('aria-hidden', 'false');
  });
}

if (filterModalClose && filterModalOverlay) {
  filterModalClose.addEventListener('click', () => {
    filterModalOverlay.classList.remove('show');
    filterModalOverlay.setAttribute('aria-hidden', 'true');
  });
}

if (filterModalOverlay) {
  filterModalOverlay.addEventListener('click', (e) => {
    if (e.target === filterModalOverlay) {
      filterModalOverlay.classList.remove('show');
      filterModalOverlay.setAttribute('aria-hidden', 'true');
    }
  });
}

if (showInstallationStationsCheckbox) {
  showInstallationStationsCheckbox.addEventListener('change', (e) => {
    showInstallationStations = e.target.checked;
    applyFilters();
    updateLegend();
  });
}

if (showInstallationLinesCheckbox) {
  showInstallationLinesCheckbox.addEventListener('change', (e) => {
    showInstallationLines = e.target.checked;
    updateSectionLinesVisibility();
    updateLegend();
  });
}

if (updateDataBtn) {
  if (EXCEL_ONLINE_URL) {
    updateDataBtn.setAttribute('href', EXCEL_ONLINE_URL);
    updateDataBtn.setAttribute('target', '_blank');
    updateDataBtn.setAttribute('rel', 'noopener');
  }
  updateDataBtn.addEventListener('click', (e) => {
    if (!EXCEL_ONLINE_URL) {
      e.preventDefault();
      alert('ยังไม่ได้ใส่ลิงก์ Excel online');
    }
  });
}
