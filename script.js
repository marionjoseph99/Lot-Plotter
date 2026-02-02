function renderAutoPlotHighlight(activeRow) {
  const rows = Array.from(document.querySelectorAll('#lines .input-row'));
  rows.forEach(row => {
    if (row === activeRow) row.classList.add('input-row--active');
    else row.classList.remove('input-row--active');
  });
}

function gatherInputs() {
  const rows = Array.from(document.querySelectorAll('#lines .input-row'));
  return rows.map(row => ({
    row,
    ns: row.querySelector('.ns'),
    deg: row.querySelector('.deg'),
    min: row.querySelector('.min'),
    ew: row.querySelector('.ew'),
    len: row.querySelector('.length')
  }));
}

function bindAutoPlotInputs() {
  const entries = gatherInputs();
  entries.forEach(({ row, ns, deg, min, ew, len }) => {
    [ns, deg, min, ew, len].forEach(field => {
      if (!field || field._autoBound) return;
      const handler = () => {
        manualMirrorsImport = false;
        schedulePlot({ reason: 'auto' });
      };
      field.addEventListener('input', handler);
      field.addEventListener('change', handler);
      field.addEventListener('focus', () => renderAutoPlotHighlight(row));
      field._autoBound = true;
    });
  });
}
function parseBearing(bearing) {
  if (!bearing || typeof bearing !== 'string') return null;
  const re = /^\s*([NS])\s*([0-9]+(?:\.[0-9]+)?)°?\s*(?:(\d+(?:\.[0-9]+)?)'?)?\s*([EW])\s*$/i;
  const m = bearing.match(re);
  if (!m) return null;
  const ns = m[1].toUpperCase();
  const deg = parseFloat(m[2]);
  const min = m[3] ? parseFloat(m[3]) : 0;
  const ew = m[4].toUpperCase();
  const theta = deg + min / 60;
  let az;
  if (ns === 'N' && ew === 'E') az = theta;
  else if (ns === 'N' && ew === 'W') az = 360 - theta;
  else if (ns === 'S' && ew === 'E') az = 180 - theta;
  else if (ns === 'S' && ew === 'W') az = 180 + theta;
  else return null;
  return (az * Math.PI) / 180;
}

function computeArea(coords) {
  if (!coords || coords.length < 3) return 0;
  let sum = 0;
  const n = coords.length;
  for (let i = 0; i < n; i++) {
    const j = (i + 1) % n;
    sum += coords[i].x * coords[j].y - coords[j].x * coords[i].y;
  }
  return Math.abs(sum) / 2;
}

const EARTH_RADIUS_METERS = 6_371_000;

function degToRad(value) {
  return (value * Math.PI) / 180;
}

function radToDeg(value) {
  return (value * 180) / Math.PI;
}

function computeGeodesicStats(pointA, pointB) {
  if (!pointA || !pointB) {
    return {
      distanceMeters: 0,
      azimuth: 0,
      bearingString: formatQuadrantBearing(0)
    };
  }
  const lat1 = degToRad(pointA.lat);
  const lat2 = degToRad(pointB.lat);
  const lon1 = degToRad(pointA.lon);
  const lon2 = degToRad(pointB.lon);
  const dLat = lat2 - lat1;
  const dLon = lon2 - lon1;

  const sinHalfDLat = Math.sin(dLat / 2);
  const sinHalfDLon = Math.sin(dLon / 2);
  const haversine = sinHalfDLat * sinHalfDLat + Math.cos(lat1) * Math.cos(lat2) * sinHalfDLon * sinHalfDLon;
  const centralAngle = 2 * Math.atan2(Math.sqrt(haversine), Math.sqrt(Math.max(0, 1 - haversine)));
  const distanceMeters = EARTH_RADIUS_METERS * centralAngle;

  const y = Math.sin(dLon) * Math.cos(lat2);
  const x = Math.cos(lat1) * Math.sin(lat2) - Math.sin(lat1) * Math.cos(lat2) * Math.cos(dLon);
  let azimuth = radToDeg(Math.atan2(y, x));
  if (!isFinite(azimuth)) azimuth = 0;
  azimuth = (azimuth + 360) % 360;
  const bearingString = formatQuadrantBearing(azimuth);

  return {
    distanceMeters,
    azimuth,
    bearingString
  };
}

function normalizeGeoRing(points) {
  if (!Array.isArray(points)) return [];
  const ring = points.slice();
  if (ring.length > 1) {
    const first = ring[0];
    const last = ring[ring.length - 1];
    if (Math.abs(first.lat - last.lat) < 1e-9 && Math.abs(first.lon - last.lon) < 1e-9) {
      ring.pop();
    }
  }
  return ring;
}

function computeGeodesicPolygonArea(points) {
  const ring = normalizeGeoRing(points);
  if (ring.length < 3) return 0;
  let sum = 0;
  for (let i = 0; i < ring.length; i++) {
    const current = ring[i];
    const next = ring[(i + 1) % ring.length];
    const lon1 = degToRad(current.lon);
    const lon2 = degToRad(next.lon);
    const lat1 = degToRad(current.lat);
    const lat2 = degToRad(next.lat);
    let deltaLon = lon2 - lon1;
    if (deltaLon > Math.PI) deltaLon -= Math.PI * 2;
    else if (deltaLon < -Math.PI) deltaLon += Math.PI * 2;
    sum += deltaLon * (Math.sin(lat1) + Math.sin(lat2));
  }
  const area = Math.abs(sum) * (EARTH_RADIUS_METERS * EARTH_RADIUS_METERS / 2);
  return area;
}

function polygonCentroid(points) {
  if (!points || points.length === 0) return { x: 0, y: 0 };
  let areaAccumulator = 0;
  let cx = 0;
  let cy = 0;
  const n = points.length;
  for (let i = 0; i < n; i++) {
    const p0 = points[i];
    const p1 = points[(i + 1) % n];
    const cross = p0.x * p1.y - p1.x * p0.y;
    areaAccumulator += cross;
    cx += (p0.x + p1.x) * cross;
    cy += (p0.y + p1.y) * cross;
  }
  const area = areaAccumulator / 2;
  if (Math.abs(area) < 1e-6) {
    return {
      x: points.reduce((sum, p) => sum + p.x, 0) / n,
      y: points.reduce((sum, p) => sum + p.y, 0) / n
    };
  }
  return {
    x: cx / (6 * area),
    y: cy / (6 * area)
  };
}

function azimuthFromDelta(dx, dy) {
  return (Math.atan2(dx, dy) * 180 / Math.PI + 360) % 360;
}

function azimuthToQuadrant(az) {
  let ns = 'N';
  let ew = 'E';
  let theta = az;
  if (az >= 0 && az <= 90) {
    ns = 'N'; ew = 'E'; theta = az;
  } else if (az > 90 && az <= 180) {
    ns = 'S'; ew = 'E'; theta = 180 - az;
  } else if (az > 180 && az <= 270) {
    ns = 'S'; ew = 'W'; theta = az - 180;
  } else {
    ns = 'N'; ew = 'W'; theta = 360 - az;
  }
  const deg = Math.floor(theta + 1e-8);
  let minutes = Math.round((theta - deg) * 60);
  let adjDeg = deg;
  if (minutes === 60) {
    adjDeg += 1;
    minutes = 0;
  }
  return { ns, ew, deg: adjDeg, minutes };
}

function formatQuadrantBearing(az) {
  const { ns, ew, deg, minutes } = azimuthToQuadrant(az);
  return `${ns} ${deg}° ${minutes}' ${ew}`;
}

function formatDistanceMeters(value) {
  return `${value.toFixed(3)} m`;
}

function clearSvg(svg) {
  while (svg.firstChild) svg.removeChild(svg.firstChild);
}

function computeViewTransform(points, svgWidth, svgHeight, padding = 20) {
  if (!points || points.length === 0) {
    return {
      scale: 1,
      tx: padding,
      ty: padding
    };
  }
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  points.forEach(p => {
    if (p.x < minX) minX = p.x;
    if (p.x > maxX) maxX = p.x;
    if (p.y < minY) minY = p.y;
    if (p.y > maxY) maxY = p.y;
  });
  if (!isFinite(minX) || !isFinite(minY) || !isFinite(maxX) || !isFinite(maxY)) {
    return {
      scale: 1,
      tx: padding,
      ty: padding
    };
  }
  let width = maxX - minX;
  let height = maxY - minY;
  if (width === 0) width = 1;
  if (height === 0) height = 1;
  const scaleX = (svgWidth - padding * 2) / width;
  const scaleY = (svgHeight - padding * 2) / height;
  const scale = Math.min(scaleX, scaleY);
  const tx = -minX * scale + (svgWidth - width * scale) / 2;
  const ty = -minY * scale + (svgHeight - height * scale) / 2;
  return { scale, tx, ty };
}

function applyTransform(point, transform) {
  return {
    x: point.x * transform.scale + transform.tx,
    y: point.y * transform.scale + transform.ty
  };
}

function applyTransformToCoords(coords, transform) {
  return coords.map(c => applyTransform(c, transform));
}

function isPolygonClosed(points, tolerance = 1e-6) {
  if (!points || points.length < 2) return false;
  const first = points[0];
  const last = points[points.length - 1];
  return Math.hypot(first.x - last.x, first.y - last.y) <= tolerance;
}

const IMPORT_STROKE = '#fb923c';
const IMPORT_FILL = 'rgba(251, 146, 60, 0.18)';
const IMPORT_POINT_FILL = '#f97316';
const IMPORT_POINT_STROKE = '#fff7ed';

let importedLayers = [];
let importedFileName = '';
let manualMirrorsImport = false;

const SVG_NS = 'http://www.w3.org/2000/svg';
const FONT_FAMILY = 'Inter, "Segoe UI", sans-serif';
let statusHideHandle = null;

function updateStatus(message = '', tone = 'info') {
  const statusEl = document.getElementById('plotStatus');
  if (!statusEl) {
    if (tone === 'error' && message) console.error(message);
    return;
  }
  if (statusHideHandle) {
    clearTimeout(statusHideHandle);
    statusHideHandle = null;
  }
  if (!message) {
    statusEl.textContent = '';
    statusEl.classList.remove('is-visible');
    statusEl.removeAttribute('data-tone');
    return;
  }
  statusEl.textContent = message;
  statusEl.classList.add('is-visible');
  statusEl.setAttribute('data-tone', tone);
  const duration = tone === 'error' ? 6000 : 3500;
  statusHideHandle = window.setTimeout(() => {
    statusHideHandle = null;
    updateStatus('');
  }, duration);
}

function showError(msg, tone = 'error') {
  updateStatus(msg, tone);
  if (tone === 'error') console.error(msg);
}

function drawSvgBackground(svg) {
  const vb = svg.viewBox && svg.viewBox.baseVal ? svg.viewBox.baseVal : null;
  const width = vb && vb.width ? vb.width : svg.clientWidth || 500;
  const height = vb && vb.height ? vb.height : svg.clientHeight || 500;

  svg.setAttribute('xmlns', SVG_NS);

  const defs = document.createElementNS(SVG_NS, 'defs');
  svg.appendChild(defs);

  const bgRect = document.createElementNS(SVG_NS, 'rect');
  bgRect.setAttribute('x', '0');
  bgRect.setAttribute('y', '0');
  bgRect.setAttribute('width', width);
  bgRect.setAttribute('height', height);
  bgRect.setAttribute('fill', '#ffffff');
  svg.appendChild(bgRect);

  const borderRect = document.createElementNS(SVG_NS, 'rect');
  borderRect.setAttribute('x', '0.5');
  borderRect.setAttribute('y', '0.5');
  borderRect.setAttribute('width', Math.max(0, width - 1));
  borderRect.setAttribute('height', Math.max(0, height - 1));
  borderRect.setAttribute('fill', 'none');
  borderRect.setAttribute('stroke', '#ffffff');
  borderRect.setAttribute('stroke-width', '1');
  svg.appendChild(borderRect);
}

let lastPlottedCoords = null;
let lastTraverseTable = [];

const plotButton = document.getElementById('plot');
const autoCloseToggle = document.getElementById('autoClose');
const closureInfoEl = document.getElementById('closureInfo');
const importInputEl = document.getElementById('importKml');
const clearImportButton = document.getElementById('clearImport');
const AUTO_PLOT_DELAY = 260;
let autoPlotTimer = null;

function schedulePlot({ reason = 'auto', immediate = false } = {}) {
  if (autoPlotTimer) {
    window.clearTimeout(autoPlotTimer);
    autoPlotTimer = null;
  }
  const invoke = () => window.requestAnimationFrame(() => plotLot({ triggeredByAuto: reason === 'auto' }));
  if (immediate) {
    invoke();
  } else {
    autoPlotTimer = window.setTimeout(invoke, AUTO_PLOT_DELAY);
  }
}

function prepareManualPlot(rows, { triggeredByAuto }) {
  if (rows.length === 0) {
    return {
      valid: false,
      error: 'No lines to plot',
      errorType: 'empty',
      suppressError: triggeredByAuto
    };
  }

  const parsed = rows.map((r, idx) => {
    const nsEl = r.querySelector('.ns');
    const degEl = r.querySelector('.deg');
    const minEl = r.querySelector('.min');
    const ewEl = r.querySelector('.ew');
    const lenEl = r.querySelector('.length');
    const ns = nsEl ? nsEl.value : null;
    const deg = degEl && degEl.value !== '' ? parseFloat(degEl.value) : NaN;
    const min = minEl && minEl.value !== '' ? parseFloat(minEl.value) : 0;
    const ew = ewEl ? ewEl.value : null;
    const len = lenEl && lenEl.value !== '' ? parseFloat(lenEl.value) : NaN;
    return { ns, deg, min, ew, len, idx };
  });

  for (const p of parsed) {
    if (!p.ns || !p.ew) {
      return {
        valid: false,
        error: `Invalid hemisphere on line ${p.idx + 1}`,
        errorType: 'validation',
        suppressError: false
      };
    }
    if (!isFinite(p.deg)) {
      return {
        valid: false,
        error: `Invalid degrees on line ${p.idx + 1}`,
        errorType: 'validation',
        suppressError: triggeredByAuto
      };
    }
    if (p.deg < 0 || p.deg > 90) {
      return {
        valid: false,
        error: `Invalid degrees on line ${p.idx + 1}`,
        errorType: 'validation',
        suppressError: false
      };
    }
    if (!isFinite(p.min) || p.min < 0 || p.min >= 60) {
      return {
        valid: false,
        error: `Invalid minutes on line ${p.idx + 1}`,
        errorType: 'validation',
        suppressError: false
      };
    }
    if (!isFinite(p.len)) {
      return {
        valid: false,
        error: `Invalid length on line ${p.idx + 1}`,
        errorType: 'validation',
        suppressError: triggeredByAuto
      };
    }
    if (p.len <= 0) {
      return {
        valid: false,
        error: `Invalid length on line ${p.idx + 1}`,
        errorType: 'validation',
        suppressError: false
      };
    }
  }

  let x = 0;
  let y = 0;
  const coords = [{ x, y }];
  for (const p of parsed) {
    const theta = p.deg + (p.min || 0) / 60;
    let az;
    if (p.ns === 'N' && p.ew === 'E') az = theta;
    else if (p.ns === 'N' && p.ew === 'W') az = 360 - theta;
    else if (p.ns === 'S' && p.ew === 'E') az = 180 - theta;
    else if (p.ns === 'S' && p.ew === 'W') az = 180 + theta;
    else {
      return {
        valid: false,
        error: `Invalid bearing on line ${p.idx + 1}`,
        errorType: 'validation',
        suppressError: false
      };
    }
    const angle = (az * Math.PI) / 180;
    x += p.len * Math.sin(angle);
    y -= p.len * Math.cos(angle);
    coords.push({ x, y });
  }

  if (coords.length <= 1) {
    return {
      valid: false,
      error: 'No traverse data available',
      errorType: 'empty',
      suppressError: triggeredByAuto
    };
  }

  const totalDistance = parsed.reduce((sum, p) => sum + p.len, 0);
  const closureDx = coords[0].x - coords[coords.length - 1].x;
  const closureDy = coords[0].y - coords[coords.length - 1].y;
  const closureDistance = Math.hypot(closureDx, closureDy);
  const closureAzimuth = closureDistance > 1e-9 ? azimuthFromDelta(closureDx, -closureDy) : 0;
  const area = computeArea(coords);

  const segments = [];
  for (let i = 0; i < coords.length - 1; i++) {
    const a = coords[i];
    const b = coords[i + 1];
    const dx = b.x - a.x;
    const dy = a.y - b.y;
    const dist = Math.hypot(dx, dy);
    const az = azimuthFromDelta(dx, dy);
    const bearingStr = formatQuadrantBearing(az);
    segments.push({ from: i + 1, to: i + 2, bearing: bearingStr, distance: dist });
  }

  return {
    valid: true,
    coords,
    segments,
    area,
    totalDistance,
    closureDistance,
    closureAzimuth,
    error: null,
    errorType: 'none',
    suppressError: false
  };
}

function clearManualOutputs() {
  const tableWrap = document.getElementById('bearingTable');
  if (tableWrap) {
    tableWrap.innerHTML = '';
    tableWrap.style.display = 'none';
    tableWrap.setAttribute('aria-hidden', 'true');
  }
  if (closureInfoEl) {
    closureInfoEl.innerHTML = '';
    closureInfoEl.setAttribute('aria-hidden', 'true');
    closureInfoEl.classList.remove('has-warning');
  }
  lastPlottedCoords = null;
  lastTraverseTable = [];
}

function drawPointMarkersAndLabels(svg, points, centroid, {
  startIndex = 0,
  circleFill = '#4c9aff',
  circleStroke = '#e8edf4',
  circleStrokeWidth = 2,
  labelColor = '#0b2b36',
  labelStroke = '#ffffff',
  labelStrokeWidth = 3,
  labelPrefix = 'P'
} = {}) {
  if (!points || points.length === 0) return startIndex;

  const vb = svg.viewBox && svg.viewBox.baseVal ? svg.viewBox.baseVal : null;
  const viewWidth = vb && vb.width ? vb.width : svg.clientWidth || 500;
  const viewHeight = vb && vb.height ? vb.height : svg.clientHeight || 500;
  const margin = 18;
  const labelRadius = 26;
  const minLabelSpacing = 20;

  const labelPositions = [];
  const withinView = (x, y) => x >= margin && x <= viewWidth - margin && y >= margin && y <= viewHeight - margin;
  const isTooClose = (x, y) => labelPositions.some(pos => Math.hypot(pos.x - x, pos.y - y) < minLabelSpacing);

  let labelIndex = startIndex;

  points.forEach(p => {
    const circle = document.createElementNS(SVG_NS, 'circle');
    circle.setAttribute('cx', p.x);
    circle.setAttribute('cy', p.y);
    circle.setAttribute('r', '5');
    circle.setAttribute('fill', circleFill);
    circle.setAttribute('stroke', circleStroke);
    circle.setAttribute('stroke-width', String(circleStrokeWidth));
    svg.appendChild(circle);

    let dirX = p.x - centroid.x;
    let dirY = p.y - centroid.y;
    if (Math.abs(dirX) < 1e-6 && Math.abs(dirY) < 1e-6) {
      dirX = 0;
      dirY = -1;
    }
    const baseAngle = Math.atan2(dirY, dirX);
    const angleStep = Math.PI / 12;
    const angleAttempts = [0];
    for (let n = 1; n <= 8; n++) {
      angleAttempts.push(n * angleStep);
      angleAttempts.push(-n * angleStep);
    }

    let finalPosition = null;
    for (const delta of angleAttempts) {
      const angle = baseAngle + delta;
      const trialX = p.x + Math.cos(angle) * labelRadius;
      const trialY = p.y + Math.sin(angle) * labelRadius;
      if (!withinView(trialX, trialY)) continue;
      if (isTooClose(trialX, trialY)) continue;
      finalPosition = { x: trialX, y: trialY };
      break;
    }

    if (!finalPosition) {
      const fallbackAngle = baseAngle;
      let trialX = p.x + Math.cos(fallbackAngle) * labelRadius;
      let trialY = p.y + Math.sin(fallbackAngle) * labelRadius;
      trialX = Math.min(viewWidth - margin, Math.max(margin, trialX));
      trialY = Math.min(viewHeight - margin, Math.max(margin, trialY));
      finalPosition = { x: trialX, y: trialY };
    }

    labelPositions.push(finalPosition);

    labelIndex += 1;
    const text = document.createElementNS(SVG_NS, 'text');
    text.setAttribute('x', finalPosition.x);
    text.setAttribute('y', finalPosition.y);
    text.setAttribute('font-size', '12');
    text.setAttribute('font-weight', '600');
    text.setAttribute('font-family', FONT_FAMILY);
    text.setAttribute('fill', labelColor);
    text.setAttribute('text-anchor', 'middle');
    text.setAttribute('dominant-baseline', 'middle');
    text.setAttribute('paint-order', 'stroke');
    text.setAttribute('stroke', labelStroke);
    text.setAttribute('stroke-width', String(labelStrokeWidth));
    text.textContent = `${labelPrefix}${labelIndex}`;
    svg.appendChild(text);
  });

  return labelIndex;
}

function renderManualPlot(svg, manualData, transform, autoCloseEnabled) {
  const fitted = applyTransformToCoords(manualData.coords, transform);
  if (autoCloseEnabled && fitted.length >= 3) {
    const polygon = document.createElementNS(SVG_NS, 'polygon');
    polygon.setAttribute('points', fitted.map(c => `${c.x},${c.y}`).join(' '));
    polygon.setAttribute('fill', 'rgba(76, 154, 255, 0.18)');
    polygon.setAttribute('stroke', 'none');
    svg.appendChild(polygon);
  }

  const polyline = document.createElementNS(SVG_NS, 'polyline');
  polyline.setAttribute('points', fitted.map(c => `${c.x},${c.y}`).join(' '));
  polyline.setAttribute('fill', 'none');
  polyline.setAttribute('stroke', '#4c9aff');
  polyline.setAttribute('stroke-width', '2.5');
  polyline.setAttribute('stroke-linejoin', 'round');
  polyline.setAttribute('stroke-linecap', 'round');
  svg.appendChild(polyline);

  if (autoCloseEnabled && fitted.length > 1 && Math.hypot(fitted[0].x - fitted[fitted.length - 1].x, fitted[0].y - fitted[fitted.length - 1].y) > 0.5) {
    const closure = document.createElementNS(SVG_NS, 'line');
    closure.setAttribute('x1', fitted[fitted.length - 1].x);
    closure.setAttribute('y1', fitted[fitted.length - 1].y);
    closure.setAttribute('x2', fitted[0].x);
    closure.setAttribute('y2', fitted[0].y);
    closure.setAttribute('stroke', '#1d4ed8');
    closure.setAttribute('stroke-width', '2');
    closure.setAttribute('stroke-dasharray', '8 6');
    closure.setAttribute('stroke-linecap', 'round');
    closure.setAttribute('stroke-opacity', '0.75');
    svg.appendChild(closure);
  }

  const hasClosingPoint = fitted.length > 1 && Math.hypot(fitted[0].x - fitted[fitted.length - 1].x, fitted[0].y - fitted[fitted.length - 1].y) < 1e-3;
  const polygonPoints = hasClosingPoint ? fitted.slice(0, -1) : fitted.slice();
  const centroid = polygonCentroid(polygonPoints.length ? polygonPoints : fitted);

  const finalLabelIndex = drawPointMarkersAndLabels(svg, fitted, centroid, {
    startIndex: 0,
    circleFill: '#4c9aff',
    circleStroke: '#e8edf4',
    labelColor: '#0b2b36',
    labelStroke: '#ffffff'
  });

  if (autoCloseEnabled && polygonPoints.length >= 3) {
    const areaText = document.createElementNS(SVG_NS, 'text');
    areaText.setAttribute('x', centroid.x);
    areaText.setAttribute('y', centroid.y);
    areaText.setAttribute('font-size', '16');
    areaText.setAttribute('font-weight', '600');
    areaText.setAttribute('font-family', FONT_FAMILY);
    areaText.setAttribute('fill', '#0b2b36');
    areaText.setAttribute('text-anchor', 'middle');
    areaText.setAttribute('dominant-baseline', 'middle');
    areaText.setAttribute('paint-order', 'stroke');
    areaText.setAttribute('stroke', '#ffffff');
    areaText.setAttribute('stroke-width', '4');
    areaText.textContent = `Area: ${manualData.area.toFixed(2)} sqm`;
    svg.appendChild(areaText);
  }

  lastPlottedCoords = manualData.coords;

  updateClosureMetrics({
    enabled: autoCloseEnabled,
    closureDistance: manualData.closureDistance,
    closureAzimuth: manualData.closureAzimuth,
    totalDistance: manualData.totalDistance
  });

  return finalLabelIndex;
}

function getImportedTotals(autoCloseEnabled) {
  if (!importedLayers.length) {
    return {
      pointCount: 0,
      area: 0,
      areaText: ''
    };
  }
  const pointCount = importedLayers.reduce((sum, layer) => sum + layer.projectedCoords.length, 0);
  let area = 0;
  if (autoCloseEnabled) {
    importedLayers.forEach(layer => {
      const areaValue = layer.geodesicArea ?? computeGeodesicPolygonArea(layer.geoCoords);
      if (areaValue && areaValue > 0) {
        layer.geodesicArea = areaValue;
        area += areaValue;
      }
    });
  }
  const areaText = autoCloseEnabled && area > 0 ? `${area.toFixed(2)} sqm` : '';
  return { pointCount, area, areaText };
}

function updateImportInfoDisplay(autoCloseEnabled) {
  const infoEl = document.getElementById('importInfo');
  if (!infoEl) return;
  if (!importedLayers.length) {
    infoEl.innerHTML = '';
    infoEl.classList.remove('is-visible');
    infoEl.setAttribute('aria-hidden', 'true');
    return;
  }

  const totals = getImportedTotals(autoCloseEnabled);
  const areaDisplay = autoCloseEnabled
    ? (totals.area > 0 ? `${totals.area.toFixed(2)} sqm` : '0.00 sqm')
    : 'N/A (auto-close disabled)';

  infoEl.innerHTML = `
    <div><strong>Source:</strong> ${importedFileName || 'Unknown file'}</div>
    <div><strong>Groups:</strong> ${importedLayers.length}</div>
    <div><strong>Imported points:</strong> ${totals.pointCount}</div>
    <div><strong>Total area:</strong> ${areaDisplay}</div>
    <div><strong>Bearing source:</strong> Calculated from KML</div>
  `;
  infoEl.classList.add('is-visible');
  infoEl.setAttribute('aria-hidden', 'false');
}

function drawImportedLayers(svg, transform, autoCloseEnabled, startIndex = 0) {
  let labelIndex = startIndex;
  importedLayers.forEach(layer => {
    if (!layer.projectedCoords || layer.projectedCoords.length === 0) return;
    const transformed = applyTransformToCoords(layer.projectedCoords, transform);
    const pointsAttr = transformed.map(c => `${c.x},${c.y}`).join(' ');

    if (autoCloseEnabled && transformed.length >= 3) {
      const polygon = document.createElementNS(SVG_NS, 'polygon');
      polygon.setAttribute('points', pointsAttr);
      polygon.setAttribute('fill', IMPORT_FILL);
      polygon.setAttribute('stroke', 'none');
      svg.appendChild(polygon);
    }

    const polyline = document.createElementNS(SVG_NS, 'polyline');
    polyline.setAttribute('points', pointsAttr);
    polyline.setAttribute('fill', 'none');
    polyline.setAttribute('stroke', IMPORT_STROKE);
    polyline.setAttribute('stroke-width', '2.5');
    polyline.setAttribute('stroke-linejoin', 'round');
    polyline.setAttribute('stroke-linecap', 'round');
    polyline.setAttribute('stroke-dasharray', '6 4');
    svg.appendChild(polyline);

    const closed = isPolygonClosed(transformed, 0.5);
    if (autoCloseEnabled && transformed.length > 1 && !closed) {
      const closure = document.createElementNS(SVG_NS, 'line');
      closure.setAttribute('x1', transformed[transformed.length - 1].x);
      closure.setAttribute('y1', transformed[transformed.length - 1].y);
      closure.setAttribute('x2', transformed[0].x);
      closure.setAttribute('y2', transformed[0].y);
      closure.setAttribute('stroke', IMPORT_STROKE);
      closure.setAttribute('stroke-width', '2');
      closure.setAttribute('stroke-dasharray', '8 6');
      closure.setAttribute('stroke-opacity', '0.75');
      svg.appendChild(closure);
    }

    const polygonPoints = closed ? transformed.slice(0, -1) : transformed.slice();
    const centroidSource = polygonPoints.length >= 3 ? polygonPoints : transformed;
    const centroid = polygonCentroid(centroidSource);

    labelIndex = drawPointMarkersAndLabels(svg, transformed, centroid, {
      startIndex: labelIndex,
      circleFill: IMPORT_POINT_FILL,
      circleStroke: IMPORT_POINT_STROKE,
      labelColor: '#7c2d12',
      labelStroke: '#fff7ed'
    });

    if (autoCloseEnabled && centroidSource.length >= 3) {
      const areaValue = layer.geodesicArea ?? computeGeodesicPolygonArea(layer.geoCoords);
      if (areaValue && areaValue > 0) {
        layer.geodesicArea = areaValue;
      const areaLabel = document.createElementNS(SVG_NS, 'text');
      areaLabel.setAttribute('x', centroid.x);
      areaLabel.setAttribute('y', centroid.y);
      areaLabel.setAttribute('font-size', '14');
      areaLabel.setAttribute('font-weight', '600');
      areaLabel.setAttribute('font-family', FONT_FAMILY);
      areaLabel.setAttribute('fill', '#7c2d12');
      areaLabel.setAttribute('text-anchor', 'middle');
      areaLabel.setAttribute('dominant-baseline', 'middle');
      areaLabel.setAttribute('paint-order', 'stroke');
      areaLabel.setAttribute('stroke', '#fff7ed');
      areaLabel.setAttribute('stroke-width', '4');
        areaLabel.textContent = `Area: ${areaValue.toFixed(2)} sqm (KML)`;
        svg.appendChild(areaLabel);
      }
    }
  });
  return labelIndex;
}

function drawNorthArrow(svg) {
  if (!svg) return;

  const legacyNorthArrow = svg.querySelector('.north-arrow');
  if (legacyNorthArrow) legacyNorthArrow.remove();

  const container = svg.parentElement;
  if (!container) return;

  const existingOverlay = container.querySelector('.north-orientation');
  if (existingOverlay) existingOverlay.remove();

  const overlay = document.createElement('div');
  overlay.className = 'north-orientation';
  overlay.setAttribute('aria-hidden', 'true');
  overlay.innerHTML = `
    <span class="north-orientation__icon" aria-hidden="true"></span>
    <span class="north-orientation__label">N</span>
  `;

  container.appendChild(overlay);
}

function renderTraverseTable({ manualData = null, autoCloseEnabled = true }) {
  const tableWrap = document.getElementById('bearingTable');
  if (!tableWrap) return;

  const rows = [];
  if (manualData && manualData.segments && manualData.segments.length) {
    rows.push({ type: 'section', label: 'Manual traverse' });
    manualData.segments.forEach(seg => {
      rows.push({
        line: `P${seg.from}-P${seg.to}`,
        bearing: seg.bearing,
        distance: seg.distance,
        source: 'manual'
      });
    });
    if (autoCloseEnabled && manualData.closureDistance > 1e-6) {
      rows.push({
        line: 'Closure (Manual)',
        bearing: formatQuadrantBearing(manualData.closureAzimuth),
        distance: manualData.closureDistance,
        source: 'manual'
      });
    }
  }

  importedLayers.forEach((layer, layerIdx) => {
    if (!layer.geoCoords || layer.geoCoords.length < 2) return;
    const baseTitle = importedLayers.length > 1 ? `Imported group ${layerIdx + 1}` : 'Imported traverse';
    const suffix = layerIdx === 0 && importedFileName ? ` — ${importedFileName}` : '';
    rows.push({ type: 'section', label: `${baseTitle}${suffix}` });
    const segments = layer.geodesicSegments || [];
    segments.forEach((seg, segIdx) => {
      rows.push({
        line: `G${layerIdx + 1}-${segIdx + 1}`,
        bearing: seg.bearingString,
        distance: seg.distanceMeters,
        source: 'imported'
      });
    });
    if (autoCloseEnabled) {
      const closure = buildClosureSegmentGeo(layer.geoCoords);
      if (closure && closure.distanceMeters > 0.01) {
        rows.push({
          line: `Closure (G${layerIdx + 1})`,
          bearing: closure.bearingString,
          distance: closure.distanceMeters,
          source: 'imported'
        });
      }
    }
  });

  if (!rows.length) {
    tableWrap.innerHTML = '';
    tableWrap.style.display = 'none';
    tableWrap.setAttribute('aria-hidden', 'true');
    return;
  }

  let html = '<table><thead><tr><th>LINE</th><th>BEARING</th><th class="small">DIS (m)</th></tr></thead><tbody>';
  rows.forEach(row => {
    if (row.type === 'section') {
      html += `<tr class="table-section"><td colspan="3">${row.label}</td></tr>`;
    } else {
      const distanceValue = typeof row.distance === 'number' ? row.distance : 0;
      const distanceText = distanceValue.toFixed(2);
      html += `<tr><td>${row.line}</td><td>${row.bearing}</td><td class="small">${distanceText}</td></tr>`;
    }
  });
  html += '</tbody></table>';
  tableWrap.innerHTML = html;
  tableWrap.style.display = 'block';
  tableWrap.setAttribute('aria-hidden', 'false');

  lastTraverseTable = rows.map(row => {
    if (row.type === 'section') {
      return { type: 'section', label: row.label };
    }
    return {
      type: 'row',
      line: row.line,
      bearing: row.bearing,
      distance: row.distance,
      source: row.source || null
    };
  });
}

function parseKmlCoordinateString(text) {
  if (!text) return [];
  const rawPoints = text
    .trim()
    .split(/\s+/)
    .map(token => token.trim())
    .filter(Boolean);
  const coords = [];
  rawPoints.forEach(token => {
    const parts = token.split(',');
    if (parts.length < 2) return;
    const lon = parseFloat(parts[0]);
    const lat = parseFloat(parts[1]);
    if (!isFinite(lon) || !isFinite(lat)) return;
    const prev = coords[coords.length - 1];
    if (prev && Math.abs(prev.lon - lon) < 1e-9 && Math.abs(prev.lat - lat) < 1e-9) return;
    coords.push({ lat, lon });
  });
  return coords;
}

function extractCoordinateGroupsFromKml(doc) {
  const nodes = Array.from(doc.getElementsByTagName('coordinates'));
  const groups = nodes
    .map(node => parseKmlCoordinateString(node.textContent || ''))
    .filter(group => group.length >= 2);
  return groups;
}

function projectKmlGroups(groups) {
  if (!groups.length || !groups[0].length) return [];
  const origin = groups[0][0];
  const lat0 = origin.lat;
  const lon0 = origin.lon;
  const lat0Rad = (lat0 * Math.PI) / 180;
  const cosLat0 = Math.cos(lat0Rad);
  const safeCos = Math.abs(cosLat0) < 1e-6 ? 1e-6 : cosLat0;
  return groups.map(group => group.map(point => ({
    x: (point.lon - lon0) * 111320 * safeCos,
    y: -(point.lat - lat0) * 110540
  })));
}

function buildGeodesicSegmentsFromGeoCoords(points) {
  if (!points || points.length < 2) return [];
  const segments = [];
  for (let i = 0; i < points.length - 1; i++) {
    const a = points[i];
    const b = points[i + 1];
    const stats = computeGeodesicStats(a, b);
    segments.push({
      index: i,
      from: i + 1,
      to: i + 2,
      distanceMeters: stats.distanceMeters,
      azimuth: stats.azimuth,
      bearingString: stats.bearingString,
      fromCoord: a,
      toCoord: b
    });
  }
  return segments;
}

function buildClosureSegmentGeo(points) {
  if (!points || points.length < 2) return null;
  const first = points[0];
  const last = points[points.length - 1];
  if (!first || !last) return null;
  if (Math.abs(first.lat - last.lat) < 1e-9 && Math.abs(first.lon - last.lon) < 1e-9) return null;
  const stats = computeGeodesicStats(last, first);
  return {
    index: points.length - 1,
    from: points.length,
    to: 1,
    distanceMeters: stats.distanceMeters,
    azimuth: stats.azimuth,
    bearingString: stats.bearingString,
    fromCoord: last,
    toCoord: first,
    isClosure: true
  };
}

function handleKmlTextContent(text, fileName) {
  try {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(text, 'text/xml');
    if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
      throw new Error('Unable to parse KML file. Please check that the file is valid.');
    }
    const groups = extractCoordinateGroupsFromKml(xmlDoc);
    if (!groups.length) {
      throw new Error('No coordinate data found in the selected KML file.');
    }
    const projectedGroups = projectKmlGroups(groups);
    importedLayers = groups.map((group, idx) => ({
      geoCoords: group,
      projectedCoords: projectedGroups[idx],
      geodesicSegments: buildGeodesicSegmentsFromGeoCoords(group),
      geodesicArea: computeGeodesicPolygonArea(group)
    }));
    importedFileName = fileName || 'KML import';
    populateTraverseInputsFromImportedLayers({ announce: true });
    plotLot({ message: `Imported ${importedFileName}`, tone: 'success' });
  } catch (err) {
    console.error('KML import failed:', err);
    showError(err.message || 'Failed to import KML file');
  }
}

function handleKmlImport(file) {
  if (!file) return;
  if (!/\.kml$/i.test(file.name)) {
    showError('Please choose a .kml file to import.');
    return;
  }
  const reader = new FileReader();
  reader.onload = () => {
    const text = typeof reader.result === 'string' ? reader.result : '';
    handleKmlTextContent(text, file.name);
  };
  reader.onerror = () => {
    console.error('Error reading KML file', reader.error);
    showError('Could not read the selected KML file.');
  };
  reader.readAsText(file);
}

function clearImportedPlot({ silent = false, skipPlot = false } = {}) {
  if (!importedLayers.length) {
    if (!silent) updateStatus('No imported plot to clear', 'info');
    return;
  }
  importedLayers = [];
  importedFileName = '';
  manualMirrorsImport = false;
  updateImportInfoDisplay(autoCloseToggle ? autoCloseToggle.checked : true);
  if (!skipPlot) plotLot({ triggeredByAuto: true });
  if (!silent) updateStatus('Cleared imported plot', 'info');
}


function plotLot(options = {}) {
  const { message = null, tone = 'info', triggeredByAuto = false } = options;
  const svg = document.getElementById('lotCanvas');
  if (!svg) return;
  if (!triggeredByAuto) updateStatus();

  const autoCloseEnabled = autoCloseToggle ? autoCloseToggle.checked : true;
  const rows = Array.from(document.querySelectorAll('#lines .input-row'));
  const manualResult = prepareManualPlot(rows, { triggeredByAuto });
  const hasImported = importedLayers.length > 0;
  const manualIsMirrored = manualMirrorsImport && hasImported;
  const hasManual = manualResult.valid && manualResult.coords && manualResult.coords.length > 1 && !manualIsMirrored;

  if (manualResult.error && !(manualResult.suppressError || (hasImported && manualResult.errorType === 'empty'))) {
    showError(manualResult.error);
  }

  const vb = svg.viewBox && svg.viewBox.baseVal ? svg.viewBox.baseVal : null;
  const viewWidth = vb && vb.width ? vb.width : svg.clientWidth || 500;
  const viewHeight = vb && vb.height ? vb.height : svg.clientHeight || 500;

  if (!hasManual && !hasImported) {
    clearSvg(svg);
    drawSvgBackground(svg);
    drawNorthArrow(svg);
    clearManualOutputs();
    updateImportInfoDisplay(autoCloseEnabled);
    return;
  }

  const combinedPoints = [];
  if (hasManual) combinedPoints.push(...manualResult.coords);
  if (hasImported) {
    importedLayers.forEach(layer => {
      if (layer.projectedCoords && layer.projectedCoords.length) {
        combinedPoints.push(...layer.projectedCoords);
      }
    });
  }
  if (combinedPoints.length === 0) {
    combinedPoints.push({ x: 0, y: 0 });
  }

  const transform = computeViewTransform(combinedPoints, viewWidth, viewHeight, 20);

  clearSvg(svg);
  drawSvgBackground(svg);

  let labelOffset = 0;
  if (hasManual) {
    labelOffset = renderManualPlot(svg, manualResult, transform, autoCloseEnabled);
  } else {
    clearManualOutputs();
  }

  drawImportedLayers(svg, transform, autoCloseEnabled, labelOffset);
  drawNorthArrow(svg);
  updateImportInfoDisplay(autoCloseEnabled);
  renderTraverseTable({ manualData: hasManual ? manualResult : null, autoCloseEnabled });

  if (!triggeredByAuto) {
    const totals = getImportedTotals(autoCloseEnabled);
    let statusMessage = message || (hasManual ? 'Plot updated' : 'Imported plot ready');
    const manualAreaDetail =
      hasManual && autoCloseEnabled && manualResult.area > 0 ? ` — Area: ${manualResult.area.toFixed(2)} sqm` : '';
    const importedAreaDetail = hasImported && totals.areaText
      ? (hasManual ? ` · Imported area: ${totals.areaText}` : ` — Imported area: ${totals.areaText}`)
      : '';
    statusMessage += manualAreaDetail;
    statusMessage += importedAreaDetail;
    updateStatus(statusMessage, tone);
  }
}

function updateClosureMetrics({ enabled, closureDistance, closureAzimuth, totalDistance }) {
  if (!closureInfoEl) return;
  if (!enabled) {
    closureInfoEl.innerHTML = '<strong>Auto-close disabled.</strong> Closure metrics are hidden.';
    closureInfoEl.classList.remove('has-warning');
    closureInfoEl.setAttribute('aria-hidden', 'false');
    return;
  }

  const bearingText = closureDistance <= 1e-9 ? 'Perfect closure' : formatQuadrantBearing(closureAzimuth);
  const distanceText = closureDistance <= 1e-9 ? '0.000 m' : formatDistanceMeters(closureDistance);
  const ratioValue = closureDistance <= 1e-9 ? Infinity : totalDistance / closureDistance;
  const ratioText = ratioValue === Infinity ? '1:∞' : `1:${Math.round(ratioValue).toLocaleString()}`;
  const warning = ratioValue !== Infinity && ratioValue < 5000 ? '⚠️ Poor closure accuracy (≤ 1:5000)' : null;

  let html = '';
  html += `<div><strong>Closure Bearing:</strong> ${bearingText}</div>`;
  html += `<div><strong>Closure Distance:</strong> ${distanceText}</div>`;
  html += `<div><strong>Error Ratio:</strong> ${ratioText}</div>`;
  if (warning) {
    html += `<div class="closure-warning">${warning}</div>`;
    closureInfoEl.classList.add('has-warning');
  } else {
    closureInfoEl.classList.remove('has-warning');
  }
  closureInfoEl.innerHTML = html;
  closureInfoEl.setAttribute('aria-hidden', 'false');
}

if (plotButton && !plotButton._bound) {
  plotButton.addEventListener('click', () => plotLot());
  plotButton._bound = true;
}

if (importInputEl && !importInputEl._bound) {
  importInputEl.addEventListener('change', (event) => {
    const input = event.currentTarget;
    const file = input && input.files ? input.files[0] : null;
    if (file) {
      handleKmlImport(file);
    }
    if (input) input.value = '';
  });
  importInputEl._bound = true;
}

if (clearImportButton && !clearImportButton._bound) {
  clearImportButton.addEventListener('click', () => clearImportedPlot({ silent: false }));
  clearImportButton._bound = true;
}

function createInputRow(defaults = {}) {
  const {
    ns = 'N',
    deg = '',
    min = '',
    ew = 'E',
    len = ''
  } = defaults;
  const div = document.createElement('div');
  div.className = 'input-row';
  div.setAttribute('role', 'listitem');
  div.innerHTML = `
    <span class="segment-label" aria-hidden="true"></span>
    <select class="ns" aria-label="North or South">
      <option value="N" ${ns === 'N' ? 'selected' : ''}>N</option>
      <option value="S" ${ns === 'S' ? 'selected' : ''}>S</option>
    </select>
    <input type="number" class="deg" placeholder="°" min="0" max="90" aria-label="Degrees" value="${deg}">
    <input type="number" class="min" placeholder="'" min="0" max="59" aria-label="Minutes" value="${min}">
    <select class="ew" aria-label="East or West">
      <option value="E" ${ew === 'E' ? 'selected' : ''}>E</option>
      <option value="W" ${ew === 'W' ? 'selected' : ''}>W</option>
    </select>
    <input type="number" step="0.001" placeholder="Length" class="length" aria-label="Length" value="${len}">
    <button type="button" class="delete-btn" title="Delete line" aria-label="Delete segment">✕</button>
  `;
  return div;
}

function bindAddLineButton() {
  const btn = document.getElementById('addLine');
  if (!btn) return;
  if (btn._bound) return;
  btn.addEventListener('click', () => {
    const container = document.getElementById('lines');
    const row = createInputRow();
    container.appendChild(row);
    updateRowControls();
    bindAutoPlotInputs();
    schedulePlot({ reason: 'auto' });
    updateStatus('Added a new segment row', 'info');
  });
  btn._bound = true;
}

const A4_WIDTH_PX = 2480;
const A4_HEIGHT_PX = 3508;

function svgToPngDataUrl(svgEl, width = A4_WIDTH_PX, height = A4_HEIGHT_PX, tableRows = null) {
  return new Promise((resolve, reject) => {
    const serializer = new XMLSerializer();
    const svgString = serializer.serializeToString(svgEl);
    const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const iw = img.width;
      const ih = img.height;
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');

      const hasTable = Array.isArray(tableRows) && tableRows.length;
      const countContentHeight = (metrics) => {
        let total = 0;
        tableRows.forEach((row, idx) => {
          total += row.type === 'section' ? metrics.sectionHeight : metrics.rowHeight;
          if (idx < tableRows.length - 1) total += metrics.rowGap;
        });
        return total;
      };

      let tableLayout = null;

      if (hasTable) {
        const marginX = Math.max(36, Math.round(width * 0.03));
        const marginTop = Math.max(24, Math.round(height * 0.02));
        const marginBottom = Math.max(12, Math.round(height * 0.015));
        const maxPanelHeight = Math.round(height * 0.5);
        const cardWidth = width - marginX * 2;

        const baseMetrics = {
          paddingX: 48,
          paddingY: 28,
          headerHeight: 56,
          rowHeight: 36,
          sectionHeight: 40,
          rowGap: 12,
          radius: 22,
          headerOverhang: 18,
          headerFontSize: 18,
          bodyFontSize: 15,
          sectionFontSize: 16
        };

        const baseContentHeight = countContentHeight(baseMetrics);
        const baseCardHeight = baseMetrics.paddingY * 2 + baseMetrics.headerHeight + baseContentHeight;

        const availableCardHeight = Math.max(200, maxPanelHeight - marginTop - marginBottom);
        const scaleFactor = Math.min(1, availableCardHeight / baseCardHeight);

        const paddingX = baseMetrics.paddingX * scaleFactor;
        const paddingY = baseMetrics.paddingY * scaleFactor;
        const headerHeight = baseMetrics.headerHeight * scaleFactor;
        const rowHeight = baseMetrics.rowHeight * scaleFactor;
        const sectionHeight = baseMetrics.sectionHeight * scaleFactor;
        const rowGap = baseMetrics.rowGap * scaleFactor;
        const radius = baseMetrics.radius * scaleFactor;
        const headerOverhang = baseMetrics.headerOverhang * scaleFactor;
        const headerFontSize = Math.max(11, baseMetrics.headerFontSize * scaleFactor);
        const bodyFontSize = Math.max(10, baseMetrics.bodyFontSize * scaleFactor);
        const sectionFontSize = Math.max(11, baseMetrics.sectionFontSize * scaleFactor);

        let contentHeight = 0;
        tableRows.forEach((row, idx) => {
          contentHeight += row.type === 'section' ? sectionHeight : rowHeight;
          if (idx < tableRows.length - 1) contentHeight += rowGap;
        });

        const cardHeight = paddingY * 2 + headerHeight + contentHeight;
  let topPanelHeight = Math.round(cardHeight + marginTop + marginBottom);
  topPanelHeight = Math.min(maxPanelHeight, Math.max(topPanelHeight, marginTop + marginBottom + 40));

        const originX = marginX;
        const originY = marginTop;
        const contentWidth = cardWidth - paddingX * 2;
        const colLine = originX + paddingX;
        const colBearing = originX + paddingX + contentWidth * 0.45;
        const colDistance = originX + cardWidth - paddingX;

        tableLayout = {
          scaleFactor,
          topPanelHeight,
          originX,
          originY,
          cardWidth,
          cardHeight,
          paddingX,
          paddingY,
          headerHeight,
          rowHeight,
          sectionHeight,
          rowGap,
          radius,
          headerOverhang,
          headerFontSize,
          bodyFontSize,
          sectionFontSize,
          contentWidth,
          colLine,
          colBearing,
          colDistance
        };
      }

      const topPanelHeight = tableLayout ? tableLayout.topPanelHeight : Math.round(height * 0.32);
      const bottomPanelHeight = height - topPanelHeight;

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, width, height);

      ctx.fillStyle = '#f3f4ff';
      ctx.fillRect(0, 0, width, topPanelHeight);
      ctx.fillStyle = 'rgba(148, 163, 184, 0.22)';
      ctx.fillRect(0, topPanelHeight - 1, width, 2);

      const plotMarginX = Math.round(width * 0.08);
      const plotMarginY = Math.round(bottomPanelHeight * 0.1);
      const plotAvailWidth = width - plotMarginX * 2;
      const plotAvailHeight = bottomPanelHeight - plotMarginY * 2;
      const scale = Math.min(plotAvailWidth / iw, plotAvailHeight / ih);
      const plotWidth = iw * scale;
      const plotHeight = ih * scale;
      const plotX = plotMarginX + (plotAvailWidth - plotWidth) / 2;
      const plotY = topPanelHeight + plotMarginY + (plotAvailHeight - plotHeight) / 2;
      ctx.drawImage(img, plotX, plotY, plotWidth, plotHeight);

      if (tableLayout) {
        const theme = {
          panelBackground: '#ffffff',
          panelBorder: 'rgba(148, 163, 184, 0.35)',
          panelShadow: 'rgba(15, 23, 42, 0.12)',
          headerBackgroundStart: '#bfdbfe',
          headerBackgroundEnd: '#dbeafe',
          headerText: '#0f172a',
          sectionText: '#1d4ed8',
          rowText: '#0f172a',
          rowAltEven: 'rgba(191, 219, 254, 0.35)',
          rowAltOdd: 'rgba(226, 232, 240, 0.45)'
        };

        const {
          scaleFactor,
          originX,
          originY,
          cardWidth,
          cardHeight,
          paddingX,
          paddingY,
          headerHeight,
          rowHeight,
          sectionHeight,
          rowGap,
          radius,
          headerOverhang,
          headerFontSize,
          bodyFontSize,
          sectionFontSize,
          contentWidth,
          colLine,
          colBearing,
          colDistance
        } = tableLayout;

        const drawRoundedRect = (ctx2, x, y, w, h, r) => {
          const rr = Math.min(r, w / 2, h / 2);
          ctx2.beginPath();
          ctx2.moveTo(x + rr, y);
          ctx2.lineTo(x + w - rr, y);
          ctx2.quadraticCurveTo(x + w, y, x + w, y + rr);
          ctx2.lineTo(x + w, y + h - rr);
          ctx2.quadraticCurveTo(x + w, y + h, x + w - rr, y + h);
          ctx2.lineTo(x + rr, y + h);
          ctx2.quadraticCurveTo(x, y + h, x, y + h - rr);
          ctx2.lineTo(x, y + rr);
          ctx2.quadraticCurveTo(x, y, x + rr, y);
          ctx2.closePath();
        };

        ctx.save();
        ctx.shadowColor = theme.panelShadow;
  const shadowScale = scaleFactor || 1;
  ctx.shadowBlur = 24 * shadowScale;
        ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 18 * shadowScale;
        drawRoundedRect(ctx, originX, originY, cardWidth, cardHeight, radius);
        ctx.fillStyle = theme.panelBackground;
        ctx.fill();
        ctx.restore();

        ctx.save();
        drawRoundedRect(ctx, originX, originY, cardWidth, cardHeight, radius);
        ctx.strokeStyle = theme.panelBorder;
        ctx.lineWidth = 1.4;
        ctx.stroke();
        ctx.restore();

  const headerRectX = originX + paddingX - headerOverhang * 0.5;
  const headerRectY = originY + paddingY - headerOverhang;
  const headerRectWidth = contentWidth + headerOverhang;
  const headerRectHeight = headerHeight + headerOverhang * 2;
        ctx.save();
        const gradient = ctx.createLinearGradient(headerRectX, headerRectY, headerRectX + headerRectWidth, headerRectY);
        gradient.addColorStop(0, theme.headerBackgroundStart);
        gradient.addColorStop(1, theme.headerBackgroundEnd);
        ctx.fillStyle = gradient;
        drawRoundedRect(ctx, headerRectX, headerRectY, headerRectWidth, headerRectHeight, Math.max(10, 16 * scaleFactor));
        ctx.fill();
        ctx.restore();

        ctx.save();
        ctx.textBaseline = 'middle';
  ctx.font = `600 ${headerFontSize}px "Inter", "Segoe UI", Arial, sans-serif`;
        ctx.fillStyle = theme.headerText;
        const headerY = originY + paddingY + headerHeight / 2;
        ctx.textAlign = 'left';
        ctx.fillText('LINE', colLine, headerY);
        ctx.fillText('BEARING', colBearing, headerY);
        ctx.textAlign = 'right';
        ctx.fillText('DIS (M)', colDistance, headerY);

  ctx.strokeStyle = theme.panelBorder;
  ctx.lineWidth = 1;
        ctx.beginPath();
        const dividerY = originY + paddingY + headerHeight;
        ctx.moveTo(originX + paddingX, dividerY);
        ctx.lineTo(originX + cardWidth - paddingX, dividerY);
        ctx.stroke();

        let currentY = dividerY + rowGap;
        let bodyRowIndex = 0;
        tableRows.forEach((row, index) => {
          const isSection = row.type === 'section';
          const blockHeight = isSection ? sectionHeight : rowHeight;
          const baseline = currentY + blockHeight / 2;
          if (isSection) {
            ctx.font = `600 ${sectionFontSize}px "Inter", "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = theme.sectionText;
            ctx.textAlign = 'left';
            ctx.fillText(String(row.label || '').toUpperCase(), colLine, baseline);
          } else {
            const rowRectX = originX + paddingX;
            const rowRectWidth = contentWidth;
            ctx.fillStyle = bodyRowIndex % 2 === 0 ? theme.rowAltEven : theme.rowAltOdd;
            ctx.fillRect(rowRectX, currentY, rowRectWidth, blockHeight);

            const distanceValue = typeof row.distance === 'number' ? row.distance : parseFloat(row.distance) || 0;
            ctx.font = `500 ${bodyFontSize}px "Inter", "Segoe UI", Arial, sans-serif`;
            ctx.fillStyle = theme.rowText;
            ctx.textAlign = 'left';
            ctx.fillText(row.line || '', colLine, baseline);
            ctx.fillText(row.bearing || '', colBearing, baseline);
            ctx.textAlign = 'right';
            ctx.fillText(distanceValue.toFixed(2), colDistance, baseline);
            bodyRowIndex += 1;
          }
          currentY += blockHeight;
          if (index < tableRows.length - 1) currentY += rowGap;
        });
        ctx.restore();
      }
      URL.revokeObjectURL(url);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = (e) => {
      URL.revokeObjectURL(url);
      reject(e);
    };
    img.src = url;
  });
}

document.getElementById('exportPng').addEventListener('click', async () => {
  const svg = document.getElementById('lotCanvas');
  try {
  const dataUrl = await svgToPngDataUrl(svg, A4_WIDTH_PX, A4_HEIGHT_PX, lastTraverseTable);
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = 'lot.png';
    a.click();
    updateStatus('PNG export started — check your downloads folder', 'success');
  } catch (e) {
    showError('PNG export failed');
  }
});

document.getElementById('exportPdf').addEventListener('click', async () => {
  const svg = document.getElementById('lotCanvas');
  if (!window.jspdf) return showError('PDF library not loaded yet — try again in a moment');
  try {
  const dataUrl = await svgToPngDataUrl(svg, A4_WIDTH_PX, A4_HEIGHT_PX, lastTraverseTable);
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const img = new Image();
    img.src = dataUrl;
    await new Promise((res, rej) => { img.onload = res; img.onerror = rej; });
    const iw = img.width, ih = img.height;
    const scale = Math.min(pageW / iw, pageH / ih);
    const w = iw * scale, h = ih * scale;
    const x = (pageW - w) / 2, y = (pageH - h) / 2;
    doc.addImage(dataUrl, 'PNG', x, y, w, h);
    doc.save('lot.pdf');
    updateStatus('PDF export complete — saved as lot.pdf', 'success');
  } catch (e) {
    console.error(e);
    showError('PDF export failed');
  }
});

function formatBearingForScript(bearingString) {
  if (!bearingString || typeof bearingString !== 'string') return '';
  const clean = bearingString.trim();
  const match = clean.match(/^([NS])\s*(\d+)(?:°)?\s*(\d+)?(?:'|′)?\s*([EW])$/i);
  if (!match) {
    return clean.replace(/\s+/g, '');
  }
  const [, ns, degRaw, minRaw = '0', ew] = match;
  const deg = String(Math.abs(parseInt(degRaw, 10) || 0));
  const minutes = Math.abs(parseInt(minRaw, 10) || 0).toString().padStart(2, '0');
  return `${ns.toUpperCase()}${deg}d${minutes}'${ew.toUpperCase()}`;
}

function formatDistanceForScript(distance) {
  const value = Number(distance);
  if (!isFinite(value) || Math.abs(value) < 1e-9) return '0';
  return value.toFixed(3).replace(/\.?(?:0)+$/, '');
}

function appendSegmentsToScript(script, segments) {
  const filtered = segments.filter(seg => seg && isFinite(seg.distance) && Math.abs(seg.distance) > 1e-6 && seg.bearing);
  if (!filtered.length) return false;
  script.push('_.pline');
  script.push('0,0');
  filtered.forEach(seg => {
    const bearing = formatBearingForScript(seg.bearing);
    const distance = formatDistanceForScript(seg.distance);
    script.push(`@${distance}<${bearing}`);
  });
  script.push('');
  return true;
}

document.getElementById('exportScr').addEventListener('click', async () => {
  const rows = Array.from(document.querySelectorAll('#lines .input-row'));
  const autoCloseEnabled = autoCloseToggle ? autoCloseToggle.checked : true;
  let hasOutput = false;
  const script = [];

  const manualResult = prepareManualPlot(rows, { triggeredByAuto: true });
  if (manualResult && manualResult.valid && manualResult.segments && manualResult.segments.length) {
    const manualSegments = manualResult.segments.map(seg => ({
      distance: seg.distance,
      bearing: seg.bearing
    }));
    if (autoCloseEnabled && manualResult.closureDistance > 1e-6) {
      manualSegments.push({
        distance: manualResult.closureDistance,
        bearing: formatQuadrantBearing(manualResult.closureAzimuth)
      });
    }
    hasOutput = appendSegmentsToScript(script, manualSegments) || hasOutput;
  }

  importedLayers.forEach(layer => {
    if (!layer || !Array.isArray(layer.geodesicSegments) || !layer.geodesicSegments.length) return;
    const layerSegments = layer.geodesicSegments.map(seg => ({
      distance: seg.distanceMeters,
      bearing: seg.bearingString
    }));
    if (autoCloseEnabled) {
      const closure = buildClosureSegmentGeo(layer.geoCoords);
      if (closure && closure.distanceMeters > 1e-6) {
        layerSegments.push({
          distance: closure.distanceMeters,
          bearing: closure.bearingString
        });
      }
    }
    hasOutput = appendSegmentsToScript(script, layerSegments) || hasOutput;
  });

  if (!hasOutput) {
    showError('Nothing to export. Plot a traverse or import a KML first.');
    return;
  }

  const scriptText = script.join('\n');

  try {
    await navigator.clipboard.writeText(scriptText);
    updateStatus('Script copied to clipboard — paste in AutoCAD command line', 'success');
  } catch (e) {
    console.error(e);
    showError('Clipboard copy blocked — copy the script manually from the dialog');
    prompt('Copy the script below', scriptText);
  }
});

const STORAGE_KEY = 'lot_plots_v1';

function parseBearingFields(bearingString) {
  if (!bearingString || typeof bearingString !== 'string') return null;
  const match = bearingString.trim().match(/^([NS])\s*([0-9]+)(?:°)?\s*(\d+)?(?:'|′)?\s*([EW])$/i);
  if (!match) return null;
  const ns = match[1].toUpperCase();
  const deg = parseInt(match[2], 10);
  const min = match[3] ? parseInt(match[3], 10) : 0;
  const ew = match[4].toUpperCase();
  return { ns, deg, min, ew };
}

function populateTraverseInputsFromImportedLayers({ layerIndex = 0, announce = false } = {}) {
  if (!importedLayers.length) return;
  const layer = importedLayers[Math.min(layerIndex, importedLayers.length - 1)];
  if (!layer || !Array.isArray(layer.geodesicSegments) || !layer.geodesicSegments.length) return;
  const container = document.getElementById('lines');
  if (!container) return;
  container.innerHTML = '';
  layer.geodesicSegments.forEach(seg => {
    const bearingFields = parseBearingFields(seg.bearingString) || {};
    const len = isFinite(seg.distanceMeters) ? seg.distanceMeters.toFixed(3) : '';
    const row = createInputRow({
      ns: bearingFields.ns || 'N',
      deg: isFinite(bearingFields.deg) ? bearingFields.deg : '',
      min: isFinite(bearingFields.min) ? bearingFields.min : '',
      ew: bearingFields.ew || 'E',
      len
    });
    container.appendChild(row);
  });
  updateRowControls();
  bindAutoPlotInputs();
  schedulePlot({ reason: 'auto', immediate: true });
  manualMirrorsImport = true;
  if (announce) updateStatus('Imported traverse added to inputs', 'success');
}

function snapshotImportedLayersForSave() {
  if (!importedLayers.length) return [];
  return importedLayers.map(layer => ({
    geoCoords: Array.isArray(layer.geoCoords) ? layer.geoCoords : []
  })).filter(entry => Array.isArray(entry.geoCoords) && entry.geoCoords.length);
}

function restoreImportedLayersFromSnapshot(snapshot, sourceName = 'Saved import') {
  if (!Array.isArray(snapshot) || !snapshot.length) {
    importedLayers = [];
    importedFileName = '';
    return;
  }
  const groups = snapshot.map(entry => entry.geoCoords || []).filter(group => Array.isArray(group) && group.length);
  if (!groups.length) {
    importedLayers = [];
    importedFileName = '';
    return;
  }
  const projectedGroups = projectKmlGroups(groups);
  importedLayers = groups.map((group, idx) => ({
    geoCoords: group,
    projectedCoords: projectedGroups[idx] || [],
    geodesicSegments: buildGeodesicSegmentsFromGeoCoords(group),
    geodesicArea: computeGeodesicPolygonArea(group)
  }));
  importedFileName = sourceName || 'Saved import';
}

function loadSavedIndex() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    console.error('Failed to read storage', e);
    return {};
  }
}

function saveSavedIndex(index) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(index));
}

function refreshSavedList() {
  const sel = document.getElementById('savedList');
  sel.innerHTML = '';
  const idx = loadSavedIndex();
  Object.keys(idx).forEach(k => {
    const opt = document.createElement('option');
    opt.value = k;
    const record = idx[k];
    const tags = [];
    if (record && (record.manualCoords || record.coords)) tags.push('manual');
    if (record && record.importedLayers && record.importedLayers.length) tags.push('imported');
    const tagText = tags.length ? ` [${tags.join('+')}]` : '';
    opt.textContent = record.name + tagText + ' — ' + new Date(record.ts).toLocaleString();
    sel.appendChild(opt);
  });
}

document.getElementById('savePlot').addEventListener('click', () => {
  const rows = Array.from(document.querySelectorAll('#lines .input-row'));
  const manualResult = prepareManualPlot(rows, { triggeredByAuto: true });
  const manualCoords = manualResult && manualResult.valid ? manualResult.coords : null;
  const importedSnapshot = snapshotImportedLayersForSave();

  if (!manualCoords && !importedSnapshot.length) {
    return showError('Plot or import a KML before saving');
  }

  const name = document.getElementById('saveName').value.trim() || 'Saved Plot';
  const idx = loadSavedIndex();
  const id = 'plot_' + Date.now();
  idx[id] = {
    name,
    ts: Date.now(),
    manualCoords,
    importedLayers: importedSnapshot,
    importedFileName
  };
  saveSavedIndex(idx);
  refreshSavedList();
  updateStatus(`Saved plot "${name}"`, 'success');
});

document.getElementById('loadPlot').addEventListener('click', () => {
  const sel = document.getElementById('savedList');
  const key = sel.value;
  if (!key) return showError('Choose a saved plot to load');
  const idx = loadSavedIndex();
  const record = idx[key];
  if (!record) return showError('Saved plot not found');
  const container = document.getElementById('lines');
  container.innerHTML = '';
  const manualCoords = record.manualCoords || record.coords || null;
  restoreImportedLayersFromSnapshot(record.importedLayers, record.importedFileName || importedFileName);
  if (importedLayers.length) {
    populateTraverseInputsFromImportedLayers({ announce: true });
  } else if (manualCoords && manualCoords.length > 1) {
    for (let i = 0; i < manualCoords.length - 1; i++) {
      const a = manualCoords[i];
      const b = manualCoords[i + 1];
      const dx = b.x - a.x;
      const dy = a.y - b.y;
      const len = Math.hypot(dx, dy);
      const az = (Math.atan2(dx, dy) * 180 / Math.PI + 360) % 360;
      let ns = 'N', ew = 'E', theta = az;
      if (az >= 0 && az <= 90) { ns = 'N'; ew = 'E'; theta = az; }
      else if (az > 90 && az <= 180) { ns = 'S'; ew = 'E'; theta = 180 - az; }
      else if (az > 180 && az <= 270) { ns = 'S'; ew = 'W'; theta = az - 180; }
      else { ns = 'N'; ew = 'W'; theta = 360 - az; }
      const deg = Math.floor(Math.abs(theta));
      const min = Math.round((Math.abs(theta) - deg) * 60);
      const row = createInputRow({
        ns,
        deg,
        min,
        ew,
        len: len.toFixed(3)
      });
      container.appendChild(row);
    }
    updateRowControls();
    bindAutoPlotInputs();
    schedulePlot({ reason: 'auto', immediate: true });
  }
  plotLot({ message: `Loaded saved plot "${record.name}"`, tone: 'success' });
});

document.getElementById('deletePlot').addEventListener('click', () => {
  const sel = document.getElementById('savedList');
  const key = sel.value;
  if (!key) return showError('Choose a saved plot to delete');
  const idx = loadSavedIndex();
  const record = idx[key];
  if (!record) return showError('Saved plot not found');
  delete idx[key];
  saveSavedIndex(idx);
  refreshSavedList();
  updateStatus(`Deleted saved plot "${record.name}"`, 'info');
});

refreshSavedList();

function updateRowControls() {
  const rows = Array.from(document.querySelectorAll('#lines .input-row'));
  const addBtn = document.getElementById('addLine');
  if (addBtn) {
    addBtn.disabled = false;
    bindAddLineButton();
  }

  rows.forEach((row, idx) => {
    const segmentLabel = `${idx + 1}-${idx + 2}`;
    let labelEl = row.querySelector('.segment-label');
    if (!labelEl) {
      labelEl = document.createElement('span');
      labelEl.className = 'segment-label';
      labelEl.setAttribute('aria-hidden', 'true');
      row.insertBefore(labelEl, row.firstElementChild || null);
    }
    labelEl.textContent = segmentLabel;
    row.dataset.segment = segmentLabel;
    row.setAttribute('aria-label', `Segment ${idx + 1}: points ${idx + 1} to ${idx + 2}`);

    const del = row.querySelector('.delete-btn');
    if (!del) return;
    del.title = `Delete segment ${segmentLabel}`;
    del.setAttribute('aria-label', `Delete segment ${segmentLabel}`);
    if (!del._bound) {
      del.addEventListener('click', () => {
        row.remove();
        updateRowControls();
        bindAutoPlotInputs();
        schedulePlot({ reason: 'auto' });
        updateStatus('Removed a segment row', 'info');
      });
      del._bound = true;
    }
  });
}

updateRowControls();

function clearAll({ silent = false } = {}) {
  clearImportedPlot({ silent: true, skipPlot: true });
  const container = document.getElementById('lines');
  if (container) {
    container.innerHTML = '';
    container.appendChild(createInputRow());
    container.appendChild(createInputRow());
    container.appendChild(createInputRow());
    updateRowControls();
    bindAutoPlotInputs();
  }
  manualMirrorsImport = false;
  lastPlottedCoords = null;
  lastTraverseTable = [];
  const svg = document.getElementById('lotCanvas');
  if (svg) {
    clearSvg(svg);
    drawSvgBackground(svg);
  }
  plotLot({ triggeredByAuto: true });
  if (!silent) updateStatus('Cleared inputs and imported plot', 'info');
}

document.getElementById('clear').addEventListener('click', () => clearAll());

document.addEventListener('DOMContentLoaded', () => {
  updateRowControls();
  bindAddLineButton();
  bindAutoPlotInputs();
  if (autoCloseToggle) {
    autoCloseToggle.addEventListener('change', () => {
      schedulePlot({ reason: 'auto', immediate: true });
      updateStatus(autoCloseToggle.checked ? 'Auto-close enabled' : 'Auto-close disabled', 'info');
    });
  }
  updateImportInfoDisplay(autoCloseToggle ? autoCloseToggle.checked : true);
  plotLot({ message: 'Initial plot ready', tone: 'info', triggeredByAuto: true });
});

/* QR Code Modal Logic - Removed as donation feature is deprecated */
/* Header Popup Logic */
document.addEventListener('DOMContentLoaded', () => {
  const wrappers = [
    { id: 'feedbackWrapper', btnSelector: '.btn-report' },
    { id: 'versionWrapper', btnSelector: '.btn-version' }
  ];

  wrappers.forEach(config => {
    const wrapper = document.getElementById(config.id);
    if (!wrapper) return;
    const btn = wrapper.querySelector(config.btnSelector);
    
    if (!btn) return;

    // Toggle on click
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const isActive = wrapper.classList.contains('is-active');
      
      // Close all others first
      document.querySelectorAll('.button-with-help').forEach(el => {
        el.classList.remove('is-active');
        const b = el.querySelector('button');
        if (b) b.setAttribute('aria-expanded', 'false');
      });

      if (!isActive) {
        wrapper.classList.add('is-active');
        btn.setAttribute('aria-expanded', 'true');
      }
    });
  });

  // Close on outside click
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.button-with-help')) {
      document.querySelectorAll('.button-with-help').forEach(el => {
        el.classList.remove('is-active');
        const b = el.querySelector('button');
        if (b) b.setAttribute('aria-expanded', 'false');
      });
    }
  });

  // Close on Escape
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      document.querySelectorAll('.button-with-help').forEach(el => {
        el.classList.remove('is-active');
        const b = el.querySelector('button');
        if (b) b.setAttribute('aria-expanded', 'false');
      });
    }
  });
});
