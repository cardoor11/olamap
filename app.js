/* DoN Congressional Map — draft
 * Data:  data/map_data.xlsx (edited by hand)  +  geo/*.topojson (boundaries)
 * Libs:  MapLibre GL, topojson-client, SheetJS, Turf
 */
(() => {
  'use strict';

  // ---------------------------------------------------------------- constants
  const STATE_NAMES = {
    AL: 'Alabama', AK: 'Alaska', AZ: 'Arizona', AR: 'Arkansas', CA: 'California', CO: 'Colorado',
    CT: 'Connecticut', DE: 'Delaware', DC: 'District of Columbia', FL: 'Florida', GA: 'Georgia',
    HI: 'Hawaii', ID: 'Idaho', IL: 'Illinois', IN: 'Indiana', IA: 'Iowa', KS: 'Kansas', KY: 'Kentucky',
    LA: 'Louisiana', ME: 'Maine', MD: 'Maryland', MA: 'Massachusetts', MI: 'Michigan', MN: 'Minnesota',
    MS: 'Mississippi', MO: 'Missouri', MT: 'Montana', NE: 'Nebraska', NV: 'Nevada', NH: 'New Hampshire',
    NJ: 'New Jersey', NM: 'New Mexico', NY: 'New York', NC: 'North Carolina', ND: 'North Dakota',
    OH: 'Ohio', OK: 'Oklahoma', OR: 'Oregon', PA: 'Pennsylvania', RI: 'Rhode Island',
    SC: 'South Carolina', SD: 'South Dakota', TN: 'Tennessee', TX: 'Texas', UT: 'Utah', VT: 'Vermont',
    VA: 'Virginia', WA: 'Washington', WV: 'West Virginia', WI: 'Wisconsin', WY: 'Wyoming',
    AS: 'American Samoa', GU: 'Guam', MP: 'Northern Mariana Islands', PR: 'Puerto Rico', VI: 'U.S. Virgin Islands'
  };
  const TERRITORIES = new Set(['DC', 'AS', 'GU', 'MP', 'PR', 'VI']);
  const PARTY_NAME = { D: 'Democrat', R: 'Republican', I: 'Independent' };
  const GROUP_COLOR = { navy: '#1d4ed8', usmc: '#b91c1c', joint: '#7c3aed', reserve: '#0d9488', minor: '#64748b' };
  const GROUP_BADGE = { navy: 'N', usmc: 'MC', joint: 'J', reserve: 'R' };
  const CONUS = [[-125, 24.3], [-66.5, 49.5]];
  const REGIONS = {
    conus: CONUS,
    AK: [[-170, 51], [-129, 71.5]],
    HI: [[-160.6, 18.8], [-154.7, 22.3]],
    PR: [[-67.4, 17.6], [-64.5, 18.6]],
    GU: [[144.5, 13.2], [146.1, 15.4]],
    AS: [[-171.2, -14.6], [-168.1, -14.1]]
  };
  const HOVER_DELAY = 500;
  const IS_TOUCH = window.matchMedia('(hover: none)').matches;
  const HOUSE_CORE = ['District', 'State', 'Name', 'Party', 'Status'];
  const SENATE_CORE = ['State', 'Name', 'Party', 'Class', 'Status'];

  // ---------------------------------------------------------------- state
  const S = {
    chamber: 'house',
    house: new Map(),          // 'VA-2' -> row
    senate: new Map(),         // 'VA'   -> [row, row]
    districts: null,           // FeatureCollection
    districtById: new Map(),
    districtBbox: new Map(),
    states: null,
    stateByAbbr: new Map(),
    inst: [],                  // installation features (polygons or points)
    instById: new Map(),
    contractors: [],
    instDistricts: new Map(),  // instId -> [districtId]
    districtInsts: new Map(),  // districtId -> [instId]
    conDistricts: new Map(),   // contractorId -> [districtId]
    districtCons: new Map(),   // districtId -> [contractor feature]
    overlapDone: false,
    hover: null,               // {kind, id}
    hoverTimer: null,
    selected: null,
    toggles: { navy: true, usmc: true, joint: true, reserve: true, minor: false, contractors: true, basemap: false }
  };

  const $ = (s) => document.querySelector(s);
  const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const setStatus = (msg, cls = '') => { const el = $('#status'); el.textContent = msg; el.className = 'status ' + cls; };

  // ---------------------------------------------------------------- data loading
  async function fetchJSON(url) {
    const r = await fetch(url); if (!r.ok) throw new Error(`${url}: ${r.status}`); return r.json();
  }

  function readSheet(wb, name) {
    const ws = wb.Sheets[name]; if (!ws) return [];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
    // normalise headers: "District*" -> "District"
    return rows.map((r) => {
      const o = {};
      for (const [k, v] of Object.entries(r)) o[String(k).replace(/\*/g, '').trim()] = typeof v === 'string' ? v.trim() : v;
      return o;
    });
  }

  const groupOf = (branch, component) => /reserve/i.test(component || '') ? 'reserve'
    : /marine/i.test(branch) ? 'usmc' : /joint/i.test(branch) ? 'joint' : 'navy';
  const isNo = (v) => /^(no|n|false|0)$/i.test(String(v).trim());

  async function load() {
    const [dTopo, sTopo, iTopo, labels, xbuf] = await Promise.all([
      fetchJSON('geo/districts.topojson'),
      fetchJSON('geo/states.topojson'),
      fetchJSON('geo/installations.topojson'),
      fetchJSON('geo/district_labels.geojson'),
      fetch('data/map_data.xlsx', { cache: 'no-cache' }).then((r) => { if (!r.ok) throw new Error('map_data.xlsx missing'); return r.arrayBuffer(); })
    ]);

    S.districts = topojson.feature(dTopo, dTopo.objects.districts);
    S.states = topojson.feature(sTopo, sTopo.objects.states);
    const instGeo = topojson.feature(iTopo, iTopo.objects.installations);

    const wb = XLSX.read(xbuf, { type: 'array' });
    for (const r of readSheet(wb, 'House')) if (r.District) S.house.set(String(r.District).toUpperCase(), r);
    for (const r of readSheet(wb, 'Senate')) {
      if (!r.State) continue; const k = String(r.State).toUpperCase();
      if (!S.senate.has(k)) S.senate.set(k, []); S.senate.get(k).push(r);
    }

    for (const f of S.districts.features) {
      const id = f.properties.id; const rep = S.house.get(id);
      f.properties.party = rep ? (rep.Party || 'U') : 'X';
      S.districtById.set(id, f); S.districtBbox.set(id, turf.bbox(f));
    }
    for (const f of S.states.features) S.stateByAbbr.set(f.properties.st, f);

    // installations: geometry from topojson joined to sheet rows by ID; rows w/o geometry use lat/lon
    const geomById = new Map(instGeo.features.map((f) => [f.properties.id, f.geometry]));
    for (const r of readSheet(wb, 'Installations')) {
      if (!r.ID || !r.Name || isNo(r.Show)) continue; // Show = No rows are left off the site entirely
      let geometry = geomById.get(r.ID);
      const lat = parseFloat(r.Latitude), lon = parseFloat(r.Longitude);
      if (!geometry && isFinite(lat) && isFinite(lon)) geometry = { type: 'Point', coordinates: [lon, lat] };
      if (!geometry) continue;
      const f = {
        type: 'Feature', geometry,
        properties: {
          id: r.ID, name: r.Name, branch: r.Branch || 'Navy', category: r.Category || '', state: r.State || '',
          notes: r.Notes || '', group: groupOf(r.Branch || '', r.Component), minor: isNo(r.Show) ? 1 : 0,
          pointOnly: geometry.type === 'Point' ? 1 : 0, quality: r['Location Quality'] || '', row: r
        }
      };
      f.properties.lon = isFinite(lon) ? lon : turf.pointOnFeature(f).geometry.coordinates[0];
      f.properties.lat = isFinite(lat) ? lat : turf.pointOnFeature(f).geometry.coordinates[1];
      S.inst.push(f); S.instById.set(r.ID, f);
    }

    // reserve centers sitting inside an active base's outline are listed as tenants of that base (no separate dot)
    const hosts = S.inst.filter((f) => f.geometry.type !== 'Point' && f.properties.group !== 'reserve' && !f.properties.minor)
      .map((f) => ({ f, bb: turf.bbox(f) }));
    for (const f of S.inst) {
      f.properties.tenants = f.properties.tenants || [];
      if (f.properties.group !== 'reserve' || f.geometry.type !== 'Point') continue;
      const [lon, lat] = f.geometry.coordinates;
      const h = hosts.find(({ f: hf, bb }) => lon >= bb[0] && lon <= bb[2] && lat >= bb[1] && lat <= bb[3] && turf.booleanPointInPolygon([lon, lat], hf));
      if (h) { f.properties.host = h.f.properties.id; f.properties.onBase = 1; (h.f.properties.tenants ||= []).push(f.properties.id); }
    }

    for (const r of readSheet(wb, 'Contractors')) {
      const lat = parseFloat(r.Latitude), lon = parseFloat(r.Longitude);
      if (!(r.Company || r.Facility) || !isFinite(lat) || !isFinite(lon) || isNo(r.Show)) continue;
      const id = r.ID || `con-${S.contractors.length}`;
      S.contractors.push({ type: 'Feature', geometry: { type: 'Point', coordinates: [lon, lat] },
        properties: { id, name: r.Facility || r.Company, company: r.Company || '', category: r.Category || '', state: String(r.State || '').trim().toUpperCase(), row: r } });
    }

    return labels;
  }

  // ---------------------------------------------------------------- map
  const map = new maplibregl.Map({
    container: 'map',
    bounds: CONUS,
    fitBoundsOptions: { padding: 20 },
    minZoom: 1.5,
    maxZoom: 15.5,
    dragRotate: false,
    pitchWithRotate: false,
    touchPitch: false,
    attributionControl: { compact: true },
    style: {
      version: 8,
      glyphs: 'https://tiles.openfreemap.org/fonts/{fontstack}/{range}.pbf',
      sources: {
        basemap: {
          type: 'raster', tileSize: 256, maxzoom: 19,
          tiles: ['a', 'b', 'c', 'd'].map((s) => `https://${s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}@2x.png`),
          attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors © <a href="https://carto.com/attributions">CARTO</a>'
        }
      },
      layers: [
        { id: 'bg', type: 'background', paint: { 'background-color': '#ffffff' } },
        { id: 'basemap', type: 'raster', source: 'basemap', layout: { visibility: 'none' }, paint: { 'raster-opacity': 0.85 } }
      ]
    }
  });
  map.touchZoomRotate.disableRotation();
  map.addControl(new maplibregl.NavigationControl({ showCompass: false }), 'bottom-right');
  map.addControl(new maplibregl.ScaleControl({ unit: 'imperial' }), 'bottom-left');

  function addLayers(labels) {
    map.addSource('states', { type: 'geojson', data: S.states, promoteId: 'st' });
    map.addSource('districts', { type: 'geojson', data: S.districts, promoteId: 'id', tolerance: 0.3 });
    map.addSource('district-labels', { type: 'geojson', data: labels });
    map.addSource('inst', { type: 'geojson', data: { type: 'FeatureCollection', features: S.inst }, promoteId: 'id' });
    map.addSource('inst-pts', {
      type: 'geojson', promoteId: 'id',
      data: { type: 'FeatureCollection', features: S.inst.map((f) => ({ type: 'Feature', geometry: { type: 'Point', coordinates: [f.properties.lon, f.properties.lat] }, properties: { id: f.properties.id, group: f.properties.group, minor: f.properties.minor, name: f.properties.name, pointOnly: f.properties.pointOnly, onBase: f.properties.onBase || 0 } })) }
    });
    map.addSource('contractors', { type: 'geojson', data: { type: 'FeatureCollection', features: S.contractors }, promoteId: 'id' });

    const fs = (key, on, off) => ['case', ['boolean', ['feature-state', key], false], on, off];

    // --- states (senate hover target + outlines)
    map.addLayer({ id: 'state-fill', type: 'fill', source: 'states',
      paint: { 'fill-color': '#1f3864', 'fill-opacity': ['case', ['boolean', ['feature-state', 'sel'], false], 0.1, ['boolean', ['feature-state', 'hover'], false], 0.06, 0] } });

    // --- districts
    map.addLayer({ id: 'district-fill', type: 'fill', source: 'districts',
      paint: { 'fill-color': '#1f3864', 'fill-opacity': ['case', ['boolean', ['feature-state', 'sel'], false], 0.1, ['boolean', ['feature-state', 'hover'], false], 0.06, 0] } });
    map.addLayer({ id: 'district-line', type: 'line', source: 'districts',
      paint: {
        'line-color': '#9aa4b2',
        'line-width': ['interpolate', ['linear'], ['zoom'], 4, 0.4, 8, 0.9, 12, 1.6],
        'line-opacity': ['interpolate', ['linear'], ['zoom'], 3.6, 0, 5, 1]
      } });

    // --- state outlines on top of district lines
    map.addLayer({ id: 'state-line', type: 'line', source: 'states',
      paint: { 'line-color': '#3b4452', 'line-width': ['interpolate', ['linear'], ['zoom'], 3, 0.7, 8, 1.6, 12, 2.4] } });

    // --- installations (semi-translucent outlines)
    const instColor = ['match', ['get', 'group'], 'usmc', GROUP_COLOR.usmc, 'joint', GROUP_COLOR.joint, 'reserve', GROUP_COLOR.reserve, GROUP_COLOR.navy];
    const isRes = ['==', ['get', 'group'], 'reserve'];
    const keepPt = ['==', ['get', 'pointOnly'], 1]; // sites with no outline keep their dot at every zoom
    const instColorMinor = ['case', ['==', ['get', 'minor'], 1], GROUP_COLOR.minor, instColor];
    map.addLayer({ id: 'inst-fill', type: 'fill', source: 'inst', filter: ['==', ['geometry-type'], 'Polygon'],
      paint: { 'fill-color': instColorMinor, 'fill-opacity': ['interpolate', ['linear'], ['zoom'], 6, 0.45, 12, 0.28] } });
    map.addLayer({ id: 'inst-line', type: 'line', source: 'inst', filter: ['==', ['geometry-type'], 'Polygon'],
      paint: { 'line-color': instColorMinor, 'line-width': fs('sel', 3, 1.2), 'line-opacity': 0.9 } });
    map.addLayer({ id: 'inst-pt', type: 'circle', source: 'inst-pts',
      paint: {
        'circle-color': instColorMinor,
        'circle-radius': ['interpolate', ['linear'], ['zoom'], 3, ['case', isRes, 2.2, 2.6], 7, ['case', isRes, 3.8, 4.5], 12, ['case', isRes, 6, 6.5]],
        'circle-stroke-color': '#fff', 'circle-stroke-width': ['case', isRes, 0.8, 1],
        'circle-opacity': ['interpolate', ['linear'], ['zoom'], 8.5, 0.9, 10, ['case', keepPt, 0.95, 0]],
        'circle-stroke-opacity': ['interpolate', ['linear'], ['zoom'], 8.5, 1, 10, ['case', keepPt, 1, 0]]
      } });

    // --- contractors
    map.addLayer({ id: 'con-pt', type: 'circle', source: 'contractors',
      paint: { 'circle-color': '#d97706', 'circle-radius': ['interpolate', ['linear'], ['zoom'], 3, 2.6, 7, 4.5, 12, 6.5],
        'circle-stroke-color': '#fff', 'circle-stroke-width': 1 } });

    // --- highlighted outlines (selected/hovered district or state)
    map.addLayer({ id: 'district-hl', type: 'line', source: 'districts',
      paint: { 'line-color': '#1f3864', 'line-width': ['case', ['boolean', ['feature-state', 'sel'], false], 3, ['boolean', ['feature-state', 'hover'], false], 2, 0] } });
    map.addLayer({ id: 'state-hl', type: 'line', source: 'states', layout: { visibility: 'none' },
      paint: { 'line-color': '#1f3864', 'line-width': ['case', ['boolean', ['feature-state', 'sel'], false], 3.2, ['boolean', ['feature-state', 'hover'], false], 2.4, 0] } });

    // --- labels
    map.addLayer({ id: 'district-label', type: 'symbol', source: 'district-labels', minzoom: 5.5,
      layout: {
        'text-field': ['concat', ['get', 'st'], '-', ['get', 'n']],
        'text-font': ['Noto Sans Regular'],
        'text-size': ['interpolate', ['linear'], ['zoom'], 5.5, 10, 9, 13, 12, 15],
        'text-letter-spacing': 0.04, 'text-padding': 4
      },
      paint: { 'text-color': '#8590a0', 'text-halo-color': 'rgba(255,255,255,0.9)', 'text-halo-width': 1.4,
        'text-opacity': ['interpolate', ['linear'], ['zoom'], 5.5, 0, 6.2, 1] } });
    map.addLayer({ id: 'inst-label', type: 'symbol', source: 'inst-pts', minzoom: 9.5,
      layout: { 'text-field': ['get', 'name'], 'text-font': ['Noto Sans Regular'], 'text-size': 11.5,
        'text-max-width': 11, 'text-offset': [0, 0.2], 'text-anchor': 'top', 'text-optional': true },
      paint: { 'text-color': '#24324a', 'text-halo-color': '#fff', 'text-halo-width': 1.4 } });
    map.addLayer({ id: 'con-label', type: 'symbol', source: 'contractors', minzoom: 8,
      layout: { 'text-field': ['get', 'name'], 'text-font': ['Noto Sans Regular'], 'text-size': 11.5,
        'text-max-width': 11, 'text-offset': [0, 0.9], 'text-anchor': 'top', 'text-optional': true },
      paint: { 'text-color': '#8a4a00', 'text-halo-color': '#fff', 'text-halo-width': 1.4 } });

    applyToggles();
  }

  // ---------------------------------------------------------------- toggles / chamber
  function applyToggles() {
    const t = S.toggles;
    const groups = ['navy', 'usmc', 'joint', 'reserve'].filter((g) => t[g]);
    const f = ['all', ['in', ['get', 'group'], ['literal', groups]]];
    if (!t.minor) f.push(['==', ['get', 'minor'], 0]);
    for (const id of ['inst-fill', 'inst-line']) map.setFilter(id, ['all', ['==', ['geometry-type'], 'Polygon'], f]);
    const fPt = [...f, ['!=', ['get', 'onBase'], 1]]; // tenants of an active base have no dot of their own
    map.setFilter('inst-pt', fPt); map.setFilter('inst-label', fPt);
    const vis = (on) => (on ? 'visible' : 'none');
    map.setLayoutProperty('con-pt', 'visibility', vis(t.contractors));
    map.setLayoutProperty('con-label', 'visibility', vis(t.contractors));
    map.setLayoutProperty('basemap', 'visibility', vis(t.basemap));
  }

  function setChamber(ch) {
    if (S.chamber === ch) return;
    S.chamber = ch;
    document.querySelectorAll('.chamber button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.chamber === ch)));
    const house = ch === 'house';
    for (const id of ['district-fill', 'district-line', 'district-label', 'district-hl']) map.setLayoutProperty(id, 'visibility', house ? 'visible' : 'none');
    map.setLayoutProperty('state-hl', 'visibility', house ? 'none' : 'visible');
    clearHover(); clearSelection(); closePanel();
    $('#q').value = '';
    $('#q').placeholder = house ? 'Search a representative, district (VA-2), or installation' : 'Search a senator, state, or installation';
  }

  // ---------------------------------------------------------------- feature-state helpers
  function setFS(target, on) {
    if (!target) return;
    const src = target.kind === 'district' ? 'districts' : target.kind === 'state' ? 'states' : 'inst';
    map.setFeatureState({ source: src, id: target.id }, on);
  }
  function clearSelection() {
    if (S.selected) { setFS(S.selected, { sel: false }); S.selected = null; }
  }
  function select(kind, id) {
    clearSelection(); S.selected = { kind, id }; setFS(S.selected, { sel: true });
  }

  // ---------------------------------------------------------------- spatial overlap (installation <-> district)
  function bboxOverlap(a, b) { return a[0] <= b[2] && a[2] >= b[0] && a[1] <= b[3] && a[3] >= b[1]; }

  // District for a single point. Points just offshore (piers, coarse shoreline) snap to the nearest district within ~3 km.
  function districtsAtPoint(lon, lat) {
    const pt = turf.point([lon, lat]); const out = [];
    for (const [did, d] of S.districtById) if (bboxOverlap([lon, lat, lon, lat], S.districtBbox.get(did)) && turf.booleanPointInPolygon(pt, d)) out.push(did);
    if (out.length) return out;
    const pad = 0.03, box = [lon - pad, lat - pad, lon + pad, lat + pad];
    let best = null, bestKm = 3;
    for (const [did, d] of S.districtById) {
      if (!bboxOverlap(box, S.districtBbox.get(did))) continue;
      const ln = turf.polygonToLine(d);
      for (const l of ln.type === 'FeatureCollection' ? ln.features : [ln]) {
        const km = turf.pointToLineDistance(pt, l, { units: 'kilometers' }); if (km < bestKm) { bestKm = km; best = did; }
      }
    }
    return best ? [best] : [];
  }

  function computeContractorDistricts() {
    S.conDistricts.clear(); S.districtCons.clear();
    for (const f of S.contractors) {
      const [lon, lat] = f.geometry.coordinates;
      const ds = districtsAtPoint(lon, lat);
      S.conDistricts.set(f.properties.id, ds);
      for (const did of ds) { if (!S.districtCons.has(did)) S.districtCons.set(did, []); S.districtCons.get(did).push(f); }
    }
  }

  function computeOverlaps() {
    computeContractorDistricts();
    const ids = [...S.districtById.keys()];
    const queue = [...S.inst];
    const step = () => {
      const t0 = performance.now();
      while (queue.length && performance.now() - t0 < 30) {
        const f = queue.shift(); const bb = turf.bbox(f); const hits = [];
        for (const did of ids) {
          if (!bboxOverlap(bb, S.districtBbox.get(did))) continue;
          const d = S.districtById.get(did);
          let hit = false;
          try { hit = f.geometry.type === 'Point' ? turf.booleanPointInPolygon(f, d) : turf.booleanIntersects(f, d); }
          catch (e) { hit = turf.booleanPointInPolygon(turf.point([f.properties.lon, f.properties.lat]), d); }
          if (hit) hits.push(did);
        }
        S.instDistricts.set(f.properties.id, hits);
        for (const did of hits) { if (!S.districtInsts.has(did)) S.districtInsts.set(did, []); S.districtInsts.get(did).push(f.properties.id); }
      }
      if (queue.length) setTimeout(step, 0);
      else { S.overlapDone = true; refreshPanel(); }
    };
    step();
  }

  // ---------------------------------------------------------------- rendering helpers
  function roleFor(id) {
    const st = id.split('-')[0];
    if (st === 'PR') return 'Resident Commissioner';
    return TERRITORIES.has(st) ? 'Delegate' : 'Representative';
  }
  function districtLabel(id) {
    const [st, n] = id.split('-');
    const name = STATE_NAMES[st] || st;
    if (n === 'AL') return TERRITORIES.has(st) ? name : `${name} at-large`;
    const v = (+n) % 100, sfx = ['th', 'st', 'nd', 'rd'];
    return `${name} ${n}${sfx[(v - 20) % 10] || sfx[v] || sfx[0]} District`;
  }
  const isUncalled = (r) => !r || /uncalled/i.test(r.Status || '') || /^uncalled$/i.test(r.Name || '');
  function badge(r) {
    if (!r) return '<div class="badge X">–</div>';
    if (isUncalled(r)) return '<div class="badge U">?</div>';
    const p = String(r.Party || '').toUpperCase().slice(0, 1);
    return `<div class="badge ${['D', 'R', 'I'].includes(p) ? p : 'X'}">${esc(p || '–')}</div>`;
  }
  function extraKV(r, core) {
    const items = Object.entries(r).filter(([k, v]) => !core.includes(k) && v !== '' && v != null);
    if (!items.length) return '';
    return '<dl class="kv">' + items.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join('') + '</dl>';
  }

  function repHTML(id, { full = false } = {}) {
    const r = S.house.get(id);
    const unc = isUncalled(r);
    const name = !r ? 'No data' : unc ? 'Race uncalled' : r.Name;
    const party = r && !unc ? (PARTY_NAME[String(r.Party).toUpperCase()] || r.Party) : '';
    let h = `<div class="prof-head">${badge(r)}<div>
      <div class="prof-name">${esc(name)}</div>
      <div class="prof-meta">${esc(id)} · ${esc(districtLabel(id))}</div>
      <div class="prof-meta">${esc([party, roleFor(id)].filter(Boolean).join(' · '))}</div>
      ${r && r.Status ? `<span class="pill ${unc ? 'unc' : ''}">${esc(r.Status)}</span>` : ''}
    </div></div>`;
    if (r) h += extraKV(r, HOUSE_CORE);
    const insts = S.districtInsts.get(id) || [];
    const visible = insts.map((i) => S.instById.get(i)).filter((f) => f && instVisible(f));
    const active = visible.filter((f) => f.properties.group !== 'reserve');
    const reserve = visible.filter((f) => f.properties.group === 'reserve');
    h += instSection('DoN installations in district', active, full, 6,
      'None inside this district. Nearby sites are visible on the map.');
    if (S.toggles.reserve && reserve.length) h += instSection('Reserve centers in district', reserve, full, 3, '');
    if (S.toggles.contractors) h += conSection('Defense contractors in district', S.districtCons.get(id) || [], full, 4);
    if (!full) h += `<div class="hint">${IS_TOUCH ? '' : 'Click the district to pin details.'}</div>`;
    return h;
  }

  function instSection(title, list, full, limit, emptyMsg) {
    let h = `<div class="sect"><h3>${title}${list.length ? ` (${list.length})` : ''}</h3>`;
    if (!S.overlapDone) return h + '<div class="none">Calculating…</div></div>';
    if (!list.length) return h + `<div class="none">${emptyMsg}</div></div>`;
    const shown = full ? list : list.slice(0, limit);
    h += '<ul>' + shown.map((f) => `<li><span class="dot" style="background:${GROUP_COLOR[f.properties.minor ? 'minor' : f.properties.group]}"></span>${full ? `<button data-inst="${esc(f.properties.id)}">${esc(f.properties.name)}</button>` : esc(f.properties.name)}</li>`).join('') + '</ul>';
    if (!full && list.length > shown.length) h += `<div class="none">+ ${list.length - shown.length} more — click the district for the full list</div>`;
    return h + '</div>';
  }

  function conSection(title, list, full, limit) {
    if (!list.length) return '';
    let h = `<div class="sect"><h3>${title} (${list.length})</h3>`;
    const shown = full ? list : list.slice(0, limit);
    h += '<ul>' + shown.map((f) => {
      const p = f.properties, label = `${p.company}${p.name !== p.company ? ` — ${p.name}` : ''}`;
      return `<li><span class="dot" style="background:var(--con);border-radius:50%"></span>${full ? `<button data-con="${esc(p.id)}">${esc(label)}</button>` : esc(label)}</li>`;
    }).join('') + '</ul>';
    if (!full && list.length > shown.length) h += `<div class="none">+ ${list.length - shown.length} more — click for the full list</div>`;
    return h + '</div>';
  }

  function senatorBlock(r) {
    const unc = isUncalled(r);
    const party = !unc ? (PARTY_NAME[String(r.Party).toUpperCase()] || r.Party) : '';
    return `<div class="prof-head">${badge(r)}<div>
      <div class="prof-name">${esc(unc ? 'Race uncalled' : r.Name)}</div>
      <div class="prof-meta">${esc([party, 'Senator', r.Class ? `Class ${r.Class}` : ''].filter(Boolean).join(' · '))}</div>
      ${r.Status ? `<span class="pill ${unc ? 'unc' : ''}">${esc(r.Status)}</span>` : ''}
    </div></div>${extraKV(r, SENATE_CORE)}`;
  }

  function stateHTML(st, { full = false } = {}) {
    const sens = S.senate.get(st) || [];
    let h = `<div class="prof-meta" style="margin-bottom:8px;font-weight:600;color:var(--ink);font-size:14px">${esc(STATE_NAMES[st] || st)}</div>`;
    if (!sens.length) h += `<div class="none" style="color:var(--muted)">No U.S. senators. Represented in the House by a non-voting ${st === 'PR' ? 'Resident Commissioner' : 'delegate'}.</div>`;
    else h += '<div class="sen-pair">' + sens.map((r) => `<div>${senatorBlock(r)}</div>`).join('') + '</div>';
    const insts = S.inst.filter((f) => f.properties.state === st && instVisible(f));
    h += `<div class="sect"><h3>DoN installations in state (${insts.length})</h3>`;
    if (!insts.length) h += '<div class="none">None listed.</div>';
    else if (full) h += '<ul>' + insts.map((f) => `<li><span class="dot" style="background:${GROUP_COLOR[f.properties.minor ? 'minor' : f.properties.group]}"></span><button data-inst="${esc(f.properties.id)}">${esc(f.properties.name)}</button></li>`).join('') + '</ul>';
    else h += '<div class="none">Click the state for the full list.</div>';
    h += '</div>';
    if (S.toggles.contractors) {
      const cons = S.contractors.filter((f) => f.properties.state === st);
      if (cons.length && !full) h += `<div class="sect"><h3>Defense contractors in state (${cons.length})</h3></div>`;
      else h += conSection('Defense contractors in state', cons, full, 0);
    }
    return h;
  }

  function instHTML(id) {
    const f = S.instById.get(id); if (!f) return '';
    const p = f.properties;
    let h = `<div class="prof-head"><div class="badge" style="background:${GROUP_COLOR[p.minor ? 'minor' : p.group]};border-radius:8px">${GROUP_BADGE[p.group]}</div><div>
      <div class="prof-name">${esc(p.name)}</div>
      <div class="prof-meta">${esc([p.category, p.branch].filter(Boolean).join(' · '))}</div>
      <div class="prof-meta">${esc(STATE_NAMES[p.state] || p.state)}</div></div></div>`;
    const kv = [];
    if (p.notes) kv.push(['Notes', p.notes]);
    if (p.host && S.instById.get(p.host)) kv.push(['Located on', S.instById.get(p.host).properties.name]);
    if (kv.length) h += '<dl class="kv">' + kv.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join('') + '</dl>';
    const ten = (p.tenants || []).map((t) => S.instById.get(t)).filter(Boolean);
    if (ten.length) h += `<div class="sect"><h3>Reserve tenants on base (${ten.length})</h3><ul>` + ten.map((t) => `<li><span class="dot" style="background:${GROUP_COLOR.reserve};border-radius:50%"></span><button data-inst="${esc(t.properties.id)}">${esc(t.properties.name)}</button></li>`).join('') + '</ul></div>';
    const ds = S.instDistricts.get(id);
    h += `<div class="sect"><h3>Congressional district${ds && ds.length > 1 ? 's' : ''}</h3>`;
    if (!ds) h += '<div class="none">Calculating…</div>';
    else if (!ds.length) h += '<div class="none">Not inside any district boundary (offshore or outside the U.S.).</div>';
    else h += '<ul>' + ds.map((d) => {
      const r = S.house.get(d); const unc = isUncalled(r);
      return `<li><button data-district="${esc(d)}"><b>${esc(d)}</b> — ${esc(unc ? 'Uncalled' : r ? `${r.Name} (${r.Party})` : 'No data')}</button></li>`;
    }).join('') + '</ul>';
    h += '</div>';
    const sens = S.senate.get(p.state) || [];
    if (sens.length) {
      h += '<div class="sect"><h3>Senators</h3><ul>' + sens.map((r) => `<li>${esc(isUncalled(r) ? 'Uncalled' : `${r.Name} (${r.Party})`)}</li>`).join('') + '</ul></div>';
    }
    return h;
  }

  function instVisible(f) {
    const t = S.toggles; const p = f.properties;
    return t[p.group] && (t.minor || !p.minor);
  }

  // ---------------------------------------------------------------- panel
  let panelState = null;
  function openPanel(kind, id) {
    panelState = { kind, id };
    refreshPanel();
    $('#panel').hidden = false;
    $('#panel').scrollTop = 0;
  }
  function refreshPanel() {
    if (!panelState) return;
    const { kind, id } = panelState;
    $('#panelBody').innerHTML = kind === 'district' ? repHTML(id, { full: true })
      : kind === 'state' ? stateHTML(id, { full: true })
      : kind === 'inst' ? instHTML(id) : '';
  }
  function closePanel() { panelState = null; $('#panel').hidden = true; }

  $('#panelBody').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.inst) goInst(b.dataset.inst);
    if (b.dataset.con) goCon(b.dataset.con);
    if (b.dataset.district) goDistrict(b.dataset.district);
  });
  $('#panelClose').addEventListener('click', () => { closePanel(); clearSelection(); });

  // ---------------------------------------------------------------- navigation
  function panelPadding() {
    const mobile = window.innerWidth <= 760;
    return mobile ? { top: 30, bottom: Math.round(window.innerHeight * 0.5), left: 30, right: 30 }
                  : { top: 40, bottom: 40, left: 380, right: 40 };
  }
  function fit(bbox, maxZoom) {
    map.fitBounds([[bbox[0], bbox[1]], [bbox[2], bbox[3]]], { padding: panelPadding(), maxZoom, duration: 900 });
  }
  function goDistrict(id) {
    if (!S.districtById.has(id)) return;
    setChamber('house');
    select('district', id); openPanel('district', id);
    fit(S.districtBbox.get(id), 12);
  }
  function goState(st) {
    const f = S.stateByAbbr.get(st); if (!f) return;
    setChamber('senate');
    select('state', st); openPanel('state', st);
    fit(st === 'AK' ? [-170, 51, -129, 71.5] : turf.bbox(f), 9);
  }
  function goInst(id) {
    const f = S.instById.get(id); if (!f) return;
    const p = f.properties;
    if (!instVisible(f)) { // make sure it is shown
      S.toggles[p.group] = true; if (p.minor) S.toggles.minor = true;
      syncToggleInputs(); applyToggles();
    }
    clearSelection(); S.selected = { kind: 'inst', id }; setFS(S.selected, { sel: true });
    openPanel('inst', id);
    const bb = turf.bbox(f);
    if (f.geometry.type === 'Point') map.flyTo({ center: f.geometry.coordinates, zoom: 13, padding: panelPadding() });
    else fit(bb, 13.5);
  }

  // ---------------------------------------------------------------- hover (desktop)
  const card = $('#hovercard'), tip = $('#tip');
  let lastPt = null;

  function clearHover() {
    clearTimeout(S.hoverTimer);
    if (S.hover) setFS(S.hover, { hover: false });
    S.hover = null; card.hidden = true; tip.hidden = true;
  }
  function place(el, pt, dx = 16, dy = 16) {
    const mw = map.getContainer().clientWidth, mh = map.getContainer().clientHeight;
    const top0 = map.getContainer().offsetTop;
    const w = el.offsetWidth, h = el.offsetHeight;
    let x = pt.x + dx, y = pt.y + dy;
    if (x + w > mw - 8) x = pt.x - w - dx;
    if (y + h > mh - 8) y = Math.max(8, mh - h - 8);
    el.style.left = `${Math.max(8, x)}px`; el.style.top = `${y + top0}px`;
  }

  function onMove(e) {
    lastPt = e.point;
    // installation / contractor tooltip (immediate)
    const pts = map.queryRenderedFeatures(e.point, { layers: ['con-pt', 'inst-pt', 'inst-fill'] });
    if (pts.length) {
      const f = pts[0];
      if (f.layer.id === 'con-pt') {
        tip.innerHTML = `<b>${esc(f.properties.company)}</b><br>${esc(f.properties.name)}${f.properties.category ? ` · ${esc(f.properties.category)}` : ''}`;
      } else {
        const inst = S.instById.get(f.properties.id);
        tip.innerHTML = `<b>${esc(inst.properties.name)}</b><br>${esc(inst.properties.category)} · ${esc(inst.properties.branch)}`;
      }
      tip.hidden = false; place(tip, e.point, 14, -42);
      map.getCanvas().style.cursor = 'pointer';
    } else { tip.hidden = true; map.getCanvas().style.cursor = ''; }

    // district / state profile card (after 1 s dwell)
    const layer = S.chamber === 'house' ? 'district-fill' : 'state-fill';
    const hit = map.queryRenderedFeatures(e.point, { layers: [layer] })[0];
    const target = hit ? { kind: S.chamber === 'house' ? 'district' : 'state', id: hit.properties[S.chamber === 'house' ? 'id' : 'st'] } : null;
    if (!target) { if (S.hover) clearHover(); return; }
    if (S.hover && S.hover.id === target.id) {
      if (!card.hidden) place(card, e.point, 18, 18);
      return;
    }
    clearTimeout(S.hoverTimer); card.hidden = true;
    if (S.hover) setFS(S.hover, { hover: false });
    S.hover = target; setFS(target, { hover: true });
    S.hoverTimer = setTimeout(() => {
      if (!S.hover || S.hover.id !== target.id) return;
      card.innerHTML = target.kind === 'district' ? repHTML(target.id) : stateHTML(target.id);
      card.hidden = false; place(card, lastPt, 18, 18);
    }, HOVER_DELAY);
  }

  function onClick(e) {
    const inst = map.queryRenderedFeatures(e.point, { layers: ['inst-pt', 'inst-fill'] })[0];
    if (inst) { goInstNoFly(inst.properties.id); return; }
    const con = map.queryRenderedFeatures(e.point, { layers: ['con-pt'] })[0];
    if (con) { showContractor(con.properties.id); return; }
    const layer = S.chamber === 'house' ? 'district-fill' : 'state-fill';
    const hit = map.queryRenderedFeatures(e.point, { layers: [layer] })[0];
    if (!hit) { closePanel(); clearSelection(); return; }
    card.hidden = true;
    if (S.chamber === 'house') { select('district', hit.properties.id); openPanel('district', hit.properties.id); }
    else { select('state', hit.properties.st); openPanel('state', hit.properties.st); }
  }
  function goInstNoFly(id) {
    clearSelection(); S.selected = { kind: 'inst', id }; setFS(S.selected, { sel: true }); openPanel('inst', id);
  }
  function showContractor(id) {
    const f = S.contractors.find((c) => c.properties.id === id); if (!f) return;
    const p = f.properties, r = p.row;
    const ds = S.conDistricts.get(id) || districtsAtPoint(...f.geometry.coordinates);
    panelState = null; clearSelection();
    const kv = [['Products / programs', r['Products / Programs']], ['Category', p.category],
      ['Notes', r.Notes]].filter(([, v]) => v);
    const shown = ['ID', 'Company', 'Facility', 'Products / Programs', 'Category', 'State', 'Latitude', 'Longitude', 'Location Precision', 'Show', 'Notes'];
    $('#panelBody').innerHTML = `<div class="prof-head"><div class="badge" style="background:var(--con);border-radius:50%">C</div><div>
      <div class="prof-name">${esc(p.name)}</div><div class="prof-meta">${esc(p.company)}</div>
      <div class="prof-meta">${esc(STATE_NAMES[p.state] || p.state)}</div></div></div>
      <dl class="kv">${kv.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join('')}</dl>
      ${extraKV(r, shown)}
      <div class="sect"><h3>Congressional district</h3><ul>${ds.map((d) => { const rr = S.house.get(d); return `<li><button data-district="${esc(d)}"><b>${esc(d)}</b> — ${esc(isUncalled(rr) ? 'Uncalled' : `${rr.Name} (${rr.Party})`)}</button></li>`; }).join('') || '<div class="none">Unknown</div>'}</ul></div>`;
    $('#panel').hidden = false;
  }
  function goCon(id) {
    const f = S.contractors.find((c) => c.properties.id === id); if (!f) return;
    if (!S.toggles.contractors) { S.toggles.contractors = true; syncToggleInputs(); applyToggles(); }
    map.flyTo({ center: f.geometry.coordinates, zoom: Math.max(map.getZoom(), 11), padding: panelPadding() });
    showContractor(id);
  }

  // ---------------------------------------------------------------- search
  let index = [];
  function buildIndex() {
    index = [];
    for (const [id, f] of S.districtById) {
      const r = S.house.get(id); const unc = isUncalled(r);
      const [st, n] = id.split('-');
      index.push({ type: 'rep', key: id,
        title: unc ? `${id} — Uncalled` : `${r ? r.Name : 'No data'}`,
        sub: `${id} · ${unc ? 'race uncalled' : (PARTY_NAME[String(r?.Party).toUpperCase()] || r?.Party || '')} · ${STATE_NAMES[st] || st}`,
        text: [r && !unc ? r.Name : '', id, `${st} ${n}`, `${STATE_NAMES[st]} ${n}`, r?.Notes || ''].join(' | ').toLowerCase() });
    }
    for (const [st, rows] of S.senate) for (const r of rows) {
      const unc = isUncalled(r);
      index.push({ type: 'sen', key: st, title: unc ? `${STATE_NAMES[st]} Senate seat — Uncalled` : r.Name,
        sub: ['Senator', unc ? 'race uncalled' : (PARTY_NAME[String(r.Party).toUpperCase()] || r.Party), STATE_NAMES[st]].filter(Boolean).join(' · '),
        text: [unc ? '' : r.Name, STATE_NAMES[st], st, r.Notes || ''].join(' | ').toLowerCase() });
    }
    for (const st of S.stateByAbbr.keys()) index.push({ type: 'state', key: st, title: STATE_NAMES[st] || st, sub: 'State — view senators', text: `${(STATE_NAMES[st] || st).toLowerCase()} | ${st.toLowerCase()}` });
    for (const f of S.inst) {
      const p = f.properties;
      index.push({ type: p.group === 'reserve' ? 'res' : 'inst', key: p.id, title: p.name, sub: `${p.category} · ${STATE_NAMES[p.state] || p.state}${p.minor ? ' · minor site' : ''}`,
        text: [p.name, p.row['Source Name'] || '', p.category, p.state].join(' | ').toLowerCase(), minor: p.minor });
    }
    for (const f of S.contractors) index.push({ type: 'con', key: f.properties.id, title: f.properties.company && !f.properties.name.includes(f.properties.company) ? `${f.properties.company} — ${f.properties.name}` : f.properties.name, sub: ['Contractor', f.properties.category, STATE_NAMES[f.properties.state] || f.properties.state].filter(Boolean).join(' · '), text: [f.properties.name, f.properties.company, f.properties.row['Products / Programs'] || '', f.properties.state].join(' | ').toLowerCase() });
  }

  const KIND = { rep: ['Rep', '#1f3864'], sen: ['Sen', '#4b5563'], state: ['State', '#64748b'], inst: ['DoN', '#1d4ed8'], res: ['Res', '#0d9488'], con: ['Ind', '#d97706'] };
  // "representative", "rep", "congressman", "senator", "sen" … narrow results to that chamber;
  // on their own they list every member of it.
  const ROLE_WORDS = [['representatives', 'rep'], ['representative', 'rep'], ['congressman', 'rep'], ['congresswoman', 'rep'],
    ['congressmember', 'rep'], ['rep.', 'rep'], ['rep', 'rep'], ['delegate', 'rep'], ['senators', 'sen'], ['senator', 'sen'], ['sen.', 'sen'], ['sen', 'sen']];
  function roleFilter(q) {
    const [w, ...rest] = q.split(/\s+/);
    for (const [word, type] of ROLE_WORDS) {
      const full = word.replace('.', '');
      // whole word, or a partial (4+ letters) of the long forms while still typing it
      if (w === word || (!rest.length && w.length >= 4 && type && !word.startsWith('del') && full.startsWith(w))) return { type, rest: rest.join(' ') };
    }
    return null;
  }
  function search(q) {
    q = q.trim().toLowerCase().replace(/\s*-\s*/g, '-');
    if (!q) return [];
    const role = roleFilter(q);
    if (role) {
      const pool = index.filter((it) => it.type === role.type);
      if (!role.rest) return pool.slice().sort((a, b) => role.type === 'rep' ? a.key.localeCompare(b.key, 'en', { numeric: true }) : a.key.localeCompare(b.key) || a.title.localeCompare(b.title));
      const saved = index; index = pool;
      try { return search(role.rest); } finally { index = saved; }
    }
    const q2 = q.replace(/-/g, ' ');
    const out = [];
    for (const it of index) {
      let s = 0; const title = it.title.toLowerCase();
      if (title === q || it.key.toLowerCase() === q) s = 100;
      else if (title.startsWith(q)) s = 80;
      else if (new RegExp(`(^|[\\s(|–-])${q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`).test(it.text)) s = 60;
      else if (it.text.includes(q) || it.text.includes(q2)) s = 40;
      if (!s) continue;
      if (s < 80 && title.includes(q)) s += 15; // name match beats a match in notes/products
      if (it.type === 'rep' && S.chamber === 'house') s += 5;
      if ((it.type === 'sen' || it.type === 'state') && S.chamber === 'senate') s += 5;
      if (it.minor) s -= 10;
      out.push([s, it]);
    }
    out.sort((a, b) => b[0] - a[0] || a[1].title.localeCompare(b[1].title));
    return out.slice(0, 10).map((x) => x[1]);
  }

  const input = $('#q'), list = $('#suggest');
  let results = [], active = -1;
  function renderSuggest() {
    if (!input.value.trim()) { list.hidden = true; return; }
    list.innerHTML = results.length ? results.map((it, i) => `<li role="option" data-i="${i}" aria-selected="${i === active}">
        <span class="kind" style="background:${KIND[it.type][1]}">${KIND[it.type][0]}</span>
        <span><div class="t">${esc(it.title)}</div><div class="s">${esc(it.sub)}</div></span></li>`).join('')
      : `<li class="empty">${index.length ? 'No matches' : 'Still loading data…'}</li>`;
    list.hidden = false;
  }
  function choose(it) {
    if (!it) return;
    input.value = it.title; list.hidden = true; input.blur();
    if (it.type === 'rep') goDistrict(it.key);
    else if (it.type === 'sen' || it.type === 'state') goState(it.key);
    else if (it.type === 'inst' || it.type === 'res') goInst(it.key);
    else if (it.type === 'con') goCon(it.key);
  }
  input.addEventListener('input', () => { results = search(input.value); active = results.length ? 0 : -1; renderSuggest(); });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') { active = Math.min(results.length - 1, active + 1); renderSuggest(); e.preventDefault(); }
    else if (e.key === 'ArrowUp') { active = Math.max(0, active - 1); renderSuggest(); e.preventDefault(); }
    else if (e.key === 'Enter') { choose(results[active]); }
    else if (e.key === 'Escape') { list.hidden = true; }
  });
  input.addEventListener('focus', () => { if (input.value.trim()) { results = search(input.value); renderSuggest(); } });
  list.addEventListener('mousedown', (e) => { const li = e.target.closest('li[data-i]'); if (li) { e.preventDefault(); choose(results[+li.dataset.i]); } });
  document.addEventListener('click', (e) => { if (!e.target.closest('.search')) list.hidden = true; });

  // ---------------------------------------------------------------- UI wiring
  document.querySelectorAll('.chamber button').forEach((b) => b.addEventListener('click', () => setChamber(b.dataset.chamber)));
  $('#layersBtn').addEventListener('click', () => {
    const open = $('#layers').hidden; $('#layers').hidden = !open; $('#layersBtn').setAttribute('aria-expanded', String(open));
  });
  function syncToggleInputs() { document.querySelectorAll('#layers input[data-layer]').forEach((i) => { i.checked = !!S.toggles[i.dataset.layer]; }); }
  document.querySelectorAll('#layers input[data-layer]').forEach((i) => i.addEventListener('change', () => {
    S.toggles[i.dataset.layer] = i.checked; applyToggles(); refreshPanel();
  }));
  $('#jump').addEventListener('change', (e) => {
    const r = REGIONS[e.target.value]; if (r) map.fitBounds(r, { padding: 30, duration: 800 }); e.target.value = '';
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && document.activeElement !== input) { closePanel(); clearSelection(); } });

  // ---------------------------------------------------------------- boot
  const dataReady = load(); // start downloading data while the map initialises
  dataReady.catch(() => {});
  map.on('load', async () => {
    try {
      const labels = await dataReady;
      addLayers(labels);
      const fitHeader = () => {
        const h = document.querySelector('.topbar').offsetHeight;
        map.getContainer().style.top = `${h}px`;
        document.documentElement.style.setProperty('--header-h', `${h}px`);
        map.resize();
      };
      fitHeader(); map.fitBounds(CONUS, { padding: 20, duration: 0 });
      window.addEventListener('resize', fitHeader);
      buildIndex();
      if (!IS_TOUCH) {
        map.on('mousemove', onMove);
        map.getCanvas().addEventListener('mouseleave', clearHover);
        map.on('movestart', () => { card.hidden = true; tip.hidden = true; });
      }
      map.on('click', onClick);
      setStatus('Ready'); setTimeout(() => $('#status').classList.add('hide'), 600);
      setTimeout(computeOverlaps, 50);
    } catch (err) {
      console.error(err);
      setStatus(`Could not load data: ${err.message}`, 'err');
    }
  });

  window.__mapState = S; window.__map = map; // for debugging in the console
})();
