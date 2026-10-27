/* 訪問ルート（地図表示・訪問順の計算）
 * 外部へ送るのは「住所の文字列」（国土地理院の住所検索）と地図タイルの取得のみ。
 * 訪問順の計算はブラウザ内で行う（直線距離）。
 */
(function () {
  'use strict';

  const GEO_KEY = 'visitPatientList.geo.v1';
  const GSI_SEARCH_URL = 'https://msearch.gsi.go.jp/address-search/AddressSearch?q=';
  const TILE_URL = 'https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png';
  const TILE_ATTRIBUTION = '<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">地理院タイル</a>';
  const GEOCODE_INTERVAL_MS = 300;
  const EXACT_SOLVE_MAX = 8; // 訪問先がこの数以下なら全通りを調べて最短を求める
  const GMAPS_MAX_WAYPOINTS = 9; // GoogleマップのURLで指定できる経由地の上限
  const DEFAULT_CENTER = [35.681, 139.767];

  // cache: 住所 → { lat, lng, title, manual }（title は住所検索で見つかった地名）
  const geo = { cache: {}, officeAddress: '', roundTrip: true, route: null };
  let map = null;
  let markerLayer = null;
  let pickAddr = null; // 地図クリックで位置を指定する住所
  let busy = false;

  const $ = (sel) => document.querySelector(sel);

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function addrKey(addr) {
    return String(addr || '').normalize('NFKC').replace(/\s+/g, '').trim();
  }

  // 位置検索に使う住所（建物名などの「住所（続き）」は含めない）
  function patientAddr(p) {
    return (p.addr1 || p.address || '').trim();
  }

  // ---------- 保存 ----------
  function saveGeo() {
    try {
      localStorage.setItem(GEO_KEY, JSON.stringify(geo));
    } catch (e) {
      console.warn(e);
    }
  }

  function loadGeo() {
    try {
      Object.assign(geo, JSON.parse(localStorage.getItem(GEO_KEY) || '{}'));
    } catch (e) {
      console.warn(e);
    }
  }

  function pointOf(addr) {
    return addr ? geo.cache[addrKey(addr)] || null : null;
  }

  // ---------- 住所 → 位置（国土地理院 住所検索） ----------
  async function searchGsi(query) {
    const res = await fetch(GSI_SEARCH_URL + encodeURIComponent(query));
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const list = await res.json();
    if (!Array.isArray(list) || !list.length) return null;
    const [lng, lat] = list[0].geometry.coordinates;
    return { lat, lng, title: list[0].properties.title || '' };
  }

  // 見つからない場合は番地までに切り詰めて再検索する
  async function geocode(addr) {
    const full = addr.normalize('NFKC').replace(/\s+/g, '');
    const hit = await searchGsi(full);
    if (hit) return hit;
    const m = full.match(/^\D*\d+(?:(?:丁目|番地|番|号|の|-|‐|－|ー|−)\d+)*/);
    if (m && m[0] !== full) return searchGsi(m[0]);
    return null;
  }

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  async function geocodeMissing() {
    if (busy) return;
    const targets = new Set();
    if (geo.officeAddress && !pointOf(geo.officeAddress)) targets.add(geo.officeAddress);
    // 訪問先ごとに1件（同じグループホームは代表の住所のみ）
    buildStops().forEach((s) => {
      if (s.addr && !s.point) targets.add(s.addr);
    });
    if (!targets.size) {
      setStatus('位置を取得していない住所はありません。');
      return;
    }
    busy = true;
    renderButtons();
    const failed = [];
    let i = 0;
    for (const addr of targets) {
      i += 1;
      setStatus(`住所から位置を取得しています… ${i} / ${targets.size}`);
      try {
        const hit = await geocode(addr);
        if (hit) geo.cache[addrKey(addr)] = Object.assign(hit, { manual: false });
        else failed.push(addr);
      } catch (e) {
        busy = false;
        saveGeo();
        render();
        setStatus('国土地理院の住所検索に接続できませんでした。インターネット接続を確認してください。地図をクリックして手動で位置を指定することもできます。', true);
        return;
      }
      if (i < targets.size) await sleep(GEOCODE_INTERVAL_MS);
    }
    busy = false;
    saveGeo();
    render();
    fitAll();
    setStatus(failed.length
      ? `${targets.size - failed.length}件の位置を取得しました。${failed.length}件は見つかりませんでした（右の一覧の「地図で指定」から指定してください）。`
      : `${targets.size}件の位置を取得しました。位置がずれている場合はピンをドラッグして直せます。`, failed.length > 0);
  }

  // ---------- 訪問先（同じグループホーム・同じ住所は1か所にまとめる） ----------
  function buildStops() {
    const stops = new Map();
    window.VisitApp.todayPatients().forEach((p) => {
      const addr = patientAddr(p);
      const id = p.gh ? `gh:${p.gh}` : `addr:${addrKey(addr) || p.key}`;
      if (!stops.has(id)) stops.set(id, { id, gh: p.gh, addr, patients: [] });
      stops.get(id).patients.push(p);
    });
    return [...stops.values()].map((s) => Object.assign(s, { point: pointOf(s.addr) }));
  }

  function stopLabel(s) {
    if (s.gh) return `${s.gh}（${s.patients.length}名）`;
    return s.patients.map((p) => p.name || p.id).join('、');
  }

  // ---------- 訪問順の計算（直線距離） ----------
  function distKm(a, b) {
    const R = 6371;
    const r = Math.PI / 180;
    const dLat = (b.lat - a.lat) * r;
    const dLng = (b.lng - a.lng) * r;
    const h = Math.sin(dLat / 2) ** 2 + Math.cos(a.lat * r) * Math.cos(b.lat * r) * Math.sin(dLng / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(h));
  }

  // points[0] が出発地（hasDepot のとき）。返り値は points の添字の並び
  function solveRoute(points, hasDepot, roundTrip) {
    const n = points.length;
    const D = points.map((a) => points.map((b) => distKm(a, b)));
    const closed = hasDepot && roundTrip;
    const cost = (t) => {
      let c = 0;
      for (let i = 0; i < t.length - 1; i++) c += D[t[i]][t[i + 1]];
      return closed ? c + D[t[t.length - 1]][t[0]] : c;
    };
    const fixed = hasDepot ? 1 : 0;
    const free = [];
    for (let i = fixed; i < n; i++) free.push(i);
    const head = hasDepot ? [0] : [];

    if (free.length <= EXACT_SOLVE_MAX) {
      let best = null;
      let bestCost = Infinity;
      permute(free, (perm) => {
        const t = head.concat(perm);
        const c = cost(t);
        if (c < bestCost - 1e-9) { bestCost = c; best = t; }
      });
      return { order: best, km: bestCost };
    }

    // 訪問先が多い場合：最近傍法と無作為な初期解を 2-opt / 1点移動 で改善し、最良のものを採用
    const starts = [];
    const nearest = (from, rest) => {
      const t = [from];
      const left = new Set(rest);
      while (left.size) {
        const cur = t[t.length - 1];
        let nx = null;
        left.forEach((j) => { if (nx == null || D[cur][j] < D[cur][nx]) nx = j; });
        t.push(nx);
        left.delete(nx);
      }
      return t;
    };
    if (hasDepot) starts.push(nearest(0, free));
    else free.forEach((s) => starts.push(nearest(s, free.filter((x) => x !== s))));
    for (let k = 0; k < 30; k++) {
      const shuffled = free.slice().sort(() => Math.random() - 0.5);
      starts.push(head.concat(shuffled));
    }
    let best = null;
    let bestCost = Infinity;
    for (const st of starts) {
      const t = improve(st, cost, fixed);
      const c = cost(t);
      if (c < bestCost - 1e-9) { bestCost = c; best = t; }
    }
    return { order: best, km: bestCost };
  }

  function improve(tour, cost, fixed) {
    let t = tour.slice();
    let c = cost(t);
    let improved = true;
    while (improved) {
      improved = false;
      // 2-opt：区間を反転
      for (let i = fixed; i < t.length - 1; i++) {
        for (let k = i + 1; k < t.length; k++) {
          const cand = t.slice(0, i).concat(t.slice(i, k + 1).reverse(), t.slice(k + 1));
          const cc = cost(cand);
          if (cc < c - 1e-9) { t = cand; c = cc; improved = true; }
        }
      }
      // 1点を別の位置へ移動
      for (let i = fixed; i < t.length; i++) {
        for (let j = fixed; j < t.length; j++) {
          if (i === j) continue;
          const cand = t.slice();
          const [v] = cand.splice(i, 1);
          cand.splice(j, 0, v);
          const cc = cost(cand);
          if (cc < c - 1e-9) { t = cand; c = cc; improved = true; }
        }
      }
    }
    return t;
  }

  function permute(arr, visit) {
    const a = arr.slice();
    const c = new Array(a.length).fill(0);
    visit(a.slice());
    let i = 0;
    while (i < a.length) {
      if (c[i] < i) {
        const j = i % 2 === 0 ? 0 : c[i];
        [a[j], a[i]] = [a[i], a[j]];
        visit(a.slice());
        c[i] += 1;
        i = 0;
      } else {
        c[i] = 0;
        i += 1;
      }
    }
  }

  function computeRoute() {
    const stops = buildStops();
    const located = stops.filter((s) => s.point);
    if (located.length < 1) {
      setStatus('位置がわかる訪問先がありません。先に「住所から位置を取得」を押してください。', true);
      return;
    }
    const office = pointOf(geo.officeAddress);
    if (geo.officeAddress && !office) {
      setStatus('事業所の位置がまだ取得できていません。「住所から位置を取得」を押すか、「地図で指定」で事業所の位置を指定してください。', true);
      return;
    }
    const points = (office ? [office] : []).concat(located.map((s) => s.point));
    const { order, km } = solveRoute(points, !!office, geo.roundTrip);
    const stopOrder = order.filter((i) => !office || i !== 0).map((i) => located[i - (office ? 1 : 0)].id);
    geo.route = { stopIds: stopOrder, km, withOffice: !!office, roundTrip: geo.roundTrip, signature: signature(stops) };
    saveGeo();
    render();
    const missing = stops.length - located.length;
    setStatus(`訪問順を計算しました（直線距離 合計 約${km.toFixed(1)}km）。${missing ? `位置が未設定の${missing}か所は含まれていません。` : ''}「この順番を当日リストに反映」で当日リスト・印刷の並びになります。`, missing > 0);
  }

  // 当日リストの訪問先が変わったかを判定するための値
  function signature(stops) {
    return stops.map((s) => s.id + ':' + (s.point ? `${s.point.lat},${s.point.lng}` : '')).sort().join('|')
      + `#${geo.officeAddress}#${geo.roundTrip}`;
  }

  function currentRoute(stops) {
    if (!geo.route) return null;
    if (geo.route.signature !== signature(stops)) return { stale: true };
    const byId = new Map(stops.map((s) => [s.id, s]));
    const ordered = geo.route.stopIds.map((id) => byId.get(id)).filter(Boolean);
    return { stale: false, stops: ordered, km: geo.route.km };
  }

  function applyRoute() {
    const stops = buildStops();
    const r = currentRoute(stops);
    if (!r || r.stale) return;
    const keys = r.stops.flatMap((s) => s.patients.map((p) => p.key));
    window.VisitApp.applyRouteOrder(keys);
    setStatus('訪問順を当日リストに反映しました。印刷もこの順番になります（グループホームの患者の直前にホーム名の行が入ります）。');
  }

  // ---------- Googleマップで経路を開く（住所のみを渡す） ----------
  function gmapsParam(addr) {
    const pt = pointOf(addr);
    if (pt && pt.manual) return `${pt.lat.toFixed(6)},${pt.lng.toFixed(6)}`;
    return addr;
  }

  function gmapsLinks(orderedStops) {
    const pts = orderedStops.map((s) => gmapsParam(s.addr));
    if (geo.officeAddress && pointOf(geo.officeAddress)) {
      pts.unshift(gmapsParam(geo.officeAddress));
      if (geo.roundTrip) pts.push(gmapsParam(geo.officeAddress));
    }
    if (pts.length < 2) return [];
    const links = [];
    const per = GMAPS_MAX_WAYPOINTS + 1; // 1区間で進める地点数
    for (let s = 0; s < pts.length - 1; s += per) {
      const seg = pts.slice(s, s + per + 1);
      const q = new URLSearchParams({ api: '1', origin: seg[0], destination: seg[seg.length - 1], travelmode: 'driving' });
      if (seg.length > 2) q.set('waypoints', seg.slice(1, -1).join('|'));
      links.push('https://www.google.com/maps/dir/?' + q.toString());
    }
    return links;
  }

  // ---------- 地図 ----------
  function ensureMap() {
    if (map || typeof L === 'undefined') return;
    map = L.map('route-map', { zoomControl: true }).setView(DEFAULT_CENTER, 11);
    L.tileLayer(TILE_URL, { attribution: TILE_ATTRIBUTION, maxZoom: 18 }).addTo(map);
    markerLayer = L.layerGroup().addTo(map);
    map.on('click', (e) => {
      if (!pickAddr) return;
      geo.cache[addrKey(pickAddr)] = { lat: e.latlng.lat, lng: e.latlng.lng, title: '地図で指定', manual: true };
      pickAddr = null;
      $('#route-map').classList.remove('picking');
      saveGeo();
      render();
      setStatus('位置を指定しました。');
    });
  }

  function pinIcon(text, cls) {
    return L.divIcon({ className: `route-pin ${cls || ''}`, html: `<span>${esc(text)}</span>`, iconSize: [30, 30], iconAnchor: [15, 15], popupAnchor: [0, -14] });
  }

  function addMarker(addr, latlng, icon, popupHtml) {
    const m = L.marker(latlng, { icon, draggable: true }).addTo(markerLayer);
    m.bindPopup(popupHtml);
    m.on('dragend', () => {
      const ll = m.getLatLng();
      geo.cache[addrKey(addr)] = { lat: ll.lat, lng: ll.lng, title: 'ピンで調整', manual: true };
      saveGeo();
      render();
      setStatus('位置を修正しました。訪問順を再計算してください。');
    });
  }

  function renderMap(stops, route) {
    ensureMap();
    if (!map) return;
    markerLayer.clearLayers();
    const office = pointOf(geo.officeAddress);
    if (office) addMarker(geo.officeAddress, [office.lat, office.lng], pinIcon('事', 'office'), `<b>事業所（出発地）</b><br>${esc(geo.officeAddress)}`);
    const numberOf = new Map();
    if (route && !route.stale) route.stops.forEach((s, i) => numberOf.set(s.id, i + 1));
    stops.forEach((s) => {
      if (!s.point) return;
      const no = numberOf.get(s.id);
      const names = s.patients.map((p) => esc(p.name || p.id)).join('<br>');
      const popup = `<b>${no ? `${no}. ` : ''}${esc(s.gh || '居宅')}</b><br>${names}<br><span class="popup-addr">${esc(s.addr)}</span>`
        + `<br><span class="popup-hit">検索結果: ${esc(s.point.title)}</span>`;
      addMarker(s.addr, [s.point.lat, s.point.lng], pinIcon(no || '・', s.gh ? 'gh' : ''), popup);
    });
    if (route && !route.stale && route.stops.length) {
      const line = route.stops.map((s) => [s.point.lat, s.point.lng]);
      if (office) {
        line.unshift([office.lat, office.lng]);
        if (geo.roundTrip) line.push([office.lat, office.lng]);
      }
      L.polyline(line, { color: '#1f6feb', weight: 3, opacity: 0.7, dashArray: '6 6' }).addTo(markerLayer);
    }
  }

  function fitAll() {
    if (!map) return;
    const pts = [];
    markerLayer.eachLayer((l) => { if (l.getLatLng) pts.push(l.getLatLng()); });
    if (pts.length === 1) map.setView(pts[0], 15);
    else if (pts.length > 1) map.fitBounds(L.latLngBounds(pts).pad(0.15));
  }

  // ---------- 一覧・ボタン ----------
  function setStatus(text, warn) {
    const el = $('#route-status');
    el.textContent = text;
    el.classList.toggle('is-warn', !!warn);
  }

  function renderButtons(stops, route) {
    stops = stops || buildStops();
    route = route === undefined ? currentRoute(stops) : route;
    const hasToday = stops.length > 0;
    $('#btn-geocode').disabled = busy || !hasToday;
    $('#btn-optimize').disabled = busy || !stops.some((s) => s.point);
    $('#btn-apply-route').disabled = busy || !route || route.stale;
  }

  function render() {
    if (!window.VisitApp) return;
    const stops = buildStops();
    const route = currentRoute(stops);
    renderButtons(stops, route);
    renderMap(stops, route);

    const list = [];
    const office = pointOf(geo.officeAddress);
    if (geo.officeAddress) {
      list.push(`<li class="office">${office ? '' : '<span class="tag">位置未設定</span>'}事業所（出発）${office ? '' : ` <button type="button" class="btn small" data-pick="${esc(geo.officeAddress)}">地図で指定</button>`}</li>`);
    }
    const ordered = route && !route.stale ? route.stops : stops;
    const unordered = route && !route.stale ? stops.filter((s) => !route.stops.includes(s)) : [];
    ordered.concat(unordered).forEach((s) => {
      const no = route && !route.stale && route.stops.includes(s) ? `${route.stops.indexOf(s) + 1}.` : '';
      const tag = s.point ? '' : '<span class="tag">位置未設定</span>';
      const pick = s.point ? '' : ` <button type="button" class="btn small" data-pick="${esc(s.addr)}"${s.addr ? '' : ' disabled'}>地図で指定</button>`;
      list.push(`<li class="${s.gh ? 'gh' : ''}"><b>${no}</b> ${tag}${esc(stopLabel(s))}${pick}<div class="route-addr">${esc(s.addr || '（住所なし）')}</div></li>`);
    });
    if (geo.officeAddress && geo.roundTrip && route && !route.stale) list.push('<li class="office">事業所（帰着）</li>');
    $('#route-list').innerHTML = list.join('') || '<li class="empty">当日リストに患者を追加すると表示されます</li>';

    const summary = route && !route.stale
      ? `直線距離 合計 約${route.km.toFixed(1)}km（実際の道のりはこれより長くなります）`
      : route && route.stale ? '当日リストまたは位置が変わりました。「訪問順を計算」をもう一度押してください。' : '';
    $('#route-summary').textContent = summary;

    const links = route && !route.stale ? gmapsLinks(route.stops) : [];
    $('#gmaps-links').innerHTML = links.length
      ? links.map((u, i) => `<a class="btn" href="${esc(u)}" target="_blank" rel="noopener noreferrer">Googleマップで経路を開く${links.length > 1 ? `（${i + 1}/${links.length}）` : ''}</a>`).join('')
        + '<p class="hint">Googleマップには住所（地図で指定した地点は座標）だけを渡します。</p>'
      : '';
  }

  function clearStorage() {
    localStorage.removeItem(GEO_KEY);
    geo.cache = {};
    geo.route = null;
    geo.officeAddress = '';
    $('#office-address').value = '';
    render();
  }

  function init() {
    loadGeo();
    $('#office-address').value = geo.officeAddress;
    $('#trip-mode').value = geo.roundTrip ? 'round' : 'oneway';
    $('#office-address').addEventListener('change', (e) => {
      geo.officeAddress = e.target.value.trim();
      saveGeo();
      render();
    });
    $('#trip-mode').addEventListener('change', (e) => {
      geo.roundTrip = e.target.value === 'round';
      saveGeo();
      render();
    });
    $('#btn-geocode').addEventListener('click', geocodeMissing);
    $('#btn-optimize').addEventListener('click', () => { computeRoute(); fitAll(); });
    $('#btn-apply-route').addEventListener('click', applyRoute);
    $('#btn-fit').addEventListener('click', fitAll);
    $('#route-list').addEventListener('click', (e) => {
      const b = e.target.closest('button[data-pick]');
      if (!b) return;
      pickAddr = b.dataset.pick;
      $('#route-map').classList.add('picking');
      setStatus(`地図上をクリックして「${pickAddr}」の位置を指定してください。`);
    });
    render();
    fitAll();
  }

  window.VisitRoute = { refresh: render, clearStorage };
  document.addEventListener('DOMContentLoaded', init);
})();
