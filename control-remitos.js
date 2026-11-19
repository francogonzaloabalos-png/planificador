import {
  DEFAULT_MODEL, headerFields, lineFields, states, decisions, normalize, text, datesFor, readWorkbook,
  differences, indexWorkbook, chosenRow, lineId, compareReceipts, transcribe, publicReceipt,
  receiveReceipt, mergePublished, exportRows, excelDay, GeminiQuotaError
} from './control-remitos-core.js';

const $ = id => document.getElementById(id);
const LEGACY_KEY = 'control_remitos_gemini_key_v1';
const LOCAL_KEYS = 'control_remitos_gemini_keys_v1';
const SHARE_KEYS = 'control_remitos_share_keys_v1';
const MODEL = 'control_remitos_gemini_model_v1';
const MAX_BATCH = 10;
const MAX_EDGE = 1600;
const PAUSE = 1800;
const SHEET_JS = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const HEIC = 'https://cdn.jsdelivr.net/npm/heic2any@0.0.4/dist/heic2any.min.js';
const html = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const sleep = delay => new Promise(resolve => setTimeout(resolve, delay));
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const newId = () => `r${Date.now().toString(36)}-${crypto.randomUUID()}`;

let workspace;
let roomGeneration = 0;
let targetScope;
let busy = false;
let loadingExcel = false;
let chooser = null;
const cards = new Map();
const photoUrls = new Map();
const libraries = new Map();

function notice(message, error = false) {
  $('notice').hidden = !message;
  $('notice').textContent = message;
  $('notice').classList.toggle('error', error);
}

function jobMessage(message, error = false) {
  $('task-status').textContent = message;
  $('task-status').classList.toggle('error', error);
}

function dependency(name, url) {
  if (window[name]) return Promise.resolve(window[name]);
  if (!libraries.has(name)) {
    const script = document.createElement('script');
    script.src = url;
    const pending = new Promise((resolve, reject) => {
      script.onload = () => window[name] ? resolve(window[name]) : reject(new Error(`No se pudo iniciar ${name}.`));
      script.onerror = () => reject(new Error(`No se pudo cargar ${name}. Revisá la conexión y reintentá.`));
      document.head.append(script);
    }).catch(error => {
      script.remove();
      libraries.delete(name);
      throw error;
    });
    libraries.set(name, pending);
  }
  return libraries.get(name);
}

// Excel, fotos, transcripciones locales y publicaciones pendientes sobreviven a recargas.
// Las claves de Gemini se administran por separado de los registros de IndexedDB.
const database = new Promise((resolve, reject) => {
  const request = indexedDB.open('control-remitos-v1', 1);
  request.onupgradeneeded = () => {
    const store = request.result.createObjectStore('records', { keyPath: ['scope', 'id'] });
    store.createIndex('scope', 'scope');
  };
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(new Error('No se pudo abrir el almacenamiento de este navegador.'));
});

async function readRecords(scope) {
  const db = await database;
  return new Promise((resolve, reject) => {
    const request = db.transaction('records').objectStore('records').index('scope').getAll(scope);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function storeRecords(records) {
  if (!records.length) return;
  const snapshot = structuredClone(records);
  const db = await database;
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('records', 'readwrite');
    transaction.oncomplete = resolve;
    transaction.onabort = () => reject(transaction.error || new Error('Se interrumpió el guardado local.'));
    const store = transaction.objectStore('records');
    for (const record of snapshot) store.put(record);
  });
}

function storeError() {
  notice('El navegador no pudo guardar los últimos cambios. Exportá los resultados antes de cerrar esta pestaña.', true);
}

function parentConnection() {
  try {
    if (parent === window) return null;
    const saved = JSON.parse(parent.localStorage.getItem('planificador_sync_v1')
      || parent.sessionStorage.getItem('planificador_sync_v1'));
    if (!saved?.config?.databaseURL || !saved.room) return null;
    const url = saved.config.databaseURL.replace(/\/+$/, '');
    const sdk = parent.firebase;
    const app = [...(sdk?.apps || [])].reverse().find(candidate => candidate.name.startsWith('planificador-')
      && candidate.options.databaseURL?.replace(/\/+$/, '') === url && candidate.options.projectId === saved.config.projectId);
    return { sdk, app, room: saved.room, scope: JSON.stringify([url, saved.room]) };
  } catch {
    // Una apertura directa o fuera del planificador también puede trabajar en local.
    return null;
  }
}

async function refreshRoom() {
  const connection = parentConnection();
  const scope = connection?.scope || 'local';
  if (scope === targetScope) {
    if (workspace?.scope === scope && workspace.sync?.app !== connection?.app) connect(workspace, connection);
    else if (workspace?.scope === scope) listenKeys(workspace);
    return;
  }
  const generation = ++roomGeneration;
  targetScope = scope;
  if (workspace) detach(workspace);
  workspace = null;
  if ($('chooser').open) $('chooser').close();
  for (const url of photoUrls.values()) URL.revokeObjectURL(url);
  photoUrls.clear();
  cards.clear();
  $('receipts').replaceChildren();
  $('sync').textContent = 'Abriendo espacio de trabajo…';
  $('progress-wrap').hidden = true;
  $('queue-wrap').hidden = true;
  updateControls();
  jobMessage('');
  keyFeedback('');
  notice('');
  try {
    const records = await readRecords(scope);
    if (generation !== roomGeneration) return;
    const workbook = records.find(record => record.id === 'excel');
    const local = new Map(records.filter(record => record.type === 'photo').map(job => [job.id, job]));
    const cached = records.find(record => record.id === 'cloud')?.values || {};
    workspace = {
      scope, workbook, excel: workbook ? indexWorkbook(workbook.rows) : null, local,
      remote: new Map(), includeAll: records.find(record => record.id === 'options')?.includeAll === true,
      comparison: null, sync: null, keyBlocks: new Map()
    };
    readCloud(workspace, cached);
    recalculate(workspace);
    connect(workspace, parentConnection());
    render();
  } catch (error) {
    if (generation !== roomGeneration) return;
    $('sync').textContent = 'Almacenamiento no disponible';
    notice(error.message, true);
  }
}

function detach(ctx) {
  const sync = ctx.sync;
  ctx.sync = null;
  if (!sync?.ref) return;
  if (sync.keysRef && sync.onKeys) {
    try { sync.keysRef.off('value', sync.onKeys); } catch { /* La app compartida puede estar cerrada. */ }
  }
  try {
    sync.ref.off('value', sync.onValue);
    sync.connected.off('value', sync.onConnected);
  } catch { /* El planificador puede haber cerrado la app compartida. */ }
}

function listenKeys(ctx) {
  const sync = ctx.sync;
  if (!sync?.keysRef || !!sync.onKeys === usesRoomKeys(ctx)) return;
  if (sync.onKeys) {
    sync.keysRef.off('value', sync.onKeys);
    sync.onKeys = null;
    sync.keys = null;
    sync.keysError = false;
    return;
  }
  const onKeys = snapshot => {
    if (ctx.sync !== sync || sync.onKeys !== onKeys) return;
    sync.keys = decodeKeys(snapshot.val());
    renderKeys();
  };
  sync.onKeys = onKeys;
  sync.keysRef.on('value', onKeys, () => {
    if (ctx.sync !== sync || sync.onKeys !== onKeys) return;
    sync.keysError = true;
    renderKeys();
  });
}

function connect(ctx, connection) {
  detach(ctx);
  if (!connection || connection.scope !== ctx.scope || !connection.app) {
    updateSync();
    renderKeys();
    return;
  }
  const sync = { app: connection.app, ready: false, online: false, writing: false, error: false };
  ctx.sync = sync;
  try {
    const db = connection.sdk.database(connection.app);
    // Hermano de «salas»: el planificador reemplaza el contenido completo de cada sala.
    const room = db.ref('controlRemitos').child(connection.room);
    sync.ref = room.child('remitos');
    sync.keysRef = room.child('config');
    listenKeys(ctx);
    sync.connected = db.ref('.info/connected');
    sync.onValue = snapshot => {
      if (ctx.sync !== sync) return;
      sync.ready = true;
      const values = snapshot.val() || {};
      readCloud(ctx, values);
      storeRecords([{ id: 'cloud', scope: ctx.scope, values }]).catch(storeError);
      recalculate(ctx);
      render();
    };
    sync.onConnected = snapshot => {
      if (ctx.sync !== sync) return;
      sync.online = snapshot.val() === true;
      updateSync();
      renderKeys();
      flush(ctx);
    };
    sync.ref.on('value', sync.onValue, () => {
      if (ctx.sync !== sync) return;
      sync.error = true;
      updateSync();
    });
    sync.connected.on('value', sync.onConnected);
  } catch {
    sync.error = true;
  }
  updateSync();
  renderKeys();
}

function readCloud(ctx, values) {
  const previous = ctx.remote;
  ctx.remote = new Map();
  for (const [id, value] of Object.entries(values)) {
    if (!/^r[a-z0-9-]+$/i.test(id) || ctx.local.has(id)) continue;
    const receipt = receiveReceipt(id, value);
    if (receipt) ctx.remote.set(id, same(previous.get(id), receipt) ? previous.get(id) : receipt);
  }
}

function updateSync() {
  const ctx = workspace;
  if (!ctx) return;
  const sync = ctx.sync;
  const pending = [...ctx.local.values()].filter(job => job.pending).length;
  $('sync-retry').hidden = !sync?.error;
  let status = 'Modo local';
  if (ctx.scope !== 'local') {
    const suffix = pending ? ` · ${pending} por publicar` : '';
    if (sync?.error) status = `Error de sincronización${suffix}`;
    else if (!sync?.ready || !sync.online) status = `Sala sin conexión${suffix}`;
    else status = pending ? `Sincronizando${suffix}` : 'Sala sincronizada';
  }
  $('sync').textContent = status;
}

async function flush(ctx) {
  const sync = ctx.sync;
  if (workspace !== ctx || !sync?.ready || !sync.online || sync.writing || sync.error) return;
  const connection = parentConnection();
  if (connection?.scope !== ctx.scope || connection.app !== sync.app) return;
  const changes = new Map([...ctx.local.values()].filter(job => job.pending).map(job => [job.id, job.payload]));
  if (!changes.size) return;
  sync.writing = true;
  try {
    const outcome = await sync.ref.transaction(current => mergePublished(current || {}, changes), undefined, false);
    if (ctx.sync !== sync) return;
    if (!outcome.committed) throw new Error('Publicación interrumpida.');
    const values = outcome.snapshot.val() || {};
    readCloud(ctx, values);
    const acknowledged = [];
    for (const [id, payload] of changes) {
      const job = ctx.local.get(id);
      // Una edición hecha durante el envío queda pendiente para la transacción siguiente.
      if (job.payload === payload) {
        job.pending = false;
        acknowledged.push(job);
      }
    }
    await storeRecords(acknowledged).catch(storeError);
  } catch {
    if (ctx.sync === sync) sync.error = true;
  } finally {
    sync.writing = false;
    if (ctx.sync === sync) {
      recalculate(ctx);
      render();
    }
  }
}

function localReceipts(ctx) {
  return [...ctx.local.values()].filter(job => job.receipt).map(job => job.receipt);
}

function recalculate(ctx) {
  const local = localReceipts(ctx);
  ctx.comparison = compareReceipts(local, ctx.remote.values(), ctx.excel, ctx.includeAll);
  const changed = [];
  for (const receipt of local) {
    const job = ctx.local.get(receipt.id);
    const payload = publicReceipt(receipt, ctx.comparison.results);
    if (same(payload, job.payload)) continue;
    receipt.lines.forEach((line, index) => {
      const result = ctx.comparison.results.get(lineId(receipt, index));
      line.estado = result.state;
      line.elegida = result.chosen;
      line.decision = result.decision || '';
    });
    job.payload = payload;
    job.pending = ctx.scope !== 'local';
    changed.push(job);
  }
  storeRecords(changed).then(() => flush(ctx)).catch(storeError);
}

const prefersSharing = () => localStorage.getItem(SHARE_KEYS) !== '0';
const usesRoomKeys = ctx => !!ctx && ctx.scope !== 'local' && prefersSharing();
const validKey = value => typeof value === 'string' && /^[A-Za-z0-9_-]{8,}$/.test(value);
const labels = { available: 'Disponible', paused: 'En pausa', daily: 'Agotada por hoy' };

// Un nodo con versión representa una lista vacía, que Firebase no guarda como array.
function encodeKeys(entries) {
  return { version: 1, keys: Object.fromEntries(entries.map((entry, i) => [`k${i}`, entry])) };
}

function decodeKeys(config) {
  const keys = config?.keys || {};
  const seen = new Set();
  return Object.keys(keys).filter(id => /^k\d+$/.test(id)).sort((a, b) => +a.slice(1) - +b.slice(1)).flatMap(id => {
    const { value, status, until } = keys[id] || {};
    if (!validKey(value) || seen.has(value)) return [];
    seen.add(value);
    return [{ value, ...((status === 'paused' || status === 'daily') && Number.isFinite(until) ? { status, until } : {}) }];
  });
}

function localKeys() {
  const stored = localStorage.getItem(LOCAL_KEYS);
  if (stored) {
    try { return decodeKeys(JSON.parse(stored)); } catch { return []; }
  }
  const value = localStorage.getItem(LEGACY_KEY);
  return validKey(value) ? [{ value }] : [];
}

function activeKeys(ctx) {
  return usesRoomKeys(ctx) ? ctx.sync?.keys || [] : localKeys();
}

function effectiveBlock(entry, other, now = Date.now()) {
  const saved = entry?.until > now ? entry : null;
  const temporary = other?.until > now ? other : null;
  if (!saved || !temporary) return saved || temporary;
  if (saved.status !== temporary.status) return saved.status === 'daily' ? saved : temporary;
  return saved.until >= temporary.until ? saved : temporary;
}

function keyState(ctx, entry) {
  return effectiveBlock(entry, ctx?.keyBlocks.get(entry.value));
}

function keyFeedback(message, error = false) {
  $('key-feedback').textContent = message;
  $('key-feedback').classList.toggle('error', error);
}

function renderKeys() {
  const ctx = workspace;
  const shared = usesRoomKeys(ctx);
  const ready = !shared || (!!ctx.sync?.online && !!ctx.sync?.keys && !ctx.sync.keysError);
  const disabled = !ctx || busy || !ready || !!ctx.savingKeys;
  const entries = activeKeys(ctx);
  $('share-keys').checked = prefersSharing();
  $('share-keys').disabled = busy || !!ctx?.savingKeys;
  $('key-warning').hidden = !prefersSharing();
  $('gemini-key').disabled = $('add-key').disabled = disabled;
  const mode = shared ? 'Claves compartidas de la sala' : 'Claves privadas de este navegador';
  const status = shared && !ready ? 'Sin acceso a la configuración de la sala'
    : !entries.length ? 'No hay claves de Gemini configuradas' : '';
  $('key-info').textContent = status ? `${mode} · ${status}` : mode;
  const items = entries.map((entry, i) => {
    const suffix = html(entry.value.slice(-4));
    const controls = (action, label, unavailable = false) => `<button type="button" class="small" data-key-action="${action}" data-key-index="${i}" aria-label="${label} clave terminada en ${suffix}" ${disabled || unavailable ? 'disabled' : ''}>${label}</button>`;
    return `<li><span>••••${suffix} · ${labels[keyState(ctx, entry)?.status || 'available']}</span><div class="actions">
      ${controls('up', 'Subir', i === 0)}${controls('down', 'Bajar', i === entries.length - 1)}${controls('delete', 'Borrar')}</div></li>`;
  }).join('');
  const list = $('gemini-keys');
  if (list.dataset.markup !== items) {
    list.dataset.markup = items;
    list.innerHTML = items;
  }
}

async function changeSharedKeys(ctx, update) {
  const sync = ctx.sync;
  const connection = parentConnection();
  if (!sync?.keysRef || connection?.scope !== ctx.scope || connection.app !== sync.app) throw new Error();
  const outcome = await sync.keysRef.transaction(current => {
    const entries = decodeKeys(current);
    const updated = update(entries);
    return updated === entries ? undefined : encodeKeys(updated);
  }, undefined, false);
  if (ctx.sync === sync && outcome.committed) {
    sync.keys = decodeKeys(outcome.snapshot.val());
    renderKeys();
  }
}

function saveLocalKeys(entries) {
  localStorage.setItem(LOCAL_KEYS, JSON.stringify(encodeKeys(entries)));
  localStorage.removeItem(LEGACY_KEY);
}

function blockKey(ctx, entry, block) {
  const update = entries => {
    const current = entries.find(key => key.value === entry.value);
    const effective = effectiveBlock(current, block);
    if (!current || !effective || effective === current) return entries;
    return entries.map(key => key === current ? { value: key.value, ...effective } : key);
  };
  ctx.keyBlocks.set(entry.value, effectiveBlock(block, ctx.keyBlocks.get(entry.value)));
  if (usesRoomKeys(ctx)) {
    changeSharedKeys(ctx, update).catch(() => {
      if (workspace === ctx && usesRoomKeys(ctx)) keyFeedback('La pausa se aplicó en este navegador. No se pudo publicar en la sala.', true);
    });
  } else {
    try { saveLocalKeys(update(localKeys())); }
    catch {
      keyFeedback('No se pudo guardar el estado de la clave en este navegador.', true);
    }
  }
  renderKeys();
}

async function editKeys(update) {
  const ctx = workspace;
  if (!ctx || busy || ctx.savingKeys) return false;
  const shared = usesRoomKeys(ctx);
  ctx.savingKeys = true;
  keyFeedback('');
  renderKeys();
  try {
    if (shared) {
      if (!ctx.sync?.online || !ctx.sync.keys || ctx.sync.keysError) throw new Error();
      await changeSharedKeys(ctx, update);
    } else saveLocalKeys(update(localKeys()));
    return workspace === ctx;
  } catch {
    if (workspace === ctx) keyFeedback(shared ? 'No se pudieron guardar las claves en la sala.' : 'No se pudieron guardar las claves en este navegador.', true);
    return false;
  } finally {
    ctx.savingKeys = false;
    renderKeys();
  }
}

function modelName() {
  return text($('gemini-model').value).replace(/^models\//, '');
}

function updateControls() {
  const ctx = workspace;
  $('choose-excel').disabled = !ctx || loadingExcel;
  $('choose-images').disabled = !ctx || busy;
  $('process').disabled = !ctx || busy || ![...ctx.local.values()].some(job => job.status === 'queued');
  $('retry').disabled = !ctx || busy || ![...ctx.local.values()].some(job => job.status === 'failed');
  $('include-all').disabled = !ctx || !ctx.excel;
  $('include-all').checked = !!ctx?.includeAll;
  $('excel-info').textContent = ctx?.workbook ? `${ctx.workbook.rows.length.toLocaleString('es-AR')} filas · ${ctx.workbook.name}` : 'Excel sin cargar';
  $('gemini-model').disabled = $('save-model').disabled = busy;
  renderKeys();
}

function queueList() {
  const jobs = [...workspace.local.values()].sort((a, b) => b.id.localeCompare(a.id));
  $('queue-wrap').hidden = !jobs.length;
  const labels = { queued: 'Pendiente', failed: 'Falló', done: 'Transcripta' };
  const failed = jobs.filter(job => job.status === 'failed').length;
  $('queue-summary').textContent = `Fotos cargadas (${jobs.length})${failed ? ` · ${failed} fallidas` : ''}`;
  $('queue').innerHTML = jobs.map(job => `<li class="${job.status === 'failed' ? 'error' : ''}"><b>${html(job.name)}</b> · ${labels[job.status]}${job.error ? ` · ${html(job.error)}` : ''}</li>`).join('');
}

function heicFile(source) {
  return /\.hei[cf]$/i.test(source.name) || /image\/hei[cf]/.test(source.type);
}

async function jpegSource(job) {
  if (!heicFile(job.source) || job.preview) return job.preview || job.source;
  // Algunos navegadores ya decodifican HEIC. Usar esa vía si el conversor lo rechaza.
  try {
    const convert = await dependency('heic2any', HEIC);
    const converted = await convert({ blob: job.source, toType: 'image/jpeg', quality: .93 });
    return Array.isArray(converted) ? converted[0] : converted;
  } catch {
    try {
      return await resizePhoto(job.source, 4032);
    } catch {
      throw new Error('No se pudo convertir el HEIC. Reintentá o agregá una versión JPG.');
    }
  }
}

async function resizePhoto(source, maxEdge = MAX_EDGE) {
  const url = URL.createObjectURL(source);
  try {
    const photo = new Image();
    photo.src = url;
    await photo.decode();
    const scale = Math.min(1, maxEdge / Math.max(photo.naturalWidth, photo.naturalHeight));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(photo.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(photo.naturalHeight * scale));
    const context = canvas.getContext('2d', { alpha: false });
    context.fillStyle = '#fff';
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(photo, 0, 0, canvas.width, canvas.height);
    const jpeg = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', .89));
    if (!jpeg) throw new Error('No se pudo crear el JPG.');
    return jpeg;
  } catch {
    throw new Error('No se pudo abrir o reducir esta imagen.');
  } finally {
    URL.revokeObjectURL(url);
  }
}

function encodePhoto(photo) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = () => reject(new Error('No se pudo leer el JPG preparado.'));
    reader.readAsDataURL(photo);
  });
}

async function addPhotos(files) {
  const ctx = workspace;
  if (!ctx || !files.length) return;
  if (files.length > MAX_BATCH) return jobMessage('Seleccioná hasta 10 imágenes por tanda.', true);
  if (files.some(file => !/\.(jpe?g|png|hei[cf])$/i.test(file.name) && !/^image\/(jpeg|png|hei[cf])$/.test(file.type))) {
    return jobMessage('Elegí solo imágenes JPG, PNG o HEIC.', true);
  }
  const jobs = files.map(source => ({ id: newId(), scope: ctx.scope, type: 'photo', name: source.name, source, status: 'queued' }));
  try {
    await storeRecords(jobs);
    for (const job of jobs) ctx.local.set(job.id, job);
    if (workspace !== ctx) return;
    jobMessage(`${jobs.length} ${jobs.length === 1 ? 'foto lista para transcribir' : 'fotos listas para transcribir'}.`);
    queueList();
    updateControls();
  } catch { jobMessage('No hay espacio o acceso para guardar las fotos en este navegador.', true); }
}

async function transcribeWithKeys(ctx, model, base64, onStatus) {
  const attempted = new Set();
  let waited = false;
  while (workspace === ctx) {
    const entries = activeKeys(ctx);
    const next = entries.find(entry => !attempted.has(entry.value) && !keyState(ctx, entry));
    if (next) {
      attempted.add(next.value);
      try {
        return { receipt: await transcribe({ key: next.value, model, base64 }) };
      } catch (error) {
        if (!(error instanceof GeminiQuotaError)) throw error;
        if (workspace !== ctx) break;
        blockKey(ctx, next, error.block);
        onStatus(`${labels[error.block.status]}. Buscando otra clave disponible…`);
        continue;
      }
    }

    if (!entries.length) return { pending: 'No hay claves de Gemini configuradas.' };
    const blocks = entries.map(entry => keyState(ctx, entry));
    if (blocks.every(block => block?.status === 'daily')) return { pending: 'Todas las claves de Gemini están agotadas por hoy.' };
    const now = Date.now();
    const soonest = Math.min(...blocks.filter(block => block?.status !== 'daily').map(block => block?.until ?? now));
    if (waited || soonest - now > 65000) return { pending: 'Las claves de Gemini están en pausa o agotadas por hoy.' };

    // Si no hay otra clave, esperar una pausa corta y volver a intentar esta misma foto una vez.
    waited = true;
    attempted.clear();
    onStatus('Esperando a que termine la pausa de Gemini…');
    while (workspace === ctx && Date.now() < soonest && activeKeys(ctx).every(entry => keyState(ctx, entry))) {
      await sleep(Math.min(1000, Math.max(0, soonest - Date.now())));
    }
  }
  return { pending: 'Cambió el espacio de trabajo.' };
}

async function processPhotos(status) {
  const ctx = workspace;
  if (busy || !ctx) return;
  const model = localStorage.getItem(MODEL) || DEFAULT_MODEL;
  if (!activeKeys(ctx).length) {
    $('settings').open = true;
    $('gemini-key').focus();
    const waiting = usesRoomKeys(ctx) && !ctx.sync?.keys;
    return jobMessage(waiting ? 'No se pudieron leer las claves de la sala. También podés usar tus claves privadas.'
      : 'Agregá una clave de Gemini en Control de remitos.', true);
  }
  busy = true;
  updateControls();
  const run = async () => {
    // Otra pestaña del mismo navegador podría haber terminado una foto mientras se esperaba el turno.
    const stored = await readRecords(ctx.scope);
    if (workspace !== ctx) return;
    for (const job of stored) {
      if (job.type === 'photo' && (!ctx.local.has(job.id) || !ctx.local.get(job.id).receipt)) ctx.local.set(job.id, job);
    }
    recalculate(ctx);
    render();
    const selected = [...ctx.local.values()].filter(job => job.status === status).slice(0, MAX_BATCH);
    if (!selected.length) return jobMessage('La tanda ya se procesó en otra pestaña.');
    const progress = $('progress');
    $('progress-wrap').hidden = false;
    progress.max = selected.length;
    progress.value = 0;
    let failures = 0, completed = 0, pending = '';
    for (const [index, job] of selected.entries()) {
      if (workspace !== ctx) break;
      const label = `${index + 1}/${selected.length} · ${job.name}`;
      jobMessage(label);
      try {
        const source = await jpegSource(job);
        if (heicFile(job.source)) job.preview = source;
        const base64 = await encodePhoto(await resizePhoto(source));
        const outcome = await transcribeWithKeys(ctx, model, base64, message => jobMessage(`${label} · ${message}`));
        if (outcome.pending) {
          pending = outcome.pending;
          break;
        }
        job.receipt = { id: job.id, ...outcome.receipt };
        job.status = 'done';
        job.error = '';
        completed++;
      } catch (error) {
        job.status = 'failed';
        job.error = error.message;
        failures++;
        await storeRecords([job]).catch(storeError);
      }
      recalculate(ctx);
      if (workspace !== ctx) break;
      progress.value = index + 1;
      render();
      if (index < selected.length - 1) await sleep(PAUSE);
    }
    if (workspace === ctx) {
      const summary = `${completed} de ${selected.length} transcritas.${failures ? ` ${failures} fallidas.` : ''}`;
      jobMessage(pending ? `${pending} Se conservaron las fotos pendientes. ${summary}` : summary, !!pending || failures > 0);
      if (pending || failures) $('queue-wrap').open = true;
    }
  };
  try {
    if (navigator.locks) {
      await navigator.locks.request(`control-remitos:${ctx.scope}`, { ifAvailable: true }, lock =>
        lock ? run() : jobMessage('Ya hay una tanda en curso en otra pestaña de este navegador.', true));
    } else await run();
  } catch (error) {
    if (workspace === ctx) jobMessage(`No se pudo completar la tanda. ${error.message}`, true);
  } finally {
    busy = false;
    updateControls();
  }
}

function receiptList() {
  return [...localReceipts(workspace), ...workspace.remote.values()].sort((a, b) => b.id.localeCompare(a.id));
}

function resultFor(receipt, index) {
  const id = lineId(receipt, index);
  const line = receipt.lines[index];
  return workspace.comparison.results.get(id) || { id, receipt, index, line,
    date: datesFor(receipt)[index].delivery, chosen: line.elegida, state: line.estado || 'pending', decision: line.decision };
}

function formatRow(row, result, canChoose = false) {
  const printable = { ...row, day: row.day ?? excelDay(row.fecha) };
  const warnings = differences(result.receipt, result.line, result.date, printable).warnings;
  return `<div class="row-detail"><b>Fila ${row.fila} · ${html(row.herramienta)}</b>
    <div>${html(row.fecha || 'Fecha vacía')} · ${html(row.persona || 'Persona vacía')} · ${html(row.obra || 'Obra vacía')}</div>
    <div>Detalle: ${html(row.detalle || '—')} · Cantidad: ${html(row.cantidad || '—')}</div>
    <div>MOVIMIENTO: ${html(row.movimiento || '—')} · Nº REMITO: ${html(row.remito || '—')}</div>
    ${warnings.length ? `<div class="warning">${warnings.map(html).join(' · ')}</div>` : ''}
    ${canChoose ? `<button type="button" class="small" data-pick="${row.fila}">Elegir fila ${row.fila}</button>` : ''}</div>`;
}

function readField(label, value, wide = false) {
  return `<label class="field${wide ? ' wide' : ''}">${label}<span class="read-field">${html(value || '—')}</span></label>`;
}

function editField(field, label, value, editable, wide = false) {
  if (!editable) return readField(label, value, wide);
  const editor = ['observacion', 'descripcion', 'nota_manuscrita'].includes(field)
    ? `<textarea rows="2" data-field="${field}" aria-label="${label}">${html(value)}</textarea>`
    : `<input type="text" data-field="${field}" aria-label="${label}" value="${html(value)}">`;
  return `<label class="field${wide ? ' wide' : ''}">${label}${editor}</label>`;
}

function lineMarkup(line, index, editable) {
  const highlight = editable ? `<label class="check"><input type="checkbox" data-field="resaltado" ${line.resaltado ? 'checked' : ''}> Resaltado rosa</label>`
    : `<span class="hint">${line.resaltado ? 'Resaltado rosa' : 'Sin resaltar'}</span>`;
  const confidence = editable ? `<label class="field">Confianza <select data-field="confianza">${['alta', 'media', 'baja'].map(value => `<option${line.confianza === value ? ' selected' : ''}>${value}</option>`).join('')}</select></label>`
    : readField('Confianza', line.confianza);
  const compact = key => editable
    ? `<input type="text" data-field="${key}" aria-label="${lineFields[key]}" value="${html(line[key])}">`
    : `<span class="read-field">${html(line[key] || '—')}</span>`;
  return `<tr data-line="${index}"><td data-label="Fecha">${compact('fecha_linea')}<span class="hint" data-effective-date></span></td>
    <td data-label="Descripción">${editable ? `<textarea rows="2" data-field="descripcion" aria-label="Descripción">${html(line.descripcion)}</textarea>` : `<span class="read-field">${html(line.descripcion || '—')}</span>`}
      ${highlight}<details><summary>Devolución y notas <span class="hint" data-return-date></span></summary><div class="edit-grid">
      ${editField('fecha_devolucion', 'Fecha devuelta', line.fecha_devolucion, editable)}
      ${editField('cantidad_devuelta', 'Cantidad devuelta', line.cantidad_devuelta, editable)}
      ${editField('nota_manuscrita', 'Nota manuscrita', line.nota_manuscrita, editable)}${confidence}
      ${editable ? '<div><button type="button" class="small" data-action="remove">Quitar línea</button></div>' : ''}
      </div></details></td><td data-label="Cantidad">${compact('cantidad')}</td><td data-label="Comparación" data-result></td></tr>`;
}

function makeCard(receipt, job) {
  const card = document.createElement('article');
  card.className = 'receipt';
  card.dataset.id = receipt.id;
  const editable = !!job;
  let photo = '<p class="hint">Foto disponible en el navegador donde se cargó.</p>';
  if (job) {
    if (!photoUrls.has(receipt.id)) photoUrls.set(receipt.id, URL.createObjectURL(job.preview || job.source));
    photo = `<figure tabindex="0" aria-label="Foto del remito"><img loading="lazy" src="${photoUrls.get(receipt.id)}" alt="Foto de ${html(job.name)}"></figure>
      <label class="check">Zoom <input type="range" min="100" max="400" step="25" value="100" aria-label="Zoom de la foto"></label>
      <p class="hint">${html(job.name)}</p>`;
  }
  card.innerHTML = `<h3></h3><div class="receipt-head">${Object.entries(headerFields).map(([field, label]) =>
    editField(field, label, receipt[field], editable, field === 'observacion')).join('')}</div>
    <div class="receipt-layout"><div class="photo-block">${photo}</div><div class="review">
      ${receipt.lines.length ? `<div class="table-wrap"><table aria-label="Líneas del remito"><thead><tr><th>Fecha</th><th>Descripción</th><th>Cant.</th><th>Comparación</th></tr></thead>
      <tbody>${receipt.lines.map((line, index) => lineMarkup(line, index, editable)).join('')}</tbody></table></div>`
    : '<p class="empty">El remito no tiene líneas transcritas.</p>'}
      ${editable ? '<div class="receipt-foot"><button type="button" class="small" data-action="add">Agregar línea</button></div>' : ''}</div></div>`;
  return card;
}

function renderStatus(cell, result, editable) {
  const signature = JSON.stringify([result.state, result.chosen, result.decision, result.conflict, result.date, result.line.cantidad,
    result.receipt.numero_remito, result.candidates?.map(({ row }) => row.fila), !!workspace.excel]);
  if (cell.dataset.signature === signature) return;
  cell.dataset.signature = signature;
  const chosen = result.chosen ? formatRow(result.chosen, result) : '';
  const action = (id, label) => `<button type="button" class="small" data-action="${id}">${label}</button>`;
  const buttons = [];
  if (editable && result.state !== 'skipped') {
    if (result.chosen && !result.decision) buttons.push(action('confirm', 'Confirmar'));
    if (workspace.excel) buttons.push(action('choose', result.chosen ? 'Elegir otra fila' : result.candidates?.length ? 'Ver candidatos' : 'Buscar en Excel'));
    if (result.decision !== 'absent') buttons.push(action('absent', 'No está en el Excel'));
    if (result.decision) buttons.push(action('undo', 'Reabrir'));
  }
  cell.innerHTML = `<span class="status ${result.state}">${states[result.state]}</span>${chosen}
    ${result.decision ? `<p class="details-text">${decisions[result.decision]}</p>` : ''}
    ${result.conflict ? '<p class="details-text">La fila ya está asignada a otra línea.</p>' : ''}
    ${result.state === 'pending' && editable ? '<p class="hint">Cargá el Excel para comparar.</p>' : ''}
    ${buttons.length ? `<div class="actions">${buttons.join('')}</div>` : ''}`;
}

function selectOptions(id, values, all) {
  const select = $(id);
  const previous = select.value;
  const items = [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es-AR'));
  const signature = JSON.stringify(items);
  if (select.dataset.options === signature) return;
  select.dataset.options = signature;
  select.replaceChildren(new Option(all, ''), ...items.map(value => new Option(value, value)));
  if (items.includes(previous)) select.value = previous;
}

function render() {
  const ctx = workspace;
  if (!ctx) return;
  const receipts = receiptList();
  const ids = new Set(receipts.map(receipt => receipt.id));
  for (const [id, entry] of cards) {
    if (!ids.has(id)) {
      entry.card.remove();
      cards.delete(id);
    }
  }
  selectOptions('person', receipts.map(receipt => receipt.persona), 'Todas');
  selectOptions('work', receipts.map(receipt => receipt.obra_o_vehiculo), 'Todas');
  const totals = { ok: 0, doubtful: 0, missing: 0, reviewed: 0 };
  let visible = 0;
  const container = $('receipts');
  receipts.forEach((receipt, position) => {
    const job = ctx.local.get(receipt.id);
    let entry = cards.get(receipt.id);
    if (!entry || entry.length !== receipt.lines.length || entry.receipt !== receipt) {
      const card = makeCard(receipt, job);
      entry?.card.remove();
      entry = { card, length: receipt.lines.length, receipt };
      cards.set(receipt.id, entry);
    }
    const { card } = entry;
    if (container.children[position] !== card) container.insertBefore(card, container.children[position] || null);
    card.querySelector('h3').textContent = receipt.numero_remito ? `Remito ${receipt.numero_remito}` : 'Remito sin número';
    const dates = datesFor(receipt);
    const rows = card.querySelectorAll('tr[data-line]');
    let compared = 0, finished = 0, rowVisible = false;
    rows.forEach((row, index) => {
      const result = resultFor(receipt, index);
      if (['ok', 'doubtful', 'missing'].includes(result.state)) totals[result.state]++;
      if (result.state !== 'skipped') compared++;
      if (result.decision) finished++;
      row.classList.toggle('highlighted', result.line.resaltado);
      row.hidden = $('only-issues').checked && !['doubtful', 'missing'].includes(result.state);
      rowVisible ||= !row.hidden;
      const label = dates[index].label;
      row.querySelector('[data-effective-date]').textContent = label !== result.line.fecha_linea ? label : '';
      row.querySelector('[data-return-date]').textContent = dates[index].returnLabel ? `· ${dates[index].returnLabel}` : '';
      renderStatus(row.querySelector('[data-result]'), result, !!job);
    });
    if (compared && finished === compared) totals.reviewed++;
    card.hidden = ($('person').value && $('person').value !== receipt.persona) || ($('work').value && $('work').value !== receipt.obra_o_vehiculo)
      || ($('only-issues').checked && !rowVisible);
    if (!card.hidden) visible++;
  });
  for (const [key, count] of Object.entries(totals)) $(`count-${key}`).textContent = count;
  $('empty').hidden = visible > 0;
  $('empty').textContent = receipts.length ? 'No hay remitos con esos filtros.' : 'Agregá fotos para empezar o abrí el control desde una sala del planificador para ver sus resultados.';
  $('export-xlsx').disabled = $('export-csv').disabled = !receipts.some(receipt => receipt.lines.length);
  updateControls();
  queueList();
  updateSync();
}

function resetLine(line) {
  line.elegida = null;
  line.decision = '';
  line.estado = '';
}

function editReceipt(event) {
  const field = event.target.dataset.field;
  const id = event.target.closest('article')?.dataset.id;
  const receipt = workspace?.local.get(id)?.receipt;
  if (!field || !receipt) return;
  const index = event.target.closest('tr')?.dataset.line;
  const value = field === 'resaltado' ? event.target.checked : text(event.target.value);
  const datesBefore = datesFor(receipt);
  if (index === undefined) {
    if (receipt[field] === value) return;
    receipt[field] = value;
    if (['numero_remito', 'persona', 'obra_o_vehiculo'].includes(field)) receipt.lines.forEach(resetLine);
  } else {
    const line = receipt.lines[index];
    if (line[field] === value) return;
    line[field] = value;
    if (['fecha_linea', 'descripcion', 'cantidad'].includes(field)) resetLine(line);
  }
  if (field === 'fecha' || field === 'fecha_linea') {
    datesFor(receipt).forEach((date, i) => {
      if (!same(date.delivery, datesBefore[i].delivery)) resetLine(receipt.lines[i]);
    });
  }
  recalculate(workspace);
  // Dejar que el clic que sacó el foco llegue al botón antes de actualizar los resultados.
  setTimeout(render, 0);
}

function doLineAction(event) {
  const button = event.target.closest('button[data-action]');
  const id = button?.closest('article')?.dataset.id;
  const receipt = workspace?.local.get(id)?.receipt;
  if (!receipt) return;
  const index = Number(button.closest('tr')?.dataset.line);
  const action = button.dataset.action;
  if (action === 'add') {
    const line = Object.fromEntries(Object.keys(lineFields).map(field => [field, field === 'resaltado' ? false : '']));
    receipt.lines.push({ ...line, confianza: 'media' });
  } else if (action === 'remove') {
    if (!window.confirm('¿Quitar esta línea de la transcripción?')) return;
    const before = datesFor(receipt);
    receipt.lines.splice(index, 1);
    datesFor(receipt).forEach((date, i) => {
      if (!same(date.delivery, before[i + (i >= index ? 1 : 0)].delivery)) resetLine(receipt.lines[i]);
    });
  } else if (action === 'choose') {
    openChooser(receipt, index);
    return;
  } else {
    const result = resultFor(receipt, index);
    if (action === 'confirm' && result.chosen) Object.assign(result.line, { decision: 'confirmed', elegida: result.chosen });
    if (action === 'absent') Object.assign(result.line, { decision: 'absent', elegida: null });
    if (action === 'undo') resetLine(result.line);
  }
  recalculate(workspace);
  render();
  if (action === 'add') cards.get(id)?.card.querySelector('tr:last-child textarea')?.focus();
}

function openChooser(receipt, index) {
  chooser = { id: receipt.id, index, page: 0 };
  $('manual-search').value = '';
  renderChooser();
  $('chooser').showModal();
  $('manual-search').focus();
}

function rowAvailable(row, result) {
  const comparison = workspace.comparison;
  return !comparison.sharedRows.has(row.fila) && (!comparison.used.has(row.fila) || comparison.used.get(row.fila) === result.id);
}

function renderChooser() {
  const receipt = workspace?.local.get(chooser?.id)?.receipt;
  if (!receipt || !workspace.excel || !receipt.lines[chooser.index]) return;
  const result = resultFor(receipt, chooser.index);
  $('chooser-line').textContent = result.line.descripcion || 'Línea sin descripción';
  $('suggestions').innerHTML = result.candidates?.length ? '<h3>Candidatos sugeridos</h3>' + result.candidates.map(({ row }) => formatRow(row, result, true)).join('') : '';
  const words = normalize($('manual-search').value).split(' ').filter(Boolean);
  const rows = workspace.excel.items.filter(row => words.every(word => row.search.includes(word)) && rowAvailable(row, result));
  const pageSize = 12;
  const pages = Math.max(1, Math.ceil(rows.length / pageSize));
  chooser.page = Math.max(0, Math.min(chooser.page, pages - 1));
  $('manual-results').innerHTML = rows.slice(chooser.page * pageSize, (chooser.page + 1) * pageSize).map(row => formatRow(row, result, true)).join('')
    || '<p class="hint">No hay filas disponibles con esa búsqueda.</p>';
  $('manual-pagination').hidden = pages <= 1;
  $('manual-prev').disabled = chooser.page === 0;
  $('manual-next').disabled = chooser.page >= pages - 1;
  $('manual-page').textContent = `${chooser.page + 1} / ${pages} · ${rows.length} filas`;
}

function chooseRow(event) {
  const number = Number(event.target.closest('[data-pick]')?.dataset.pick);
  const row = workspace?.excel?.byRow.get(number);
  const receipt = workspace?.local.get(chooser?.id)?.receipt;
  if (!row || !receipt) return;
  const result = resultFor(receipt, chooser.index);
  if (!rowAvailable(row, result)) {
    renderChooser();
    return notice('Esa fila acaba de asignarse a otra línea. Elegí una fila disponible.', true);
  }
  Object.assign(result.line, { decision: 'manual', elegida: chosenRow(row) });
  $('chooser').close();
  recalculate(workspace);
  render();
}

async function loadWorkbook(file) {
  const ctx = workspace;
  if (!ctx || !file || loadingExcel) return;
  loadingExcel = true;
  updateControls();
  $('excel-info').textContent = 'Leyendo Excel…';
  notice('');
  try {
    if (!/\.xlsx$/i.test(file.name)) throw new Error('Elegí un archivo .xlsx.');
    const xlsx = await dependency('XLSX', SHEET_JS);
    const rows = readWorkbook(xlsx, await file.arrayBuffer());
    const workbook = { id: 'excel', scope: ctx.scope, name: file.name, rows };
    await storeRecords([workbook]);
    const excel = indexWorkbook(rows);
    for (const receipt of localReceipts(ctx)) {
      for (const line of receipt.lines) {
        if (!line.elegida) continue;
        const refreshed = excel.byRow.get(line.elegida.fila);
        const identity = row => ['fecha', 'codigo', 'herramienta', 'persona', 'obra'].map(key => normalize(row?.[key]));
        if (!refreshed || !same(identity(refreshed), identity(line.elegida))) resetLine(line);
        else line.elegida = chosenRow(refreshed);
      }
    }
    ctx.workbook = workbook;
    ctx.excel = excel;
    recalculate(ctx);
    render();
  } catch (error) {
    if (workspace === ctx) notice(`No se pudo cargar el Excel. ${error.message}`, true);
  } finally {
    loadingExcel = false;
    updateControls();
  }
}

async function download(format) {
  if (!workspace) return;
  const rows = exportRows(receiptList(), workspace.comparison.results);
  if (rows.length <= 1) return;
  const name = `control-remitos-${new Intl.DateTimeFormat('sv-SE', { timeZone: 'America/Argentina/Buenos_Aires' }).format(new Date())}`;
  try {
    if (format === 'xlsx') {
      const xlsx = await dependency('XLSX', SHEET_JS);
      const book = xlsx.utils.book_new();
      const sheet = xlsx.utils.aoa_to_sheet(rows);
      sheet['!cols'] = [14, 26, 24, 15, 48, 13, 24, 12, 23].map(wch => ({ wch }));
      xlsx.utils.book_append_sheet(book, sheet, 'Control de remitos');
      xlsx.writeFile(book, `${name}.xlsx`);
    } else {
      // El BOM conserva acentos en Excel. El escape evita que un texto transcripto se interprete como fórmula.
      const csvCell = value => {
        const original = String(value ?? '');
        const safe = /^[\s\uFEFF]*[=+\-@]|^[\t\r\n]/.test(original) ? `'${original}` : original;
        return `"${safe.replace(/"/g, '""')}"`;
      };
      const csv = rows.map(row => row.map(csvCell).join(';')).join('\r\n');
      const url = URL.createObjectURL(new Blob(['\uFEFF', csv], { type: 'text/csv;charset=utf-8' }));
      const link = document.createElement('a');
      link.href = url;
      link.download = `${name}.csv`;
      document.body.append(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    }
  } catch (error) { notice(`No se pudo exportar. ${error.message}`, true); }
}

$('settings-form').addEventListener('submit', async event => {
  event.preventDefault();
  const value = text($('gemini-key').value);
  if (!validKey(value)) return keyFeedback('Pegá una clave de Gemini válida.', true);
  if (activeKeys(workspace).some(entry => entry.value === value)) return keyFeedback('Esa clave ya está en la lista.', true);
  if (await editKeys(entries => entries.some(entry => entry.value === value) ? entries : [...entries, { value }])) {
    $('gemini-key').value = '';
  }
});
$('gemini-keys').addEventListener('click', event => {
  const button = event.target.closest('[data-key-action]');
  if (!button) return;
  const selected = activeKeys(workspace)[Number(button.dataset.keyIndex)];
  if (!selected) return;
  editKeys(entries => {
    const index = entries.findIndex(entry => entry.value === selected.value);
    if (index < 0) return entries;
    if (button.dataset.keyAction === 'delete') return entries.filter(entry => entry.value !== selected.value);
    const position = index + (button.dataset.keyAction === 'up' ? -1 : 1);
    if (position < 0 || position >= entries.length) return entries;
    [entries[index], entries[position]] = [entries[position], entries[index]];
    return [...entries];
  });
});
$('share-keys').onchange = event => {
  try {
    const enabled = event.target.checked;
    localStorage.setItem(SHARE_KEYS, enabled ? '1' : '0');
    if (workspace) listenKeys(workspace);
    keyFeedback(!enabled && workspace?.scope !== 'local' ? 'Las claves ya publicadas siguen disponibles para el resto de la sala.' : '');
  } catch { keyFeedback('No se pudo guardar el modo de claves en este navegador.', true); }
  renderKeys();
};
$('save-model').onclick = () => {
  const model = modelName();
  if (!/^[a-z0-9][a-z0-9._-]*$/i.test(model)) return keyFeedback('Ingresá un nombre de modelo de Gemini válido.', true);
  try {
    localStorage.setItem(MODEL, model);
    $('gemini-model').value = model;
    keyFeedback('Modelo guardado.');
  } catch { keyFeedback('No se pudo guardar el modelo en este navegador.', true); }
};
$('choose-excel').onclick = () => $('excel-file').click();
$('excel-file').onchange = event => { loadWorkbook(event.target.files[0]); event.target.value = ''; };
$('choose-images').onclick = () => $('images').click();
$('images').onchange = event => { addPhotos([...event.target.files]); event.target.value = ''; };
$('process').onclick = () => processPhotos('queued');
$('retry').onclick = () => processPhotos('failed');
$('sync-retry').onclick = () => { if (workspace) connect(workspace, parentConnection()); };
$('include-all').onchange = event => {
  const ctx = workspace;
  ctx.includeAll = event.target.checked;
  storeRecords([{ id: 'options', scope: ctx.scope, includeAll: ctx.includeAll }]).catch(storeError);
  recalculate(ctx);
  render();
};
for (const id of ['person', 'work', 'only-issues']) $(id).onchange = render;
$('receipts').addEventListener('change', editReceipt);
$('receipts').addEventListener('click', doLineAction);
$('receipts').addEventListener('input', event => {
  if (event.target.type === 'range') event.target.closest('.photo-block').querySelector('img').style.width = `${event.target.value}%`;
});
$('chooser').addEventListener('click', chooseRow);
$('chooser').addEventListener('close', () => { chooser = null; });
$('manual-search').oninput = () => { chooser.page = 0; renderChooser(); };
$('manual-prev').onclick = () => { chooser.page--; renderChooser(); };
$('manual-next').onclick = () => { chooser.page++; renderChooser(); };
$('export-xlsx').onclick = () => download('xlsx');
$('export-csv').onclick = () => download('csv');

$('gemini-model').value = localStorage.getItem(MODEL) || DEFAULT_MODEL;
refreshRoom();
window.addEventListener('storage', refreshRoom);
window.addEventListener('focus', refreshRoom);
// El evento storage no se emite cuando el planificador cambia sessionStorage en este documento superior.
setInterval(() => { refreshRoom(); renderKeys(); }, 1500);
