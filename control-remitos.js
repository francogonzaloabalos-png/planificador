import {
  headerFields, lineFields, states, decisions, normalize, text, datesFor, readWorkbook,
  differences, indexWorkbook, chosenRow, lineId, compareReceipts, transcribe, publicReceipt,
  receiveReceipt, exportRows, excelDay, GeminiQuotaError, publicLine, newLineKey, excelBlocks, GEMINI_MODELS, materialSimilarity
} from './control-remitos-core.js?v=revision-remitos';

const $ = id => document.getElementById(id);
const LEGACY_KEY = 'control_remitos_gemini_key_v1';
const LOCAL_KEYS = 'control_remitos_gemini_keys_v1';
const SHARE_KEYS = 'control_remitos_share_keys_v1';
const MODEL_CHOICE = 'control_remitos_model_choice_v1';
const REVIEWER = 'control_remitos_reviewer_v1';
const VERIFIED_FILTER = 'control_remitos_verified_filter_v1';
const MAX_BATCH = 10;
const MAX_EDGE = 1600;
const PAUSE = 1800;
const SHEET_JS = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const HEIC = 'https://cdn.jsdelivr.net/npm/heic2any@0.0.4/dist/heic2any.min.js';
const html = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const sleep = delay => new Promise(resolve => setTimeout(resolve, delay));
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const newId = () => `r${Date.now().toString(36)}-${crypto.randomUUID()}`;
const loadTime = new Intl.DateTimeFormat('es-AR', {
  timeZone: 'America/Argentina/Buenos_Aires', dateStyle: 'short', timeStyle: 'short', hour12: false
});

let workspace;
let roomGeneration = 0;
let targetScope;
let busy = false;
let loadingExcel = false;
let chooser = null;
const cards = new Map();
const photoUrls = new Map();
let previewUrl;
let previewTicket = 0;
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

async function storeRecords(records, deletes = []) {
  if (!records.length && !deletes.length) return;
  const snapshot = structuredClone(records);
  const db = await database;
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('records', 'readwrite');
    transaction.oncomplete = resolve;
    transaction.onabort = () => reject(transaction.error || new Error('Se interrumpió el guardado local.'));
    const store = transaction.objectStore('records');
    for (const record of snapshot) store.put(record);
    for (const key of deletes) store.delete(key);
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
  excelFeedback('');
  if ($('chooser').open) $('chooser').close();
  if ($('photo-preview').open) $('photo-preview').close();
  for (const [id] of photoUrls) forgetPhoto(id);
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
    const excelUpload = records.find(record => record.id === 'excel-queued');
    const workbook = excelUpload || records.find(record => record.id === 'excel' && record.name !== 'BASE DE DATOS (CSV)');
    const local = new Map(records.filter(record => record.type === 'photo').map(job => [job.id, job]));
    const values = records.find(record => record.id === 'cloud')?.values || {};
    workspace = {
      scope, workbook, excel: workbook ? indexWorkbook(workbook.rows) : null, local, selectedPhotos: new Set(),
      values, remote: new Map(), excelUpload, patches: new Map(records.filter(record => record.type === 'patch')
        .sort((a, b) => a.id.localeCompare(b.id)).map(record => [record.id, record])),
      includeAll: records.find(record => record.id === 'options')?.includeAll === true,
      comparison: null, sync: null, keyBlocks: new Map()
    };
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
  try {
    if (sync.onKeys) sync.configRef.off('value', sync.onKeys);
    if (sync.onExcel) sync.excelRef.child('actual').off('value', sync.onExcel);
    sync.ref.off('value', sync.onValue);
    sync.connected.off('value', sync.onConnected);
  } catch { /* El planificador puede haber cerrado la app compartida. */ }
}

function listenKeys(ctx) {
  const sync = ctx.sync;
  if (!sync?.configRef || !!sync.onKeys === usesRoomKeys(ctx)) return;
  if (sync.onKeys) {
    sync.configRef.off('value', sync.onKeys);
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
  sync.configRef.on('value', onKeys, () => {
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
    const room = sync.room = db.ref('controlRemitos').child(connection.room);
    sync.ref = room.child('remitos');
    sync.configRef = room.child('config');
    sync.excelRef = room.child('excel');
    sync.photoWork = new Set();
    sync.photoDone = new Set();
    listenKeys(ctx);
    sync.onExcel = snapshot => {
      if (ctx.sync === sync) receiveExcel(ctx, sync, snapshot.val());
    };
    sync.excelRef.child('actual').on('value', sync.onExcel, () => {
      if (ctx.sync === sync) excelFeedback('No se pudo leer el Excel de la sala.', true);
    });
    sync.connected = db.ref('.info/connected');
    sync.onValue = snapshot => {
      if (ctx.sync !== sync) return;
      sync.ready = true;
      const values = snapshot.val() || {};
      ctx.values = values;
      storeRecords([{ id: 'cloud', scope: ctx.scope, values }]).catch(storeError);
      recalculate(ctx);
      render();
      flush(ctx);
      flushPhotos(ctx);
    };
    sync.onConnected = snapshot => {
      if (ctx.sync !== sync) return;
      sync.online = snapshot.val() === true;
      updateSync();
      renderKeys();
      flush(ctx);
      flushExcel(ctx);
      flushPhotos(ctx);
      updateControls();
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

function withReceiptChanges(values, fields) {
  const next = { ...values };
  for (const [path, value] of Object.entries(fields)) {
    const parts = path.split('/');
    const field = parts.pop();
    let node = next;
    for (const part of parts) node = node[part] = { ...node[part] };
    if (value === null) delete node[field];
    else node[field] = value;
  }
  return next;
}

function recalculate(ctx) {
  let values = { ...ctx.values };
  for (const job of ctx.local.values()) {
    if (job.receipt && !values[job.id]) values[job.id] = job.payload || publicReceipt(job.receipt);
  }
  for (const patch of ctx.patches.values()) values = withReceiptChanges(values, patch.fields);
  const receipts = new Map();
  for (const [id, value] of Object.entries(values)) {
    if (!/^r[a-z0-9-]+$/i.test(id)) continue;
    const receipt = receiveReceipt(id, value);
    if (receipt) receipts.set(id, receipt);
  }
  ctx.remote = receipts;
  ctx.comparison = compareReceipts(receipts.values(), [], ctx.excel, ctx.includeAll);
}

function updatesAllowed(ctx) {
  const current = parentConnection();
  return workspace === ctx && ctx.sync?.online && current?.scope === ctx.scope && current.app === ctx.sync.app;
}

function updateSync() {
  const ctx = workspace;
  if (!ctx) return;
  const sync = ctx.sync;
  const pending = [...ctx.local.values()].filter(job => job.pending).length + ctx.patches.size + Number(!!ctx.excelUpload);
  const error = sync?.error || sync?.photoError || sync?.excelError;
  $('sync-retry').hidden = !error;
  let status = 'Modo local';
  if (ctx.scope !== 'local') {
    const suffix = pending ? ' · ' + pending + ' por publicar' : '';
    if (error) status = 'Error de sincronización' + suffix;
    else if (!sync?.ready || !sync.online) status = 'Sala sin conexión' + suffix;
    else status = pending ? 'Sincronizando' + suffix : 'Sala sincronizada';
  }
  $('sync').textContent = status;
}

async function flush(ctx) {
  const sync = ctx.sync;
  if (!updatesAllowed(ctx) || !sync?.ready || sync.writing || sync.error) return;
  const creations = [...ctx.local.values()].filter(job => job.pending);
  const patches = [...ctx.patches.values()];
  if (!creations.length && !patches.length) return;
  sync.writing = true;
  try {
    for (const job of creations) {
      if (!updatesAllowed(ctx) || ctx.sync !== sync) break;
      const result = await sync.ref.child(job.id).transaction(current => current == null ? job.payload : undefined, undefined, false);
      if (ctx.sync !== sync) return;
      if (!ctx.values[job.id] && result.snapshot.val()) ctx.values = { ...ctx.values, [job.id]: result.snapshot.val() };
      await storeRecords([{ ...job, pending: false }]);
      job.pending = false;
    }
    for (const patch of patches) {
      if (!updatesAllowed(ctx) || ctx.sync !== sync) break;
      const previous = ctx.values;
      await sync.ref.update(patch.fields);
      if (ctx.sync !== sync) return;
      if (ctx.values === previous) ctx.values = withReceiptChanges(ctx.values, patch.fields);
      await storeRecords([], [[ctx.scope, patch.id]]);
      ctx.patches.delete(patch.id);
    }
  } catch {
    if (ctx.sync === sync) sync.error = true;
  } finally {
    sync.writing = false;
    if (ctx.sync === sync) {
      recalculate(ctx);
      render();
      flushPhotos(ctx);
      flush(ctx);
    }
  }
}

async function changeReceipt(fields) {
  const ctx = workspace;
  try {
    if (ctx.scope === 'local') {
      ctx.values = withReceiptChanges(ctx.values, fields);
      const receiptId = Object.keys(fields)[0].split('/')[0];
      // Los trabajos antiguos pueden tener el remito únicamente en IndexedDB.
      if (!ctx.values[receiptId].version) {
        const job = ctx.local.get(receiptId);
        ctx.values = withReceiptChanges({ ...ctx.values, [receiptId]: job.payload || publicReceipt(job.receipt) }, fields);
      }
      recalculate(ctx);
      await storeRecords([{ id: 'cloud', scope: ctx.scope, values: ctx.values }]);
    } else {
      const patch = { id: 'patch-' + newId(), scope: ctx.scope, type: 'patch', fields };
      ctx.patches.set(patch.id, patch);
      recalculate(ctx);
      await storeRecords([patch]);
      flush(ctx);
    }
  } catch { storeError(); }
  if (workspace === ctx) render();
}

const prefersSharing = () => localStorage.getItem(SHARE_KEYS) !== '0';
const usesRoomKeys = ctx => !!ctx && ctx.scope !== 'local' && prefersSharing();
// Las claves se envían como tokens opacos. Las de autorización también pueden incluir puntos.
const validKey = value => typeof value === 'string' && /^[A-Za-z0-9._~+/-]{8,}={0,2}$/.test(value);
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
  $('gemini-model').disabled = busy;
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

async function updateRoomConfig(ctx, update) {
  const sync = ctx.sync;
  const connection = parentConnection();
  if (!sync?.configRef || connection?.scope !== ctx.scope || connection.app !== sync.app) throw new Error();
  return sync.configRef.transaction(current => update(current || {}), undefined, false);
}

async function changeSharedKeys(ctx, update) {
  const sync = ctx.sync;
  const outcome = await updateRoomConfig(ctx, current => {
    const entries = decodeKeys(current);
    const updated = update(entries);
    return updated === entries ? undefined : { ...current, ...encodeKeys(updated) };
  });
  if (ctx.sync === sync && usesRoomKeys(ctx) && outcome.committed) {
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

function excelFeedback(message, error = false) {
  const feedback = $('excel-feedback');
  feedback.textContent = message;
  feedback.hidden = !message;
  feedback.classList.toggle('error', error);
}

function updateControls() {
  const ctx = workspace;
  $('choose-excel').disabled = !ctx || loadingExcel || !!ctx.sync?.writingExcel;
  $('choose-excel').textContent = ctx?.workbook ? 'Actualizar (subir de nuevo)' : 'Cargar Excel (.xlsx)';
  $('remove-excel').hidden = !ctx?.workbook && !ctx?.sync?.excelMeta;
  $('remove-excel').disabled = loadingExcel || !!ctx?.sync?.writingExcel;
  $('remove-excel').textContent = ctx?.scope === 'local' ? 'Quitar Excel local' : 'Quitar Excel de la sala';
  $('choose-images').disabled = !ctx || busy;
  $('process').disabled = !ctx || busy || !ctx.selectedPhotos.size;
  for (const input of $('queue').querySelectorAll('[data-queue-select]')) input.disabled = busy;
  $('include-all').disabled = !ctx || !ctx.excel;
  $('include-all').checked = !!ctx?.includeAll;
  const workbook = ctx?.workbook;
  const time = workbook?.loadedAt ? ` · ${loadTime.format(workbook.loadedAt)}` : '';
  const count = workbook?.rows.length;
  const status = ctx?.excelUpload ? ' · pendiente de publicar' : '';
  $('excel-info').textContent = loadingExcel || ctx?.sync?.readingExcel ? 'Cargando Excel…'
    : workbook ? `${count.toLocaleString('es-AR')} ${count === 1 ? 'fila' : 'filas'} · ${workbook.name}${time}${status}` : 'Excel sin cargar';
  renderKeys();
}

function queueList() {
  const ctx = workspace;
  const jobs = [...ctx.local.values()].sort((a, b) => b.id.localeCompare(a.id));
  $('queue-wrap').hidden = !jobs.length;
  const markup = jobs.map(job => {
    const name = html(job.name);
    const selectable = job.status !== 'done';
    const canPreview = job.source && ctx.remote.get(job.id)?.photo !== false;
    const title = selectable ? `<label class="check queue-name"><input type="checkbox" data-queue-select="${job.id}" aria-label="Seleccionar ${name} para transcribir"> <strong>${name}</strong></label>`
      : `<span class="queue-name"><strong>${name}</strong> · Transcripta</span>`;
    return `<li><div class="queue-item">${title}${canPreview ? `<button type="button" class="small" data-queue-preview="${job.id}" aria-label="Ver foto ${name}">Ver foto</button>` : ''}</div>
      ${job.error ? `<p class="hint error">${html(job.error)}</p>` : ''}</li>`;
  }).join('');
  const queue = $('queue');
  if (queue.dataset.markup !== markup) {
    queue.dataset.markup = markup;
    queue.innerHTML = markup;
  }
  for (const input of queue.querySelectorAll('[data-queue-select]')) {
    input.checked = ctx.selectedPhotos.has(input.dataset.queueSelect);
    input.disabled = busy;
  }
}

function selectQueuedPhoto(event) {
  const input = event.target;
  const ctx = workspace;
  const id = input.dataset.queueSelect;
  const job = ctx?.local.get(id);
  if (!job || busy || job.status === 'done') return;
  if (!input.checked) ctx.selectedPhotos.delete(id);
  else if (ctx.selectedPhotos.size < MAX_BATCH) ctx.selectedPhotos.add(id);
  else {
    input.checked = false;
    return jobMessage('Marcá hasta 10 fotos por tanda.', true);
  }
  updateControls();
}

async function previewQueuedPhoto(event) {
  const id = event.target.closest('[data-queue-preview]')?.dataset.queuePreview;
  const job = workspace?.local.get(id);
  if (!job?.source) return;
  const dialog = $('photo-preview');
  const image = $('preview-image');
  const status = $('preview-status');
  const ticket = ++previewTicket;
  $('preview-title').textContent = job.name;
  image.hidden = true;
  status.textContent = 'Cargando foto…';
  dialog.showModal();
  try {
    const source = await jpegSource(job);
    if (ticket !== previewTicket || !dialog.open) return;
    previewUrl = URL.createObjectURL(source);
    image.onload = () => { image.hidden = false; status.textContent = ''; };
    image.onerror = () => { status.textContent = 'No se pudo abrir la foto.'; };
    image.src = previewUrl;
  } catch (error) {
    if (ticket === previewTicket && dialog.open) status.textContent = error.message;
  }
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

async function resizePhoto(source, maxEdge = MAX_EDGE, quality = .89) {
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
    const jpeg = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', quality));
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

function flushPhotos(ctx) {
  const sync = ctx.sync;
  if (!updatesAllowed(ctx) || !sync?.ready || sync.photoWork.size || sync.photoError) return;
  const job = [...ctx.local.values()].find(job => job.source && ctx.values[job.id]?.version === 1
    && ctx.values[job.id].photo === undefined && !sync.photoDone.has(job.id));
  if (!job) return;
  sync.photoWork.add(job.id);
  (async () => {
    try {
      const jpeg = await encodePhoto(await resizePhoto(await jpegSource(job), 1000, .6));
      if (ctx.sync !== sync || !updatesAllowed(ctx) || ctx.values[job.id]?.photo !== undefined) return;
      await sync.room.update({ [`imagenes/${job.id}`]: jpeg, [`remitos/${job.id}/photo`]: true });
      sync.photoDone.add(job.id);
    } catch {
      if (ctx.sync === sync) sync.photoError = true;
    } finally {
      sync.photoWork.delete(job.id);
      updateSync();
      if (sync.photoDone.has(job.id)) flushPhotos(ctx);
    }
  })();
}

function forgetPhoto(id) {
  const url = photoUrls.get(id);
  if (url?.startsWith('blob:')) URL.revokeObjectURL(url);
  photoUrls.delete(id);
}

async function photoAction(event) {
  const button = event.target.closest('button[data-photo]');
  const id = button?.closest('article')?.dataset.id;
  const ctx = workspace;
  const receipt = ctx?.remote.get(id);
  if (!receipt) return;
  const job = ctx.local.get(id);
  const shared = ctx.scope !== 'local';
  if (button.dataset.photo === 'remove' && !window.confirm('¿Quitar solo la foto de este remito? Su transcripción se conserva.')) return;
  const feedback = button.closest('.photo-block').querySelector('[data-photo-feedback]');
  button.disabled = true;
  feedback.textContent = '';
  try {
    if (button.dataset.photo === 'remove') {
      if (shared) {
        if (!updatesAllowed(ctx)) throw new Error('Conectate a la sala para quitar la foto.');
        await ctx.sync.room.update({ [`imagenes/${id}`]: null, [`remitos/${id}/photo`]: false });
      }
      if (job) {
        delete job.source;
        delete job.preview;
        await storeRecords([job]);
      }
      forgetPhoto(id);
    } else if (!photoUrls.has(id)) {
      if (shared) {
        if (!ctx.sync?.online) throw new Error('Conectate a la sala para ver la foto.');
        const snapshot = await ctx.sync.room.child('imagenes').child(id).once('value');
        const jpeg = snapshot.val();
        if (typeof jpeg !== 'string' || !/^[A-Za-z0-9+/]+={0,2}$/.test(jpeg)) throw new Error('La foto no está disponible.');
        if (workspace !== ctx || ctx.remote.get(id)?.photo !== true) return;
        photoUrls.set(id, `data:image/jpeg;base64,${jpeg}`);
      } else if (job?.source) photoUrls.set(id, URL.createObjectURL(job.preview || job.source));
    }
    if (workspace === ctx) render();
  } catch (error) {
    feedback.textContent = error.message || 'No se pudo acceder a la foto.';
  } finally { button.disabled = false; }
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
    jobMessage('Marcá las fotos que querés transcribir.');
    queueList();
    $('queue-wrap').open = true;
    updateControls();
  } catch { jobMessage('No hay espacio o acceso para guardar las fotos en este navegador.', true); }
}

async function transcribeWithKeys(ctx, base64, modelId, onStatus) {
  const attempted = new Set();
  let waited = false;
  while (workspace === ctx) {
    const entries = activeKeys(ctx);
    const next = entries.find(entry => !attempted.has(entry.value) && !keyState(ctx, entry));
    if (next) {
      attempted.add(next.value);
      try {
        return { receipt: await transcribe({ key: next.value, base64, modelId }) };
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

async function processPhotos() {
  const ctx = workspace;
  if (busy || !ctx?.selectedPhotos.size) return;
  if (!activeKeys(ctx).length) {
    $('settings').open = true;
    $('gemini-key').focus();
    const waiting = usesRoomKeys(ctx) && !ctx.sync?.keys;
    return jobMessage(waiting ? 'No se pudieron leer las claves de la sala. También podés usar tus claves privadas.'
      : 'Agregá una clave de Gemini en Control de remitos.', true);
  }
  const modelId = $('gemini-model').value;
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
    const selected = [];
    for (const id of ctx.selectedPhotos) {
      const job = ctx.local.get(id);
      if (job && job.status !== 'done') selected.push(job);
      else ctx.selectedPhotos.delete(id);
    }
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
        const outcome = await transcribeWithKeys(ctx, base64, modelId, message => jobMessage(`${label} · ${message}`));
        if (outcome.pending) {
          pending = outcome.pending;
          break;
        }
        job.receipt = { id: job.id, ...outcome.receipt, nombre_archivo: job.name,
          lines: outcome.receipt.lines.map((line, order) => ({ ...line, key: newLineKey(), order })) };
        job.payload = publicReceipt(job.receipt);
        job.pending = ctx.scope !== 'local';
        job.status = 'done';
        job.error = '';
        ctx.selectedPhotos.delete(job.id);
        await storeRecords([job]).catch(storeError);
        completed++;
      } catch (error) {
        job.status = 'failed';
        job.error = error.message;
        failures++;
        await storeRecords([job]).catch(storeError);
      }
      recalculate(ctx);
      if (workspace !== ctx) break;
      flush(ctx);
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
  return [...workspace.remote.values()].sort((a, b) => b.id.localeCompare(a.id));
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
    <div>Similitud del material: ${materialSimilarity(result.line.descripcion, row.herramienta)}%</div>
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

function lineMarkup(line, editable) {
  const highlight = editable ? `<label class="check"><input type="checkbox" data-field="resaltado" ${line.resaltado ? 'checked' : ''}> Resaltado rosa</label>`
    : `<span class="hint">${line.resaltado ? 'Resaltado rosa' : 'Sin resaltar'}</span>`;
  const confidence = editable ? `<label class="field">Confianza <select data-field="confianza">${['alta', 'media', 'baja'].map(value => `<option${line.confianza === value ? ' selected' : ''}>${value}</option>`).join('')}</select></label>`
    : readField('Confianza', line.confianza);
  const compact = key => editable
    ? `<input type="text" data-field="${key}" aria-label="${lineFields[key]}" value="${html(line[key])}">`
    : `<span class="read-field">${html(line[key] || '—')}</span>`;
  return `<tr data-line="${line.key}"><td data-label="Fecha">${compact('fecha_linea')}<span class="hint" data-effective-date></span></td>
    <td data-label="Descripción">${editable ? `<textarea rows="2" data-field="descripcion" aria-label="Descripción">${html(line.descripcion)}</textarea>
      <button type="button" class="small" data-action="recompare" hidden>Volver a buscar en Excel</button>` : `<span class="read-field">${html(line.descripcion || '—')}</span>`}
      ${highlight}<details><summary>Devolución y notas <span class="hint" data-return-date></span></summary><div class="edit-grid">
      ${editField('fecha_devolucion', 'Fecha devuelta', line.fecha_devolucion, editable)}
      ${editField('cantidad_devuelta', 'Cantidad devuelta', line.cantidad_devuelta, editable)}
      ${editField('nota_manuscrita', 'Nota manuscrita', line.nota_manuscrita, editable)}${confidence}
      ${editable ? '<div><button type="button" class="small" data-action="remove">Quitar línea</button></div>' : ''}
      </div></details></td><td data-label="Cantidad">${compact('cantidad')}</td><td data-label="Comparación" data-result></td></tr>`;
}

function makeCard(receipt) {
  const card = document.createElement('article');
  card.className = 'receipt';
  card.dataset.id = receipt.id;
  card.innerHTML = `<details class="receipt-toggle"><summary class="receipt-summary"><span data-receipt-title></span><span class="hint" data-receipt-match></span></summary>
    <div class="receipt-head">${Object.entries(headerFields).map(([field, label]) =>
    editField(field, label, receipt[field], true, field === 'observacion')).join('')}</div>
    <div class="receipt-layout"><div class="photo-block">
      <div class="actions"><button type="button" class="small" data-photo="open" hidden>Ver foto</button>
      <button type="button" class="small" data-photo="remove" hidden>Quitar foto</button></div>
      <p class="hint" data-photo-status></p><p class="hint" data-photo-feedback role="status"></p>
      <figure tabindex="0" aria-label="Foto del remito" hidden><img alt="Foto del remito"></figure>
      <label class="check" data-photo-zoom hidden>Zoom <input type="range" min="100" max="400" step="25" value="100" aria-label="Zoom de la foto"></label>
    </div><div class="review">
      ${receipt.lines.length ? `<div class="table-wrap"><table aria-label="Líneas del remito"><thead><tr><th>Fecha</th><th>Descripción</th><th>Cant.</th><th>Comparación</th></tr></thead>
      <tbody>${receipt.lines.map(line => lineMarkup(line, true)).join('')}</tbody></table></div>`
    : '<p class="empty">El remito no tiene líneas transcritas.</p>'}
      <div class="receipt-foot"><button type="button" class="small" data-action="add">Agregar línea</button></div></div></div></details>`;
  return card;
}

function renderPhoto(card, receipt, job) {
  const cloud = workspace.scope !== 'local';
  const available = cloud ? receipt.photo === true : !!job?.source;
  if (!available) forgetPhoto(receipt.id);
  const url = photoUrls.get(receipt.id);
  card.querySelector('[data-photo="open"]').hidden = !available || !!url;
  card.querySelector('[data-photo="remove"]').hidden = !available;
  card.querySelector('[data-photo-status]').textContent = available ? ''
    : cloud && receipt.photo === undefined && job?.source ? 'Compartiendo foto…' : 'Foto no disponible';
  card.querySelector('figure').hidden = card.querySelector('[data-photo-zoom]').hidden = !url;
  const image = card.querySelector('figure img');
  if (url && image.getAttribute('src') !== url) image.src = url;
  if (!url) image.removeAttribute('src');
}

function renderStatus(cell, result) {
  const mark = result.line.verificado;
  const best = result.candidates?.[0]?.row;
  const signature = JSON.stringify([result.state, result.chosen, result.decision, result.conflict, result.stale, result.date, result.line.cantidad,
    result.line.descripcion, result.receipt.numero_remito, best, !!workspace.excel, mark]);
  if (cell.dataset.signature === signature) return;
  cell.dataset.signature = signature;
  const chosen = result.chosen ? formatRow(result.chosen, result) : '';
  const action = (id, label) => `<button type="button" class="small" data-action="${id}">${label}</button>`;
  const buttons = [];
  if (result.state !== 'skipped') {
    if (result.chosen && !result.decision) buttons.push(action('confirm', 'Confirmar'));
    if (workspace.excel) buttons.push(action('choose', result.chosen ? 'Elegir otra fila' : result.candidates?.length ? 'Ver candidatos' : 'Buscar en Excel'));
    if (result.decision !== 'absent') buttons.push(action('absent', 'No está en el Excel'));
    if (result.decision) buttons.push(action('undo', 'Reabrir'));
  }
  const canVerify = mark || ['ok', 'doubtful', 'missing'].includes(result.state);
  cell.innerHTML = `<span class="status ${result.state}">${states[result.state]}</span>${chosen}
    ${!chosen && best ? `<p class="hint">Mejor candidato: ${materialSimilarity(result.line.descripcion, best.herramienta)}% de similitud</p>` : ''}
    ${result.decision ? `<p class="details-text">${decisions[result.decision]}</p>` : ''}
    ${result.conflict ? '<p class="details-text">La fila ya está asignada a otra línea.</p>' : ''}
    ${result.stale ? '<p class="details-text">La fila elegida cambió en el Excel de la sala.</p>' : ''}
    ${result.state === 'pending' ? '<p class="hint">Cargá el Excel para comparar.</p>' : ''}
    ${canVerify ? `<label class="check"><input type="checkbox" data-verify ${mark ? 'checked' : ''}> ✓ Verificado</label>` : ''}
    ${mark ? `<p class="details-text">${html(mark.por)} · ${loadTime.format(mark.fecha)}</p>` : ''}
    ${buttons.length ? `<div class="actions">${buttons.join('')}</div>` : ''}`;
}

function renderFields(container, data) {
  for (const field of container.querySelectorAll('[data-field]')) {
    if (field === document.activeElement) continue;
    const value = data[field.dataset.field];
    if (field.type === 'checkbox') field.checked = !!value;
    else if (field.value !== value) field.value = value;
  }
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
  const totals = { ok: 0, doubtful: 0, missing: 0, reviewed: 0, verified: 0 };
  const filter = $('verified-filter').value;
  let visible = 0;
  const container = $('receipts');
  receipts.forEach((receipt, position) => {
    const job = ctx.local.get(receipt.id);
    let entry = cards.get(receipt.id);
    const lineKeys = receipt.lines.map(line => line.key).join(' ');
    const editing = document.activeElement?.matches('[data-field]') && entry?.card.contains(document.activeElement);
    if (!entry || (entry.lineKeys !== lineKeys && !editing)) {
      const card = makeCard(receipt);
      card.querySelector('.receipt-toggle').open = entry?.card.querySelector('.receipt-toggle').open ?? false;
      entry?.card.remove();
      entry = { card, lineKeys };
      cards.set(receipt.id, entry);
    }
    const { card } = entry;
    if (container.children[position] !== card) container.insertBefore(card, container.children[position] || null);
    const title = receipt.numero_remito ? `Remito ${receipt.numero_remito}` : 'Remito sin número';
    const fileName = receipt.nombre_archivo || job?.name;
    card.querySelector('[data-receipt-title]').textContent = fileName ? `${title} · ${fileName}` : title;
    renderFields(card.querySelector('.receipt-head'), receipt);
    renderPhoto(card, receipt, job);
    const dates = datesFor(receipt);
    const rows = new Map([...card.querySelectorAll('tr[data-line]')].map(row => [row.dataset.line, row]));
    const matches = { ok: 0, doubtful: 0, missing: 0, pending: 0 };
    let compared = 0, finished = 0, rowVisible = false;
    receipt.lines.forEach((line, index) => {
      const result = resultFor(receipt, index);
      if (['ok', 'doubtful', 'missing'].includes(result.state)) totals[result.state]++;
      if (Object.hasOwn(matches, result.state)) matches[result.state]++;
      if (result.state !== 'skipped') compared++;
      if (result.decision) finished++;
      if (line.verificado) totals.verified++;
      const row = rows.get(line.key);
      if (!row) return;
      renderFields(row, line);
      row.querySelector('[data-action="recompare"]').hidden = !ctx.excel;
      row.classList.toggle('highlighted', result.line.resaltado);
      row.hidden = ($('only-issues').checked && !['doubtful', 'missing'].includes(result.state))
        || (filter !== 'all' && !!line.verificado !== (filter === 'yes'));
      rowVisible ||= !row.hidden;
      const label = dates[index].label;
      row.querySelector('[data-effective-date]').textContent = label !== result.line.fecha_linea ? label : '';
      row.querySelector('[data-return-date]').textContent = dates[index].returnLabel ? `· ${dates[index].returnLabel}` : '';
      renderStatus(row.querySelector('[data-result]'), result);
    });
    const checked = matches.ok + matches.doubtful + matches.missing;
    const summary = [checked && `Coinciden ${matches.ok} de ${checked} líneas comparadas`,
      matches.doubtful && `${matches.doubtful} ${matches.doubtful === 1 ? 'dudosa' : 'dudosas'}`,
      matches.missing && `${matches.missing} sin coincidencia`,
      matches.pending && `${matches.pending} ${matches.pending === 1 ? 'pendiente' : 'pendientes'} de comparar`].filter(Boolean).join(' · ');
    card.querySelector('[data-receipt-match]').textContent = summary || (receipt.lines.length ? 'Sin líneas comparadas' : 'Sin renglones transcritos');
    if (compared && finished === compared) totals.reviewed++;
    card.hidden = ($('person').value && $('person').value !== receipt.persona) || ($('work').value && $('work').value !== receipt.obra_o_vehiculo)
      || (($('only-issues').checked || filter !== 'all') && !rowVisible);
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

const linePath = (receipt, key) => receipt.id + '/lines/' + key;

function changeLine(receipt, key, fields) {
  const path = linePath(receipt, key);
  return changeReceipt(Object.fromEntries(Object.entries(fields).map(([field, value]) => [path + '/' + field, value])));
}

function editReceipt(event) {
  const control = event.target;
  const id = control.closest('article')?.dataset.id;
  const receipt = workspace?.remote.get(id);
  if (!receipt) return;
  const key = control.closest('tr')?.dataset.line;
  const line = receipt.lines.find(line => line.key === key);
  if (control.hasAttribute('data-verify') && line) {
    let mark = null;
    if (control.checked) {
      const stored = localStorage.getItem(REVIEWER);
      const name = text(stored || window.prompt('¿Con qué nombre querés verificar remitos?')).slice(0, 120);
      if (!name) { control.checked = false; return; }
      if (!stored) localStorage.setItem(REVIEWER, name);
      mark = { por: name, fecha: Date.now() };
    }
    changeLine(receipt, key, { verificado: mark });
    return;
  }
  const field = control.dataset.field;
  if (!field) return;
  const value = field === 'resaltado' ? control.checked : text(control.value);
  if (line && line[field] !== value) changeLine(receipt, key, { [field]: value });
  else if (!key && receipt[field] !== value) changeReceipt({ [id + '/' + field]: value });
}

async function doLineAction(event) {
  const button = event.target.closest('button[data-action]');
  const id = button?.closest('article')?.dataset.id;
  const receipt = workspace?.remote.get(id);
  if (!receipt) return;
  const key = button.closest('tr')?.dataset.line;
  const action = button.dataset.action;
  if (action === 'add') {
    const newKey = newLineKey();
    const order = receipt.lines.length ? receipt.lines.at(-1).order + 1 : 0;
    await changeReceipt({ [linePath(receipt, newKey)]: publicLine({ confianza: 'media' }, order) });
    cards.get(id)?.card.querySelector('[data-line="' + newKey + '"] textarea')?.focus();
    return;
  }
  const index = receipt.lines.findIndex(line => line.key === key);
  if (index < 0) return;
  if (action === 'remove') {
    if (window.confirm('¿Quitar esta línea de la transcripción?')) changeLine(receipt, key, { deleted: true });
    return;
  }
  if (action === 'choose') { openChooser(receipt, key); return; }
  const result = resultFor(receipt, index);
  if (action === 'recompare' && workspace.excel) {
    const line = result.line;
    const fields = {};
    if (!line.comparar) fields.comparar = true;
    if (line.decision) fields.decision = '';
    if (line.elegida) fields.elegida = null;
    if (line.estado !== 'pending') fields.estado = 'pending';
    if (Object.keys(fields).length) await changeLine(receipt, key, fields);
    else { recalculate(workspace); render(); }
    return;
  }
  if (action === 'confirm' && result.chosen) changeLine(receipt, key, { decision: 'confirmed', elegida: result.chosen, estado: 'ok' });
  if (action === 'absent') changeLine(receipt, key, { decision: 'absent', elegida: null, estado: 'missing' });
  if (action === 'undo') changeLine(receipt, key, { decision: '', elegida: null, estado: 'pending' });
}

function openChooser(receipt, key) {
  chooser = { id: receipt.id, key, page: 0 };
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
  const receipt = workspace?.remote.get(chooser?.id);
  const index = receipt?.lines.findIndex(line => line.key === chooser.key) ?? -1;
  if (index < 0 || !workspace.excel) return;
  const result = resultFor(receipt, index);
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
  const receipt = workspace?.remote.get(chooser?.id);
  const index = receipt?.lines.findIndex(line => line.key === chooser.key) ?? -1;
  if (!row || index < 0) return;
  const result = resultFor(receipt, index);
  if (!rowAvailable(row, result)) {
    renderChooser();
    return notice('Esa fila acaba de asignarse a otra línea. Elegí una fila disponible.', true);
  }
  $('chooser').close();
  changeLine(receipt, result.line.key, { decision: 'manual', elegida: chosenRow(row), estado: 'ok' });
}

function useWorkbook(ctx, workbook) {
  ctx.workbook = workbook;
  ctx.excel = workbook ? indexWorkbook(workbook.rows) : null;
  recalculate(ctx);
  if (workspace === ctx) render();
}

async function receiveExcel(ctx, sync, meta) {
  sync.excelMeta = meta;
  if (ctx.excelUpload) return;
  if (!meta) {
    if (!ctx.workbook) return;
    useWorkbook(ctx, null);
    await storeRecords([], [[ctx.scope, 'excel']]).catch(storeError);
    return;
  }
  if (ctx.workbook?.revision === meta.revision || sync.readingExcel === meta.revision) return;
  const revision = meta.revision;
  if (!/^r[a-z0-9-]+$/i.test(revision) || !Number.isSafeInteger(meta.blocks) || meta.blocks < 0) {
    excelFeedback('Los datos del Excel de la sala son inválidos.', true);
    return;
  }
  sync.readingExcel = revision;
  updateControls();
  try {
    const ref = sync.excelRef.child('bloques').child(revision);
    const rows = [];
    for (let i = 0; i < meta.blocks; i++) {
      const snapshot = await ref.child('b' + i).once('value');
      if (ctx.sync !== sync || sync.excelMeta?.revision !== revision || ctx.excelUpload) return;
      const block = snapshot.val();
      if (!block) throw new Error();
      rows.push(...Object.values(block));
    }
    if (rows.length !== meta.count) throw new Error();
    rows.sort((a, b) => a.fila - b.fila);
    const workbook = { id: 'excel', scope: ctx.scope, ...meta, rows };
    await storeRecords([workbook]);
    if (ctx.sync === sync && sync.excelMeta?.revision === revision && !ctx.excelUpload) useWorkbook(ctx, workbook);
  } catch {
    if (ctx.sync === sync && sync.excelMeta?.revision === revision && !ctx.excelUpload) {
      sync.excelError = true;
      excelFeedback('No se pudo descargar el Excel de la sala. Podés reintentar la sincronización.', true);
    }
  } finally {
    if (sync.readingExcel === revision) sync.readingExcel = null;
    if (workspace === ctx) { updateControls(); updateSync(); }
  }
}

async function discardExcelRevision(sync, meta, keep) {
  if (/^r[a-z0-9-]+$/i.test(meta?.revision || '') && meta.revision !== keep) {
    await sync.excelRef.child('bloques').child(meta.revision).remove();
  }
}

async function flushExcel(ctx) {
  const sync = ctx.sync;
  const upload = ctx.excelUpload;
  if (!upload || !updatesAllowed(ctx) || sync.writingExcel || sync.excelError) return;
  sync.writingExcel = true;
  updateControls();
  try {
    const blocks = excelBlocks(upload.rows);
    const ref = sync.excelRef.child('bloques').child(upload.revision);
    for (const [index, block] of blocks.entries()) {
      if (!updatesAllowed(ctx) || ctx.sync !== sync) return;
      await ref.child('b' + index).set(block);
    }
    if (!updatesAllowed(ctx) || ctx.sync !== sync) return;
    const meta = { revision: upload.revision, name: upload.name, loadedAt: upload.loadedAt, count: upload.rows.length, blocks: blocks.length };
    let previous;
    const outcome = await sync.excelRef.child('actual').transaction(current => {
      previous = current;
      return meta;
    }, undefined, false);
    if (ctx.sync !== sync || !outcome.committed) return;
    const workbook = { ...upload, id: 'excel' };
    await storeRecords([workbook], [[ctx.scope, 'excel-queued']]);
    ctx.excelUpload = null;
    useWorkbook(ctx, workbook);
    if (sync.excelMeta?.revision !== workbook.revision) {
      const latest = (await sync.excelRef.child('actual').once('value')).val();
      if (ctx.sync === sync) receiveExcel(ctx, sync, latest);
    }
    discardExcelRevision(sync, previous, upload.revision).catch(() => {});
    if (workspace === ctx) excelFeedback('');
  } catch {
    if (ctx.sync === sync) {
      sync.excelError = true;
      excelFeedback('El Excel está disponible en este navegador. Reintentá la sincronización para compartirlo.', true);
    }
  } finally {
    sync.writingExcel = false;
    if (workspace === ctx) { updateControls(); updateSync(); }
  }
}

async function loadWorkbook(file) {
  const ctx = workspace;
  if (!ctx || !file || loadingExcel || ctx.sync?.writingExcel) return;
  if (!/\.xlsx$/i.test(file.name)) return excelFeedback('Elegí un archivo .xlsx.', true);
  loadingExcel = true;
  updateControls();
  excelFeedback('');
  try {
    const xlsx = await dependency('XLSX', SHEET_JS);
    const rows = readWorkbook(xlsx, await file.arrayBuffer());
    if (workspace !== ctx) return;
    const shared = ctx.scope !== 'local';
    const workbook = { id: shared ? 'excel-queued' : 'excel', scope: ctx.scope, name: file.name,
      rows: shared ? excelBlocks(rows).flatMap(Object.values) : rows, loadedAt: Date.now(),
      ...(shared && { revision: newId() }) };
    await storeRecords([workbook]);
    const previous = ctx.excelUpload;
    ctx.excelUpload = shared ? workbook : null;
    useWorkbook(ctx, workbook);
    if (shared && ctx.sync) {
      ctx.sync.excelError = false;
      discardExcelRevision(ctx.sync, previous, ctx.sync.excelMeta?.revision).catch(() => {});
      flushExcel(ctx);
    }
  } catch (error) {
    if (workspace === ctx) excelFeedback('No se pudo cargar el Excel. ' + error.message, true);
  } finally {
    loadingExcel = false;
    updateControls();
    updateSync();
  }
}

async function removeWorkbook() {
  const ctx = workspace;
  if (!ctx || loadingExcel || ctx.sync?.writingExcel) return;
  const local = ctx.scope === 'local';
  if (!local && !updatesAllowed(ctx)) return excelFeedback('Conectate a la sala para quitar el Excel.', true);
  if (!window.confirm(local ? '¿Quitar el Excel de este navegador?' : '¿Quitar el Excel para todos en la sala?')) return;
  loadingExcel = true;
  updateControls();
  try {
    const sync = ctx.sync;
    let previous;
    const queued = ctx.excelUpload;
    if (!local) {
      const outcome = await sync.excelRef.child('actual').transaction(current => {
        previous = current;
        return null;
      }, undefined, false);
      if (!outcome.committed || ctx.sync !== sync) return;
      sync.excelError = false;
    }
    ctx.excelUpload = null;
    await storeRecords([], [[ctx.scope, 'excel'], [ctx.scope, 'excel-queued']]);
    useWorkbook(ctx, null);
    if (!local) await Promise.all([discardExcelRevision(sync, previous), discardExcelRevision(sync, queued)]);
    excelFeedback('');
  } catch { excelFeedback('No se pudo completar la eliminación del Excel.', true); }
  finally { loadingExcel = false; updateControls(); updateSync(); }
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
  if (!validKey(value)) return keyFeedback('Copiá el valor completo de la clave desde Google AI Studio y pegalo sin espacios.', true);
  if (activeKeys(workspace).some(entry => entry.value === value)) return keyFeedback('Esa clave ya está en la lista.', true);
  if (await editKeys(entries => entries.some(entry => entry.value === value) ? entries : [...entries, { value }])) {
    $('gemini-key').value = '';
    keyFeedback('Clave agregada. Podés transcribir las fotos.');
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
$('choose-excel').onclick = () => $('excel-file').click();
$('remove-excel').onclick = removeWorkbook;
$('excel-file').onchange = event => { loadWorkbook(event.target.files[0]); event.target.value = ''; };
$('choose-images').onclick = () => $('images').click();
$('images').onchange = event => { addPhotos([...event.target.files]); event.target.value = ''; };
$('process').onclick = processPhotos;
$('queue').addEventListener('change', selectQueuedPhoto);
$('queue').addEventListener('click', previewQueuedPhoto);
$('photo-preview').addEventListener('close', () => {
  previewTicket++;
  const image = $('preview-image');
  image.onload = image.onerror = null;
  image.removeAttribute('src');
  if (previewUrl) URL.revokeObjectURL(previewUrl);
  previewUrl = null;
});
$('sync-retry').onclick = () => { if (workspace) connect(workspace, parentConnection()); };
$('include-all').onchange = event => {
  const ctx = workspace;
  ctx.includeAll = event.target.checked;
  storeRecords([{ id: 'options', scope: ctx.scope, includeAll: ctx.includeAll }]).catch(storeError);
  recalculate(ctx);
  render();
};
for (const id of ['person', 'work', 'only-issues']) $(id).onchange = render;
$('verified-filter').value = ['all', 'yes', 'no'].includes(localStorage.getItem(VERIFIED_FILTER)) ? localStorage.getItem(VERIFIED_FILTER) : 'all';
$('verified-filter').onchange = event => { localStorage.setItem(VERIFIED_FILTER, event.target.value); render(); };
$('receipts').addEventListener('change', editReceipt);
$('receipts').addEventListener('click', doLineAction);
$('receipts').addEventListener('click', photoAction);
$('receipts').addEventListener('focusout', event => { if (event.target.matches('[data-field]')) setTimeout(render, 0); });
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

const modelSelect = $('gemini-model');
modelSelect.options[0].text = `Automático (prioriza ${GEMINI_MODELS[0].label})`;
modelSelect.append(...GEMINI_MODELS.map(({ id, label }) => new Option(label, id)));
const savedModel = localStorage.getItem(MODEL_CHOICE);
modelSelect.value = GEMINI_MODELS.some(({ id }) => id === savedModel) ? savedModel : '';
modelSelect.onchange = () => localStorage.setItem(MODEL_CHOICE, modelSelect.value);

refreshRoom();
window.addEventListener('storage', refreshRoom);
window.addEventListener('focus', refreshRoom);
// El evento storage no se emite cuando el planificador cambia sessionStorage en este documento superior.
setInterval(() => { refreshRoom(); renderKeys(); }, 1500);
