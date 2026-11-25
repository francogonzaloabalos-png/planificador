import {
  headerFields, lineFields, states, decisions, normalize, text, datesFor, readWorkbook,
  compareRowFields, reviewReturn, indexWorkbook, chosenRow, lineId, compareReceipts, transcribe, publicReceipt,
  receiveReceipt, exportRows, GeminiQuotaError, publicLine, newLineKey, excelBlocks, GEMINI_MODELS, productCategories,
  excelCorrection, excelEditLabel, excelEditLabels, excelEditValue, excelRowEdits, isReturnMovement, sameWorkbookRows
} from './control-remitos-core.js?v=fix-records-20261120';
import { buildRemitosExcel } from './control-remitos-export.js?v=fix-records-20261120';

const $ = id => document.getElementById(id);
const LEGACY_KEY = 'control_remitos_gemini_key_v1';
const LOCAL_KEYS = 'control_remitos_gemini_keys_v1';
const SHARE_KEYS = 'control_remitos_share_keys_v1';
const MODEL_CHOICE = 'control_remitos_model_choice_v1';
const REVIEWER = 'control_remitos_reviewer_v1';
const VERIFIED_FILTER = 'control_remitos_verified_filter_v1';
const MAX_UPLOAD = 10;
const MAX_EDGE = 1600;
const SHARED_PHOTO_EDGE = 1250;
const SHARED_PHOTO_QUALITY = .72;
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
const fileNameOrder = new Intl.Collator('es-AR', { numeric: true, sensitivity: 'base' });

let workspace;
let roomGeneration = 0;
let targetScope;
let busy = false;
let loadingExcel = false;
let chooser = null;
const cards = new Map();
const reviewFolders = new Map();
const photoUrls = new Map();
let removal;
let receiptRemoval;
let move;
let folderEdit;
let photoUpload;
let verification;
let movementEdit;
let excelFieldEdit;
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

async function readRecords(scope, id) {
  const db = await database;
  return new Promise((resolve, reject) => {
    const records = db.transaction('records').objectStore('records');
    const request = id === undefined ? records.index('scope').getAll(scope) : records.get([scope, id]);
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
  for (const id of ['remove-photo-dialog', 'move-photo-dialog', 'remove-receipt-dialog', 'folder-name-dialog', 'verify-dialog', 'movement-dialog',
    'excel-edit-dialog']) if ($(id).open) $(id).close();
  for (const [id] of photoUrls) forgetPhoto(id);
  photoUrls.clear();
  cards.clear();
  reviewFolders.clear();
  $('receipts').replaceChildren();
  $('queue').replaceChildren();
  delete $('queue').dataset.markup;
  $('photo-storage').hidden = true;
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
    const folders = records.find(record => record.id === 'folders-cloud')?.values || {};
    const localFolders = new Map(records.filter(record => record.type === 'folder').map(folder => [folder.id, folder]));
    workspace = {
      scope, workbook, excel: workbook ? indexWorkbook(workbook.rows, workbook.base) : null, local, selectedPhotos: new Set(),
      folders, localFolders, uploadFolder: records.find(record => record.id === 'upload-folder')?.folder || '', deletingPhotos: new Set(),
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
    if (sync.onFolders) sync.foldersRef.off('value', sync.onFolders);
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
    sync.foldersRef = room.child('carpetas');
    sync.photoWork = new Set();
    sync.photoDone = new Set();
    listenKeys(ctx);
    sync.onExcel = snapshot => {
      if (ctx.sync === sync) receiveExcel(ctx, sync, snapshot.val());
    };
    sync.excelRef.child('actual').on('value', sync.onExcel, () => {
      if (ctx.sync === sync) excelFeedback('No se pudo leer el Excel de la sala.', true);
    });
    sync.onFolders = snapshot => {
      if (ctx.sync !== sync) return;
      ctx.folders = snapshot.val() || {};
      storeRecords([{ id: 'folders-cloud', scope: ctx.scope, values: ctx.folders }]).catch(storeError);
      render();
    };
    sync.foldersRef.on('value', sync.onFolders, () => {
      if (ctx.sync !== sync) return;
      sync.folderError = true;
      updateSync();
    });
    sync.connected = db.ref('.info/connected');
    sync.onValue = snapshot => {
      if (ctx.sync !== sync) return;
      sync.ready = true;
      const values = snapshot.val() || {};
      ctx.values = values;
      saveReceiptCache(ctx).catch(storeError);
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
      flushFolders(ctx);
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

function patchReceiptId(patch) {
  return Object.keys(patch.fields)[0].split('/')[0];
}

function canApplyReceiptPatch(value, patch) {
  if (patch.onlyUnfiled) return !!value && (value.version === 1 || value.transcripcion_eliminada) && !validFolderId(value.carpeta_id);
  return value?.version === 1 && !value.transcripcion_eliminada;
}

function removedTranscript(value) {
  return { transcripcion_eliminada: true,
    ...(value.nombre_archivo && { nombre_archivo: value.nombre_archivo }),
    ...(value.carpeta_id && { carpeta_id: value.carpeta_id }),
    ...(typeof value.photo === 'boolean' && { photo: value.photo }) };
}

// El marcador conserva el vínculo con la foto y evita recuperar una transcripción vieja de IndexedDB.
function saveReceiptCache(ctx) {
  const records = [{ id: 'cloud', scope: ctx.scope, values: ctx.values }];
  const deletes = [];
  for (const [id, job] of ctx.local) {
    if (!job.receipt || !ctx.values[id]?.transcripcion_eliminada) continue;
    if (job.source) {
      const { receipt, payload, ...photo } = job;
      photo.pending = false;
      ctx.local.set(id, photo);
      records.push(photo);
    } else {
      ctx.local.delete(id);
      deletes.push([ctx.scope, id]);
    }
    ctx.selectedPhotos.delete(id);
  }
  for (const [id, patch] of ctx.patches) {
    if (patch.onlyUnfiled || !ctx.values[patchReceiptId(patch)]?.transcripcion_eliminada) continue;
    ctx.patches.delete(id);
    deletes.push([ctx.scope, id]);
  }
  return storeRecords(records, deletes);
}

function recalculate(ctx) {
  let values = { ...ctx.values };
  for (const job of ctx.local.values()) {
    if (job.receipt && !values[job.id]) values[job.id] = job.payload || publicReceipt(job.receipt);
  }
  for (const patch of ctx.patches.values()) {
    if (canApplyReceiptPatch(values[patchReceiptId(patch)], patch)) values = withReceiptChanges(values, patch.fields);
  }
  const receipts = new Map();
  for (const [id, value] of Object.entries(values)) {
    if (!/^r[a-z0-9-]+$/i.test(id)) continue;
    const receipt = receiveReceipt(id, value);
    if (receipt) receipts.set(id, receipt);
  }
  ctx.remote = receipts;
  ctx.receiptValues = values;
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
  const pending = [...ctx.local.values(), ...ctx.localFolders.values()].filter(record => record.pending).length
    + ctx.patches.size + Number(!!ctx.excelUpload);
  const error = sync?.error || sync?.photoError || sync?.excelError || sync?.folderError;
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
  const creations = [...ctx.local.values()].filter(job => job.pending && job.id !== ctx.removingReceipt && !ctx.deletingPhotos.has(job.id));
  const patches = [...ctx.patches.values()];
  if (!creations.length && !patches.length) return;
  sync.writing = true;
  try {
    for (const job of creations) {
      if (!updatesAllowed(ctx) || ctx.sync !== sync) break;
      if (ctx.local.get(job.id) !== job || job.id === ctx.removingReceipt || ctx.deletingPhotos.has(job.id)) continue;
      const result = await sync.ref.child(job.id).transaction(current => current == null ? job.payload : undefined, undefined, false);
      if (ctx.sync !== sync) return;
      if (!ctx.values[job.id] && result.snapshot.val()) ctx.values = { ...ctx.values, [job.id]: result.snapshot.val() };
      if (ctx.local.get(job.id) !== job || job.id === ctx.removingReceipt || ctx.deletingPhotos.has(job.id)) continue;
      await storeRecords([{ ...job, pending: false }]);
      job.pending = false;
    }
    for (const patch of patches) {
      if (!updatesAllowed(ctx) || ctx.sync !== sync) break;
      if (!ctx.patches.has(patch.id)) continue;
      const id = patchReceiptId(patch);
      const fields = Object.fromEntries(Object.entries(patch.fields).map(([path, value]) => [path.slice(id.length + 1), value]));
      const previous = ctx.values;
      // Una edición atrasada no puede volver a crear datos de un remito eliminado.
      const result = await sync.ref.child(id).transaction(current =>
        canApplyReceiptPatch(current, patch) ? withReceiptChanges(current, fields) : undefined, undefined, false);
      if (ctx.sync !== sync) return;
      const saved = result.snapshot.val();
      if (!result.committed && canApplyReceiptPatch(saved, patch)) throw new Error('Edición pendiente de sincronizar.');
      if (ctx.values === previous) ctx.values = withReceiptChanges(ctx.values, { [id]: saved });
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

async function changeReceipt(fields, ctx = workspace) {
  const receiptId = Object.keys(fields)[0].split('/')[0];
  if (ctx.removingReceipt === receiptId || ctx.values[receiptId]?.transcripcion_eliminada) return false;
  try {
    if (ctx.scope === 'local') {
      const cached = (await readRecords(ctx.scope, 'cloud'))?.values?.[receiptId];
      if (cached?.transcripcion_eliminada) {
        ctx.values = { ...ctx.values, [receiptId]: cached };
        await saveReceiptCache(ctx);
        recalculate(ctx);
        if (workspace === ctx) render();
        return false;
      }
      ctx.values = withReceiptChanges(ctx.values, fields);
      // Los trabajos antiguos pueden tener el remito únicamente en IndexedDB.
      if (!ctx.values[receiptId].version) {
        const job = ctx.local.get(receiptId);
        ctx.values = withReceiptChanges({ ...ctx.values, [receiptId]: job.payload || publicReceipt(job.receipt) }, fields);
      }
      recalculate(ctx);
      await storeRecords([{ id: 'cloud', scope: ctx.scope, values: ctx.values }]);
    } else {
      const patch = { id: 'patch-' + newId(), scope: ctx.scope, type: 'patch', fields };
      await storeRecords([patch]);
      ctx.patches.set(patch.id, patch);
      recalculate(ctx);
      flush(ctx);
    }
  } catch {
    storeError();
    return false;
  }
  if (workspace === ctx) render();
  return true;
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

const validFolderId = id => typeof id === 'string' && /^f[a-z0-9-]+$/i.test(id);
const folderId = value => validFolderId(value) ? value : '';

function photoFolder(ctx, id) {
  const receipt = ctx.receiptValues[id];
  return folderId(receipt ? receipt.carpeta_id : ctx.local.get(id)?.carpeta_id);
}

function folderNames(ctx) {
  const folders = new Map();
  for (const [id, folder] of [...Object.entries(ctx.folders), ...ctx.localFolders]) {
    if (validFolderId(id) && text(folder?.nombre)) folders.set(id, text(folder.nombre));
  }
  // La asociación sigue siendo válida aunque todavía no haya llegado el nombre de la carpeta.
  for (const [id, receipt] of Object.entries(ctx.receiptValues)) {
    const assigned = folderId(receipt?.carpeta_id);
    if (assigned && !folders.has(assigned) && (receipt.version === 1 || receipt.photo === true || ctx.local.get(id)?.source)) {
      folders.set(assigned, 'Carpeta sin nombre');
    }
  }
  for (const [id, job] of ctx.local) {
    const assigned = photoFolder(ctx, id);
    if (job.source && assigned && !folders.has(assigned)) folders.set(assigned, 'Carpeta sin nombre');
  }
  if (validFolderId(ctx.uploadFolder) && !folders.has(ctx.uploadFolder)) folders.set(ctx.uploadFolder, 'Carpeta sin nombre');
  return new Map([['', 'Sin carpeta'], ...[...folders].sort(([, a], [, b]) => a.localeCompare(b, 'es-AR'))]);
}

function folderOptions(select, folders, value) {
  const markup = [...folders].map(([id, name]) => `<option value="${id}">${html(name)}</option>`).join('');
  if (select.dataset.markup !== markup) {
    select.dataset.markup = markup;
    select.innerHTML = markup;
  }
  select.value = folders.has(value) ? value : '';
}

async function flushFolders(ctx) {
  const sync = ctx.sync;
  if (!updatesAllowed(ctx) || !sync?.foldersRef || sync.writingFolders || sync.folderError) return;
  const pending = [...ctx.localFolders.values()].filter(folder => folder.pending);
  if (!pending.length) return;
  sync.writingFolders = true;
  const previous = ctx.folders;
  try {
    await sync.foldersRef.update(Object.fromEntries(pending.map(({ id, nombre }) => [`${id}/nombre`, nombre])));
    if (ctx.sync !== sync) return;
    const published = { ...ctx.folders };
    for (const { id, nombre } of pending) {
      if (published[id]?.nombre === previous[id]?.nombre) published[id] = { ...published[id], nombre };
    }
    ctx.folders = published;
    const completed = pending.filter(folder => ctx.localFolders.get(folder.id) === folder);
    await storeRecords([{ id: 'folders-cloud', scope: ctx.scope, values: published }], completed.map(({ id }) => [ctx.scope, id]));
    for (const folder of completed) {
      if (ctx.localFolders.get(folder.id) === folder) ctx.localFolders.delete(folder.id);
    }
  } catch {
    if (ctx.sync === sync) sync.folderError = true;
  } finally {
    sync.writingFolders = false;
    if (ctx.sync === sync) {
      render();
      flushFolders(ctx);
    }
  }
}

const folderSelection = (ctx, folder) => ({ id: 'upload-folder', scope: ctx.scope, folder });

function selectUploadFolder(ctx, folder) {
  ctx.uploadFolder = folder;
  folderOptions($('upload-folder'), folderNames(ctx), folder);
  storeRecords([folderSelection(ctx, folder)]).catch(storeError);
}

const folderRecord = (ctx, id, nombre) => ({ id, scope: ctx.scope, type: 'folder', nombre, pending: ctx.scope !== 'local' });
const newFolderId = () => `f${crypto.randomUUID()}`;
const folderWithName = (ctx, nombre) => [...folderNames(ctx)].find(([, name]) => name.toLocaleLowerCase('es-AR') === nombre.toLocaleLowerCase('es-AR'))?.[0];

async function createFolder() {
  const ctx = workspace;
  if (!ctx || busy) return;
  const nombre = window.prompt('Nombre de la carpeta')?.trim();
  if (!nombre) return;
  if (nombre.length > 80) return jobMessage('Usá hasta 80 caracteres para el nombre de la carpeta.', true);
  if (folderWithName(ctx, nombre) !== undefined) return jobMessage('Ya hay una carpeta con ese nombre.', true);
  const folder = folderRecord(ctx, newFolderId(), nombre);
  try {
    await storeRecords([folder, folderSelection(ctx, folder.id)]);
    ctx.localFolders.set(folder.id, folder);
    ctx.uploadFolder = folder.id;
    if (workspace !== ctx) return;
    render();
    $('queue-wrap').open = true;
    queueList(folder.id);
    jobMessage(`Carpeta «${nombre}» lista para agregar fotos.`);
    flushFolders(ctx);
  } catch { storeError(); }
}

const folderDetailId = id => `folder-detail-${id || 'sin-carpeta'}`;

function folderHeading(id, name, showDetail = false) {
  return `<span class="folder-heading"><span data-folder-title>${html(name)}</span>
    <span class="actions">
      ${showDetail ? `<button type="button" class="small" data-folder-detail aria-controls="${folderDetailId(id)}" aria-expanded="false">Ver detalle</button>` : ''}
      <button type="button" class="small" data-folder-edit="${id}" aria-label="Editar nombre de ${html(name)}">Editar nombre</button>
    </span></span>`;
}

function askFolderName(event) {
  const button = event.target.closest('[data-folder-edit]');
  const ctx = workspace;
  if (!button || !ctx || busy) return;
  event.preventDefault();
  const id = button.dataset.folderEdit;
  const name = folderNames(ctx).get(id);
  if (!name) return;
  folderEdit = { ctx, id, review: !!button.closest('#receipts') };
  $('folder-name-title').textContent = id ? 'Editar nombre de la carpeta' : 'Poner nombre a «Sin carpeta»';
  $('folder-name-info').hidden = !!id;
  $('folder-name-info').textContent = 'Se moverán todas las fotos de este navegador y los remitos de esta sala que estén sin carpeta. Podés escribir el nombre de una carpeta existente.';
  $('folder-name').value = id ? text(ctx.localFolders.get(id)?.nombre ?? ctx.folders[id]?.nombre) : '';
  $('folder-name-error').hidden = true;
  $('folder-name-dialog').showModal();
  $('folder-name').select();
}

async function groupUnfiled(ctx, destination, records) {
  const stored = await readRecords(ctx.scope);
  const cloud = ctx.scope === 'local' ? stored.find(record => record.id === 'cloud')?.values : null;
  const values = { ...ctx.receiptValues, ...cloud };
  const unfiled = new Set(Object.entries(values).filter(([id, value]) => /^r[a-z0-9-]+$/i.test(id) && !validFolderId(value?.carpeta_id)
    && (value?.version === 1 || (value?.transcripcion_eliminada && (value.photo !== false || ctx.local.get(id)?.source)))).map(([id]) => id));
  const jobs = new Map();
  for (const job of stored) {
    if (job.type !== 'photo' || !(job.source || unfiled.has(job.id))) continue;
    const receipt = values[job.id] || job.payload || job.receipt;
    if (validFolderId(receipt ? receipt.carpeta_id : job.carpeta_id)) continue;
    if (job.receipt) unfiled.add(job.id);
    const updated = { ...job, carpeta_id: destination };
    if (job.receipt) updated.receipt = { ...job.receipt, carpeta_id: destination };
    if (job.payload) updated.payload = { ...job.payload, carpeta_id: destination };
    jobs.set(job.id, updated);
  }
  const patches = [];
  for (const id of unfiled) {
    if (ctx.scope === 'local') {
      if (values[id]) values[id] = { ...values[id], carpeta_id: destination };
    } else patches.push({ id: 'patch-' + newId(), scope: ctx.scope, type: 'patch', onlyUnfiled: true,
      fields: { [`${id}/carpeta_id`]: destination } });
  }
  const previous = new Map([...jobs.keys()].map(id => [id, ctx.local.get(id)]));
  for (const [id, job] of jobs) ctx.local.set(id, job);
  try {
    if (ctx.scope === 'local') records.push({ id: 'cloud', scope: ctx.scope, values });
    await storeRecords([...records, ...jobs.values(), ...patches, folderSelection(ctx, destination)]);
    if (ctx.scope === 'local') ctx.values = values;
    for (const patch of patches) ctx.patches.set(patch.id, patch);
    ctx.uploadFolder = destination;
  } catch (error) {
    for (const [id, job] of previous) {
      if (ctx.local.get(id) !== jobs.get(id)) continue;
      if (job) ctx.local.set(id, job);
      else ctx.local.delete(id);
    }
    throw error;
  }
}

async function saveFolderName(event) {
  event.preventDefault();
  const { ctx, id, review } = folderEdit || {};
  if (!ctx || workspace !== ctx || busy) return;
  const save = $('save-folder-name');
  const cancel = $('cancel-folder-name');
  const errorBox = $('folder-name-error');
  if (save.disabled) return;
  const nombre = text($('folder-name').value);
  const existing = folderWithName(ctx, nombre);
  if (!nombre || (!id && existing === '') || (id && existing !== undefined && existing !== id)) {
    errorBox.textContent = nombre ? 'Ya hay una carpeta con ese nombre.' : 'Escribí un nombre para la carpeta.';
    errorBox.hidden = false;
    return;
  }
  const destination = id || existing || newFolderId();
  const folder = id || !existing ? folderRecord(ctx, destination, nombre) : null;
  const previous = ctx.localFolders.get(destination);
  save.disabled = cancel.disabled = true;
  errorBox.hidden = true;
  try {
    const change = async () => {
      if (id) ctx.localFolders.set(destination, folder);
      try {
        if (id) await storeRecords([folder]);
        else {
          await groupUnfiled(ctx, destination, folder ? [folder] : []);
          if (folder) ctx.localFolders.set(destination, folder);
        }
      } catch (error) {
        if (folder && ctx.localFolders.get(destination) === folder) {
          if (previous) ctx.localFolders.set(destination, previous);
          else ctx.localFolders.delete(destination);
        }
        throw error;
      }
    };
    if (!id && navigator.locks) await navigator.locks.request(`control-remitos:${ctx.scope}`, { ifAvailable: true }, lock => {
      if (!lock) throw new Error('Hay una tanda en curso en otra pestaña. Volvé a intentar cuando termine.');
      return change();
    });
    else await change();
    if (workspace !== ctx) return;
    $('folder-name-dialog').close();
    recalculate(ctx);
    render();
    if (review) {
      reviewFolders.get(destination).open = true;
      notice(`Carpeta «${folderNames(ctx).get(destination)}» guardada.`);
    } else {
      queueList(destination);
      jobMessage(`Carpeta «${folderNames(ctx).get(destination)}» guardada.`);
    }
    flushFolders(ctx);
    flush(ctx);
  } catch (error) {
    if (workspace === ctx) {
      errorBox.textContent = error.message || 'No se pudo guardar el nombre de la carpeta.';
      errorBox.hidden = false;
    }
  } finally { save.disabled = cancel.disabled = false; }
}

function updateControls() {
  const ctx = workspace;
  $('choose-excel').disabled = !ctx || loadingExcel || !!ctx.sync?.writingExcel;
  $('choose-excel').textContent = ctx?.workbook ? 'Actualizar (subir de nuevo)' : 'Cargar Excel (.xlsx)';
  $('remove-excel').hidden = !ctx?.workbook && !ctx?.sync?.excelMeta;
  $('remove-excel').disabled = loadingExcel || !!ctx?.sync?.writingExcel;
  $('remove-excel').textContent = ctx?.scope === 'local' ? 'Quitar Excel local' : 'Quitar Excel de la sala';
  $('choose-images').disabled = !ctx || busy;
  $('upload-folder').disabled = $('new-folder').disabled = !ctx || busy;
  $('process').disabled = !ctx || busy || !ctx.selectedPhotos.size;
  const selectable = ctx ? pendingPhotos(ctx) : [];
  $('select-all-photos').disabled = busy || !selectable.length;
  $('select-all-photos').textContent = selectable.length && selectable.every(job => ctx.selectedPhotos.has(job.id))
    ? 'Deseleccionar todas' : 'Seleccionar todas';
  $('remove-all-photos').disabled = busy || !ctx || !loadedPhotos(ctx).size;
  for (const input of $('queue').querySelectorAll('[data-queue-select], [data-queue-remove], [data-queue-move], [data-folder-edit], [data-folder-upload]')) input.disabled = busy;
  for (const button of $('receipts').querySelectorAll('[data-folder-edit]')) button.disabled = busy;
  for (const button of $('receipts').querySelectorAll('[data-photo="move"], [data-photo="remove"], [data-action="remove-receipt"]')) button.disabled = busy;
  $('include-all').disabled = !ctx || !ctx.excel;
  $('include-all').checked = !!ctx?.includeAll;
  const workbook = ctx?.workbook;
  const time = workbook?.loadedAt ? ` · ${loadTime.format(workbook.loadedAt)}` : '';
  const count = workbook?.rows.length;
  const status = ctx?.excelUpload ? ' · pendiente de publicar' : '';
  $('excel-info').textContent = loadingExcel || ctx?.sync?.readingExcel ? 'Cargando Excel…'
    : workbook ? `${count.toLocaleString('es-AR')} ${count === 1 ? 'fila' : 'filas'} · ${workbook.name}${time}${status}` : 'Excel sin cargar';
  const needsOriginal = workbook && !workbook.base;
  $('excel-copy-hint').hidden = !needsOriginal;
  $('export-xlsx').textContent = needsOriginal ? 'Completar columnas para exportar' : 'Exportar XLSX';
  renderKeys();
}

function compareFileNames(left, right) {
  if (!left || !right) return Number(!left) - Number(!right);
  return fileNameOrder.compare(left, right);
}

function loadedPhotos(ctx) {
  const jobs = new Map([...ctx.local].filter(([, job]) => !job.removedFromQueue && job.source));
  for (const [id, value] of Object.entries(ctx.values)) {
    if (value?.transcripcion_eliminada && value.photo === true && !jobs.has(id) && /^r[a-z0-9-]+$/i.test(id)) {
      jobs.set(id, { id, name: value.nombre_archivo || 'Foto del remito', status: 'done' });
    }
  }
  return jobs;
}

function pendingPhotos(ctx) {
  return [...ctx.local.values()].filter(job => job.source && !job.removedFromQueue && job.status !== 'done');
}

function queueList(openFolder) {
  const ctx = workspace;
  const folders = folderNames(ctx);
  const groups = new Map([...folders.keys()].map(id => [id, []]));
  const jobs = loadedPhotos(ctx);
  for (const job of [...jobs.values()].sort((a, b) => compareFileNames(a.name, b.name) || b.id.localeCompare(a.id))) {
    const id = photoFolder(ctx, job.id);
    groups.get(folders.has(id) ? id : '').push(job);
  }
  folderOptions($('upload-folder'), folders, ctx.uploadFolder);
  $('queue-wrap').hidden = !jobs.size && folders.size === 1;
  const bytes = [...jobs.values()].reduce((total, job) => total + (job.source?.size || 0) + (job.preview?.size || 0), 0);
  $('photo-storage').hidden = !bytes;
  $('photo-storage').textContent = `${(bytes / 1e6).toLocaleString('es-AR', { maximumFractionDigits: 1 })} MB de fotos de esta sala en este navegador`;
  const photoRow = job => {
    const name = html(job.name);
    const selectable = job.status !== 'done';
    const transcriptRemoved = ctx.values[job.id]?.transcripcion_eliminada;
    const canPreview = transcriptRemoved ? job.source || ctx.values[job.id].photo === true : ctx.remote.get(job.id)?.photo !== false;
    const status = transcriptRemoved ? 'Sin transcripción' : 'Transcripta';
    const title = selectable ? `<label class="check queue-name"><input type="checkbox" data-queue-select="${job.id}" aria-label="Seleccionar ${name} para transcribir"> <strong>${name}</strong></label>`
      : `<span class="queue-name"><strong>${name}</strong> · ${status}</span>`;
    return `<li><div class="queue-item">${title}<div class="actions">
      ${canPreview ? `<button type="button" class="small" data-queue-preview="${job.id}" aria-label="Ver foto ${name}">Ver foto</button>` : ''}
      ${folders.size > 1 && !transcriptRemoved ? `<button type="button" class="small" data-queue-move="${job.id}" aria-label="Mover ${name}">Mover</button>` : ''}
      <button type="button" class="small remove-photo" data-queue-remove="${job.id}" aria-label="Eliminar foto ${name}">×</button></div></div>
      ${job.error ? `<p class="hint error">${html(job.error)}</p>` : ''}</li>`;
  };
  const markup = [...groups].filter(([id, group]) => id || group.length).map(([id, group]) =>
    `<details class="folder" data-queue-folder="${id}"><summary>${folderHeading(id, folders.get(id))}</summary>
    ${group.length ? `<ul class="queue">${group.map(photoRow).join('')}</ul>` : ''}
    <button type="button" class="small folder-upload" data-folder-upload="${id}" aria-label="Agregar fotos a ${html(folders.get(id))}">Agregar fotos acá</button></details>`).join('');
  const queue = $('queue');
  if (queue.dataset.markup !== markup) {
    const wasOpen = new Map([...queue.children].map(group => [group.dataset.queueFolder, group.open]));
    queue.dataset.markup = markup;
    queue.innerHTML = markup;
    for (const group of queue.children) {
      const id = group.dataset.queueFolder;
      const count = groups.get(id).length;
      group.querySelector('[data-folder-title]').textContent = `${folders.get(id)} · ${count} ${count === 1 ? 'foto' : 'fotos'}`;
      group.open = wasOpen.get(id) ?? (id === ctx.uploadFolder || id === '');
    }
  }
  for (const group of queue.children) if (group.dataset.queueFolder === openFolder) group.open = true;
  for (const input of queue.querySelectorAll('[data-queue-select]')) {
    input.checked = ctx.selectedPhotos.has(input.dataset.queueSelect);
  }
  for (const input of queue.querySelectorAll('[data-queue-select], [data-queue-remove], [data-queue-move], [data-folder-edit], [data-folder-upload]')) input.disabled = busy;
}

function selectQueuedPhoto(event) {
  const input = event.target;
  const ctx = workspace;
  const id = input.dataset.queueSelect;
  const job = ctx?.local.get(id);
  if (!job || busy || job.status === 'done') return;
  if (input.checked) ctx.selectedPhotos.add(id);
  else ctx.selectedPhotos.delete(id);
  updateControls();
}

function toggleAllPhotos() {
  const ctx = workspace;
  if (!ctx || busy) return;
  const jobs = pendingPhotos(ctx);
  const selected = jobs.every(job => ctx.selectedPhotos.has(job.id));
  for (const { id } of jobs) {
    if (selected) ctx.selectedPhotos.delete(id);
    else ctx.selectedPhotos.add(id);
  }
  queueList();
  if (jobs.length) $('queue-wrap').open = true;
  updateControls();
}

async function previewQueuedPhoto(event) {
  const id = event.target.closest('[data-queue-preview]')?.dataset.queuePreview;
  const ctx = workspace;
  const job = ctx?.local.get(id);
  const saved = ctx?.values[id];
  if (!job?.source && !(saved?.transcripcion_eliminada && saved.photo === true)) return;
  const dialog = $('photo-preview');
  const image = $('preview-image');
  const figure = $('preview-figure');
  const status = $('preview-status');
  const ticket = ++previewTicket;
  $('preview-title').textContent = job?.name || saved?.nombre_archivo || 'Foto del remito';
  figure.hidden = $('preview-controls').hidden = true;
  $('preview-zoom').value = 100;
  zoomPhoto(figure, 100);
  status.textContent = 'Cargando foto…';
  dialog.showModal();
  try {
    const source = job?.source ? await jpegSource(job) : await sharedPhotoUrl(ctx, id);
    if (ticket !== previewTicket || !dialog.open) return;
    previewUrl = job?.source ? URL.createObjectURL(source) : source;
    image.onload = () => {
      figure.hidden = $('preview-controls').hidden = false;
      updatePanCursor(figure);
      status.textContent = '';
    };
    image.onerror = () => { status.textContent = 'No se pudo abrir la foto.'; };
    image.src = previewUrl;
  } catch (error) {
    if (ticket === previewTicket && dialog.open) status.textContent = error.message;
  }
}

function updatePanCursor(figure) {
  figure.classList.toggle('pannable', figure.scrollHeight > figure.clientHeight || figure.scrollWidth > figure.clientWidth);
}

function zoomPhoto(figure, value) {
  figure.classList.toggle('zoomed', Number(value) > 100);
  figure.querySelector('img').style.width = `${value}%`;
  updatePanCursor(figure);
}

function enablePhotoPan(figure) {
  let drag;
  figure.addEventListener('pointerdown', event => {
    if (event.pointerType !== 'mouse' || event.button !== 0 || event.target.tagName !== 'IMG') return;
    updatePanCursor(figure);
    if (!figure.classList.contains('pannable')) return;
    drag = { id: event.pointerId, x: event.clientX, y: event.clientY, left: figure.scrollLeft, top: figure.scrollTop };
    figure.setPointerCapture(event.pointerId);
    figure.classList.add('is-dragging');
    event.preventDefault();
  });
  figure.addEventListener('pointermove', event => {
    if (drag?.id !== event.pointerId) return;
    figure.scrollLeft = drag.left + drag.x - event.clientX;
    figure.scrollTop = drag.top + drag.y - event.clientY;
  });
  for (const type of ['pointerup', 'pointercancel', 'lostpointercapture']) figure.addEventListener(type, () => {
    drag = null;
    figure.classList.remove('is-dragging');
  });
  figure.querySelector('img').addEventListener('load', () => updatePanCursor(figure));
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
  const job = [...ctx.local.values()].find(job => job.source && !ctx.deletingPhotos.has(job.id)
    && (ctx.values[job.id]?.version === 1 || ctx.values[job.id]?.transcripcion_eliminada)
    && ctx.values[job.id].photo === undefined && !sync.photoDone.has(job.id));
  if (!job) return;
  sync.photoWork.add(job.id);
  (async () => {
    try {
      const jpeg = await encodePhoto(await resizePhoto(await jpegSource(job), SHARED_PHOTO_EDGE, SHARED_PHOTO_QUALITY));
      if (ctx.sync !== sync || !updatesAllowed(ctx) || !job.source || ctx.deletingPhotos.has(job.id)
        || ctx.values[job.id]?.photo !== undefined) return;
      await sync.room.update({ [`imagenes/${job.id}`]: jpeg, [`remitos/${job.id}/photo`]: true });
      sync.photoDone.add(job.id);
    } catch {
      if (ctx.sync === sync) sync.photoError = true;
    } finally {
      sync.photoWork.delete(job.id);
      updateSync();
      if (ctx.sync === sync) flushPhotos(ctx);
    }
  })();
}

function forgetPhoto(id) {
  const url = photoUrls.get(id);
  if (url?.startsWith('blob:')) URL.revokeObjectURL(url);
  photoUrls.delete(id);
}

function photoRemovalOptions(ctx, id) {
  const job = ctx.local.get(id);
  const photo = ctx.values[id]?.photo;
  const sharing = ctx.scope !== 'local' && !!job?.source && job.status === 'done' && photo === undefined;
  return { shared: photo === true, local: !!job?.source && !sharing, room: photo === true || sharing };
}

function setPhotoRemovalButtons(choices, disabled = false) {
  $('confirm-release-photo').disabled = disabled || !choices.some(option => option.local);
  $('confirm-remove-photo').disabled = disabled || !choices.some(option => option.room);
  $('cancel-remove-photo').disabled = disabled;
}

function askRemovePhoto(id, all = false) {
  const ctx = workspace;
  if (!ctx || busy) return;
  const ids = all ? [...loadedPhotos(ctx).keys()] : [id];
  const choices = ids.map(id => ({ id, ...photoRemovalOptions(ctx, id) })).filter(item => item.local || item.room);
  if (!choices.length) return;
  removal = { ctx, choices, all };
  const job = ctx.local.get(id), receipt = ctx.remote.get(id), saved = ctx.values[id];
  const options = choices[0];
  $('remove-photo-title').textContent = all ? '¿Dónde querés eliminar todas las fotos cargadas?' : '¿Dónde querés eliminar la foto?';
  $('remove-photo-name').textContent = all ? `${choices.length} ${choices.length === 1 ? 'foto cargada' : 'fotos cargadas'}`
    : job?.name || receipt?.nombre_archivo || saved?.nombre_archivo || 'Foto del remito';
  let info = 'La foto está guardada solo en este navegador.';
  if (options.shared) info = options.local
    ? 'Si elegís solo este navegador, los demás podrán seguir viendo la foto en la sala.'
    : 'La foto está compartida en la sala. Este navegador no tiene una copia original guardada.';
  else if (options.room) info = 'La copia compartida todavía se está guardando. Esperá a que esté disponible si querés conservarla.';
  if (all) {
    const local = choices.filter(option => option.local).length;
    info = `Podés quitar ${local} ${local === 1 ? 'copia local' : 'copias locales'} o eliminar las fotos cargadas también de la sala. Las transcripciones se conservan.`;
  }
  $('remove-photo-info').textContent = !all && receipt ? `${info} La transcripción se conserva.` : info;
  $('remove-photo-error').hidden = true;
  setPhotoRemovalButtons(choices);
  $('remove-photo-dialog').showModal();
}

function askRemoveReceipt(receipt, card) {
  receiptRemoval = { ctx: workspace, id: receipt.id };
  $('remove-receipt-name').textContent = card.querySelector('[data-receipt-title]').textContent;
  $('remove-receipt-info').textContent = workspace.scope === 'local'
    ? 'Se borrarán las líneas y verificaciones de este remito. La foto, si está guardada, se conserva.'
    : 'Se borrarán las líneas y verificaciones para todos en la sala. La foto, si está guardada, se conserva.';
  $('remove-receipt-error').hidden = true;
  $('remove-receipt-dialog').showModal();
}

async function confirmRemoveReceipt() {
  const { ctx, id } = receiptRemoval || {};
  if (!ctx || workspace !== ctx || busy || ctx.removingReceipt) return;
  const button = $('confirm-remove-receipt');
  const cancel = $('cancel-remove-receipt');
  const errorBox = $('remove-receipt-error');
  const receipt = ctx.remote.get(id);
  if (!receipt) {
    $('remove-receipt-dialog').close();
    return;
  }
  ctx.removingReceipt = id;
  button.disabled = cancel.disabled = true;
  errorBox.hidden = true;
  try {
    const remove = async () => {
      const photo = removedTranscript(receipt);
      if (ctx.scope === 'local') {
        const cached = (await readRecords(ctx.scope, 'cloud'))?.values;
        const next = { ...ctx, values: { ...ctx.values, ...cached, [id]: photo }, local: new Map(ctx.local),
          patches: new Map(ctx.patches), selectedPhotos: new Set(ctx.selectedPhotos) };
        await saveReceiptCache(next);
        Object.assign(ctx, { values: next.values, local: next.local, patches: next.patches, selectedPhotos: next.selectedPhotos });
      } else {
        if (!updatesAllowed(ctx) || !ctx.sync.ready) throw new Error('Conectate a la sala para eliminar la transcripción.');
        const result = await ctx.sync.ref.child(id).transaction(current => {
          if (current?.transcripcion_eliminada) return;
          return current?.version === 1 ? removedTranscript(current) : photo;
        }, undefined, false);
        const marker = result.snapshot.val();
        if (!marker?.transcripcion_eliminada) throw new Error('No se pudo confirmar la eliminación en la sala.');
        ctx.values = { ...ctx.values, [id]: ctx.values[id]?.transcripcion_eliminada ? ctx.values[id] : marker };
        await saveReceiptCache(ctx).catch(storeError);
      }
    };
    if (navigator.locks) await navigator.locks.request(`control-remitos:${ctx.scope}`, { ifAvailable: true }, lock => {
      if (!lock) throw new Error('Hay una tanda en curso en otra pestaña. Volvé a intentar cuando termine.');
      return remove();
    });
    else await remove();
    if (workspace === ctx) {
      $('remove-receipt-dialog').close();
      notice('Transcripción eliminada.');
    }
  } catch (error) {
    if (workspace === ctx) {
      errorBox.textContent = error.message || 'No se pudo eliminar la transcripción.';
      errorBox.hidden = false;
    }
  } finally {
    ctx.removingReceipt = null;
    button.disabled = cancel.disabled = false;
    if (workspace === ctx) {
      recalculate(ctx);
      render();
      flush(ctx);
      flushPhotos(ctx);
    }
  }
}

async function confirmRemovePhoto(localOnly) {
  const { ctx, choices, all } = removal || {};
  if (!ctx || workspace !== ctx || busy || ctx.deletingPhotos.size) return;
  const targets = localOnly ? choices.filter(option => option.local) : choices;
  if (!targets.length) return;
  const errorBox = $('remove-photo-error');
  for (const { id } of targets) ctx.deletingPhotos.add(id);
  setPhotoRemovalButtons(choices, true);
  errorBox.hidden = true;
  try {
    const remove = async () => {
      const records = all ? await readRecords(ctx.scope) : [await readRecords(ctx.scope, targets[0].id)];
      const jobs = new Map(records.filter(job => job?.type === 'photo').map(job => [job.id, job]));
      const roomTargets = ctx.scope === 'local' ? [] : targets.filter(({ id, shared, room }) => localOnly
        ? shared || (jobs.get(id)?.status === 'done' && ctx.values[id]?.photo !== false)
        : room || (jobs.get(id)?.status === 'done' && ctx.values[id]?.photo !== false));
      const photos = {}, metadata = {};
      if (roomTargets.length) {
        const sync = ctx.sync;
        const offline = localOnly ? 'Conectate a la sala para comprobar que la foto reducida está guardada.'
          : 'Conectate a la sala para eliminar las fotos compartidas.';
        if (!updatesAllowed(ctx)) throw new Error(offline);
        // La marca se publica de forma atómica con el JPEG. Leer metadatos evita descargar las fotos.
        const snapshot = all ? (await sync.ref.once('value')).val() || {}
          : { [targets[0].id]: (await sync.ref.child(targets[0].id).once('value')).val() };
        if (ctx.sync !== sync || !updatesAllowed(ctx)) throw new Error(offline);
        for (const { id } of roomTargets) {
          const saved = snapshot[id];
          if (localOnly && (saved?.photo !== true || ctx.values[id]?.photo !== true)) {
            throw new Error(all ? 'No se pudieron comprobar todas las copias compartidas. Las fotos locales se conservaron.'
              : 'La copia compartida no está disponible. Se conservó la foto local.');
          }
          if (!localOnly && (saved?.version === 1 || saved?.transcripcion_eliminada) && saved.photo !== false) {
            photos[`imagenes/${id}`] = null;
            photos[`remitos/${id}/photo`] = false;
            if (ctx.values[id]) metadata[`${id}/photo`] = false;
          }
        }
        if (Object.keys(photos).length) {
          await sync.room.update(photos);
          ctx.values = withReceiptChanges(ctx.values, metadata);
        }
      }
      const stripped = new Map(), deletes = [];
      for (const { id } of targets) {
        const job = jobs.get(id);
        if (job?.receipt) {
          // La foto y la transcripción tienen ciclos independientes, incluso en modo local.
          const record = { ...job, source: null, preview: null, removedFromQueue: true };
          if (!localOnly || ctx.values[id]?.photo === false) record.payload = { ...(job.payload || publicReceipt(job.receipt)), photo: false };
          stripped.set(id, record);
        } else deletes.push([ctx.scope, id]);
      }
      await storeRecords([...stripped.values()], deletes);
      for (const { id } of targets) {
        if (stripped.has(id)) ctx.local.set(id, stripped.get(id));
        else ctx.local.delete(id);
        ctx.selectedPhotos.delete(id);
        forgetPhoto(id);
      }
    };
    if (navigator.locks) await navigator.locks.request(`control-remitos:${ctx.scope}`, { ifAvailable: true }, lock => {
      if (!lock) throw new Error('Hay una tanda en curso en otra pestaña. Volvé a intentar cuando termine.');
      return remove();
    });
    else await remove();
    if (workspace === ctx) {
      $('remove-photo-dialog').close();
      jobMessage(all ? `${targets.length} ${targets.length === 1 ? 'foto eliminada' : 'fotos eliminadas'}${localOnly ? ' de este navegador' : ''}.`
        : localOnly && ctx.values[targets[0].id]?.photo === true ? 'Copia local liberada. La foto reducida sigue disponible en la sala.' : 'Foto eliminada.');
    }
  } catch (error) {
    if (workspace === ctx) {
      errorBox.textContent = error.message || 'No se pudo eliminar la foto.';
      errorBox.hidden = false;
    }
  } finally {
    for (const { id } of targets) ctx.deletingPhotos.delete(id);
    if (workspace === ctx) {
      setPhotoRemovalButtons(choices);
      recalculate(ctx);
      render();
      flush(ctx);
      flushPhotos(ctx);
    }
  }
}

function askMovePhoto(id) {
  const ctx = workspace;
  const job = ctx?.local.get(id);
  const receipt = ctx?.remote.get(id);
  if ((!job && !receipt) || busy) return;
  move = { ctx, id };
  folderOptions($('move-folder'), folderNames(ctx), receipt?.carpeta_id ?? job?.carpeta_id);
  $('move-photo-error').hidden = true;
  $('move-photo-dialog').showModal();
}

async function movePhoto(event) {
  event.preventDefault();
  const { ctx, id } = move || {};
  if (!ctx || workspace !== ctx || busy) return;
  const folder = $('move-folder').value;
  const job = ctx.local.get(id);
  const button = $('confirm-move-photo');
  button.disabled = true;
  try {
    const save = async () => {
      const fresh = job && await readRecords(ctx.scope, id);
      if (fresh?.receipt && !ctx.remote.has(id)) {
        ctx.local.set(id, fresh);
        recalculate(ctx);
      }
      if (ctx.remote.has(id)) {
        if (!await changeReceipt({ [`${id}/carpeta_id`]: folder || null }, ctx)) throw new Error();
      } else if (fresh) {
        const updated = { ...fresh, carpeta_id: folder };
        await storeRecords([updated]);
        ctx.local.set(id, updated);
      } else throw new Error('La foto ya fue eliminada en otra pestaña.');
    };
    if (navigator.locks) await navigator.locks.request(`control-remitos:${ctx.scope}`, { ifAvailable: true }, lock => {
      if (!lock) throw new Error('Hay una tanda en curso en otra pestaña. Volvé a intentar cuando termine.');
      return save();
    });
    else await save();
    if (workspace === ctx) {
      $('move-photo-dialog').close();
      queueList(folder);
      jobMessage('Foto movida.');
    }
  } catch (error) {
    $('move-photo-error').textContent = error.message || 'No se pudo mover la foto.';
    $('move-photo-error').hidden = false;
  } finally { button.disabled = false; }
}

async function sharedPhotoUrl(ctx, id) {
  if (!ctx.sync?.online) throw new Error('Conectate a la sala para ver la foto.');
  const snapshot = await ctx.sync.room.child('imagenes').child(id).once('value');
  const jpeg = snapshot.val();
  if (typeof jpeg !== 'string' || !/^[A-Za-z0-9+/]+={0,2}$/.test(jpeg)) throw new Error('La foto no está disponible.');
  return `data:image/jpeg;base64,${jpeg}`;
}

async function photoAction(event) {
  const button = event.target.closest('button[data-photo]');
  const id = button?.closest('article')?.dataset.id;
  const ctx = workspace;
  const receipt = ctx?.remote.get(id);
  if (!receipt) return;
  if (button.dataset.photo === 'remove') return askRemovePhoto(id);
  if (button.dataset.photo === 'move') return askMovePhoto(id);
  const job = ctx.local.get(id);
  const feedback = button.closest('.photo-block').querySelector('[data-photo-feedback]');
  const details = button.closest('.receipt-toggle');
  button.disabled = true;
  feedback.textContent = '';
  try {
    if (!photoUrls.has(id)) {
      if (job?.source) {
        const source = await jpegSource(job);
        if (workspace !== ctx || !details.open || ctx.remote.get(id)?.photo === false || !job.source || ctx.deletingPhotos.has(id)) return;
        photoUrls.set(id, URL.createObjectURL(source));
      } else {
        const url = await sharedPhotoUrl(ctx, id);
        if (workspace !== ctx || !details.open || ctx.remote.get(id)?.photo !== true || ctx.deletingPhotos.has(id)) return;
        photoUrls.set(id, url);
      }
    }
    if (workspace === ctx) render();
  } catch (error) {
    feedback.textContent = error.message || 'No se pudo acceder a la foto.';
  } finally { button.disabled = false; }
}

function choosePhotos(folder) {
  const ctx = workspace;
  if (!ctx || busy) return;
  if (folder !== undefined) selectUploadFolder(ctx, folder);
  photoUpload = { ctx, folder: ctx.uploadFolder };
  $('images').click();
}

async function addPhotos(files, folder = workspace?.uploadFolder) {
  const ctx = workspace;
  if (!ctx || !files.length) return;
  if (files.length > MAX_UPLOAD) return jobMessage('Seleccioná hasta 10 imágenes por carga.', true);
  if (files.some(file => !/\.(jpe?g|png|hei[cf])$/i.test(file.name) && !/^image\/(jpeg|png|hei[cf])$/.test(file.type))) {
    return jobMessage('Elegí solo imágenes JPG, PNG o HEIC.', true);
  }
  const jobs = files.map(source => ({ id: newId(), scope: ctx.scope, type: 'photo', name: source.name, source, status: 'queued', carpeta_id: folder }));
  try {
    await storeRecords(jobs);
    for (const job of jobs) ctx.local.set(job.id, job);
    if (workspace !== ctx) return;
    const name = folderNames(ctx).get(folder);
    jobMessage(`${files.length} ${files.length === 1 ? 'foto cargada' : 'fotos cargadas'} en «${name}». Marcá las que querés transcribir.`);
    queueList(folder);
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
    const savedPhotos = new Map(stored.filter(record => record.type === 'photo').map(job => [job.id, job]));
    for (const [id, job] of ctx.local) if (!job.receipt && !savedPhotos.has(id)) ctx.local.delete(id);
    for (const [id, job] of savedPhotos) if (!ctx.local.get(id)?.receipt) ctx.local.set(id, job);
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
        job.receipt = { id: job.id, ...outcome.receipt, nombre_archivo: job.name, carpeta_id: job.carpeta_id,
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
  const result = workspace.comparison.results.get(id);
  if (result) return result;
  const { delivery: date, returnDay, returnLabel } = datesFor(receipt)[index];
  return { id, receipt, index, line, date, returnDay, returnLabel,
    chosen: line.elegida, state: line.estado || 'pending', decision: line.decision };
}

function similarityBadge(value) {
  const level = value === 100 ? 'full' : value >= 80 ? 'high' : value >= 50 ? 'mid' : 'low';
  return `<span class="similarity ${level}"><strong>${value}%</strong> similitud del material</span>`;
}

const comparisonLabels = {
  match: 'Coincide', similar: 'Similar', different: 'No coincide', unavailable: 'Sin comparar'
};

function formatRow(row, result, { canChoose = false, assigned = false, suggested = false, similarity,
  returnReview = reviewReturn(result, row), editFields = false } = {}) {
  const correction = result.correccionExcel?.fila === row.fila ? result.correccionExcel : null;
  const movement = correction?.movimiento ?? row.movimiento;
  const fields = compareRowFields(result, row, similarity);
  const edits = excelRowEdits(result, row);
  if (returnReview) fields.push(...returnReview.fields);
  const cells = fields.map(field => {
    const score = field.similarity != null && field.status !== 'unavailable' ? ` · ${field.similarity}%` : '';
    const days = field.days > 0 ? `${field.days} ${field.days === 1 ? 'día' : 'días'} de diferencia` : '';
    const action = editFields && field.key ? `<button type="button" class="small field-edit" data-action="excel-field"
      data-excel-field="${field.key}" aria-label="Cambiar ${html(field.label)} en Excel">${field.key === 'fecha_devolucion' ? 'Cargar fecha' : 'Cambiar'}</button>` : '';
    return `<tr><th scope="row">${field.label}</th><td>${html(field.remito || '—')}</td>
      <td>${html(field.excel || '—')}${field.detail ? `<small>${html(field.detail)}</small>` : ''}</td>
      <td><span class="field-match ${field.status}">${comparisonLabels[field.status]}${score}</span>${days ? `<small>${days}</small>` : ''}
      ${edits[field.key] && field.key !== 'fecha_devolucion' ? '<small>Guardado para exportar</small>' : ''}${action}</td></tr>`;
  }).join('');
  return `<div class="row-detail"><b>${suggested ? 'Mejor sugerencia · ' : ''}Fila ${row.fila}</b>
    <table class="comparison-fields" aria-label="Comparación con la fila ${row.fila}">
      <thead><tr><th scope="col">Dato</th><th scope="col">Remito</th><th scope="col">Excel</th><th scope="col">Resultado</th></tr></thead>
      <tbody>${cells}</tbody>
    </table>
    ${!returnReview && movement ? `<div>Movimiento en Excel · ${html(movement)}</div>` : ''}
    ${correction ? `<div class="movement-correction">Corregido para exportar · Original ${html(row.movimiento || 'sin movimiento')}</div>` : ''}
    ${canChoose && returnReview?.movementWarning ? '<div class="warning">Posible error de movimiento entre el remito y el Excel.</div>' : ''}
    ${assigned ? '<div class="warning">Fila ya asignada a otra línea.</div>' : ''}
    ${canChoose ? `<button type="button" class="small" data-pick="${row.fila}">Elegir fila ${row.fila}</button>` : ''}</div>`;
}

function readField(label, value, wide = false) {
  return `<label class="field${wide ? ' wide' : ''}">${label}<span class="read-field">${html(value || '—')}</span></label>`;
}

function editField(field, label, value, editable, wide = false) {
  if (!editable) return readField(label, value, wide);
  const editor = ['observacion', 'descripcion', 'nota_manuscrita'].includes(field)
    ? `<textarea rows="2" data-field="${field}" aria-label="${label}">${html(value)}</textarea>`
    : `<input type="text" data-field="${field}" aria-label="${label}" value="${html(value)}"${field === 'fecha_devolucion' ? ' placeholder="dd/mm/aa"' : ''}>`;
  return `<label class="field${wide ? ' wide' : ''}">${label}${editor}</label>`;
}

function rubroMarkup(line) {
  const options = Object.entries(productCategories).map(([id, name]) => `<option value="${id}"${line.rubro === id ? ' selected' : ''}>${html(name)}</option>`).join('');
  return `<label class="field">Rubro sugerido <select data-field="rubro" aria-label="Rubro sugerido">${options}</select></label>
    <label class="field" data-other-rubro${line.rubro === 'otros' ? '' : ' hidden'}>Otro rubro <input data-field="otro_rubro"
      type="text" maxlength="80" value="${html(line.otro_rubro || '')}" aria-label="Otro rubro"></label>`;
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
      ${rubroMarkup(line)}<button type="button" class="small" data-action="recompare" hidden>Volver a buscar en Excel</button>` : `<span class="read-field">${html(line.descripcion || '—')}</span>`}
      ${highlight}<details data-return-details><summary>Devolución y notas <span class="return-date" data-return-date></span></summary><div class="edit-grid">
      ${editable ? `<label class="check"><input type="checkbox" data-field="devolucion_marcada" ${line.devolucion_marcada ? 'checked' : ''}> Devolución marcada en el remito</label>` : ''}
      ${editField('fecha_devolucion', 'Fecha de devolución', line.fecha_devolucion, editable)}
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
      <button type="button" class="small" data-photo="move" hidden>Mover de carpeta</button>
      <button type="button" class="small remove-photo" data-photo="remove" aria-label="Eliminar foto" hidden>×</button></div>
      <p class="hint" data-photo-status></p><p class="hint" data-photo-feedback role="status"></p>
      <figure class="pan-photo" tabindex="0" aria-label="Foto del remito" hidden><img alt="Foto del remito" draggable="false"></figure>
      <label class="check" data-photo-zoom hidden>Zoom <input type="range" min="100" max="400" step="25" value="100" aria-label="Zoom de la foto"></label>
    </div><div class="review">
      ${receipt.lines.length ? `<div class="table-wrap"><table class="review-table" aria-label="Líneas del remito"><thead><tr><th>Fecha</th><th>Descripción</th><th>Cant.</th><th>Comparación</th></tr></thead>
      <tbody>${receipt.lines.map(line => lineMarkup(line, true)).join('')}</tbody></table></div>`
    : '<p class="empty">El remito no tiene líneas transcritas.</p>'}
      <div class="receipt-foot actions"><button type="button" class="small" data-action="add">Agregar línea</button>
      <button type="button" class="small danger" data-action="remove-receipt">Eliminar transcripción</button></div></div></div></details>`;
  enablePhotoPan(card.querySelector('figure'));
  const details = card.querySelector('.receipt-toggle');
  details.addEventListener('toggle', () => {
    if (details.open || !card.isConnected || !photoUrls.has(receipt.id)) return;
    forgetPhoto(receipt.id);
    renderPhoto(card, workspace.remote.get(receipt.id), workspace.local.get(receipt.id));
  });
  return card;
}

function renderPhoto(card, receipt, job) {
  const cloud = workspace.scope !== 'local';
  const available = receipt.photo === true || (receipt.photo !== false && !!job?.source);
  if (!available) forgetPhoto(receipt.id);
  const url = photoUrls.get(receipt.id);
  card.querySelector('[data-photo="open"]').hidden = !available || !!url;
  card.querySelector('[data-photo="remove"]').hidden = !available && !job?.source;
  card.querySelector('[data-photo-status]').textContent = cloud && receipt.photo === undefined && job?.source
    ? 'Compartiendo foto…' : available ? '' : 'Foto no disponible';
  card.querySelector('figure').hidden = card.querySelector('[data-photo-zoom]').hidden = !url;
  const image = card.querySelector('figure img');
  if (url && image.getAttribute('src') !== url) image.src = url;
  if (!url) image.removeAttribute('src');
}

function renderStatus(cell, result, returnReview) {
  const mark = result.line.verificado;
  const best = result.candidates?.[0];
  const signature = JSON.stringify([result.state, result.chosen, result.decision, result.conflict, result.stale, result.date, result.line.fecha_linea,
    result.line.cantidad, result.line.descripcion, result.receipt.numero_remito, result.receipt.persona, result.receipt.obra_o_vehiculo,
    best, !!workspace.excel, mark, returnReview, result.correccionExcel, result.ajustesExcel]);
  if (cell.dataset.signature === signature) return;
  cell.dataset.signature = signature;
  const chosen = result.chosen ? formatRow(result.chosen, result, { returnReview, editFields: !!workspace.excel }) : '';
  const suggestion = !chosen && best ? formatRow(best.row, result, { suggested: true, similarity: best.similarity,
    assigned: best.assigned, returnReview, editFields: !!workspace.excel && !best.assigned }) : '';
  const noSuggestion = !chosen && !best && workspace.excel && result.decision !== 'absent' && ['missing', 'doubtful'].includes(result.state);
  const action = (id, label) => `<button type="button" class="small" data-action="${id}">${label}</button>`;
  const buttons = [];
  if (returnReview?.needsDate) buttons.push(action('return-date', 'Completar fecha de devolución'));
  else if (returnReview?.missingExcelDate) buttons.push(action('return-date', 'Revisar fecha de devolución'));
  if (result.state !== 'skipped') {
    if (result.chosen && !result.decision) buttons.push(action('confirm', 'Confirmar'));
    if (workspace.excel && result.chosen) buttons.push(action('excel-movement', 'Corregir movimiento'));
    if (workspace.excel) buttons.push(action('choose', result.chosen ? 'Elegir otra fila' : result.candidates?.length ? 'Ver candidatos' : 'Buscar en Excel'));
    if (result.decision !== 'absent') buttons.push(action('absent', 'No está en el Excel'));
    if (result.decision) buttons.push(action('undo', 'Reabrir'));
  }
  const canVerify = mark || ['ok', 'doubtful', 'missing'].includes(result.state);
  const verificationControl = canVerify ? `<label class="check"><input type="checkbox" data-verify ${mark ? 'checked' : ''}> ✓ Verificado</label>` : '';
  const verificationBadge = mark ? `<div class="verification-mark">${verificationControl}
    <p>Última verificación <time datetime="${new Date(mark.fecha).toISOString()}">${loadTime.format(mark.fecha)}</time></p>
    ${mark.por ? `<p>${html(mark.por)}</p>` : ''}</div>` : '';
  cell.innerHTML = `${verificationBadge}<span class="status ${result.state}">${states[result.state]}</span>${chosen}${suggestion}
    ${returnReview?.needsDate ? '<p class="return-alert">Falta una fecha de devolución completa en el remito.</p>' : ''}
    ${returnReview?.movementWarning ? '<p class="return-alert">Posible error de movimiento. El remito indica devolución y el Excel no la registra como tal.</p>' : ''}
    ${noSuggestion ? similarityBadge(0) : ''}
    ${result.state === 'skipped' && workspace.excel ? '<p class="hint">Esta línea está fuera de la comparación de resaltadas.</p>' : ''}
    ${result.decision ? `<p class="details-text">${decisions[result.decision]}</p>` : ''}
    ${result.conflict ? '<p class="details-text">La fila ya está asignada a otra línea.</p>' : ''}
    ${result.stale ? '<p class="details-text">La fila elegida cambió en el Excel de la sala.</p>' : ''}
    ${result.state === 'pending' ? '<p class="hint">Cargá el Excel para comparar.</p>' : ''}
    ${mark ? '' : verificationControl}
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

function lineCounts() {
  return { total: 0, ok: 0, doubtful: 0, missing: 0, skipped: 0, pending: 0 };
}

function renderFolderDetail(panel, count, totals) {
  const signature = JSON.stringify([count, totals]);
  if (panel.dataset.signature === signature) return;
  panel.dataset.signature = signature;
  const group = (label, lines, className = '') => `<div class="${className}">
    <h3>${label} · ${lines.total} ${lines.total === 1 ? 'línea' : 'líneas'}</h3>
    <dl class="folder-counts">
      <div class="ok"><dt>Coinciden</dt><dd data-folder-count="ok">${lines.ok}</dd></div>
      <div class="doubtful"><dt>Dudosas</dt><dd data-folder-count="doubtful">${lines.doubtful}</dd></div>
      <div class="missing"><dt>No encontradas</dt><dd data-folder-count="missing">${lines.missing}</dd></div>
      <div><dt>Sin comparar</dt><dd data-folder-count="skipped">${lines.skipped + lines.pending}</dd></div>
    </dl></div>`;
  const total = totals.highlighted.total + totals.plain.total;
  panel.innerHTML = `<p class="hint">${count} ${count === 1 ? 'remito' : 'remitos'} · ${total} ${total === 1 ? 'línea de materiales' : 'líneas de materiales'}</p>
    <div class="folder-stats-grid">${group('Resaltadas en rosa', totals.highlighted, 'pink-lines')}${group('Sin resaltar', totals.plain, 'plain-lines')}</div>
    ${totals.plain.skipped ? '<p class="hint">Para comparar también las líneas sin resaltar, activá «Comparar todas las líneas».</p>' : ''}`;
}

function setFolderDetail(group, visible) {
  group.querySelector('[data-folder-stats]').hidden = !visible;
  const button = group.querySelector('[data-folder-detail]');
  button.setAttribute('aria-expanded', String(visible));
  button.textContent = visible ? 'Ocultar detalle' : 'Ver detalle';
}

function toggleFolderDetail(event) {
  const button = event.target.closest('[data-folder-detail]');
  if (!button) return;
  event.preventDefault();
  const group = button.closest('[data-review-folder]');
  const visible = group.querySelector('[data-folder-stats]').hidden;
  if (visible) group.open = true;
  setFolderDetail(group, visible);
}

function render() {
  const ctx = workspace;
  if (!ctx) return;
  const fileName = receipt => receipt.nombre_archivo || ctx.local.get(receipt.id)?.name;
  const receipts = receiptList().sort((a, b) => compareFileNames(fileName(a), fileName(b)));
  const ids = new Set(receipts.map(receipt => receipt.id));
  for (const [id, entry] of cards) {
    if (!ids.has(id)) {
      forgetPhoto(id);
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
  const folders = folderNames(ctx);
  const folderPositions = new Map();
  const visibleInFolder = new Map();
  const countsByFolder = new Map();
  for (const [id] of folders) {
    if (!reviewFolders.has(id)) {
      const group = document.createElement('details');
      group.className = 'folder';
      group.dataset.reviewFolder = id;
      group.open = true;
      group.innerHTML = `<summary>${folderHeading(id, folders.get(id), true)}</summary>
        <section class="folder-stats" id="${folderDetailId(id)}" data-folder-stats hidden></section>
        <div class="folder-cards"></div><p class="hint folder-cards">Sin remitos transcritos</p>`;
      group.addEventListener('toggle', () => { if (!group.open) setFolderDetail(group, false); });
      reviewFolders.set(id, group);
    }
    folderPositions.set(id, 0);
    countsByFolder.set(id, { highlighted: lineCounts(), plain: lineCounts() });
  }
  receipts.forEach(receipt => {
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
    const folder = photoFolder(ctx, receipt.id);
    const group = reviewFolders.get(folder).querySelector(':scope > div.folder-cards');
    const position = folderPositions.get(folder);
    if (group.children[position] !== card) group.insertBefore(card, group.children[position] || null);
    folderPositions.set(folder, position + 1);
    const title = receipt.numero_remito ? `Remito ${receipt.numero_remito}` : 'Remito sin número';
    const photoName = fileName(receipt);
    card.querySelector('[data-receipt-title]').textContent = photoName ? `${title} · ${photoName}` : title;
    renderFields(card.querySelector('.receipt-head'), receipt);
    renderPhoto(card, receipt, job);
    card.querySelector('[data-photo="move"]').hidden = folders.size === 1;
    const dates = datesFor(receipt);
    const rows = new Map([...card.querySelectorAll('tr[data-line]')].map(row => [row.dataset.line, row]));
    const matches = { ok: 0, doubtful: 0, missing: 0, pending: 0 };
    let compared = 0, finished = 0, rowVisible = false;
    receipt.lines.forEach((line, index) => {
      const result = resultFor(receipt, index);
      const returnReview = reviewReturn(result, result.chosen || result.candidates?.[0]?.row);
      const counts = countsByFolder.get(folder)[line.resaltado ? 'highlighted' : 'plain'];
      counts.total++;
      if (Object.hasOwn(counts, result.state)) counts[result.state]++;
      if (['ok', 'doubtful', 'missing'].includes(result.state)) totals[result.state]++;
      if (Object.hasOwn(matches, result.state)) matches[result.state]++;
      if (result.state !== 'skipped') compared++;
      if (result.decision) finished++;
      if (line.verificado) totals.verified++;
      const row = rows.get(line.key);
      if (!row) return;
      renderFields(row, line);
      const otherRubro = row.querySelector('[data-other-rubro]');
      if (otherRubro) otherRubro.hidden = line.rubro !== 'otros';
      const compare = row.querySelector('[data-action="recompare"]');
      compare.hidden = !ctx.excel;
      compare.textContent = result.state === 'skipped' ? 'Comparar esta línea' : 'Volver a buscar en Excel';
      row.classList.toggle('highlighted', result.line.resaltado);
      const hasIssue = ['doubtful', 'missing'].includes(result.state) || returnReview?.needsDate
        || returnReview?.missingExcelDate || returnReview?.movementWarning || returnReview?.dateMismatch;
      row.hidden = ($('only-issues').checked && !hasIssue)
        || (filter !== 'all' && !!line.verificado !== (filter === 'yes'));
      rowVisible ||= !row.hidden;
      const label = dates[index].label;
      row.querySelector('[data-effective-date]').textContent = label !== result.line.fecha_linea ? label : '';
      row.querySelector('[data-return-date]').textContent = dates[index].returnLabel ? `· ${dates[index].returnLabel}` : '';
      renderStatus(row.querySelector('[data-result]'), result, returnReview);
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
    if (!card.hidden) {
      visible++;
      visibleInFolder.set(folder, (visibleInFolder.get(folder) || 0) + 1);
    }
  });
  let folderPosition = 0;
  for (const [id, name] of folders) {
    const group = reviewFolders.get(id);
    const count = folderPositions.get(id);
    const shown = visibleInFolder.get(id) || 0;
    group.hidden = id ? count > 0 && shown === 0 : shown === 0;
    group.querySelector('[data-folder-title]').textContent = `${name} · ${shown} ${shown === 1 ? 'remito' : 'remitos'}`;
    group.querySelector('[data-folder-edit]').setAttribute('aria-label', `Editar nombre de ${name}`);
    group.querySelector('[data-folder-detail]').hidden = !count;
    if (count) renderFolderDetail(group.querySelector('[data-folder-stats]'), count, countsByFolder.get(id));
    else setFolderDetail(group, false);
    group.querySelector(':scope > p').hidden = count > 0;
    if (container.children[folderPosition] !== group) container.insertBefore(group, container.children[folderPosition] || null);
    folderPosition++;
  }
  for (const [id, group] of reviewFolders) if (!folders.has(id)) {
    group.remove();
    reviewFolders.delete(id);
  }
  for (const [key, count] of Object.entries(totals)) $(`count-${key}`).textContent = count;
  $('empty').hidden = visible > 0 || (!receipts.length && folders.size > 1);
  $('empty').textContent = receipts.length ? 'No hay remitos con esos filtros.' : 'Agregá fotos para empezar o abrí el control desde una sala del planificador para ver sus resultados.';
  $('export-csv').disabled = !receipts.some(receipt => receipt.lines.length);
  $('export-xlsx').disabled = $('export-csv').disabled && !ctx.workbook;
  updateControls();
  queueList();
  updateSync();
}

const linePath = (receipt, key) => receipt.id + '/lines/' + key;

function changeLine(receipt, key, fields, ctx = workspace) {
  const path = linePath(receipt, key);
  return changeReceipt(Object.fromEntries(Object.entries(fields).map(([field, value]) => [path + '/' + field, value])), ctx);
}

function excelFieldRow(result) {
  const candidate = result?.candidates?.[0];
  return result?.chosen || (!candidate?.assigned ? chosenRow(candidate?.row) : null);
}

function openExcelField(receipt, result, field) {
  const row = excelFieldRow(result);
  if (!workspace.excel || !row || !Object.hasOwn(excelEditLabels, field)) return;
  const returnDate = field === 'fecha_devolucion';
  const entry = result.ajustesExcel?.[field] ?? null;
  const remito = [...compareRowFields(result, row), ...(reviewReturn(result, row)?.fields || [])].find(item => item.key === field)?.remito;
  const copied = excelEditValue(field, remito);
  const current = entry?.valor ?? (returnDate ? result.correccionExcel?.fecha_devolucion : row[field]);
  const input = $('excel-edit-value');
  excelFieldEdit = { ctx: workspace, id: receipt.id, key: result.line.key, field, row, entry, copied,
    legacyDate: returnDate ? result.correccionExcel?.fecha_devolucion : undefined };
  $('excel-edit-row').textContent = `${excelEditLabels[field]} · Fila ${row.fila}`;
  $('excel-edit-original').textContent = `Excel · ${excelEditLabel(field, current) || 'Sin dato'}`;
  $('excel-edit-remito').textContent = `Remito · ${remito || 'Sin dato'}`;
  $('excel-edit-info').textContent = result.chosen ? 'El cambio se aplica al XLSX exportado.'
    : `Al guardar, la fila ${row.fila} queda vinculada a esta línea y el cambio se aplica al XLSX exportado.`;
  input.value = excelEditLabel(field, current || (returnDate ? copied : ''));
  input.placeholder = field === 'fecha' || returnDate ? 'dd/mm/aaaa' : '';
  input.inputMode = field === 'cantidad' ? 'decimal' : 'text';
  $('excel-edit-copy').disabled = !copied;
  $('excel-edit-restore').hidden = returnDate ? !current : !entry;
  $('excel-edit-error').hidden = true;
  $('excel-edit-dialog').showModal();
}

async function saveExcelField(restore = false) {
  const edit = excelFieldEdit;
  if (!edit || workspace !== edit.ctx || $('excel-edit-save').disabled) return;
  const { ctx, id, key, field, row } = edit;
  const receipt = ctx.remote.get(id);
  const index = receipt?.lines.findIndex(line => line.key === key) ?? -1;
  const result = index < 0 ? null : resultFor(receipt, index);
  const error = $('excel-edit-error');
  const fail = message => { error.textContent = message; error.hidden = false; };
  if (!ctx.excel || !same(excelFieldRow(result), row) || !same(result?.ajustesExcel?.[field] ?? null, edit.entry)
    || (field === 'fecha_devolucion' && result.correccionExcel?.fecha_devolucion !== edit.legacyDate)) {
    fail('La fila o el valor cambió. Cerrá el editor y volvé a abrirlo.');
    return;
  }
  const value = excelEditValue(field, $('excel-edit-value').value);
  if (!restore && !value) {
    fail('Ingresá un valor válido. Para fechas, incluí día, mes y año.');
    return;
  }
  const entry = restore && field !== 'fecha_devolucion' ? null : { fila: row.fila, valor: restore ? '' : value };
  const newEdits = { ...(!result.chosen && field !== 'fecha_devolucion' && { fecha_devolucion: { fila: row.fila, valor: '' } }),
    ...(entry && { [field]: entry }) };
  const fields = result.decision ? { ['ajustes_excel/' + field]: entry }
    : { ajustes_excel: newEdits, correccion_excel: null,
      decision: result.chosen ? 'confirmed' : 'manual', elegida: row, estado: 'ok' };
  const controls = ['excel-edit-save', 'excel-edit-cancel', 'excel-edit-restore', 'excel-edit-copy', 'excel-edit-value'].map($);
  for (const control of controls) control.disabled = true;
  error.hidden = true;
  try {
    if (!await changeLine(receipt, key, fields, ctx)) fail('No se pudo guardar el cambio. Volvé a intentar.');
    else if (excelFieldEdit === edit) $('excel-edit-dialog').close();
  } finally {
    for (const control of controls) control.disabled = false;
    $('excel-edit-copy').disabled = !edit.copied;
  }
}

function updateMovementDate() {
  const isReturn = isReturnMovement($('movement-value').value);
  $('movement-date-field').hidden = !isReturn;
  $('movement-date').required = isReturn;
}

function openMovement(receipt, result) {
  const row = result.chosen;
  if (!workspace.excel || !row) return;
  const correction = result.correccionExcel ?? null;
  movementEdit = { ctx: workspace, id: receipt.id, key: result.line.key, row, correction };
  $('movement-row').textContent = `Fila ${row.fila} · ${row.herramienta}`;
  $('movement-original').textContent = `Movimiento original · ${row.movimiento || 'Sin movimiento'}`;
  const values = new Set(['A CARGO', 'DEVUELTO', 'PARA OBRA', 'ENTREGADO', 'REPARACIÓN', 'PARA SERVICIO', 'BAJA', 'CANCELADO',
    row.movimiento, correction?.movimiento].filter(Boolean));
  const select = $('movement-value');
  select.replaceChildren(new Option('Elegí un movimiento', ''), ...[...values].map(value => new Option(value, value)));
  select.value = correction?.movimiento ?? row.movimiento ?? '';
  $('movement-date').value = excelRowEdits(result, row).fecha_devolucion ?? correction?.fecha_devolucion
    ?? (result.returnDay == null ? '' : new Date(result.returnDay * 86400000).toISOString().slice(0, 10));
  $('restore-movement').hidden = !correction;
  $('movement-error').hidden = true;
  updateMovementDate();
  $('movement-dialog').showModal();
}

async function saveMovement(restore = false) {
  const edit = movementEdit;
  if (!edit || workspace !== edit.ctx || $('save-movement').disabled) return;
  const { ctx, id, key, row } = edit;
  const error = $('movement-error');
  const fail = message => { error.textContent = message; error.hidden = false; };
  const receipt = ctx.remote.get(id);
  const index = receipt?.lines.findIndex(line => line.key === key) ?? -1;
  const result = index < 0 ? null : resultFor(receipt, index);
  if (!ctx.excel || !same(result?.chosen, row) || !same(result.correccionExcel ?? null, edit.correction)) {
    fail('La fila o su corrección cambió. Cerrá el editor y volvé a abrirlo.');
    return;
  }
  const correction = restore ? null : excelCorrection({
    fila: row.fila, movimiento: $('movement-value').value, fecha_devolucion: $('movement-date').value
  });
  if (!restore && !correction) {
    fail('Elegí un movimiento e ingresá una fecha válida si es una devolución.');
    return;
  }
  const controls = ['save-movement', 'restore-movement', 'cancel-movement', 'movement-value', 'movement-date'].map($);
  for (const control of controls) control.disabled = true;
  error.hidden = true;
  try {
    const fields = { correccion_excel: correction };
    if (correction?.fecha_devolucion && result.ajustesExcel?.fecha_devolucion) fields['ajustes_excel/fecha_devolucion'] = null;
    if (!result.decision) Object.assign(fields, { decision: 'confirmed', elegida: row, estado: 'ok' });
    if (!await changeLine(receipt, key, fields, ctx)) fail('No se pudo guardar la corrección. Volvé a intentar.');
    else if (movementEdit === edit) $('movement-dialog').close();
  } finally {
    for (const control of controls) control.disabled = false;
  }
}

function openVerification(receipt, line) {
  verification = { ctx: workspace, id: receipt.id, key: line.key };
  $('verify-line').textContent = line.descripcion || 'Línea sin descripción';
  $('verify-person').value = localStorage.getItem(REVIEWER) || '';
  $('verify-error').hidden = true;
  $('verify-dialog').showModal();
}

async function confirmVerification(withName) {
  const { ctx, id, key } = verification || {};
  if (!ctx || workspace !== ctx || $('verify-with-name').disabled) return;
  const receipt = ctx.remote.get(id);
  const line = receipt?.lines.find(line => line.key === key);
  if (!line || line.verificado) {
    $('verify-dialog').close();
    return;
  }
  const input = $('verify-person');
  const error = $('verify-error');
  const por = withName ? text(input.value).slice(0, 120) : '';
  if (withName && !por) {
    error.textContent = 'Escribí un nombre o elegí «Solo fecha».';
    error.hidden = false;
    input.focus();
    return;
  }
  const controls = [input, $('verify-with-name'), $('verify-date-only'), $('cancel-verify')];
  for (const control of controls) control.disabled = true;
  error.hidden = true;
  try {
    const mark = { fecha: Date.now(), ...(por && { por }) };
    const saved = await changeLine(receipt, key, { verificado: mark }, ctx);
    if (workspace !== ctx) return;
    if (!saved) {
      error.textContent = 'No se pudo guardar la verificación. Volvé a intentar.';
      error.hidden = false;
      return;
    }
    if (por) {
      try { localStorage.setItem(REVIEWER, por); }
      catch { notice('Se guardó la verificación. El navegador no pudo recordar el nombre.', true); }
    }
    $('verify-dialog').close();
  } finally {
    for (const control of controls) control.disabled = false;
  }
}

function editReceipt(event) {
  const control = event.target;
  const id = control.closest('article')?.dataset.id;
  const receipt = workspace?.remote.get(id);
  if (!receipt) return;
  const key = control.closest('tr')?.dataset.line;
  const line = receipt.lines.find(line => line.key === key);
  if (control.hasAttribute('data-verify') && line) {
    const checked = control.checked;
    control.checked = !!line.verificado;
    if (checked && !line.verificado) openVerification(receipt, line);
    if (!checked && line.verificado) changeLine(receipt, key, { verificado: null });
    return;
  }
  const field = control.dataset.field;
  if (!field) return;
  const value = control.type === 'checkbox' ? control.checked : text(control.value);
  if (line && line[field] !== value) changeLine(receipt, key, { [field]: value });
  else if (!key && receipt[field] !== value) changeReceipt({ [id + '/' + field]: value });
}

async function doLineAction(event) {
  const button = event.target.closest('button[data-action]');
  const id = button?.closest('article')?.dataset.id;
  const receipt = workspace?.remote.get(id);
  if (!receipt) return;
  const key = button.closest('tr[data-line]')?.dataset.line;
  const action = button.dataset.action;
  if (action === 'remove-receipt') {
    if (!busy) askRemoveReceipt(receipt, button.closest('article'));
    return;
  }
  if (action === 'add') {
    const newKey = newLineKey();
    const order = receipt.lines.length ? receipt.lines.at(-1).order + 1 : 0;
    await changeReceipt({ [linePath(receipt, newKey)]: publicLine({ confianza: 'media' }, order) });
    cards.get(id)?.card.querySelector('[data-line="' + newKey + '"] textarea')?.focus();
    return;
  }
  const index = receipt.lines.findIndex(line => line.key === key);
  if (index < 0) return;
  if (action === 'return-date') {
    const details = button.closest('tr').querySelector('[data-return-details]');
    details.open = true;
    details.querySelector('[data-field="fecha_devolucion"]').focus();
    return;
  }
  if (action === 'remove') {
    if (window.confirm('¿Quitar esta línea de la transcripción?')) changeLine(receipt, key, { deleted: true });
    return;
  }
  if (action === 'choose') { openChooser(receipt, key); return; }
  const result = resultFor(receipt, index);
  if (action === 'excel-movement') { openMovement(receipt, result); return; }
  if (action === 'excel-field') { openExcelField(receipt, result, button.dataset.excelField); return; }
  if (action === 'recompare' && workspace.excel) {
    const line = result.line;
    const fields = {};
    if (!line.comparar) fields.comparar = true;
    if (line.decision) fields.decision = '';
    if (line.elegida) fields.elegida = null;
    if (line.correccion_excel) fields.correccion_excel = null;
    if (Object.keys(line.ajustes_excel || {}).length) fields.ajustes_excel = null;
    if (line.estado !== 'pending') fields.estado = 'pending';
    if (Object.keys(fields).length) await changeLine(receipt, key, fields);
    else { recalculate(workspace); render(); }
    return;
  }
  if (action === 'confirm' && result.chosen) changeLine(receipt, key, { decision: 'confirmed', elegida: result.chosen, estado: 'ok' });
  if (action === 'absent') changeLine(receipt, key, { decision: 'absent', elegida: null, estado: 'missing', correccion_excel: null, ajustes_excel: null });
  if (action === 'undo') changeLine(receipt, key, { decision: '', elegida: null, estado: 'pending', correccion_excel: null, ajustes_excel: null });
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
  $('suggestions').innerHTML = result.candidates?.length
    ? '<h3>Candidatos sugeridos</h3>' + result.candidates.map(({ row, similarity, assigned }) =>
      formatRow(row, result, { canChoose: !assigned, similarity, assigned })).join('') : '';
  const words = normalize($('manual-search').value).split(' ').filter(Boolean);
  const rows = workspace.excel.items.filter(row => words.every(word => row.search.includes(word)) && rowAvailable(row, result));
  const pageSize = 12;
  const pages = Math.max(1, Math.ceil(rows.length / pageSize));
  chooser.page = Math.max(0, Math.min(chooser.page, pages - 1));
  $('manual-results').innerHTML = rows.slice(chooser.page * pageSize, (chooser.page + 1) * pageSize)
    .map(row => formatRow(row, result, { canChoose: true })).join('')
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
  changeLine(receipt, result.line.key, { decision: 'manual', elegida: chosenRow(row), estado: 'ok',
    ...(result.chosen?.fila !== number && { correccion_excel: null, ajustes_excel: null }) });
}

function useWorkbook(ctx, workbook) {
  ctx.workbook = workbook;
  ctx.excel = workbook ? indexWorkbook(workbook.rows, workbook.base) : null;
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
  if ((ctx.workbook?.revision === meta.revision && (!meta.base || ctx.workbook.base)) || sync.readingExcel === meta.revision) return;
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
    if (meta.base) {
      const { blocks, count, ...layout } = meta.base;
      if (!Number.isSafeInteger(blocks) || blocks < 0) throw new Error();
      const records = [];
      for (let i = 0; i < blocks; i++) {
        const snapshot = await ref.child('base').child('b' + i).once('value');
        if (ctx.sync !== sync || sync.excelMeta?.revision !== revision || ctx.excelUpload) return;
        const block = snapshot.val();
        if (!block) throw new Error();
        records.push(...Object.values(block));
      }
      if (records.length !== count) throw new Error();
      records.sort((a, b) => a.fila - b.fila);
      workbook.base = { ...layout, records };
    }
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
    const { records, ...layout } = upload.base || { records: [] };
    const baseBlocks = excelBlocks(records, true);
    for (const [index, block] of baseBlocks.entries()) {
      if (!updatesAllowed(ctx) || ctx.sync !== sync) return;
      await ref.child('base').child('b' + index).set(block);
    }
    if (!updatesAllowed(ctx) || ctx.sync !== sync) return;
    const meta = { revision: upload.revision, name: upload.name, loadedAt: upload.loadedAt, count: upload.rows.length, blocks: blocks.length };
    if (upload.base) meta.base = { ...layout, blocks: baseBlocks.length, count: records.length };
    let previous;
    const outcome = await sync.excelRef.child('actual').transaction(current => {
      previous = current;
      if (upload.completesRevision && current?.revision !== upload.completesRevision) return;
      return meta;
    }, undefined, false);
    if (ctx.sync !== sync) return;
    if (!outcome.committed) {
      ctx.excelUpload = null;
      await storeRecords([], [[ctx.scope, 'excel-queued']]);
      await receiveExcel(ctx, sync, outcome.snapshot.val());
      discardExcelRevision(sync, upload, ctx.workbook?.revision).catch(() => {});
      if (workspace === ctx) excelFeedback('El Excel de la sala cambió mientras se agregaban las columnas.', true);
      return;
    }
    const { completesRevision, ...saved } = upload;
    const workbook = { ...saved, id: 'excel' };
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

async function loadWorkbook(file, complete = false) {
  const ctx = workspace;
  if (!ctx || !file || loadingExcel || ctx.sync?.writingExcel) return;
  const current = ctx.workbook;
  if (complete && (!current || current.base || ctx.sync?.readingExcel)) return;
  if (!/\.xlsx$/i.test(file.name)) return excelFeedback('Elegí un archivo .xlsx.', true);
  loadingExcel = true;
  updateControls();
  excelFeedback('');
  try {
    const xlsx = await dependency('XLSX', SHEET_JS);
    const { rows, base } = readWorkbook(xlsx, await file.arrayBuffer(), { includeSource: true });
    if (workspace !== ctx) return;
    if (complete && (ctx.workbook !== current || !sameWorkbookRows(current.rows, rows))) {
      return excelFeedback('Las filas del archivo deben coincidir con las de la sala. Elegí el Excel original de esta revisión.', true);
    }
    const shared = ctx.scope !== 'local';
    const workbook = { id: shared ? 'excel-queued' : 'excel', scope: ctx.scope, name: complete ? current.name : file.name,
      rows: complete ? current.rows : shared ? excelBlocks(rows).flatMap(Object.values) : rows, base,
      loadedAt: complete ? current.loadedAt : Date.now(), ...(shared && { revision: newId() }),
      ...(complete && shared && !ctx.excelUpload && { completesRevision: current.revision }) };
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
  const ctx = workspace;
  if (!ctx) return;
  const { workbook } = ctx;
  if (format === 'xlsx' && workbook && !workbook.base) {
    if ($('choose-excel').disabled || ctx.sync?.readingExcel) {
      return notice('Esperá a que termine la carga del Excel.');
    }
    $('excel-complete-file').click();
    return;
  }
  const comparisons = ctx.comparison.results;
  const rows = exportRows(receiptList(), comparisons);
  if (rows.length <= 1 && (format !== 'xlsx' || !workbook)) return;
  const name = `control-remitos-${new Intl.DateTimeFormat('sv-SE', { timeZone: 'America/Argentina/Buenos_Aires' }).format(new Date())}`;
  try {
    if (format === 'xlsx') {
      const xlsx = await dependency('XLSX', SHEET_JS);
      if (workspace !== ctx) return;
      const bytes = buildRemitosExcel(xlsx, rows, workbook, comparisons);
      downloadFile(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${name}.xlsx`);
    } else {
      // El BOM conserva acentos en Excel. El escape evita que un texto transcripto se interprete como fórmula.
      const csvCell = value => {
        const original = String(value ?? '');
        const safe = /^[\s\uFEFF]*[=+\-@]|^[\t\r\n]/.test(original) ? `'${original}` : original;
        return `"${safe.replace(/"/g, '""')}"`;
      };
      const csv = rows.map(row => row.map(csvCell).join(';')).join('\r\n');
      downloadFile(new Blob(['\uFEFF', csv], { type: 'text/csv;charset=utf-8' }), `${name}.csv`);
    }
  } catch (error) { notice(`No se pudo exportar. ${error.message}`, true); }
}

function downloadFile(file, name) {
  const url = URL.createObjectURL(file);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
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
$('excel-complete-file').onchange = event => { loadWorkbook(event.target.files[0], true); event.target.value = ''; };
$('choose-images').onclick = () => choosePhotos();
$('images').onchange = event => {
  const target = photoUpload;
  photoUpload = null;
  if (!target || target.ctx === workspace) addPhotos([...event.target.files], target?.folder);
  event.target.value = '';
};
$('images').addEventListener('cancel', () => { photoUpload = null; });
$('upload-folder').onchange = event => { if (workspace) selectUploadFolder(workspace, event.target.value); };
$('new-folder').onclick = createFolder;
$('process').onclick = processPhotos;
$('select-all-photos').onclick = toggleAllPhotos;
$('remove-all-photos').onclick = () => askRemovePhoto(null, true);
$('queue').addEventListener('change', selectQueuedPhoto);
$('queue').addEventListener('click', event => {
  const button = event.target.closest('button');
  if (button?.hasAttribute('data-folder-edit')) askFolderName(event);
  else if (button?.hasAttribute('data-folder-upload')) choosePhotos(button.dataset.folderUpload);
  else if (button?.dataset.queueRemove) askRemovePhoto(button.dataset.queueRemove);
  else if (button?.dataset.queueMove) askMovePhoto(button.dataset.queueMove);
  else if (button?.dataset.queuePreview) previewQueuedPhoto(event);
});
enablePhotoPan($('preview-figure'));
$('preview-zoom').oninput = event => zoomPhoto($('preview-figure'), event.target.value);
$('photo-preview').addEventListener('close', () => {
  previewTicket++;
  const image = $('preview-image');
  image.onload = image.onerror = null;
  image.removeAttribute('src');
  if (previewUrl?.startsWith('blob:')) URL.revokeObjectURL(previewUrl);
  previewUrl = null;
});
$('confirm-release-photo').onclick = () => confirmRemovePhoto(true);
$('confirm-remove-photo').onclick = () => confirmRemovePhoto(false);
$('remove-photo-dialog').addEventListener('cancel', event => {
  if ($('cancel-remove-photo').disabled) event.preventDefault();
});
$('confirm-remove-receipt').onclick = confirmRemoveReceipt;
$('remove-receipt-dialog').addEventListener('cancel', event => {
  if ($('confirm-remove-receipt').disabled) event.preventDefault();
});
$('move-photo-form').onsubmit = event => {
  if (event.submitter?.formMethod !== 'dialog') movePhoto(event);
};
$('folder-name-form').onsubmit = saveFolderName;
$('cancel-folder-name').onclick = () => $('folder-name-dialog').close();
$('folder-name-dialog').addEventListener('cancel', event => {
  if ($('save-folder-name').disabled) event.preventDefault();
});
$('verify-form').onsubmit = event => { event.preventDefault(); confirmVerification(true); };
$('verify-date-only').onclick = () => confirmVerification(false);
$('cancel-verify').onclick = () => $('verify-dialog').close();
$('verify-dialog').addEventListener('cancel', event => {
  if ($('verify-with-name').disabled) event.preventDefault();
});
$('verify-dialog').addEventListener('close', () => { if (!$('verify-dialog').open) verification = null; });
$('movement-value').onchange = updateMovementDate;
$('movement-form').onsubmit = event => { event.preventDefault(); saveMovement(); };
$('restore-movement').onclick = () => saveMovement(true);
$('cancel-movement').onclick = () => $('movement-dialog').close();
$('movement-dialog').addEventListener('cancel', event => {
  if ($('save-movement').disabled) event.preventDefault();
});
$('movement-dialog').addEventListener('close', () => { if (!$('movement-dialog').open) movementEdit = null; });
$('excel-edit-copy').onclick = () => {
  $('excel-edit-value').value = excelEditLabel(excelFieldEdit.field, excelFieldEdit.copied);
  $('excel-edit-value').focus();
};
$('excel-edit-form').onsubmit = event => { event.preventDefault(); saveExcelField(); };
$('excel-edit-restore').onclick = () => saveExcelField(true);
$('excel-edit-cancel').onclick = () => $('excel-edit-dialog').close();
$('excel-edit-dialog').addEventListener('cancel', event => {
  if ($('excel-edit-save').disabled) event.preventDefault();
});
$('excel-edit-dialog').addEventListener('close', () => { if (!$('excel-edit-dialog').open) excelFieldEdit = null; });
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
$('receipts').addEventListener('click', askFolderName);
$('receipts').addEventListener('click', toggleFolderDetail);
$('receipts').addEventListener('focusout', event => { if (event.target.matches('[data-field]')) setTimeout(render, 0); });
$('receipts').addEventListener('input', event => {
  if (event.target.type === 'range') zoomPhoto(event.target.closest('.photo-block').querySelector('figure'), event.target.value);
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
