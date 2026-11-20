import { excelColumns, excelDay, productCategories, productLabel, suggestProductCategory } from './control-remitos-core.js?v=excel-rubros';

const workbookNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const featureNamespace = 'http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag';
const featureFile = 'xl/featurePropertyBag/featurePropertyBag.xml';
const utf8 = new TextEncoder();
const dateFormat = 'dd/mm/yy';

function excelDate(value, date1904) {
  const day = excelDay(value);
  if (day == null) return value;
  return { t: 'n', v: day + (date1904 ? 24107 : 25569), z: dateFormat };
}

function baseCopy(workbook) {
  if (workbook?.base) return workbook.base;
  const columns = Object.entries(excelColumns);
  const records = [
    { fila: 1, cells: { c0: { v: workbook ? 'BASE DE DATOS · Volvé a cargar el Excel para copiar todas sus columnas' : 'BASE DE DATOS' } } },
    { fila: 2, cells: Object.fromEntries(columns.map(([, label], column) => ['c' + column, { v: label }])) }
  ];
  for (const row of workbook?.rows || []) records.push({ fila: row.fila,
    cells: Object.fromEntries(columns.map(([field], column) => ['c' + column, { v: row[field] || '' }])) });
  return { dateColumn: 0, colCount: columns.length, records, merges: [{ start: 0, end: columns.length - 1 }] };
}

function buildBaseSheet(xlsx, workbook, base, comparisons) {
  const column = index => index === base.dateColumn ? 1 : index + (index < base.dateColumn ? 3 : 2);
  const put = (sheet, r, c, value) => { sheet[xlsx.utils.encode_cell({ r, c })] = value; };
  const lastRow = base.records.reduce((row, record) => Math.max(row, record.fila), 2);
  const width = base.colCount + 2;
  const sheet = { '!ref': xlsx.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow - 1, c: width - 1 } }) };
  const matched = new Map();
  for (const result of comparisons.values()) {
    if (result.state === 'ok' && result.chosen) matched.set(result.chosen.fila, result);
  }
  const sourceRows = new Map((workbook?.rows || []).map(row => [row.fila, row]));
  for (const source of base.records) {
    const r = source.fila - 1;
    for (const [key, cell] of Object.entries(source.cells)) {
      const c = column(Number(key.slice(1)));
      const v = cell.v;
      const t = typeof v === 'number' ? 'n' : typeof v === 'boolean' ? 'b' : 's';
      const value = c === 1 && r >= 2 && typeof v === 'string' ? excelDate(v, base.date1904) : v;
      put(sheet, r, c, typeof value === 'object' ? value : { t, v: value, ...(cell.z && { z: cell.z }) });
    }
    const sourceRow = sourceRows.get(source.fila);
    if (!sourceRow) continue;
    const result = matched.get(source.fila);
    const rubro = result?.line.rubro && result.line.rubro !== 'sin_clasificar'
      ? productLabel(result.line) : productCategories[suggestProductCategory(sourceRow.herramienta, sourceRow.detalle)];
    put(sheet, r, 0, { t: 'b', v: !!result?.line.verificado });
    put(sheet, r, 2, { t: 's', v: rubro });
  }
  put(sheet, 1, 0, { t: 's', v: 'Verificado con remito' });
  put(sheet, 1, 2, { t: 's', v: 'Rubro sugerido' });
  const fallbackWidths = [13, 14, 34, 17, 15, 12, 22, 20, 17, 16, 18, 14, 32, 13, 20, 17, 30];
  const cols = [{ wch: 24 }, { wch: 13 }, { wch: 31 }];
  for (let index = 0; index < base.colCount; index++) {
    if (index !== base.dateColumn) cols[column(index)] = { wch: Math.max(12, base.widths?.[index] || fallbackWidths[index] || 14) };
  }
  sheet['!cols'] = cols;
  const merges = (base.merges || []).filter(({ start, end }) => start >= 0 && end > start && end < base.colCount
    && (base.dateColumn === 0 || base.dateColumn < start || base.dateColumn > end));
  sheet['!merges'] = merges.map(({ start, end }) => ({ s: { r: 0, c: column(start) }, e: { r: 0, c: column(end) } }));
  for (let r = 0; r < 2; r++) for (let c = 0; c < width; c++) {
    const address = xlsx.utils.encode_cell({ r, c });
    if (!sheet[address]) sheet[address] = { t: 's', v: '' };
  }
  sheet['!autofilter'] = { ref: xlsx.utils.encode_range({ s: { r: 1, c: 0 }, e: { r: lastRow - 1, c: width - 1 } }) };
  return { sheet, secondBand: sheet['!merges'][1]?.s.c ?? width };
}

function node(parent, name, attributes = {}, namespace = parent.namespaceURI) {
  const element = parent.ownerDocument.createElementNS(namespace, name);
  for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
  parent.append(element);
  return element;
}

function indexedNode(parent, name, attributes) {
  const index = parent.childElementCount;
  node(parent, name, attributes);
  parent.setAttribute('count', parent.childElementCount);
  return index;
}

function styleWorkbook(xlsx, archive, secondBand) {
  const serializer = new XMLSerializer();
  const decode = new TextDecoder();
  const parser = new DOMParser();
  const entry = path => xlsx.CFB.find(archive, '/' + path);
  const read = path => parser.parseFromString(decode.decode(entry(path).content), 'application/xml');
  const write = (path, doc) => {
    const part = entry(path);
    part.content = utf8.encode(serializer.serializeToString(doc));
    part.size = part.content.byteLength;
  };
  const styles = read('xl/styles.xml');
  const fonts = styles.getElementsByTagName('fonts')[0], fills = styles.getElementsByTagName('fills')[0];
  const xfs = styles.getElementsByTagName('cellXfs')[0];
  const baseStyles = [...xfs.children];
  const font = (rgb, size, bold = false) => {
    const index = indexedNode(fonts, 'font');
    const entry = fonts.lastElementChild;
    if (bold) node(entry, 'b');
    node(entry, 'sz', { val: size });
    node(entry, 'color', { rgb: 'FF' + rgb });
    node(entry, 'name', { val: 'Arial' });
    return index;
  };
  const fill = rgb => {
    const index = indexedNode(fills, 'fill');
    const pattern = node(fills.lastElementChild, 'patternFill', { patternType: 'solid' });
    node(pattern, 'fgColor', { rgb: 'FF' + rgb });
    node(pattern, 'bgColor', { indexed: 64 });
    return index;
  };
  const fontsByType = { title: font('214336', 14, true), header: font('FFFFFF', 10, true),
    lightTitle: font('1D3550', 12, true), darkTitle: font('FFFFFF', 12, true), body: font('203129', 10) };
  const fillsByType = { header: fill('245D4B'), titleBase: fill('FFF2CB'), titleBlue: fill('4472C4'),
    stripe: fill('F2F7F4'), verified: fill('E1F4E9') };
  const styleIds = new Map();
  const style = (numFmtId, fontId, fillId, alignment, checkbox) => {
    const key = [numFmtId, fontId, fillId, alignment, checkbox].join(':');
    if (styleIds.has(key)) return styleIds.get(key);
    const index = indexedNode(xfs, 'xf', { numFmtId, fontId, fillId, borderId: 0, xfId: 0,
      applyFont: 1, applyFill: 1, applyNumberFormat: 1, applyAlignment: 1 });
    const xf = xfs.lastElementChild;
    node(xf, 'alignment', { vertical: 'center', ...(['header', 'center'].includes(alignment) && { horizontal: 'center' }),
      ...(['header', 'wrap'].includes(alignment) && { wrapText: 1 }) });
    if (checkbox) {
      const ext = node(node(xf, 'extLst'), 'ext', { uri: '{C7286773-470A-42A8-94C5-96B5CB345126}' });
      node(ext, 'xfpb:xfComplement', { i: 0 }, featureNamespace);
    }
    styleIds.set(key, index);
    return index;
  };
  let hasCheckbox = false;
  for (const sheetNumber of [1, 2]) {
    const path = `xl/worksheets/sheet${sheetNumber}.xml`;
    const sheet = read(path);
    const view = sheet.getElementsByTagName('sheetView')[0];
    view.setAttribute('showGridLines', '0');
    const pane = node(view, 'pane', { ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' }, workbookNamespace);
    view.prepend(pane);
    const widths = new Map([...sheet.getElementsByTagName('col')].map(col =>
      [Number(col.getAttribute('min')) - 1, Number(col.getAttribute('width'))]));
    for (const row of sheet.getElementsByTagName('row')) {
      const r = Number(row.getAttribute('r'));
      let lines = 1;
      for (const cell of row.children) {
        const c = xlsx.utils.decode_col(cell.getAttribute('r').replace(/\d+$/, ''));
        const numFmt = baseStyles[Number(cell.getAttribute('s'))]?.getAttribute('numFmtId') || 0;
        const checkbox = sheetNumber === 2 && r > 2 && c === 0 && cell.getAttribute('t') === 'b';
        hasCheckbox ||= checkbox;
        let fontId = fontsByType.body, fillId = r % 2 === 0 ? fillsByType.stripe : 0, alignment = checkbox ? 'center' : 'body';
        if (r === 1) {
          fontId = sheetNumber === 1 ? fontsByType.title : c >= secondBand ? fontsByType.darkTitle : fontsByType.lightTitle;
          fillId = sheetNumber === 1 ? 0 : c >= secondBand ? fillsByType.titleBlue : fillsByType.titleBase;
        } else if (r === 2) {
          fontId = fontsByType.header;
          fillId = fillsByType.header;
          alignment = 'header';
        } else if (checkbox && cell.textContent === '1') fillId = fillsByType.verified;
        else if (cell.getAttribute('t') === 'str') {
          const width = Math.max(1, Math.floor(widths.get(c) || 14) - 2);
          const wrapped = cell.textContent.split('\n').reduce((count, line) => count + Math.max(1, Math.ceil(line.length / width)), 0);
          if (wrapped > 1) {
            alignment = 'wrap';
            lines = Math.max(lines, wrapped);
          }
        }
        cell.setAttribute('s', style(numFmt, fontId, fillId, alignment, checkbox));
      }
      row.setAttribute('ht', r === 1 ? 26 : r === 2 ? 30 : Math.min(408, 17 * lines));
      row.setAttribute('customHeight', '1');
    }
    write(path, sheet);
  }
  write('xl/styles.xml', styles);
  if (hasCheckbox) {
    // Extensión de Excel para mostrar booleanos como casillas dentro de la celda.
    const bag = `<FeaturePropertyBags xmlns="${featureNamespace}"><bag type="Checkbox"/><bag type="XFControls"><bagId k="CellControl">0</bagId></bag>
      <bag type="XFComplement"><bagId k="XFControls">1</bagId></bag><bag type="XFComplements" extRef="XFComplementsMapperExtRef">
      <a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag></FeaturePropertyBags>`;
    xlsx.CFB.utils.cfb_add(archive, featureFile, utf8.encode(bag));
    const types = read('[Content_Types].xml');
    node(types.documentElement, 'Override', { PartName: '/' + featureFile, ContentType: 'application/vnd.ms-excel.featurepropertybag+xml' });
    write('[Content_Types].xml', types);
    const relationships = read('xl/_rels/workbook.xml.rels');
    node(relationships.documentElement, 'Relationship', { Id: 'rIdCheckboxes', Target: 'featurePropertyBag/featurePropertyBag.xml',
      Type: 'http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag' });
    write('xl/_rels/workbook.xml.rels', relationships);
  }
}

export function buildRemitosExcel(xlsx, rows, workbook, comparisons) {
  const base = baseCopy(workbook);
  const book = xlsx.utils.book_new();
  if (base.date1904) book.Workbook = { WBProps: { date1904: true } };
  const sheet = xlsx.utils.aoa_to_sheet([['Control de remitos'], ...rows.map((row, index) =>
    index === 0 ? row : row.map((value, column) => column === 0 ? { t: 's', v: value, z: '@' }
      : column === 3 ? excelDate(value, base.date1904) : value))]);
  sheet['!cols'] = [14, 24, 24, 15, 48, 14, 24, 14, 23, 31].map(wch => ({ wch }));
  sheet['!autofilter'] = { ref: `A2:J${rows.length + 1}` };
  xlsx.utils.book_append_sheet(book, sheet, 'Control de remitos');
  const copy = buildBaseSheet(xlsx, workbook, base, comparisons);
  xlsx.utils.book_append_sheet(book, copy.sheet, 'BASE DE DATOS');
  const file = xlsx.write(book, { type: 'array', bookType: 'xlsx' });
  const archive = xlsx.CFB.read(new Uint8Array(file), { type: 'array' });
  styleWorkbook(xlsx, archive, copy.secondBand);
  return xlsx.CFB.write(archive, { type: 'array', fileType: 'zip', compression: true });
}
