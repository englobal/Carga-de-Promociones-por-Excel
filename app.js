/* TODO: acá va TODO el contenido que hoy está dentro del <script>...</script>
   de index_FINAL_AREA_CORRECTO.html, sin cambios. */
   const CFG = {
  SHEET_BY_TIPO: {
    NOMINAL: "Precio Fijo a Productos",
    PORCENTUAL: "Porcentaje a Productos",
    "DESCUENTO 2DA UNIDAD (%)": "Porcentaje a Productos"
  },
  LISTA_SHEET_NAME: "Lista de Productos",
  LISTA_LOCALES_SHEET_NAME: "Lista de Locales",
  EVENTOS_TARGET_SHEET: "Porcentaje a Productos",
  CODIGO_MARCA_SHEET: "CODIGO-MARCA",
  TIPO_PROMOCION_DEFAULT: "FS-1000",
  DEFAULT_TIPO_DOCUMENTO: "Boleta",
  DEBUG: true
};

const SOURCE_MODE = { NORMAL: 'normal', EVENTOS: 'eventos' };

// ===== AREA RESPONSABLE DESDE INPUT Y3 =====
let AREA_RESPONSABLE = '';

function detectAreaFromOrigenWB(wb){
  try{
    const sheet = wb.SheetNames.find(n => normalizeHeader(n)==='input');
    if(!sheet) return '';
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {header:1});
    return String(aoa?.[2]?.[24] || '').trim().toUpperCase();
  }catch(e){
    return '';
  }
}
// ===== END =====

const SUPPORTED_TIPOS = ["NOMINAL", "PORCENTUAL", "PACK 2X1", "DESCUENTO 2DA UNIDAD (%)"];

const CLUB_CFG = {
  GRUPO_PROMOCIONES: "CLUB_NECESIDAD",
  INSCRIPTO: 1,
  DEFAULT_CONVENIOS: [
    "20003","20078","7830","7774","7587","6090","3801","3566","2308","882","880","879","878","877","5877"
  ]
};
CLUB_CFG.DEFAULT_CONVENIOS_CSV = CLUB_CFG.DEFAULT_CONVENIOS.join(',');

const LIST_RULES = {
  MAX_LEN: 15,
  NORMAL_PREFIX: 'FS_',
  NORMAL_BASE_MAX: 12,
  NORMAL_SEQ_LEN: 3,
  EVENT_INITIALS_LEN: 2,
  EVENT_BRAND_MAX: 5
};

const FIELD = {
  EANS: ['2-EANS','EANS'],
  LISTA_REF: ['2-Lista de Productos','Lista de Productos','2-ID Lista Productos','ID Lista Productos','2-ID Lista Producto','ID Lista Producto','2-Id Lista Productos','Id Lista Productos','2-Id Lista Producto','Id Lista Producto'],
  POR_UNIDAD: ['2-Por Unidad','Por Unidad'],
  FECHA_INICIO: ['2-Fecha Inicio (aaaaMMdd)','Fecha Inicio (aaaaMMdd)','Fecha Inicio'],
  FECHA_FIN: ['2-Fecha Fin (aaaaMMdd)','Fecha Fin (aaaaMMdd)','Fecha Fin'],
  TIPO_PROMO: ['2-Tipo de Promoción','Tipo de Promoción','Tipo de promoción'],
  NOMBRE_GENERAL: ['2-Nombre General','Nombre General'],
  TIPO_DOCUMENTO: ['2-Tipo Documento','Tipo Documento','2-Tipo de Documento','Tipo de Documento'],
  NOMBRE: ['2-Nombre','Nombre'],
  DESCRIPCION: ['2-Descripción','Descripción','2-Descripcion','Descripcion'],
  PRECIO_FINAL: ['2-Precio Final','Precio Final'],
  BP_PORCENTAJE: ['2-Porcentaje','Porcentaje','BP Porcentaje','2-BP Porcentaje'],
  OP_CANTIDAD: ['Operador Cantidad (1.Cada,2.Mayor Igual,3.Mayor,4.Menor Igual,5.Menor)','Operador Cantidad','BI Operador Cantidad'],
  CANTIDAD: ['Cantidad','BJ Cantidad'],
  CANTIDAD_BENEFICIO: ['Cantidad Beneficio','BK Cantidad Beneficio','Cantidad de Beneficio'],
  PROD_ESTRATEGIA: ['Productos / Estrategia (1-Menor,2-Mayor)','Productos Estrategia (1-Menor,2-Mayor)','Estrategia (1-Menor,2-Mayor)','Productos','Estrategia','BN'],
  LOCALES_POSITIVOS: ['Locales Positivos'],
  LOCALES_NEGATIVOS: ['Locales / Lista Locales Negativos','Locales Lista Locales Negativos','Lista Locales Negativos','AZ'],
  LISTA_LOCALES_POSITIVOS: ['Lista Locales Positivos'],
  LP_NOMBRE: ['nombre de la lista','nombre lista','nombre de lista','2-nombre de la lista','2-nombre lista'],
  LP_DESDE: ['desde (aaaammdd)','desde (aaaaMMdd)','desde','fecha desde','2-desde (aaaammdd)','2-desde'],
  LP_HASTA: ['hasta (aaaammdd)','hasta (aaaaMMdd)','hasta','fecha hasta','2-hasta (aaaammdd)','2-hasta'],
  LP_ITEM: ['item','items','2-item','2-items'],
  COMPITE: ['Compite (0-No Compite,1-Producto,2-Unidades,3-Promociones)','Compite'],
  GRUPO_PROMOCIONES: ['Grupo de Promociones (Grupo1|Grupo2)','Grupo de Promociones','Grupo1|Grupo2'],
  INSCRIPTO: ['Inscripto'],
  CONVENIO_SELECCIONADO: ['Convenio Seleccionado (separado por ,)','Convenio Seleccionado'],
  CODIGO_CONVENIO: ['Código Convenio (separado por ,)','Codigo Convenio (separado por ,)','Código Convenio','Codigo Convenio']
};

const ORIGEN_COLS = {
  NUMERO: ['n°','nº','n','numero','número'],
  CODIGO_PRODUCTO: ['código producto','codigo producto','codigoproducto'],
  PVP_OFERTA_PACK: ['pvp oferta pack','pvpofertapack','precio oferta pack'],
  DESC_PCT: ['descuento porcentual','descuentoporcentual','% descuento','descuento %','porcentaje','porcentaje descuento','bp porcentaje','bp porcentual'],
  INICIO: ['f. inicio','f inicio','fecha inicio'],
  TERMINO: ['f. término','f término','f termino','fecha término','fecha termino','f. termino'],
  TIPO_DESC: ['tipo de descuento','tipodedescuento'],
  UNIDADES_PACK: ['# unidades pack','unidades pack','n° unidades pack','numero unidades pack','cantidad pack'],
  COBERTURA_LOCALES: ['cobertura locales','cobertura de locales']
};

const EVENT_COLS = {
  RC: ['rc'],
  NUMERO: ['n°cam','ncam','n° cam'],
  LOCAL: ['local','loal'],
  INICIO: ['feha de inicio evento','fecha de inicio evento'],
  FIN: ['feha termino evento','fecha termino evento','fecha término evento'],
  MARCA: ['marca'],
  DESCUENTO: ['descuento']
};

const fileOrigen = document.getElementById('fileOrigen');
const fileBase = document.getElementById('fileBase');
const modoSelect = document.getElementById('modoSelect');
const nombreGeneralInput = document.getElementById('nombreGeneral');
const chkNombreGeneralDefault = document.getElementById('chkNombreGeneralDefault');
const chkListas = document.getElementById('chkListas');
const chkClub = document.getElementById('chkClub');
const chkClubConveniosEditar = document.getElementById('chkClubConveniosEditar');
const clubConveniosInput = document.getElementById('clubConveniosInput');
const umbralEans = document.getElementById('umbralEans');
const nombreListaBase = document.getElementById('nombreListaBase');
const btnProcesar = document.getElementById('btnProcesar');
const btnPreview = document.getElementById('btnPreview'); // 👈 MOVER AQUÍ

const optNormal = document.getElementById('optNormal');
const optEventos = document.getElementById('optEventos');
const eventUserSelect = document.getElementById('eventUserSelect');
const eventInitials = document.getElementById('eventInitials');

const normalModoField = document.getElementById('normalModoField');
const eventUserField = document.getElementById('eventUserField');
const initialsField = document.getElementById('initialsField');
const normalListToggle = document.getElementById('normalListToggle');
const clubToggle = document.getElementById('clubToggle');
const clubInfoField = document.getElementById('clubInfoField');
const normalListNameField = document.getElementById('normalListNameField');
const eventsInfoField = document.getElementById('eventsInfoField');
const subOrigen = document.getElementById('subOrigen');
const umbralLabel = document.getElementById('umbralLabel');

const pillModo = document.getElementById('pillModo');
const pillListas = document.getElementById('pillListas');
const pillArea = document.getElementById('pillArea');

const statusText = document.getElementById('statusText');
const statusHint = document.getElementById('statusHint');

const debugDetails = document.getElementById('debugDetails');
const debugLog = document.getElementById('debugLog');
const debugMeta = document.getElementById('debugMeta');
const debugSummary = document.getElementById('debugSummary');

const summaryWrap = document.getElementById('summaryWrap');
const summaryWhen = document.getElementById('summaryWhen');
const summaryOutName = document.getElementById('summaryOutName');
const summaryTipos = document.getElementById('summaryTipos');
const summaryPromos = document.getElementById('summaryPromos');
const summaryListas = document.getElementById('summaryListas');
const summaryTable = document.getElementById('summaryTable');

const st1 = document.getElementById('st1');
const st2 = document.getElementById('st2');
const st3 = document.getElementById('st3');
const st4 = document.getElementById('st4');
const st5 = document.getElementById('st5');

const procOverlay = document.getElementById('procOverlay');
const procTitle = document.getElementById('procTitle');
const procOrigen = document.getElementById('procOrigen');
const procBase = document.getElementById('procBase');
const procPlan = document.getElementById('procPlan');
const procOut = document.getElementById('procOut');
const procFootMeta = document.getElementById('procFootMeta');

const regenModalOverlay = document.getElementById('regenModalOverlay');
const regenMsg = document.getElementById('regenMsg');
const regenMeta = document.getElementById('regenMeta');
const regenCloseBtn = document.getElementById('regenCloseBtn');
const regenCancelBtn = document.getElementById('regenCancelBtn');
const regenYesBtn = document.getElementById('regenYesBtn');

const listPreviewPill = document.getElementById('listPreviewPill');
const listPreviewText = document.getElementById('listPreviewText');

let currentSourceMode = SOURCE_MODE.NORMAL;
let wbOrigen = null;
let wbBase = null;
let cacheOrigenParsed = null;
let cacheEventosParsed = null;
let cacheCodigoMarca = null;
let origenFileName = '';
let lastOrigenMeta = null;
let lastBaseMeta = null;
let regenAction = null;

const LOGS = [];
function log(line){
  const ts = new Date().toISOString().slice(11, 19);
  const msg = `[${ts}] ${line}`;
  LOGS.push(msg);
  debugLog.textContent = LOGS.join('\n');
  debugSummary.textContent = `${LOGS.length} líneas`;
}
function clearLogs(){
  LOGS.length = 0;
  debugLog.textContent = '';
  debugSummary.textContent = '0 líneas';
  debugMeta.textContent = '';
}
function setStatus(msg, type='', hint=''){
  statusText.textContent = msg;
  statusHint.textContent = hint || '';
  const el = document.getElementById('status');
  el.className = 'status' + (type === 'ok' ? ' ok' : type === 'err' ? ' err' : type === 'warn' ? ' warn' : '');
  log(`STATUS(${type||'info'}): ${msg}`);
  if (type === 'err' || type === 'warn') debugDetails.open = true;
}
function pop(el){
  el.classList.remove('kpi-pop');
  void el.offsetWidth;
  el.classList.add('kpi-pop');
}

function normalizeHeader(h){
  return String(h ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toLowerCase();
}
function stripAccents(s){
  return String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function sanitizeEAN(v){
  let s = String(v ?? '').trim();
  if (!s) return '';
  s = s.replace(/^["']|["']$/g,'').trim();
  if (/^\d+\.0$/.test(s)) s = s.replace(/\.0$/,'');
  s = s.replace(/\s+/g,'');
  if (/^\d+$/.test(s)){
    const stripped = s.replace(/^0+(?=\d)/, '');
    return stripped === '' ? '0' : stripped;
  }
  return s;
}
function sanitizeLetters2(v){
  return stripAccents(v).toUpperCase().replace(/[^A-Z]/g,'').slice(0,2);
}
function sanitizeNameChunk(v, maxLen=24){
  let s = stripAccents(v).toUpperCase();
  s = s.replace(/[^A-Z0-9]+/g,'_').replace(/^_+|_+$/g,'').replace(/_+/g,'_');
  return s.slice(0, maxLen) || 'ITEM';
}
function escapeHtml(s){
  return String(s ?? '')
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'","&#039;");
}
function toNumberFlexible(v){
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).trim().replace(/\s+/g,'').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : null;
}
function toYYYYMMDD(value){
  if (value == null || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)){
    const y=value.getFullYear();
    const m=String(value.getMonth()+1).padStart(2,'0');
    const d=String(value.getDate()).padStart(2,'0');
    return `${y}${m}${d}`;
  }
  if (typeof value === 'number' && window.XLSX?.SSF){
    const d=XLSX.SSF.parse_date_code(value);
    if (d && d.y && d.m && d.d) return `${d.y}${String(d.m).padStart(2,'0')}${String(d.d).padStart(2,'0')}`;
  }
  const s=String(value).trim();
  let m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}${m[2]}${m[3]}`;
  m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}${String(m[2]).padStart(2,'0')}${String(m[1]).padStart(2,'0')}`;
  m=s.match(/^(\d{8})$/);
  if (m) return m[1];
  return '';
}
function normalizePercentNumber(v){
  if (v == null || v === '') return null;
  if (typeof v === 'number'){
    if (!isFinite(v)) return null;
    if (v > 0 && v < 1) return +(v * 100).toFixed(6);
    return +v.toFixed(6);
  }
  let s = String(v).trim();
  if (!s) return null;
  s = s.replace(/%/g,'').trim().replace(/\s+/g,'').replace(',', '.');
  const n = Number(s);
  if (!isFinite(n)) return null;
  if (n > 0 && n < 1) return +(n * 100).toFixed(6);
  return +n.toFixed(6);
}
function buildTimestampDDmmYYhhmm(d = new Date()){
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yy = String(d.getFullYear()).slice(-2);
  const hh = String(d.getHours()).padStart(2,'0');
  const mi = String(d.getMinutes()).padStart(2,'0');
  return `${dd}${mm}${yy}${hh}${mi}`;
}

function sanitizeToNameToken(v){
  let s = stripAccents(v).toUpperCase();
  s = s.replace(/[^A-Z0-9]+/g,'_').replace(/^_+|_+$/g,'').replace(/_+/g,'_');
  return s;
}
function hardCut(name, maxLen, reason=''){
  const s = String(name ?? '');
  if (s.length <= maxLen) return s;
  const cut = s.slice(0, maxLen);
  log(`WARN: truncado (${reason||'len'}): "${s}" -> "${cut}"`);
  return cut;
}
function toBase36Fixed(n, len){
  const s = n.toString(36).toUpperCase();
  return s.padStart(len, '0').slice(-len);
}

function parseConveniosInput(raw){
  const text = String(raw || '').trim();
  if (!text) return [];

  if (/[\n\r;|]/.test(text)) {
    throw new Error('Los convenios CLUB solo permiten separación por coma.');
  }

  return Array.from(new Set(
    text
      .split(',')
      .map(x => x.trim())
      .filter(Boolean)
  ));
}

function getClubConveniosData(){
  const arr = parseConveniosInput(clubConveniosInput.value);
  return {
    list: arr,
    csv: arr.join(',')
  };
}

function getClubConveniosDataSafe(){
  try{
    const data = getClubConveniosData();
    return { ok:true, ...data, error:'' };
  }catch(err){
    return { ok:false, list:[], csv:'', error: err?.message || String(err) };
  }
}

function applyClubConveniosEditMode(){
  const canEdit = chkClub.checked && chkClubConveniosEditar.checked;
  clubConveniosInput.disabled = !canEdit;
}

function buildNormalListName({ baseInput, promoIndex=1, usedListNames }){
  const prefix = LIST_RULES.NORMAL_PREFIX;
  const seqLen = LIST_RULES.NORMAL_SEQ_LEN;
  const idxStr = String(promoIndex || '').replace(/\D+/g,'') || '';

  let base = sanitizeToNameToken(baseInput).replace(/_/g,'');
  base = hardCut(base, LIST_RULES.NORMAL_BASE_MAX, 'normal-base');
  if (!base) base = 'LISTA';

  const reserved = 1 + seqLen;
  const maxStemLen = LIST_RULES.MAX_LEN - reserved;

  let candidateBase = base;

  if (idxStr){
    const maxBaseLenKeepingPromo = Math.max(1, maxStemLen - prefix.length - idxStr.length);
    candidateBase = hardCut(base, maxBaseLenKeepingPromo, 'normal-base-fit-promo');
  } else {
    const maxBaseLenNoPromo = Math.max(1, maxStemLen - prefix.length);
    candidateBase = hardCut(base, maxBaseLenNoPromo, 'normal-base-fit');
  }

  let stem = `${prefix}${candidateBase}${idxStr}`;
  if (stem.length > maxStemLen){
    stem = hardCut(stem, maxStemLen, 'normal-stem');
  }

  let counter = 0;
  while (counter < 36**seqLen){
    const seq = toBase36Fixed(counter, seqLen);
    const candidate = `${stem}_${seq}`;
    if (!usedListNames.has(candidate)){
      usedListNames.add(candidate);
      return candidate;
    }
    counter++;
  }
  throw new Error('No fue posible generar un nombre de lista único (normal).');
}

function getEventMMYYFromInicio(inicioValue){
  const ymd = toYYYYMMDD(inicioValue);
  if (!ymd || ymd.length !== 8) return { mm:'00', yy:'00' };
  const mm = ymd.slice(4,6);
  const yy = ymd.slice(2,4);
  return { mm, yy };
}

function normalizeEventBrandToken(marcaRaw){
  const raw = stripAccents(String(marcaRaw ?? '')).toUpperCase().trim();

  const hasVichy = raw.includes('VICHY');
  const hasLRP = raw.includes('LRP') || raw.includes('LA ROCHE') || raw.includes('LAROCHE') || raw.includes('ROCHE');

  if (hasVichy && hasLRP) return 'V.LRP';
  if (raw.startsWith('ISISPHARMA')) return 'ISISP';
  if (raw.startsWith('LA ROCHE') || raw.startsWith('LAROCHE')) return 'LRP';

  let tok = sanitizeToNameToken(raw);
  tok = tok.replace(/_/g,'');
  if (!tok) tok = 'MARCA';
  return hardCut(tok, LIST_RULES.EVENT_BRAND_MAX, 'event-brand');
}

function buildEventosListName({ initialsInput, marcaRaw, inicioEvento, usedListNames }){
  let ini = sanitizeLetters2(initialsInput);
  if (ini.length !== 2){
    ini = (ini + 'XX').slice(0,2);
    log(`WARN: Iniciales ajustadas a 2 letras: "${initialsInput}" -> "${ini}"`);
  }

  const brand = normalizeEventBrandToken(marcaRaw);
  const { mm, yy } = getEventMMYYFromInicio(inicioEvento);

  let candidate = `${ini}_${brand}_${mm}${yy}`;
  candidate = hardCut(candidate, LIST_RULES.MAX_LEN, 'event-final');

  if (!usedListNames.has(candidate)){
    usedListNames.add(candidate);
    return candidate;
  }

  for (let i=0; i<36; i++){
    const suffix = toBase36Fixed(i,1);
    let alt = candidate.slice(0, LIST_RULES.MAX_LEN-1) + suffix;
    if (!usedListNames.has(alt)){
      usedListNames.add(alt);
      log(`WARN: colisión nombre eventos, usando alternativo: ${alt}`);
      return alt;
    }
  }

  throw new Error(`No fue posible generar un nombre único para lista eventos: ${candidate}`);
}

function updateNormalListPreview(){
  const base = nombreListaBase?.value ?? '';
  if (!listPreviewPill || !listPreviewText) return;
  if (!base.trim()){
    listPreviewPill.style.display = 'none';
    return;
  }
  try{
    const tmpSet = new Set();
    const example = buildNormalListName({ baseInput: base, promoIndex: 35, usedListNames: tmpSet });
    listPreviewText.textContent = example;
    listPreviewPill.style.display = 'inline-flex';
  }catch(e){
    listPreviewText.textContent = '—';
    listPreviewPill.style.display = 'inline-flex';
  }
}

function applyNombreGeneralDefaultMode(forceDefault=false){
  const useDefault = chkNombreGeneralDefault.checked;

  if (useDefault || forceDefault){
    nombreGeneralInput.value = 'DESCUENTO FCV';
    nombreGeneralInput.disabled = true;
    nombreGeneralInput.placeholder = 'DESCUENTO FCV';
  } else {
    if (nombreGeneralInput.value.trim().toUpperCase() === 'DESCUENTO FCV'){
      nombreGeneralInput.value = '';
    }
    nombreGeneralInput.disabled = false;
    nombreGeneralInput.placeholder = 'Ej: Campaña Marzo 2026 - Higiene';
  }
  refreshUI();
}

function showProcessing({ title='Procesando…', origenName, baseName, plan, outName }){
  procTitle.textContent = title || 'Procesando…';
  procOrigen.textContent = origenName || '-';
  procBase.textContent = baseName || '-';
  procPlan.textContent = plan || '-';
  procOut.textContent = outName || '-';
  procFootMeta.textContent = `inicio=${new Date().toLocaleTimeString()} | plan=${plan || '-'}`;
  procOverlay.style.display = 'flex';
  procOverlay.setAttribute('aria-hidden','false');
}
function hideProcessing(){
  procTitle.textContent = 'Procesando…';
  procOverlay.style.display = 'none';
  procOverlay.setAttribute('aria-hidden','true');
}

async function readFileAsWorkbook(file, label){
  if (!window.XLSX) throw new Error('No se cargó XLSX. Verifica xlsx.full.min.js en la misma carpeta.');
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type:'array', cellDates:true });
  log(`${label} cargado. Hojas: ${wb.SheetNames.join(' | ')}`);
  return wb;
}
function readSheetAsAOA(wb, sheetName){
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, raw:true, defval:'' });
}
function writeAOAToSheet(wb, sheetName, aoa){
  wb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(aoa);
}
function downloadWorkbook(wb, filename){
  log(`Descargando archivo: ${filename}`);
  const out = XLSX.write(wb, { bookType:'xlsx', type:'array' });
  const blob = new Blob([out], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
function cloneWorkbookAllSheets(wb){
  const wbOut = XLSX.utils.book_new();
  wb.SheetNames.forEach(n=>{
    const wsCopy = XLSX.utils.aoa_to_sheet(
      XLSX.utils.sheet_to_json(wb.Sheets[n], { header:1, raw:true, defval:'' })
    );
    XLSX.utils.book_append_sheet(wbOut, wsCopy, n);
  });
  return wbOut;
}
function getSheetInsensitive(wb, desiredName){
  const target = normalizeHeader(desiredName);
  return wb.SheetNames.find(n => normalizeHeader(n) === target) || null;
}
function buildHeaderMapFromTemplate(aoa, headerRowIndexGuess=1){
  let headerRowIdx = headerRowIndexGuess;
  for (let i=0; i<Math.min(120, aoa.length); i++){
    const joined = (aoa[i]||[]).map(x=>String(x||'')).join(' | ').toLowerCase();
    if (
      joined.includes('tipo de promoción') ||
      joined.includes('2-') ||
      joined.includes('eans') ||
      joined.includes('fecha inicio')
    ){
      headerRowIdx=i;
      break;
    }
  }

  const headerRow = aoa[headerRowIdx] || [];
  const headerRow2 = aoa[headerRowIdx + 1] || [];
  const map = new Map();

  const setKey = (k, idx) => {
    const key = normalizeHeader(k);
    if (!key) return;
    if (!map.has(key)) map.set(key, idx);
  };

  headerRow.forEach((h, idx)=> setKey(h, idx));
  headerRow2.forEach((h, idx)=> setKey(h, idx));

  const max = Math.max(headerRow.length, headerRow2.length);
  for (let idx=0; idx<max; idx++){
    const h1 = headerRow[idx];
    const h2 = headerRow2[idx];
    if (String(h1||'').trim() && String(h2||'').trim()){
      setKey(`${h1} / ${h2}`, idx);
      setKey(`${h1} ${h2}`, idx);
      setKey(`${h2} / ${h1}`, idx);
      setKey(`${h2} ${h1}`, idx);
    }
  }
  return { map, headerRowIdx };
}
function getCellByCandidates(headerMap, candidates){
  for (const c of candidates){
    const k = normalizeHeader(c);
    if (headerMap.has(k)) return headerMap.get(k);
  }
  return -1;
}

function hideUnusedSheets(wbOut, usedSheetNames){
  const usedNorm = new Set(Array.from(usedSheetNames).map(n => normalizeHeader(n)));
  wbOut.Workbook = wbOut.Workbook || {};
  wbOut.Workbook.Sheets = wbOut.SheetNames.map(n => ({
    name: n,
    Hidden: usedNorm.has(normalizeHeader(n)) ? 0 : 1
  }));
  const visibles = wbOut.SheetNames.filter(n => usedNorm.has(normalizeHeader(n)));
  const ocultas = wbOut.SheetNames.filter(n => !usedNorm.has(normalizeHeader(n)));
  log(`SHEETS visibles: ${visibles.join(' | ') || '(ninguna)'}`);
  log(`SHEETS ocultas: ${ocultas.join(' | ') || '(ninguna)'}`);
}

const SKIP_ROW_TOKENS = ['NO CONSIDERAR', 'PRECIO FIJO'];
function normalizeSkipToken(v){
  return String(v ?? '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g,' ')
    .replace(/[.\-_/]+/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}
function isSkipRowValue(v){
  const t = normalizeSkipToken(v);
  if (!t) return false;
  return SKIP_ROW_TOKENS.some(tok => t.includes(tok));
}
function filterOrigenSkipRowsAnyCell(parsed){
  const headers = parsed?.headers || [];
  const rows = parsed?.aoaRows || [];
  let skipped = 0;
  const keptRows = [];
  for (const row of rows){
    const shouldSkip = (row || []).some(cell => isSkipRowValue(cell));
    if (shouldSkip){ skipped++; continue; }
    keptRows.push(row);
  }
  const data = keptRows.map(r=>{
    const o = {};
    headers.forEach((h, idx)=>{ o[h] = r[idx]; });
    return o;
  });
  return { kept: data, skipped };
}

function parseOrigenSheetFirst(wb){
  const desired = "Completar";
  const picked = getSheetInsensitive(wb, desired) || wb.SheetNames[0];
  const sheetName = picked;
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:'' });

  let headerIdx = -1;
  for (let i=0; i<Math.min(100, aoa.length); i++){
    const joined = (aoa[i]||[]).map(x=>normalizeHeader(x)).join(' | ');
    if (joined.includes('codigo') && joined.includes('producto') && (joined.includes('inicio') || joined.includes('termino'))){
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) headerIdx = 0;

  const headers = (aoa[headerIdx]||[]).map(h=>String(h||'').trim());
  const aoaRows = aoa.slice(headerIdx+1).filter(r=>r && r.some(x=>String(x||'').trim()!==''));
  const data = aoaRows.map(r=>{
    const o = {};
    headers.forEach((h, idx)=>{ o[h] = r[idx]; });
    return o;
  });

  log(`Origen parseado. Hoja="${sheetName}", headers=${headers.length}, filas=${data.length}`);
  return { sheetName, headers, data, aoaRows, headerIdx };
}

function findEventSheetCandidates(wb){
  const candidates = [];
  for (const sheetName of wb.SheetNames){
    if (normalizeHeader(sheetName) === normalizeHeader(CFG.CODIGO_MARCA_SHEET)) continue;
    const aoa = readSheetAsAOA(wb, sheetName);
    const headerRow = aoa[1] || [];
    const headerMap = new Map();
    headerRow.forEach((h, idx)=>{
      const key = normalizeHeader(h);
      if (key) headerMap.set(key, idx);
    });
    const required = [EVENT_COLS.RC, EVENT_COLS.NUMERO, EVENT_COLS.LOCAL, EVENT_COLS.INICIO, EVENT_COLS.FIN, EVENT_COLS.MARCA, EVENT_COLS.DESCUENTO];
    const ok = required.every(group => getCellByCandidates(headerMap, group) >= 0);
    if (ok) candidates.push({ sheetName, aoa, headerRow, headerMap });
  }
  return candidates;
}
function pickEventSheetCandidate(wb, fileName){
  const candidates = findEventSheetCandidates(wb);
  if (!candidates.length) throw new Error('No encontré una hoja de Eventos con columnas RC, N°CAM, LOCAL, FEHA DE INICIO EVENTO, FEHA TERMINO EVENTO, MARCA y DESCUENTO.');
  const fileToken = normalizeHeader(String(fileName || '').replace(/\.[^.]+$/,''));
  const byName = candidates.find(c => fileToken.includes(normalizeHeader(c.sheetName)) || normalizeHeader(c.sheetName).includes(fileToken));
  if (byName) return byName;
  return candidates[candidates.length - 1];
}
function parseEventosWorkbook(wb, fileName){
  const chosen = pickEventSheetCandidate(wb, fileName);
  const sheetName = chosen.sheetName;
  const aoa = chosen.aoa;
  const headers = (aoa[1] || []).map(h => String(h || '').trim());
  const headerMap = new Map();
  headers.forEach((h, idx)=>{
    const key = normalizeHeader(h);
    if (key) headerMap.set(key, idx);
  });
  const rows = [];
  for (let i=2; i<aoa.length; i++){
    const row = aoa[i] || [];
    if (!row.some(v => String(v ?? '').trim() !== '')) continue;
    const obj = {};
    headers.forEach((h, idx)=>{ obj[h] = row[idx]; });
    rows.push(obj);
  }
  log(`Eventos parseado. Hoja="${sheetName}", filas=${rows.length}`);
  return { sheetName, headers, headerMap, data: rows };
}
function parseCodigoMarcaSheet(wb){
  const sheetName = getSheetInsensitive(wb, CFG.CODIGO_MARCA_SHEET);
  if (!sheetName) throw new Error(`No encontré la hoja "${CFG.CODIGO_MARCA_SHEET}" en el Excel de Eventos.`);
  const aoa = readSheetAsAOA(wb, sheetName);
  const headers = (aoa[0] || []).map(h=>String(h||'').trim());
  const rows = [];
  for (let i=1; i<aoa.length; i++){
    const row = aoa[i] || [];
    if (!row.some(v => String(v ?? '').trim() !== '')) continue;
    const obj = {};
    headers.forEach((h, idx)=>{ obj[h] = row[idx]; });
    rows.push(obj);
  }
  log(`CODIGO-MARCA parseado. filas=${rows.length}`);
  return { sheetName, headers, data: rows };
}

function detectKey(origenData, candidates){
  const keys = Object.keys(origenData[0]||{}).map(k=>({raw:k, norm: normalizeHeader(k)}));
  for (const c of candidates){
    const cn = normalizeHeader(c);
    const hit = keys.find(k=>k.norm===cn) || keys.find(k=>k.norm.includes(cn));
    if (hit) return hit.raw;
  }
  return null;
}

function detectTipoDescuentoOptions(origenData){
  if (!origenData?.length) return { options: [], keyTipo: null };
  const keyTipo = detectKey(origenData, ORIGEN_COLS.TIPO_DESC);
  if (!keyTipo) return { options: [], keyTipo: null };

  const set = new Set();
  for (const r of origenData){
    const v = String(r[keyTipo] ?? '').trim().toUpperCase();
    if (v && SUPPORTED_TIPOS.includes(v)) set.add(v);
  }
  return { options: Array.from(set), keyTipo };
}

function detectCoberturaLocalesTodos(origenData){
  if (!origenData?.length) return false;
  const kCob = detectKey(origenData, ORIGEN_COLS.COBERTURA_LOCALES);
  if (!kCob) return false;
  return origenData.some(r => String(r[kCob] ?? '').trim().toUpperCase() === 'TODOS');
}

function groupPromosByNumero(origenData, tipoDescuento){
  if (!origenData?.length) return { error:'Excel Origen sin datos.' };

  const kNumero = detectKey(origenData, ORIGEN_COLS.NUMERO);
  const kCodigoProd = detectKey(origenData, ORIGEN_COLS.CODIGO_PRODUCTO);
  const kPvpOfertaPack = detectKey(origenData, ORIGEN_COLS.PVP_OFERTA_PACK);
  const kDescPct = detectKey(origenData, ORIGEN_COLS.DESC_PCT);
  const kInicio = detectKey(origenData, ORIGEN_COLS.INICIO);
  const kTermino = detectKey(origenData, ORIGEN_COLS.TERMINO);
  const kTipoDesc = detectKey(origenData, ORIGEN_COLS.TIPO_DESC);
  const kUnidadesPack = detectKey(origenData, ORIGEN_COLS.UNIDADES_PACK);

  const missing = [];
  if (!kNumero) missing.push('N°');
  if (!kCodigoProd) missing.push('Código Producto');
  if (!kInicio) missing.push('F. Inicio');
  if (!kTermino) missing.push('F. Término');
  if (!kTipoDesc) missing.push('Tipo de Descuento');
  if (missing.length) return { error:'Faltan columnas en Excel Origen: ' + missing.join(', ') };

  const filtered = origenData.filter(r => String(r[kTipoDesc] ?? '').trim().toUpperCase() === tipoDescuento);
  if (!filtered.length) return { promos: [] };

  if (tipoDescuento === 'NOMINAL' && !kPvpOfertaPack) return { error:'Para NOMINAL falta "PVP Oferta Pack".' };
  if ((tipoDescuento === 'PORCENTUAL' || tipoDescuento === 'DESCUENTO 2DA UNIDAD (%)') && !kDescPct){
    return { error:`Para ${tipoDescuento} falta columna de descuento.` };
  }
  if (tipoDescuento === 'PACK 2X1'){
    if (!kPvpOfertaPack) return { error:'Para PACK 2X1 falta "PVP Oferta Pack".' };
    if (!kUnidadesPack) return { error:'Para PACK 2X1 falta "# Unidades Pack".' };
  }

  const groups = new Map();
  for (const row of filtered){
    const n = String(row[kNumero]).trim();
    if (!n) continue;
    if (!groups.has(n)) groups.set(n, []);
    groups.get(n).push(row);
  }

  const promos = [];
  for (const [n, rows] of groups.entries()){
    const first = rows[0];
    const fIni = toYYYYMMDD(first[kInicio]);
    const fFin = toYYYYMMDD(first[kTermino]);
    if (!fIni || !fFin) return { error:`Promo N° ${n}: Fecha Inicio/Fin inválida.` };

    const eans = rows.map(r => sanitizeEAN(r[kCodigoProd])).filter(Boolean);

    if (tipoDescuento === 'PORCENTUAL'){
      const pct = normalizePercentNumber(first[kDescPct]);
      if (pct == null) return { error:`Promo N° ${n}: % inválido.` };
      promos.push({ tipo: tipoDescuento, numero: n, eans, fechaInicio: fIni, fechaFin: fFin, descuentoPct: pct });
    } else if (tipoDescuento === 'DESCUENTO 2DA UNIDAD (%)'){
      const pct = normalizePercentNumber(first[kDescPct]);
      if (pct == null) return { error:`Promo N° ${n}: % inválido para 2DA UNIDAD.` };
      promos.push({ tipo: tipoDescuento, numero: n, eans, fechaInicio: fIni, fechaFin: fFin, descuentoPct: pct });
    } else if (tipoDescuento === 'NOMINAL'){
      const precio = toNumberFlexible(first[kPvpOfertaPack]);
      if (precio == null) return { error:`Promo N° ${n}: "PVP Oferta Pack" inválido.` };
      promos.push({ tipo: tipoDescuento, numero: n, eans, fechaInicio: fIni, fechaFin: fFin, precioFinal: precio });
    } else if (tipoDescuento === 'PACK 2X1'){
      const pvp = toNumberFlexible(first[kPvpOfertaPack]);
      if (pvp == null) return { error:`Promo N° ${n}: PVP Oferta Pack inválido.` };
      const und = parseInt(String(first[kUnidadesPack] ?? '').trim(), 10);
      if (!isFinite(und) || und <= 0) return { error:`Promo N° ${n}: # Unidades Pack inválido.` };
      const precioUnit = +(pvp / 2).toFixed(6);
      promos.push({ tipo: tipoDescuento, numero: n, eans, fechaInicio: fIni, fechaFin: fFin, precioFinal: precioUnit, unidadesPack: und });
    }
  }
  return { promos };
}

function ensureUniqueName(name, usedSet){
  let out = name;
  let i = 2;
  while (usedSet.has(out)) out = `${name}_${i++}`;
  usedSet.add(out);
  return out;
}

function fillSimpleListSheet(wbOut, sheetNameInput, listRequests){
  if (!listRequests?.length) return null;
  const sheetName = getSheetInsensitive(wbOut, sheetNameInput);
  if (!sheetName) throw new Error(`No encontré la hoja "${sheetNameInput}" en el Excel Base.`);

  const aoa = readSheetAsAOA(wbOut, sheetName);
  let headerIdx = 0;
  for (let i=0; i<Math.min(60, aoa.length); i++){
    const joined = (aoa[i] || []).map(x => normalizeHeader(x)).join(' | ');
    if (joined.includes('nombre') && joined.includes('lista') && (joined.includes('item') || joined.includes('desde') || joined.includes('hasta'))){
      headerIdx = i; break;
    }
  }

  const headers = aoa[headerIdx] || [];
  const headerMap = new Map();
  headers.forEach((h, idx)=>{
    const key = normalizeHeader(h);
    if (key) headerMap.set(key, idx);
  });

  const colNombre = getCellByCandidates(headerMap, FIELD.LP_NOMBRE);
  const colDesde  = getCellByCandidates(headerMap, FIELD.LP_DESDE);
  const colHasta  = getCellByCandidates(headerMap, FIELD.LP_HASTA);
  const colItem   = getCellByCandidates(headerMap, FIELD.LP_ITEM);

  const miss = [];
  if (colNombre < 0) miss.push('Nombre de la Lista');
  if (colDesde < 0) miss.push('Desde');
  if (colHasta < 0) miss.push('Hasta');
  if (colItem < 0) miss.push('Item');
  if (miss.length) throw new Error(`Hoja "${sheetNameInput}" sin columnas: ${miss.join(', ')}`);

  let writeRow = headerIdx + 1;
  while (aoa[writeRow] && aoa[writeRow].some(v => String(v||'').trim() !== '')) writeRow++;

  for (const req of listRequests){
    for (const item of req.items){
      const maxCol = Math.max(colNombre, colDesde, colHasta, colItem);
      const row = Array.from({ length: maxCol + 1 }, () => '');
      row[colNombre] = req.listName;
      row[colDesde]  = req.desde;
      row[colHasta]  = req.hasta;
      row[colItem]   = item;
      aoa[writeRow++] = row;
    }
  }
  writeAOAToSheet(wbOut, sheetName, aoa);
  return sheetName;
}

function fillMechanicSheet(wbOut, sheetName, promos, nombreGeneral, descripcion, tipo, listaByPromoNumero, coberturaTodos, clubData){
  const aoa = readSheetAsAOA(wbOut, sheetName);
  const { map: headerMap, headerRowIdx } = buildHeaderMapFromTemplate(aoa, 1);
  const col = (cands) => getCellByCandidates(headerMap, cands);

  const colEANS = col(FIELD.EANS);
  const colListaRef = col(FIELD.LISTA_REF);
  const colIni = col(FIELD.FECHA_INICIO);
  const colFin = col(FIELD.FECHA_FIN);
  const colTipoPromo = col(FIELD.TIPO_PROMO);
  const colNombre = col(FIELD.NOMBRE);
  const colDesc = col(FIELD.DESCRIPCION);
  const colNombreGeneral = col(FIELD.NOMBRE_GENERAL);
  const colTipoDoc = col(FIELD.TIPO_DOCUMENTO);

  const colPorUnidad = col(FIELD.POR_UNIDAD);
  const colPrecioFinal = col(FIELD.PRECIO_FINAL);
  const colPct = col(FIELD.BP_PORCENTAJE);

  const colOpCant = col(FIELD.OP_CANTIDAD);
  const colCant = col(FIELD.CANTIDAD);
  const colCantBenef = col(FIELD.CANTIDAD_BENEFICIO);
  const colProdEstrategia = col(FIELD.PROD_ESTRATEGIA);
  const colCompite = col(FIELD.COMPITE);

  const colLocalesNeg = col(FIELD.LOCALES_NEGATIVOS);

  const colGrupoPromos = col(FIELD.GRUPO_PROMOCIONES);
  const colInscripto = col(FIELD.INSCRIPTO);
  const colConvenioSel = col(FIELD.CONVENIO_SELECCIONADO);
  const colCodigoConvenio = col(FIELD.CODIGO_CONVENIO);

  const required = [];
  if (colIni < 0) required.push('Fecha Inicio');
  if (colFin < 0) required.push('Fecha Fin');
  if (colTipoPromo < 0) required.push('Tipo de Promoción');
  if (colEANS < 0 && colListaRef < 0) required.push('EANS o Lista de Productos');

  if (tipo === 'NOMINAL'){
    if (colPorUnidad < 0) required.push('Por Unidad');
    if (colPrecioFinal < 0) required.push('Precio Final');
  } else if (tipo === 'PORCENTUAL'){
    if (colPct < 0) required.push('Porcentaje');
  } else if (tipo === 'DESCUENTO 2DA UNIDAD (%)'){
    if (colPct < 0) required.push('Porcentaje');
    if (colOpCant < 0) required.push('Operador Cantidad');
    if (colCant < 0) required.push('Cantidad');
    if (colCantBenef < 0) required.push('Cantidad Beneficio');
    if (colProdEstrategia < 0) required.push('Estrategia');
    if (colCompite < 0) required.push('Compite');
  } else if (tipo === 'PACK 2X1'){
    if (colPorUnidad < 0) required.push('Por Unidad');
    if (colPrecioFinal < 0) required.push('Precio Final');
    if (colOpCant < 0) required.push('Operador Cantidad');
    if (colCant < 0) required.push('Cantidad');
    if (colCantBenef < 0) required.push('Cantidad Beneficio');
    if (colProdEstrategia < 0) required.push('Estrategia');
  }

  if (coberturaTodos && colLocalesNeg < 0) required.push('AZ Locales / Lista Locales Negativos');

  if (clubData?.enabled){
    if (colGrupoPromos < 0) required.push('Grupo de Promociones (Grupo1|Grupo2)');
    if (colInscripto < 0) required.push('Inscripto');
    if (colConvenioSel < 0) required.push('Convenio Seleccionado');
    if (colCodigoConvenio < 0) required.push('Código Convenio');
  }

  if (required.length){
    throw new Error(`La hoja "${sheetName}" no tiene columnas requeridas para ${tipo}: ${required.join(', ')}`);
  }

  let writeRow = headerRowIdx + 1;
  while (aoa[writeRow] && aoa[writeRow].some(v => String(v||'').trim() !== '')) writeRow++;

  for (let i=0; i<promos.length; i++){
    const p = promos[i];
    const rowDescripcion = `N${String(p.numero).trim()} | ${descripcion || 'Excel Origen'}`;

    const touched = [
      colEANS, colListaRef, colIni, colFin, colTipoPromo,
      colNombre, colDesc, colNombreGeneral, colTipoDoc,
      colPorUnidad, colPrecioFinal, colPct,
      colOpCant, colCant, colCantBenef, colProdEstrategia, colCompite,
      colLocalesNeg,
      colGrupoPromos, colInscripto, colConvenioSel, colCodigoConvenio
    ].filter(c => c >= 0);

    const maxCol = Math.max(...touched);
    const row = Array.from({ length: maxCol + 1 }, () => '');

    if (colNombre >= 0) row[colNombre] = nombreGeneral;
    if (colDesc >= 0) row[colDesc] = rowDescripcion;
    if (colNombreGeneral >= 0) row[colNombreGeneral] = nombreGeneral;
    if (colTipoDoc >= 0) row[colTipoDoc] = CFG.DEFAULT_TIPO_DOCUMENTO;
    if (coberturaTodos && colLocalesNeg >= 0) row[colLocalesNeg] = 'EXC_LOCALES';

    if (clubData?.enabled){
      if (colGrupoPromos >= 0) row[colGrupoPromos] = CLUB_CFG.GRUPO_PROMOCIONES;
      if (colInscripto >= 0) row[colInscripto] = CLUB_CFG.INSCRIPTO;
      if (colConvenioSel >= 0) row[colConvenioSel] = clubData.csv;
      if (colCodigoConvenio >= 0) row[colCodigoConvenio] = clubData.csv;
    }

    const listName = listaByPromoNumero?.get(p.numero) || '';
    if (listName && colListaRef >= 0){
      if (colEANS >= 0) row[colEANS] = '';
      row[colListaRef] = listName;
    } else {
      if (colEANS >= 0) row[colEANS] = (p.eans || []).join(',');
    }

    row[colIni] = p.fechaInicio;
    row[colFin] = p.fechaFin;
    row[colTipoPromo] = CFG.TIPO_PROMOCION_DEFAULT;

    if (tipo === 'NOMINAL'){
      row[colPorUnidad] = 1;
      row[colPrecioFinal] = p.precioFinal;
    } else if (tipo === 'PORCENTUAL'){
      row[colPct] = p.descuentoPct;
    } else if (tipo === 'DESCUENTO 2DA UNIDAD (%)'){
      row[colPct] = p.descuentoPct;
      row[colCompite] = 3;
      row[colOpCant] = 1;
      row[colCant] = 2;
      row[colCantBenef] = 1;
      row[colProdEstrategia] = 1;
    } else if (tipo === 'PACK 2X1'){
      row[colPorUnidad] = 1;
      row[colPrecioFinal] = p.precioFinal;
      row[colOpCant] = 1;
      row[colCant] = p.unidadesPack;
      row[colCantBenef] = p.unidadesPack;
      row[colProdEstrategia] = 1;
    }

    aoa[writeRow + i] = row;
  }
  writeAOAToSheet(wbOut, sheetName, aoa);
  return sheetName;
}

function renderSummaryNormal({ outName, planTipos, promosByTipo, totalListas, coberturaTodos, skippedSkipRows, isClub }){
  const when = new Date();
  summaryWhen.textContent = when.toLocaleString();
  summaryOutName.textContent = outName;
  summaryTipos.textContent = `Normal${isClub ? ' · CLUB' : ''} · ${planTipos.join(' + ')}`;

  const totalPromos = planTipos.reduce((acc, t)=> acc + ((promosByTipo.get(t) || []).length), 0);
  summaryPromos.innerHTML = `${totalPromos} <small>(total)</small>`;
  summaryListas.innerHTML = `${totalListas} <small>(creadas)</small>`;

  const rows = [];
  for (const tipo of planTipos){
    const promos = promosByTipo.get(tipo) || [];
    const count = promos.length;
    const eans = promos.reduce((a, p)=> a + ((p.eans || []).length), 0);
    const sample = promos.slice(0, 3).map(p => `N${p.numero}`).join(', ') + (count > 3 ? '…' : '');
    rows.push({ tipo, count, eans, sample });
  }

  summaryTable.innerHTML = `
    <thead>
      <tr>
        <th>Tipo</th>
        <th>Promos</th>
        <th>EANS</th>
        <th>Muestra</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(r => `
        <tr>
          <td><b>${escapeHtml(r.tipo)}</b></td>
          <td>${r.count}</td>
          <td>${r.eans}</td>
          <td>${escapeHtml(r.sample || '-')}</td>
        </tr>
      `).join('')}
      <tr>
        <td colspan="4" style="color:#475569;">
          Cobertura Locales = ${coberturaTodos ? '<b>TODOS</b> → AZ="EXC_LOCALES"' : '<b>(no aplica)</b>'}
          &nbsp;|&nbsp; CLUB = <b>${isClub ? 'SI' : 'NO'}</b>
          &nbsp;|&nbsp; Filas ignoradas (NO CONSIDERAR / PRECIO FIJO): <b>${skippedSkipRows}</b>
        </td>
      </tr>
    </tbody>
  `;
  summaryWrap.style.display = 'block';
  setTimeout(()=> summaryWrap.scrollIntoView({ behavior:'smooth', block:'nearest' }), 0);
}

function renderSummaryEventos({ outName, usuario, promos, productListRequests, localListRequests, warnings }){
  const when = new Date();
  summaryWhen.textContent = when.toLocaleString();
  summaryOutName.textContent = outName;
  summaryTipos.textContent = `Eventos · ${usuario}`;
  summaryPromos.innerHTML = `${promos.length} <small>(campañas)</small>`;
  summaryListas.innerHTML = `${productListRequests.length + localListRequests.length} <small>(prod+loc)</small>`;

  summaryTable.innerHTML = `
    <thead>
      <tr>
        <th>N°CAM</th>
        <th>Marca</th>
        <th>Descuento</th>
        <th>Locales</th>
        <th>Lista productos</th>
        <th>Lista locales</th>
      </tr>
    </thead>
    <tbody>
      ${promos.map(p => `
        <tr>
          <td><b>${escapeHtml(p.numero)}</b></td>
          <td>${escapeHtml(p.marcaRaw)}</td>
          <td>${escapeHtml(p.discount.label)}</td>
          <td>${p.locales.length}</td>
          <td>${escapeHtml(p.productListName || '-')}</td>
          <td>${escapeHtml(p.localListName || (p.locales.length ? p.locales.join(',') : '-'))}</td>
        </tr>
      `).join('')}
      <tr>
        <td colspan="6" style="color:#475569;">
          Listas de productos: <b>${productListRequests.length}</b> &nbsp;|&nbsp;
          Listas de locales: <b>${localListRequests.length}</b> &nbsp;|&nbsp;
          Avisos: <b>${warnings.length}</b>
        </td>
      </tr>
    </tbody>
  `;
  summaryWrap.style.display = 'block';
  setTimeout(()=> summaryWrap.scrollIntoView({ behavior:'smooth', block:'nearest' }), 0);
}

const LS_KEY = 'promos_builder_generated_v17';
function stableHash(str){
  let h = 2166136261;
  for (let i=0; i<str.length; i++){
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(16);
}
function getSessionKey(){
  const o = lastOrigenMeta ? `${lastOrigenMeta.name}|${lastOrigenMeta.size}|${lastOrigenMeta.lastModified}` : '';
  const b = lastBaseMeta ? `${lastBaseMeta.name}|${lastBaseMeta.size}|${lastBaseMeta.lastModified}` : '';
  const convHash = stableHash(clubConveniosInput.value || '');
  const extra = currentSourceMode === SOURCE_MODE.EVENTOS
    ? `|${eventUserSelect.value}|${sanitizeLetters2(eventInitials.value)}`
    : `|${modoSelect.value}|club=${chkClub.checked ? 1 : 0}|edit=${chkClubConveniosEditar.checked ? 1 : 0}|conv=${convHash}`;
  return stableHash(`${currentSourceMode}||${o}||${b}${extra}`);
}
function loadGeneratedMap(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return {};
    const obj = JSON.parse(raw);
    return obj && typeof obj === 'object' ? obj : {};
  } catch { return {}; }
}
function saveGeneratedMap(mapObj){ localStorage.setItem(LS_KEY, JSON.stringify(mapObj)); }
function getGeneratedEntry(sessionKey){
  const m = loadGeneratedMap();
  return m[sessionKey] || null;
}
function markGenerated(sessionKey, tag, info){
  const m = loadGeneratedMap();
  if (!m[sessionKey]) m[sessionKey] = {};
  m[sessionKey][tag] = { at: new Date().toISOString(), ...info };
  saveGeneratedMap(m);
  log(`TRACK: marcado tag=${tag} sessionKey=${sessionKey}`);
}
function wasGenerated(sessionKey, tag){
  const e = getGeneratedEntry(sessionKey);
  return !!e?.[tag];
}

function openRegenModal({ label, previousInfo, onYes }){
  regenAction = onYes;
  const when = previousInfo?.at ? new Date(previousInfo.at).toLocaleString() : '(desconocido)';
  const prevFile = previousInfo?.outName ? `\nÚltimo archivo: ${previousInfo.outName}` : '';

  regenMsg.textContent =
    `Este mismo Origen/Base ya se procesó antes.\n` +
    `${label}\n\n¿Quieres volver a generarlo?`;

  regenMeta.textContent =
    `Origen: ${lastOrigenMeta?.name || '-'} (size=${lastOrigenMeta?.size || '-'}, lm=${lastOrigenMeta?.lastModified || '-'})\n` +
    `Base:   ${lastBaseMeta?.name || '-'} (size=${lastBaseMeta?.size || '-'}, lm=${lastBaseMeta?.lastModified || '-'})\n` +
    `Procesado: ${when}${prevFile}`;

  regenModalOverlay.style.display = 'flex';
  regenModalOverlay.setAttribute('aria-hidden', 'false');
}
function closeRegenModal(clearAction=true){
  regenModalOverlay.style.display = 'none';
  regenModalOverlay.setAttribute('aria-hidden', 'true');
  if (clearAction) regenAction = null;
}
regenCloseBtn.addEventListener('click', ()=> closeRegenModal(true));
regenCancelBtn.addEventListener('click', ()=> closeRegenModal(true));
regenModalOverlay.addEventListener('click', (e)=>{ if (e.target === regenModalOverlay) closeRegenModal(true); });
document.addEventListener('keydown', (e)=>{ if (e.key === 'Escape' && regenModalOverlay.style.display === 'flex') closeRegenModal(true); });
regenYesBtn.addEventListener('click', ()=>{
  const act = regenAction;
  closeRegenModal(false);
  regenAction = null;
  if (typeof act === 'function') act();
});

function computeListCountForTipo(tipo){
  if (!cacheOrigenParsed) return 0;
  const { kept } = filterOrigenSkipRowsAnyCell(cacheOrigenParsed);
  const grouped = groupPromosByNumero(kept, tipo);
  if (grouped.error) return 0;
  const umbral = Math.max(1, parseInt(umbralEans.value || '5', 10));
  return (grouped.promos || []).filter(p=> (p.eans || []).length > umbral).length;
}

function detectEventUsers(data){
  const keyRC = detectKey(data, EVENT_COLS.RC);
  if (!keyRC) return [];
  return Array.from(new Set(data.map(r => String(r[keyRC] ?? '').trim()).filter(Boolean))).sort();
}

function classifyEventDiscount(value){
  if (value == null || value === '') return { kind:'UNSUPPORTED', label:'', raw:'' };
  if (typeof value === 'number'){
    const pct = normalizePercentNumber(value);
    if (pct == null) return { kind:'UNSUPPORTED', label:String(value), raw:value };
    return { kind:'PERCENT', percent: Math.round(pct), label: `${Math.round(pct)}%` };
  }
  const raw = String(value).trim();
  const s = stripAccents(raw).toUpperCase().replace(/\s+/g,' ').trim();

  const secondUnit = s.match(/^(\d+(?:[.,]\d+)?)\s*%\s*2DA\s+UNIDAD$/);
  if (secondUnit){
    const pct = Math.round(Number(secondUnit[1].replace(',', '.')));
    return { kind:'SECOND_UNIT_PERCENT', percent: pct, label: `${pct}% 2DA UNIDAD` };
  }
  const simplePct = s.match(/^(\d+(?:[.,]\d+)?)\s*%?$/);
  if (simplePct){
    let pct = Number(simplePct[1].replace(',', '.'));
    if (pct > 0 && pct < 1) pct = pct * 100;
    return { kind:'PERCENT', percent: Math.round(pct), label: `${Math.round(pct)}%` };
  }
  return { kind:'UNSUPPORTED', label: raw, raw };
}

function splitEventBrandTokens(rawBrand){
  let s = stripAccents(rawBrand).toUpperCase().trim();
  s = s.replace(/\s*-\s*/g, '/').replace(/[|,;+]/g, '/').replace(/\s+Y\s+/g, '/');
  const alias = {
    'LRP': 'LA ROCHE POSAY',
    'LA ROCHE-POSAY': 'LA ROCHE POSAY'
  };
  return Array.from(new Set(
    s.split('/').map(t => t.trim()).filter(Boolean).map(t => alias[t] || t)
  ));
}

function buildCodigoMarcaIndex(data){
  const keyCode = detectKey(data, ['codigo']);
  const keyBrand = detectKey(data, ['marca']);
  if (!keyCode || !keyBrand) throw new Error('La hoja CODIGO-MARCA debe tener columnas "codigo" y "MARCA".');
  const map = new Map();
  for (const row of data){
    const code = sanitizeEAN(row[keyCode]);
    const brand = stripAccents(String(row[keyBrand] ?? '').trim()).toUpperCase();
    if (!code || !brand) continue;
    if (!map.has(brand)) map.set(brand, new Set());
    map.get(brand).add(code);
  }
  return map;
}

function buildEventosModel(eventRows, codigoMarcaRows, selectedUser, initials, umbralLocales){
  if (!eventRows?.length) throw new Error('El Excel de Eventos no tiene filas.');
  if (!selectedUser) throw new Error('Debes seleccionar el usuario RC.');
  const init = sanitizeLetters2(initials);
  if (init.length !== 2) throw new Error('Las iniciales para Eventos deben contener exactamente 2 letras.');

  const keyRC = detectKey(eventRows, EVENT_COLS.RC);
  const keyNumero = detectKey(eventRows, EVENT_COLS.NUMERO);
  const keyLocal = detectKey(eventRows, EVENT_COLS.LOCAL);
  const keyInicio = detectKey(eventRows, EVENT_COLS.INICIO);
  const keyFin = detectKey(eventRows, EVENT_COLS.FIN);
  const keyMarca = detectKey(eventRows, EVENT_COLS.MARCA);
  const keyDescuento = detectKey(eventRows, EVENT_COLS.DESCUENTO);

  const filtered = eventRows.filter(r => String(r[keyRC] ?? '').trim() === selectedUser);
  if (!filtered.length) throw new Error(`No encontré filas para el usuario RC "${selectedUser}".`);

  const groups = new Map();
  for (const row of filtered){
    const numero = String(row[keyNumero] ?? '').trim();
    if (!numero) continue;
    if (!groups.has(numero)) groups.set(numero, []);
    groups.get(numero).push(row);
  }

  const stamp = buildTimestampDDmmYYhhmm(new Date());
  const brandIndex = buildCodigoMarcaIndex(codigoMarcaRows);
  const warnings = [];
  const promos = [];
  const productListRequests = [];
  const localListRequests = [];
  const usedListNames = new Set();
  const productListByBrand = new Map();

  const brandUsage = new Map();
  for (const [numero, rows] of groups.entries()){
    const first = rows[0];
    const marcaRaw = String(first[keyMarca] ?? '').trim();
    const discount = classifyEventDiscount(first[keyDescuento]);

    if (discount.kind === 'UNSUPPORTED'){
      throw new Error(`Descuento no soportado en N°CAM ${numero}: ${discount.label || first[keyDescuento]}`);
    }

    const fechasInicio = rows.map(r => toYYYYMMDD(r[keyInicio])).filter(Boolean);
    const fechasFin = rows.map(r => toYYYYMMDD(r[keyFin])).filter(Boolean);
    if (!fechasInicio.length || !fechasFin.length) throw new Error(`N°CAM ${numero}: fechas inválidas.`);
    const fechaInicio = fechasInicio.sort()[0];
    const fechaFin = fechasFin.sort().slice(-1)[0];

    const localSet = new Set(rows.map(r => sanitizeEAN(r[keyLocal])).filter(Boolean));
    if (localSet.size) localSet.add('9999');
    const locales = Array.from(localSet);

    const promo = {
      numero,
      usuario: selectedUser,
      marcaRaw,
      fechaInicio,
      fechaFin,
      locales,
      discount
    };
    promos.push(promo);

    if (!brandUsage.has(marcaRaw)){
      brandUsage.set(marcaRaw, { desde: fechaInicio, hasta: fechaFin });
    } else {
      const current = brandUsage.get(marcaRaw);
      if (fechaInicio < current.desde) current.desde = fechaInicio;
      if (fechaFin > current.hasta) current.hasta = fechaFin;
    }
  }

  for (const [marcaRaw, usage] of brandUsage.entries()){
    const brandTokens = splitEventBrandTokens(marcaRaw);
    const codes = new Set();
    const unresolved = [];

    for (const token of brandTokens){
      const hit = brandIndex.get(token);
      if (!hit || !hit.size){
        unresolved.push(token);
        continue;
      }
      for (const c of hit) codes.add(c);
    }

    if (!codes.size){
      throw new Error(`No encontré códigos en CODIGO-MARCA para la marca "${marcaRaw}".`);
    }

    if (unresolved.length){
      warnings.push(`Marca "${marcaRaw}": sin match para ${unresolved.join(', ')}.`);
      log(`WARN MARCA: ${marcaRaw} -> faltan ${unresolved.join(', ')}`);
    }

    const listName = buildEventosListName({
      initialsInput: init,
      marcaRaw,
      inicioEvento: usage.desde,
      usedListNames
    });

    productListByBrand.set(marcaRaw, listName);
    productListRequests.push({
      listName,
      desde: usage.desde,
      hasta: usage.hasta,
      items: Array.from(codes).sort((a,b)=> String(a).localeCompare(String(b), 'es'))
    });
  }

  for (const promo of promos){
    promo.productListName = productListByBrand.get(promo.marcaRaw) || '';
    if ((promo.locales || []).length > umbralLocales){
      const listNameBase = `${init}_LOC_N${promo.numero}_${stamp}`;
      const listName = ensureUniqueName(listNameBase, usedListNames);
      promo.localListName = listName;
      localListRequests.push({
        listName,
        desde: promo.fechaInicio,
        hasta: promo.fechaFin,
        items: promo.locales.slice()
      });
    } else {
      promo.localListName = '';
    }
  }
  return { promos, productListRequests, localListRequests, warnings };
}

function fillEventosPorcentajeSheet(wbOut, promos, nombreGeneral){
  const sheetName = getSheetInsensitive(wbOut, CFG.EVENTOS_TARGET_SHEET);
  if (!sheetName) throw new Error(`No encontré hoja "${CFG.EVENTOS_TARGET_SHEET}" en el Excel Base.`);
  const aoa = readSheetAsAOA(wbOut, sheetName);
  const { map: headerMap, headerRowIdx } = buildHeaderMapFromTemplate(aoa, 1);
  const col = (cands) => getCellByCandidates(headerMap, cands);

  const colTipoPromo = col(FIELD.TIPO_PROMO);
  const colNombre = col(FIELD.NOMBRE);
  const colDesc = col(FIELD.DESCRIPCION);
  const colNombreGeneral = col(FIELD.NOMBRE_GENERAL);
  const colTipoDoc = col(FIELD.TIPO_DOCUMENTO);
  const colIni = col(FIELD.FECHA_INICIO);
  const colFin = col(FIELD.FECHA_FIN);
  const colListaProd = col(FIELD.LISTA_REF);
  const colLocalesPos = col(FIELD.LOCALES_POSITIVOS);
  const colListaLocalesPos = col(FIELD.LISTA_LOCALES_POSITIVOS);
  const colPct = col(FIELD.BP_PORCENTAJE);
  const colOpCant = col(FIELD.OP_CANTIDAD);
  const colCant = col(FIELD.CANTIDAD);
  const colCantBenef = col(FIELD.CANTIDAD_BENEFICIO);
  const colProdEstrategia = col(FIELD.PROD_ESTRATEGIA);
  const colCompite = col(FIELD.COMPITE);

  const required = [];
  if (colTipoPromo < 0) required.push('Tipo de Promoción');
  if (colNombre < 0) required.push('Nombre');
  if (colDesc < 0) required.push('Descripción');
  if (colNombreGeneral < 0) required.push('Nombre General');
  if (colTipoDoc < 0) required.push('Tipo Documento');
  if (colIni < 0) required.push('Fecha Inicio');
  if (colFin < 0) required.push('Fecha Fin');
  if (colListaProd < 0) required.push('Lista de Productos');
  if (colLocalesPos < 0 && colListaLocalesPos < 0) required.push('Locales Positivos o Lista Locales Positivos');
  if (colPct < 0) required.push('Porcentaje');
  if (colCompite < 0) required.push('Compite');
  if (colOpCant < 0) required.push('Operador Cantidad');
  if (colCant < 0) required.push('Cantidad');
  if (colCantBenef < 0) required.push('Cantidad Beneficio');
  if (colProdEstrategia < 0) required.push('Estrategia');
  if (required.length) throw new Error(`La hoja "${sheetName}" no tiene columnas requeridas para Eventos: ${required.join(', ')}`);

  let writeRow = headerRowIdx + 1;
  while (aoa[writeRow] && aoa[writeRow].some(v => String(v||'').trim() !== '')) writeRow++;

  for (let i=0; i<promos.length; i++){
    const p = promos[i];
    const touched = [
      colTipoPromo, colNombre, colDesc, colNombreGeneral, colTipoDoc, colIni, colFin,
      colListaProd, colLocalesPos, colListaLocalesPos, colPct,
      colOpCant, colCant, colCantBenef, colProdEstrategia, colCompite
    ].filter(c => c >= 0);
    const maxCol = Math.max(...touched);
    const row = Array.from({ length: maxCol + 1 }, () => '');

    row[colTipoPromo] = CFG.TIPO_PROMOCION_DEFAULT;
    row[colNombre] = nombreGeneral;
    row[colDesc] = `N${p.numero} | ${p.usuario} | ${p.marcaRaw}`;
    row[colNombreGeneral] = nombreGeneral;
    row[colTipoDoc] = CFG.DEFAULT_TIPO_DOCUMENTO;
    row[colIni] = p.fechaInicio;
    row[colFin] = p.fechaFin;
    row[colListaProd] = p.productListName;

    if (p.localListName && colListaLocalesPos >= 0){
      row[colListaLocalesPos] = p.localListName;
    } else if (colLocalesPos >= 0){
      row[colLocalesPos] = (p.locales || []).join(',');
    }

    if (p.discount.kind === 'PERCENT'){
      row[colPct] = p.discount.percent;
    } else if (p.discount.kind === 'SECOND_UNIT_PERCENT'){
      row[colCompite] = 3;
      row[colPct] = p.discount.percent;
      row[colOpCant] = 1;
      row[colCant] = 2;
      row[colCantBenef] = 1;
      row[colProdEstrategia] = 1;
    }

    aoa[writeRow + i] = row;
  }
  writeAOAToSheet(wbOut, sheetName, aoa);
  return sheetName;
}

function computeEventListCounts(){
  if (!cacheEventosParsed || !cacheCodigoMarca || !eventUserSelect.value || sanitizeLetters2(eventInitials.value).length !== 2) return { prod:0, loc:0 };
  try{
    const model = buildEventosModel(
      cacheEventosParsed.data,
      cacheCodigoMarca.data,
      eventUserSelect.value,
      eventInitials.value,
      Math.max(1, parseInt(umbralEans.value || '5', 10))
    );
    return { prod: model.productListRequests.length, loc: model.localListRequests.length };
  } catch {
    return { prod:0, loc:0 };
  }
}

function prettyModoValue(v){
  if (!v) return '-';
  if (v === '__ALL__') return 'Todos los tipos detectados';
  return `Solo ${v}`;
}
function setStepClass(el, cls){
  el.classList.remove('pending','active','done');
  el.classList.add(cls);
}

function refreshSteps(){
  const nombreGeneralOk = !!nombreGeneralInput.value.trim();
  const clubState = chkClub.checked ? getClubConveniosDataSafe() : { ok:true, list:['default'], error:'' };
  const clubOk = !chkClub.checked || (clubState.ok && clubState.list.length > 0);

  const step1Done = !!wbOrigen;
  const step2Done = !!wbBase;
  const step3Done = currentSourceMode === SOURCE_MODE.NORMAL
    ? !!modoSelect.value
    : !!eventUserSelect.value && sanitizeLetters2(eventInitials.value).length === 2;
  const step4Done = nombreGeneralOk && clubOk;
  const step5Done = currentSourceMode === SOURCE_MODE.NORMAL
    ? (!!wbOrigen && !!wbBase && !!modoSelect.value && nombreGeneralOk && clubOk)
    : (!!wbOrigen && !!wbBase && !!eventUserSelect.value && sanitizeLetters2(eventInitials.value).length === 2 && nombreGeneralOk);

  setStepClass(st1, step1Done ? 'done' : 'pending');
  setStepClass(st2, step2Done ? 'done' : 'pending');
  setStepClass(st3, step3Done ? 'done' : 'pending');
  setStepClass(st4, step4Done ? 'done' : 'pending');
  setStepClass(st5, step5Done ? 'done' : 'pending');

  const steps = [
    { el: st1, done: step1Done },
    { el: st2, done: step2Done },
    { el: st3, done: step3Done },
    { el: st4, done: step4Done },
    { el: st5, done: step5Done }
  ];
  const firstPending = steps.find(s => !s.done);
  if (firstPending) setStepClass(firstPending.el, 'active');
}

function refreshUI(){
  const nombreGeneralOk = !!nombreGeneralInput.value.trim();
  const clubState = chkClub.checked ? getClubConveniosDataSafe() : { ok:true, list:['default'], error:'' };
  const clubOk = !chkClub.checked || (clubState.ok && clubState.list.length > 0);

  const modeOk = currentSourceMode === SOURCE_MODE.NORMAL
    ? !!modoSelect.value
    : !!eventUserSelect.value && sanitizeLetters2(eventInitials.value).length === 2;
  const ok = !!wbOrigen && !!wbBase && nombreGeneralOk && modeOk && clubOk;
  btnProcesar.disabled = !ok;
  btnPreview.disabled = false;

  if (currentSourceMode === SOURCE_MODE.NORMAL){
    const modoText = `Flujo: ${prettyModoValue(modoSelect.value)}${chkClub.checked ? ' · CLUB' : ''}`;
    if (pillModo.textContent !== modoText) pop(pillModo);
    pillModo.textContent = modoText;

    if (!wbOrigen || !wbBase || !modoSelect.value || !cacheOrigenParsed){
      const listasText = 'Listas: 0';
      if (pillListas.textContent !== listasText) pop(pillListas);
      pillListas.textContent = listasText;
    } else {
      let tipos = [];
      if (modoSelect.value === '__ALL__') tipos = detectTipoDescuentoOptions(cacheOrigenParsed.data).options;
      else tipos = [modoSelect.value];
      const totalListas = tipos.reduce((acc, t)=> acc + computeListCountForTipo(t), 0);
      const listasText = `Listas: ${chkListas.checked ? totalListas : 0}`;
      if (pillListas.textContent !== listasText) pop(pillListas);
      pillListas.textContent = listasText;
    }
  } else {
    const userTxt = eventUserSelect.value ? ` · ${eventUserSelect.value}` : '';
    const modoText = `Flujo: Eventos${userTxt}`;
    if (pillModo.textContent !== modoText) pop(pillModo);
    pillModo.textContent = modoText;

    const counts = computeEventListCounts();
    const listasText = `Listas: P${counts.prod} / L${counts.loc}`;
    if (pillListas.textContent !== listasText) pop(pillListas);
    pillListas.textContent = listasText;
  }

  if(pillArea){
    const area = AREA_RESPONSABLE || '-';
    pillArea.innerHTML = '<strong>Área:</strong> ' + area;

    // 🎨 Color dinámico moderno
    if(area.includes('BYCP')){
      pillArea.style.background = 'linear-gradient(135deg,#ef4444,#b91c1c)';
    }else if(area.includes('FARMA')){
      pillArea.style.background = 'linear-gradient(135deg,#22c55e,#15803d)';
    }else if(area.includes('BIENESTAR')){
      pillArea.style.background = 'linear-gradient(135deg,#7c3aed,#5b21b6)';
    }else{
      pillArea.style.background = 'linear-gradient(135deg,#64748b,#334155)';
    }
  }
  if (pillArea && pillArea.textContent !== (AREA_RESPONSABLE || '-')) {
    pop(pillArea);
  }
  updateNormalListPreview();

  clubInfoField.style.display = (currentSourceMode === SOURCE_MODE.NORMAL && chkClub.checked) ? 'block' : 'none';
  applyClubConveniosEditMode();

  if (chkClub.checked && !clubState.ok){
    statusHint.textContent = clubState.error;
  }

  refreshSteps();
}

function resetOrigenState(soft=false){
  wbOrigen = null;
  cacheOrigenParsed = null;
  cacheEventosParsed = null;
  cacheCodigoMarca = null;
  origenFileName = '';
  lastOrigenMeta = null;

  if (!soft) fileOrigen.value = '';

  modoSelect.disabled = true;
  modoSelect.innerHTML = '<option value="">(carga Origen)</option>';

  eventUserSelect.disabled = true;
  eventUserSelect.innerHTML = '<option value="">(carga Excel de Eventos)</option>';
  eventInitials.value = '';

  refreshUI();
}

function resetAllAfterSuccess(){
  wbOrigen = null;
  wbBase = null;
  cacheOrigenParsed = null;
  cacheEventosParsed = null;
  cacheCodigoMarca = null;
  origenFileName = '';
  lastOrigenMeta = null;
  lastBaseMeta = null;

  fileOrigen.value = '';
  fileBase.value = '';

  modoSelect.disabled = true;
  modoSelect.innerHTML = '<option value="">(carga Origen)</option>';

  eventUserSelect.disabled = true;
  eventUserSelect.innerHTML = '<option value="">(carga Excel de Eventos)</option>';
  eventInitials.value = '';

  chkListas.checked = false;
  chkClub.checked = false;
  chkClubConveniosEditar.checked = false;
  clubConveniosInput.value = CLUB_CFG.DEFAULT_CONVENIOS_CSV;
  clubConveniosInput.disabled = true;

  umbralEans.value = 5;
  nombreListaBase.value = '';

  chkNombreGeneralDefault.checked = true;
  summaryWrap.style.display = 'none';

  clearLogs();
  applyNombreGeneralDefaultMode(true);
  setStatus('Listo. Pantalla reiniciada.', 'ok', 'Carga nuevos archivos para continuar.');
  refreshUI();
}

function applySourceModeUI(mode){
  currentSourceMode = mode;
  optNormal.classList.toggle('active', mode === SOURCE_MODE.NORMAL);
  optEventos.classList.toggle('active', mode === SOURCE_MODE.EVENTOS);

  normalModoField.classList.toggle('hidden', mode !== SOURCE_MODE.NORMAL);
  eventUserField.classList.toggle('hidden', mode !== SOURCE_MODE.EVENTOS);
  initialsField.classList.toggle('hidden', mode !== SOURCE_MODE.EVENTOS);
  normalListToggle.classList.toggle('hidden', mode !== SOURCE_MODE.NORMAL);
  clubToggle.classList.toggle('hidden', mode !== SOURCE_MODE.NORMAL);
  normalListNameField.classList.toggle('hidden', mode !== SOURCE_MODE.NORMAL);
  eventsInfoField.classList.toggle('hidden', mode !== SOURCE_MODE.EVENTOS);

  if (mode !== SOURCE_MODE.NORMAL && listPreviewPill){
    listPreviewPill.style.display = 'none';
  }

  clubInfoField.style.display = 'none';

  subOrigen.innerHTML = mode === SOURCE_MODE.NORMAL
    ? 'Modo normal: se lee la hoja <b>Completar</b> y se ignoran filas “NO CONSIDERAR” / “PRECIO FIJO”.'
    : 'Modo Eventos: se detecta la hoja del evento, se filtra por <b>RC</b> y se usa <b>CODIGO-MARCA</b> para crear listas de productos.';
  umbralLabel.textContent = mode === SOURCE_MODE.NORMAL ? 'Umbral EANS' : 'Umbral Locales';

  resetOrigenState(true);
  setStatus(
    mode === SOURCE_MODE.NORMAL ? 'Flujo normal activo.' : 'Flujo Eventos activo.',
    'ok',
    mode === SOURCE_MODE.NORMAL ? 'Carga un Origen tipo Completar.' : 'Carga el Excel de Eventos para detectar RC y marcas.'
  );
  refreshUI();
}

optNormal.addEventListener('click', ()=> applySourceModeUI(SOURCE_MODE.NORMAL));
optEventos.addEventListener('click', ()=> applySourceModeUI(SOURCE_MODE.EVENTOS));

nombreGeneralInput.addEventListener('input', refreshUI);
chkNombreGeneralDefault.addEventListener('change', ()=> applyNombreGeneralDefaultMode(false));
chkListas.addEventListener('change', refreshUI);
chkClub.addEventListener('change', ()=>{
  if (!chkClub.checked){
    chkClubConveniosEditar.checked = false;
  }
  applyClubConveniosEditMode();
  refreshUI();
});
chkClubConveniosEditar.addEventListener('change', ()=>{
  applyClubConveniosEditMode();
  refreshUI();
});
clubConveniosInput.addEventListener('input', refreshUI);
umbralEans.addEventListener('input', refreshUI);
nombreListaBase.addEventListener('input', ()=>{
  if(pillArea){
    pillArea.innerHTML = '<strong>Área:</strong> '+(AREA_RESPONSABLE || '-');
  }
  updateNormalListPreview();

  refreshUI();
});
modoSelect.addEventListener('change', refreshUI);
eventUserSelect.addEventListener('change', ()=>{
  if (!eventInitials.value.trim()){
    eventInitials.value = sanitizeLetters2(eventUserSelect.value);
  }
  refreshUI();
});
eventInitials.addEventListener('input', ()=>{
  const clean = sanitizeLetters2(eventInitials.value);
  if (eventInitials.value !== clean) eventInitials.value = clean;
  refreshUI();
});

fileOrigen.addEventListener('change', async (e)=>{
	log('cacheOrigenParsed OK:', !!cacheOrigenParsed);
    console.log('PARSE RESULT:', cacheOrigenParsed);
  const f = e.target.files?.[0];
  if (!f) return;

  showProcessing({
    title: 'Leyendo Origen…',
    origenName: f.name,
    baseName: lastBaseMeta?.name || '-',
    plan: currentSourceMode === SOURCE_MODE.NORMAL ? 'Parseando hoja Completar' : 'Detectando hoja de Eventos',
    outName: 'Preparando datos'
  });

  try{
    clearLogs();
    summaryWrap.style.display = 'none';
    lastOrigenMeta = { name: f.name, size: f.size, lastModified: f.lastModified };
    origenFileName = f.name || '';
    debugMeta.textContent = `Origen: ${origenFileName}`;
    wbOrigen = null;
    cacheOrigenParsed = null;
    cacheEventosParsed = null;
    cacheCodigoMarca = null;

    setStatus('Leyendo Origen...', '', currentSourceMode === SOURCE_MODE.NORMAL ? 'Parseando Completar…' : 'Detectando hoja de Eventos…');
    wbOrigen = await readFileAsWorkbook(f, 'ORIGEN');
    AREA_RESPONSABLE = detectAreaFromOrigenWB(wbOrigen);
    log('AREA RESPONSABLE (INPUT Y3): '+AREA_RESPONSABLE);
    // AUTO FLOW POR AREA
    if(AREA_RESPONSABLE){
      const newMode = ['EVENTOS','EVENTO'].some(x => AREA_RESPONSABLE.includes(x))
        ? SOURCE_MODE.EVENTOS
        : SOURCE_MODE.NORMAL;

      // SOLO cambiar UI, NO resetear data
      currentSourceMode = newMode;

      optNormal.classList.toggle('active', newMode === SOURCE_MODE.NORMAL);
      optEventos.classList.toggle('active', newMode === SOURCE_MODE.EVENTOS);

      normalModoField.classList.toggle('hidden', newMode !== SOURCE_MODE.NORMAL);
      eventUserField.classList.toggle('hidden', newMode !== SOURCE_MODE.EVENTOS);
      initialsField.classList.toggle('hidden', newMode !== SOURCE_MODE.EVENTOS);
      normalListToggle.classList.toggle('hidden', newMode !== SOURCE_MODE.NORMAL);
      clubToggle.classList.toggle('hidden', newMode !== SOURCE_MODE.NORMAL);
      normalListNameField.classList.toggle('hidden', newMode !== SOURCE_MODE.NORMAL);
      eventsInfoField.classList.toggle('hidden', newMode !== SOURCE_MODE.EVENTOS);

      log('AUTO FLOW (FIXED): ' + newMode);
    }

    if (currentSourceMode === SOURCE_MODE.NORMAL){
      cacheOrigenParsed = parseOrigenSheetFirst(wbOrigen);
	  console.log('PARSE OK:', cacheOrigenParsed);
      log('PARSE OK filas:', cacheOrigenParsed?.data?.length);
      const { options, keyTipo } = detectTipoDescuentoOptions(cacheOrigenParsed.data);
      const { skipped } = filterOrigenSkipRowsAnyCell(cacheOrigenParsed);
      log(`SKIP(any cell): se detectaron ${skipped} filas para ignorar (NO CONSIDERAR / PRECIO FIJO)`);

      modoSelect.innerHTML = '';
      if (!keyTipo || !options.length){
        modoSelect.disabled = true;
        modoSelect.innerHTML = '<option value="">(sin tipos soportados)</option>';
        setStatus('Origen OK, pero no encontré tipos soportados en "Tipo de Descuento".', 'err');
      } else {
        modoSelect.disabled = false;
        const optAll = document.createElement('option');
        optAll.value = '__ALL__';
        optAll.textContent = `Todos los tipos detectados (${options.join(' + ')})`;
        modoSelect.appendChild(optAll);

        const optSep = document.createElement('option');
        optSep.disabled = true;
        optSep.textContent = '──────────';
        modoSelect.appendChild(optSep);

        options.forEach(v=>{
          const opt=document.createElement('option');
          opt.value=v;
          opt.textContent=`Solo ${v}`;
          modoSelect.appendChild(opt);
        });

        modoSelect.value = '__ALL__';
        setStatus(
          `Origen cargado (hoja: ${cacheOrigenParsed.sheetName}). Tipos detectados: ${options.join(' | ')}`,
          'ok',
          skipped ? `Se ignorarán ${skipped} filas (NO CONSIDERAR / PRECIO FIJO).` : ''
        );
      }
    } else {
      cacheEventosParsed = parseEventosWorkbook(wbOrigen, f.name);
      cacheCodigoMarca = parseCodigoMarcaSheet(wbOrigen);
      const users = detectEventUsers(cacheEventosParsed.data);
      eventUserSelect.innerHTML = '<option value="">(selecciona usuario RC)</option>';
      users.forEach(u=>{
        const opt = document.createElement('option');
        opt.value = u;
        opt.textContent = u;
        eventUserSelect.appendChild(opt);
      });
      eventUserSelect.disabled = !users.length;
      if (users.length){
        eventUserSelect.value = users[0];
        eventInitials.value = sanitizeLetters2(users[0]);
      }
      const keyDescuento = detectKey(cacheEventosParsed.data, EVENT_COLS.DESCUENTO);
      const sampleDiscounts = Array.from(new Set(cacheEventosParsed.data.map(r => String(r[keyDescuento] ?? '').trim()).filter(Boolean)));
      setStatus(
        `Eventos cargado (hoja: ${cacheEventosParsed.sheetName}). Usuarios RC: ${users.join(' | ') || '(ninguno)'}`,
        'ok',
        `Descuentos detectados: ${sampleDiscounts.join(' | ')}`
      );
    }
  } catch(err){
    resetOrigenState(true);
    setStatus('Error leyendo Origen: ' + (err?.message || err), 'err');
    log(err?.stack || String(err));
  } finally {
    hideProcessing();
    refreshUI();
  }
});

fileBase.addEventListener('change', async (e)=>{
  const f = e.target.files?.[0];
  if (!f) return;

  showProcessing({
    title: 'Leyendo Base…',
    origenName: lastOrigenMeta?.name || '-',
    baseName: f.name,
    plan: 'Validando hojas del Excel base',
    outName: 'Preparando plantillas'
  });

  try{
    lastBaseMeta = { name: f.name, size: f.size, lastModified: f.lastModified };
    setStatus('Leyendo Base...', '', 'Clonando plantillas…');
    wbBase = await readFileAsWorkbook(f, 'BASE');
    setStatus('Base cargada.', 'ok');
  } catch(err){
    wbBase = null;
    lastBaseMeta = null;
    setStatus('Error leyendo Base: ' + (err?.message || err), 'err');
    log(err?.stack || String(err));
  } finally {
    hideProcessing();
    refreshUI();
  }
});

function buildGenerationPlan(){
  if (!cacheOrigenParsed) return [];
  const detected = detectTipoDescuentoOptions(cacheOrigenParsed.data).options;
  if (modoSelect.value === '__ALL__') return detected.slice();
  return [modoSelect.value];
}

async function runNormalGeneration(forceRegenerate=false){
  const nombreGeneral = nombreGeneralInput.value.trim();
  const isClub = !!chkClub.checked;
  const clubState = isClub ? getClubConveniosDataSafe() : { ok:true, list:[], csv:'' };

  if (!wbOrigen || !wbBase || !cacheOrigenParsed) throw new Error('Faltan archivos: carga Origen + Base.');
  if (!nombreGeneral) throw new Error('Falta completar: Nombre General.');
  if (isClub && !clubState.ok) throw new Error(clubState.error);
  if (isClub && !clubState.list.length) throw new Error('Debes ingresar al menos un convenio para CLUB.');

  const clubData = isClub ? { enabled:true, list: clubState.list, csv: clubState.csv } : { enabled:false, list:[], csv:'' };

  const planTipos = buildGenerationPlan().map(t => String(t).toUpperCase());
  if (!planTipos.length) throw new Error('No hay tipos para generar.');

  const sessionKey = getSessionKey();
  const tag = `NORMAL:${planTipos.join('+')}:CLUB=${isClub ? 1 : 0}:CONV=${stableHash(clubData.csv)}`;
  if (!forceRegenerate && wasGenerated(sessionKey, tag)){
    const prev = getGeneratedEntry(sessionKey)?.[tag];
    openRegenModal({
      label: `Tipos: ${planTipos.join(' + ')}${isClub ? ' · CLUB' : ''}`,
      previousInfo: prev,
      onYes: () => runGeneration({ forceRegenerate:true })
    });
    setStatus('Estos archivos ya se procesaron antes. Confirma en el popup si quieres regenerar.', 'warn');
    return;
  }

  const outName = `Promos_${planTipos.join('+').replace(/[^\w+]+/g,'_')}${isClub ? '_CLUB' : ''}.xlsx`;
  showProcessing({ origenName:lastOrigenMeta?.name, baseName:lastBaseMeta?.name, plan:`${planTipos.join(' + ')}${isClub ? ' · CLUB' : ''}`, outName });

  await new Promise(r => setTimeout(r, 30));
  setStatus('Procesando...', '', 'Generando el Excel…');

  const { kept: origenFiltrado, skipped: skippedSkipRows } = filterOrigenSkipRowsAnyCell(cacheOrigenParsed);
  log(`SKIP(any cell): skipped=${skippedSkipRows}, kept=${origenFiltrado.length}`);

  const coberturaTodos = detectCoberturaLocalesTodos(origenFiltrado);
  log(`Cobertura Locales TODOS = ${coberturaTodos}`);
  log(`CLUB = ${isClub}`);
  if (isClub) log(`Convenios CLUB = ${clubData.csv}`);

  const wbOut = cloneWorkbookAllSheets(wbBase);
  const usarListas = chkListas.checked;
  const umbral = Math.max(1, parseInt(umbralEans.value || '5', 10));

  const listaByPromoNumeroByTipo = new Map();
  const listRequests = [];
  const usedListNames = new Set();
  const promosByTipo = new Map();
  const usedSheets = new Set();

  for (const tipo of planTipos){
    const grouped = groupPromosByNumero(origenFiltrado, tipo);
    if (grouped.error) throw new Error(grouped.error);

    const promos = grouped.promos || [];
    promosByTipo.set(tipo, promos);

    const listaByPromoNumero = new Map();
    if (usarListas){
      const baseName = (nombreListaBase.value || '').trim();
      for (const p of promos){
        const eansCount = (p.eans || []).length;
        if (eansCount > umbral){
          const unique = buildNormalListName({
            baseInput: baseName,
            promoIndex: Number(p.numero) || 1,
            usedListNames
          });
          listaByPromoNumero.set(p.numero, unique);
          listRequests.push({
            listName: unique,
            desde: p.fechaInicio,
            hasta: p.fechaFin,
            items: (p.eans || []).slice()
          });
        }
      }
    }
    listaByPromoNumeroByTipo.set(tipo, listaByPromoNumero);
  }

  if (listRequests.length){
    const listaSheetName = fillSimpleListSheet(wbOut, CFG.LISTA_SHEET_NAME, listRequests);
    if (listaSheetName) usedSheets.add(listaSheetName);
  }

  if (planTipos.includes('NOMINAL')){
    const promos = promosByTipo.get('NOMINAL') || [];
    if (promos.length){
      const sheetName = getSheetInsensitive(wbOut, CFG.SHEET_BY_TIPO.NOMINAL);
      if (!sheetName) throw new Error(`No encontré hoja "${CFG.SHEET_BY_TIPO.NOMINAL}".`);
      const used = fillMechanicSheet(
        wbOut,
        sheetName,
        promos,
        nombreGeneral,
        origenFileName,
        'NOMINAL',
        listaByPromoNumeroByTipo.get('NOMINAL'),
        coberturaTodos,
        clubData
      );
      usedSheets.add(used);
    }
  }

  if (planTipos.includes('PORCENTUAL')){
    const promos = promosByTipo.get('PORCENTUAL') || [];
    if (promos.length){
      const sheetName = getSheetInsensitive(wbOut, CFG.SHEET_BY_TIPO.PORCENTUAL);
      if (!sheetName) throw new Error(`No encontré hoja "${CFG.SHEET_BY_TIPO.PORCENTUAL}".`);
      const used = fillMechanicSheet(
        wbOut,
        sheetName,
        promos,
        nombreGeneral,
        origenFileName,
        'PORCENTUAL',
        listaByPromoNumeroByTipo.get('PORCENTUAL'),
        coberturaTodos,
        clubData
      );
      usedSheets.add(used);
    }
  }

  if (planTipos.includes('DESCUENTO 2DA UNIDAD (%)')){
    const promos = promosByTipo.get('DESCUENTO 2DA UNIDAD (%)') || [];
    if (promos.length){
      const sheetName = getSheetInsensitive(wbOut, CFG.SHEET_BY_TIPO['DESCUENTO 2DA UNIDAD (%)']);
      if (!sheetName) throw new Error(`No encontré hoja "${CFG.SHEET_BY_TIPO['DESCUENTO 2DA UNIDAD (%)']}".`);
      const used = fillMechanicSheet(
        wbOut,
        sheetName,
        promos,
        nombreGeneral,
        origenFileName,
        'DESCUENTO 2DA UNIDAD (%)',
        listaByPromoNumeroByTipo.get('DESCUENTO 2DA UNIDAD (%)'),
        coberturaTodos,
        clubData
      );
      usedSheets.add(used);
    }
  }

  if (planTipos.includes('PACK 2X1')){
    const promos = promosByTipo.get('PACK 2X1') || [];
    if (promos.length){
      const sheetName = getSheetInsensitive(wbOut, CFG.SHEET_BY_TIPO.NOMINAL);
      if (!sheetName) throw new Error(`No encontré hoja "${CFG.SHEET_BY_TIPO.NOMINAL}" para PACK 2X1.`);
      const used = fillMechanicSheet(
        wbOut,
        sheetName,
        promos,
        nombreGeneral,
        origenFileName,
        'PACK 2X1',
        listaByPromoNumeroByTipo.get('PACK 2X1'),
        coberturaTodos,
        clubData
      );
      usedSheets.add(used);
    }
  }

  hideUnusedSheets(wbOut, usedSheets);

  downloadWorkbook(wbOut, outName);
  markGenerated(sessionKey, tag, { outName });

  renderSummaryNormal({
    outName,
    planTipos,
    promosByTipo,
    totalListas: usarListas ? (new Set(listRequests.map(r=>r.listName))).size : 0,
    coberturaTodos,
    skippedSkipRows,
    isClub
  });

  setStatus(`OK: ${planTipos.join(' + ')}${isClub ? ' · CLUB' : ''}`, 'ok', `Descargado: ${outName}`);

  setTimeout(() => {
    resetAllAfterSuccess();
  }, 350);
}

async function runEventosGeneration(forceRegenerate=false){
  const nombreGeneral = nombreGeneralInput.value.trim();
  const selectedUser = eventUserSelect.value;
  const initials = eventInitials.value;

  if (!wbOrigen || !wbBase || !cacheEventosParsed || !cacheCodigoMarca) throw new Error('Faltan archivos: carga Excel de Eventos + Base.');
  if (!nombreGeneral) throw new Error('Falta completar: Nombre General.');
  if (!selectedUser) throw new Error('Debes seleccionar el usuario RC.');
  if (sanitizeLetters2(initials).length !== 2) throw new Error('Las iniciales de Eventos deben contener exactamente 2 letras.');

  const sessionKey = getSessionKey();
  const tag = `EVENTOS:${selectedUser}:${sanitizeLetters2(initials)}`;
  if (!forceRegenerate && wasGenerated(sessionKey, tag)){
    const prev = getGeneratedEntry(sessionKey)?.[tag];
    openRegenModal({
      label: `Eventos · Usuario RC: ${selectedUser}`,
      previousInfo: prev,
      onYes: () => runGeneration({ forceRegenerate:true })
    });
    setStatus('Estos archivos ya se procesaron antes. Confirma en el popup si quieres regenerar.', 'warn');
    return;
  }

  const outName = `Promos_EVENTOS_${sanitizeNameChunk(selectedUser, 16)}.xlsx`;
  showProcessing({ origenName:lastOrigenMeta?.name, baseName:lastBaseMeta?.name, plan:`EVENTOS · ${selectedUser}`, outName });

  await new Promise(r => setTimeout(r, 30));
  setStatus('Procesando Eventos...', '', 'Armando listas de productos y locales…');

  const model = buildEventosModel(
    cacheEventosParsed.data,
    cacheCodigoMarca.data,
    selectedUser,
    initials,
    Math.max(1, parseInt(umbralEans.value || '5', 10))
  );

  const wbOut = cloneWorkbookAllSheets(wbBase);
  const usedSheets = new Set();

  const prodSheetUsed = fillSimpleListSheet(wbOut, CFG.LISTA_SHEET_NAME, model.productListRequests);
  if (prodSheetUsed) usedSheets.add(prodSheetUsed);

  const locSheetUsed = fillSimpleListSheet(wbOut, CFG.LISTA_LOCALES_SHEET_NAME, model.localListRequests);
  if (locSheetUsed) usedSheets.add(locSheetUsed);

  const targetSheetUsed = fillEventosPorcentajeSheet(wbOut, model.promos, nombreGeneral);
  if (targetSheetUsed) usedSheets.add(targetSheetUsed);

  hideUnusedSheets(wbOut, usedSheets);

  downloadWorkbook(wbOut, outName);
  markGenerated(sessionKey, tag, { outName });

  if (model.warnings.length){
    setStatus(`OK con avisos: Eventos ${selectedUser}`, 'warn', model.warnings.join(' | '));
  } else {
    setStatus(`OK: Eventos ${selectedUser}`, 'ok', `Descargado: ${outName}`);
  }

  renderSummaryEventos({
    outName,
    usuario: selectedUser,
    promos: model.promos,
    productListRequests: model.productListRequests,
    localListRequests: model.localListRequests,
    warnings: model.warnings
  });

  setTimeout(() => {
    resetAllAfterSuccess();
  }, 350);
}

async function runGeneration({ forceRegenerate=false } = {}){
  log(`RUN: generation forceRegenerate=${forceRegenerate} mode=${currentSourceMode}`);
  try{
    summaryWrap.style.display = 'none';
    if (currentSourceMode === SOURCE_MODE.NORMAL){
      await runNormalGeneration(forceRegenerate);
    } else {
      await runEventosGeneration(forceRegenerate);
    }
  } catch (err){
    console.error(err);
    setStatus('Error: ' + (err?.message || err), 'err');
    log(err?.stack || String(err));
  } finally {
    hideProcessing();
    refreshUI();
  }
}

btnProcesar.addEventListener('click', ()=> runGeneration({ forceRegenerate:false }));

if (window.XLSX) setStatus('XLSX listo (offline).', 'ok');
else setStatus('No se cargó XLSX (xlsx.full.min.js).', 'err');

clubConveniosInput.value = CLUB_CFG.DEFAULT_CONVENIOS_CSV;
clubConveniosInput.disabled = true;
applyNombreGeneralDefaultMode(true);
applySourceModeUI(SOURCE_MODE.NORMAL);
// ================= PREVIEW =================
// ================= PREVIEW =================
const previewModalHTML = `
  <div class="modalOverlay" id="previewOverlay" aria-hidden="true">
    <div class="modal" role="dialog" aria-modal="true" aria-labelledby="previewTitle">
      <div class="modalH">
        <strong id="previewTitle">Visualización</strong>
        <button class="btn secondary" id="previewCloseBtn">Cerrar</button>
      </div>
      <div class="modalB">
        <p style="margin:0 0 10px 0;">
          Esto es solo una <b>previsualización</b>. No descarga archivos.
        </p>

        <div class="card" style="margin-bottom:10px;">
          <div class="k">Origen</div>
          <div class="v" id="previewOrigenName">-</div>
        </div>

        <div class="card" style="margin-bottom:10px;">
          <div class="k">Modo</div>
          <div class="v" id="previewModo">-</div>
        </div>

        <div class="card" style="margin-bottom:10px;">
          <div class="k">Nombre General</div>
          <div class="v" id="previewNombreGeneral">-</div>
        </div>

        <div class="card" style="margin-bottom:10px;">
          <div class="k">Cobertura Locales</div>
          <div class="v" id="previewCobertura">-</div>
        </div>

        <div class="card" style="margin-bottom:10px;">
          <div class="k">Club</div>
          <div class="v" id="previewClub">-</div>
        </div>

        <div class="card" style="margin-bottom:10px;">
          <div class="k">Listas</div>
          <div class="v" id="previewListas">-</div>
        </div>

        <div class="tableWrap" style="margin-top:10px;">
          <table id="previewTable"></table>
        </div>

        <div class="hintCard" style="margin-top:10px;">
          <b>Tip:</b> Si faltan columnas o datos, revisa el panel de logs.
        </div>
      </div>
      <div class="modalF">
        <button class="btn secondary" id="previewCloseBtn2">Cerrar</button>
      </div>
    </div>
  </div>
`;

document.body.insertAdjacentHTML('beforeend', previewModalHTML);

const previewOverlay = document.getElementById('previewOverlay');
const previewCloseBtn = document.getElementById('previewCloseBtn');
const previewCloseBtn2 = document.getElementById('previewCloseBtn2');

const previewOrigenName = document.getElementById('previewOrigenName');
const previewModo = document.getElementById('previewModo');
const previewNombreGeneral = document.getElementById('previewNombreGeneral');
const previewCobertura = document.getElementById('previewCobertura');
const previewClub = document.getElementById('previewClub');
const previewListas = document.getElementById('previewListas');
const previewTable = document.getElementById('previewTable');

function openPreviewModal(){
  previewOverlay.style.display = 'flex';
  previewOverlay.setAttribute('aria-hidden','false');
}
function closePreviewModal(){
  previewOverlay.style.display = 'none';
  previewOverlay.setAttribute('aria-hidden','true');
}
previewCloseBtn.addEventListener('click', closePreviewModal);
previewCloseBtn2.addEventListener('click', closePreviewModal);
previewOverlay.addEventListener('click', (e)=>{ if (e.target === previewOverlay) closePreviewModal(); });
document.addEventListener('keydown', (e)=>{ if (e.key === 'Escape' && previewOverlay.style.display === 'flex') closePreviewModal(); });

function buildPreviewRowsNormal({ planTipos, promosByTipo, umbral, usarListas }){
  const rows = [];
  for (const tipo of planTipos){
    const promos = promosByTipo.get(tipo) || [];
    for (const p of promos){
      const eansCount = (p.eans || []).length;
      const usesList = usarListas && eansCount > umbral;
      rows.push({
        tipo,
        numero: p.numero,
        inicio: p.fechaInicio,
        fin: p.fechaFin,
        eans: eansCount,
        lista: usesList ? 'SI' : 'NO',
        extra:
          tipo === 'NOMINAL' || tipo === 'PACK 2X1'
            ? `Precio=${p.precioFinal}${tipo === 'PACK 2X1' ? ` | UnidPack=${p.unidadesPack}` : ''}`
            : `Pct=${p.descuentoPct}`
      });
    }
  }
  return rows;
}

function renderPreviewTableNormal(rows){
  previewTable.innerHTML = `
    <thead>
      <tr>
        <th>Tipo</th>
        <th>N°</th>
        <th>Inicio</th>
        <th>Fin</th>
        <th>EANS</th>
        <th>Lista</th>
        <th>Extra</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(r => `
        <tr>
          <td><b>${escapeHtml(r.tipo)}</b></td>
          <td>${escapeHtml(r.numero)}</td>
          <td>${escapeHtml(r.inicio)}</td>
          <td>${escapeHtml(r.fin)}</td>
          <td>${r.eans}</td>
          <td>${escapeHtml(r.lista)}</td>
          <td>${escapeHtml(r.extra)}</td>
        </tr>
      `).join('')}
    </tbody>
  `;
}

function buildPreviewRowsEventos({ promos, umbralLocales }){
  return promos.map(p => ({
    numero: p.numero,
    usuario: p.usuario,
    marca: p.marcaRaw,
    inicio: p.fechaInicio,
    fin: p.fechaFin,
    descuento: p.discount.label,
    locales: (p.locales || []).length,
    listaLocales: (p.locales || []).length > umbralLocales ? 'SI' : 'NO',
    listaProductos: p.productListName
  }));
}

function renderPreviewTableEventos(rows){
  previewTable.innerHTML = `
    <thead>
      <tr>
        <th>N°CAM</th>
        <th>RC</th>
        <th>Marca</th>
        <th>Descuento</th>
        <th>Inicio</th>
        <th>Fin</th>
        <th>Locales</th>
        <th>Lista locales</th>
        <th>Lista productos</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(r => `
        <tr>
          <td><b>${escapeHtml(r.numero)}</b></td>
          <td>${escapeHtml(r.usuario)}</td>
          <td>${escapeHtml(r.marca)}</td>
          <td>${escapeHtml(r.descuento)}</td>
          <td>${escapeHtml(r.inicio)}</td>
          <td>${escapeHtml(r.fin)}</td>
          <td>${r.locales}</td>
          <td>${escapeHtml(r.listaLocales)}</td>
          <td>${escapeHtml(r.listaProductos || '-') }</td>
        </tr>
      `).join('')}
    </tbody>
  `;
}

async function runPreview(){
  currentSourceMode = SOURCE_MODE.NORMAL; // 🔥 FIX CLAVE

  log('RUN PREVIEW');
  console.log('RUN PREVIEW START');
  console.log('ANTES DE VALIDACIONES');
  try{
    if (!wbOrigen) throw new Error('Carga primero el Excel Origen.');
    if (currentSourceMode === SOURCE_MODE.NORMAL && !cacheOrigenParsed) throw new Error('Origen no está parseado todavía.');
    if (currentSourceMode === SOURCE_MODE.EVENTOS && (!cacheEventosParsed || !cacheCodigoMarca)) throw new Error('Eventos no está parseado todavía.');
    if (!nombreGeneralInput.value.trim()) throw new Error('Completa el Nombre General.');

    previewOrigenName.textContent = lastOrigenMeta?.name || origenFileName || '(sin nombre)';
    previewNombreGeneral.textContent = nombreGeneralInput.value.trim();

    if (currentSourceMode === SOURCE_MODE.NORMAL){
      const planTipos = buildGenerationPlan().map(t => String(t).toUpperCase());
      const { kept: origenFiltrado } = filterOrigenSkipRowsAnyCell(cacheOrigenParsed);

      const promosByTipo = new Map();
      for (const tipo of planTipos){
        const grouped = groupPromosByNumero(origenFiltrado, tipo);
        if (grouped.error) throw new Error(grouped.error);
        promosByTipo.set(tipo, grouped.promos || []);
      }

      const umbral = Math.max(1, parseInt(umbralEans.value || '5', 10));
      const usarListas = !!chkListas.checked;

      const coberturaTodos = detectCoberturaLocalesTodos(origenFiltrado);
      previewCobertura.textContent = coberturaTodos ? 'TODOS (AZ=EXC_LOCALES)' : '(no aplica)';
      previewClub.textContent = chkClub.checked ? `SI (${getClubConveniosDataSafe().list.length} convenios)` : 'NO';
      previewListas.textContent = usarListas ? `SI (umbral=${umbral})` : 'NO';
      previewModo.textContent = `Normal · ${planTipos.join(' + ')}`;

      const rows = buildPreviewRowsNormal({ planTipos, promosByTipo, umbral, usarListas });
      if (planTipos.length === 1 && planTipos[0] === 'NOMINAL') {
      const promos = promosByTipo.get('NOMINAL') || [];
      const coberturaTodos = detectCoberturaLocalesTodos(origenFiltrado);
      
        renderPreviewCardsNominal(promos, coberturaTodos);
      } else {
        renderPreviewTableNormal(rows); // fallback
      }
      openPreviewModal();
      return;
    }

    // EVENTOS
    const selectedUser = eventUserSelect.value;
    const initials = eventInitials.value;

    if (!selectedUser) throw new Error('Selecciona un usuario RC.');
    if (sanitizeLetters2(initials).length !== 2) throw new Error('Iniciales inválidas (2 letras).');

    const umbralLocales = Math.max(1, parseInt(umbralEans.value || '5', 10));

    const model = buildEventosModel(
      cacheEventosParsed.data,
      cacheCodigoMarca.data,
      selectedUser,
      initials,
      umbralLocales
    );

    previewCobertura.textContent = '(n/a eventos)';
    previewClub.textContent = '(n/a eventos)';
    previewListas.textContent = `Prod=${model.productListRequests.length} / Loc=${model.localListRequests.length}`;
    previewModo.textContent = `Eventos · ${selectedUser} · ini=${sanitizeLetters2(initials)}`;

    const rows = buildPreviewRowsEventos({ promos: model.promos, umbralLocales });
    renderPreviewTableEventos(rows);
    openPreviewModal();
  } catch(err){
	  console.error('ERROR PREVIEW:', err);
      alert(err.message);
    setStatus('Preview: ' + (err?.message || err), 'err');
    log(err?.stack || String(err));
  }
}
// ================= END PREVIEW =================

// (Nada más del código va aquí: este archivo original termina con el JS que ya venías usando.)
// Nota: La Parte 4/4 incluirá el cierre final y cualquier “bootstrap” restante si existiera en el original.
// ================= END PREVIEW =================

// Si en tu HTML original había más código después de:
// btnPreview.addEventListener('click', runPreview);
// entonces va aquí exactamente igual.

// (fin de app.js)
// ===== Preview (versión original que pegaste) =====

// OJO: si ya agregaste otro listener antes, evitá duplicarlo.
btnPreview.addEventListener('click', () => {
  console.log('CLICK PREVIEW');

  try{
    runPreview();
  }catch(err){
    console.error('ERROR CLICK:', err);
    alert(err.message);
  }
});

function renderPreviewCardsNominal(promos, coberturaTodos){
  const cards = promos.map(p => {
    const ean = (p.eans || [])[0] || '-';

    return `
      <div style="
        border-radius:18px;
        padding:18px;
        margin-bottom:18px;
        background:#ffffff;
        border:1px solid #e2e8f0;
        box-shadow:0 18px 40px rgba(0,0,0,0.08);
      ">

        <!-- HEADER -->
        <div style="
          display:flex;
          justify-content:space-between;
          align-items:center;
          margin-bottom:14px;
        ">
          <div style="font-weight:900; font-size:16px;">
            Promo N${p.numero}
          </div>
          <div style="
            font-size:11px;
            font-weight:800;
            padding:4px 8px;
            border-radius:999px;
            background:#dbeafe;
            color:#1d4ed8;
          ">
            PRECIO FIJO
          </div>
        </div>

        <!-- DESCRIPCIÓN -->
        <div style="
          font-size:12px;
          color:#475569;
          margin-bottom:14px;
        ">
          Vigencia: ${p.fechaInicio} → ${p.fechaFin}
        </div>

        <!-- CONDICIONES -->
        <div style="
          border-radius:12px;
          padding:12px;
          margin-bottom:12px;
          background:#f8fafc;
          border:1px solid #e2e8f0;
        ">
          <div style="
            font-size:12px;
            font-weight:900;
            margin-bottom:8px;
            color:#334155;
          ">
            CONDICIONES
          </div>

          <div style="display:flex; flex-direction:column; gap:6px; font-size:13px;">

            <div>
              📄 <b>Documento</b><br>
              <span style="color:#475569">Tipo = Boleta</span>
            </div>

            <div>
              🏪 <b>Locales</b><br>
              <span style="color:#475569">
                ${coberturaTodos ? 'Excluye lista EXC_LOCALES' : 'Todos los locales'}
              </span>
            </div>

            <div>
              📦 <b>Producto</b><br>
              <span style="color:#475569">EAN = ${ean}</span>
            </div>

          </div>
        </div>

        <!-- APLICADORES -->
        <div style="
          border-radius:12px;
          padding:12px;
          background:#f0fdf4;
          border:1px solid #bbf7d0;
        ">
          <div style="
            font-size:12px;
            font-weight:900;
            margin-bottom:8px;
            color:#166534;
          ">
            APLICADORES
          </div>

          <div style="display:flex; flex-direction:column; gap:6px; font-size:13px;">

            <div>
              💰 <b>Beneficio</b><br>
              <span style="color:#166534">
                Precio fijo: $${p.precioFinal}
              </span>
            </div>

            <div>
			  📦 <b>Aplicado a</b><br>
			  <span style="color:#166534">EAN = ${ean}</span>
			</div>

			<div>
			  🔢 <b>Por unidad</b><br>
			  <span style="color:#166534">SI</span>
			</div>

          </div>
        </div>

      </div>
    `;
  }).join('');

  previewTable.innerHTML = `<div style="display:flex; flex-direction:column;">${cards}</div>`;
}

// Con la separación en archivos, es mejor enganchar el botón por JS:
document.addEventListener('click', (e) => {
  const btn = e.target?.closest?.('[data-preview-close]');
  if (btn) closePreview();
});
