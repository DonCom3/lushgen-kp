<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ЛучСтрой КП-Генератор</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
<style>
:root{
  --graphite:#3f444b; --graphite-dark:#2b2f33; --graphite-soft:#4b5158;
  --yellow:#ffc107; --yellow-dark:#e0a800;
  --bg:#eef0f3; --white:#fff; --text:#1f2328; --muted:#6b7280;
  --border:#e5e7eb; --ok:#16a34a; --warn:#f59e0b; --err:#dc2626;
}
*{box-sizing:border-box}
html,body{height:100%}
body{margin:0;font-family:'Segoe UI',Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);line-height:1.5}

/* ============ LAYOUT ============ */
.app{display:grid;grid-template-columns:420px 1fr;height:100vh;overflow:hidden}
@media (max-width:900px){.app{grid-template-columns:1fr;height:auto;overflow:auto}}

/* ============ SIDEBAR ============ */
.sidebar{background:var(--graphite-dark);color:#fff;overflow-y:auto;padding:24px;border-right:3px solid var(--yellow)}
.sidebar h2{font-size:18px;margin:0 0 4px;letter-spacing:.5px}
.sidebar h2 span{color:var(--yellow)}
.sidebar .tagline{color:#9ca3af;font-size:12px;margin-bottom:24px}
.panel{background:#353a40;border-radius:14px;padding:16px;margin-bottom:16px;border:1px solid #454b52}
.panel h3{font-size:13px;text-transform:uppercase;letter-spacing:1px;color:var(--yellow);margin:0 0 12px}
.field{margin-bottom:10px}
.field label{display:block;font-size:12px;color:#cbd5e1;margin-bottom:4px}
.field input,.field select,.field textarea{
  width:100%;padding:9px 11px;border-radius:8px;border:1px solid #4b5158;
  background:#2b2f33;color:#fff;font-size:13px;font-family:inherit;
}
.field input:focus,.field select:focus,.field textarea:focus{outline:none;border-color:var(--yellow)}
.field.row{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.field.row .field{margin:0}

/* Dropzone */
.dropzone{
  border:2px dashed #5b6269;border-radius:12px;padding:22px 14px;text-align:center;
  cursor:pointer;transition:.2s;background:#2b2f33;
}
.dropzone:hover,.dropzone.drag{border-color:var(--yellow);background:#33383e}
.dropzone .icon{font-size:28px;color:var(--yellow)}
.dropzone .text{font-size:13px;color:#cbd5e1;margin-top:6px}
.dropzone .hint{font-size:11px;color:#8b9198;margin-top:4px}

.btn{
  width:100%;padding:11px 16px;border:none;border-radius:10px;
  font-weight:700;font-size:13px;cursor:pointer;transition:.15s;
  display:inline-flex;align-items:center;justify-content:center;gap:8px;
}
.btn.primary{background:var(--yellow);color:var(--graphite-dark)}
.btn.primary:hover{background:var(--yellow-dark)}
.btn.ghost{background:transparent;color:#cbd5e1;border:1px solid #4b5158}
.btn.ghost:hover{border-color:var(--yellow);color:var(--yellow)}
.btn:disabled{opacity:.4;cursor:not-allowed}
.btn-group{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px}

/* Status message */
.status{font-size:12px;padding:8px 12px;border-radius:8px;margin-top:10px;display:none}
.status.ok{background:rgba(22,163,74,.15);color:#4ade80;display:block}
.status.warn{background:rgba(245,158,11,.15);color:#fbbf24;display:block}
.status.err{background:rgba(220,38,38,.15);color:#f87171;display:block}

/* Mapping */
.mapping{font-size:11px;color:#9ca3af;margin-top:8px;display:none}
.mapping.show{display:block}
.mapping .map-row{display:flex;justify-content:space-between;padding:2px 0;border-bottom:1px dashed #454b52}
.mapping .map-row:last-child{border:none}
.mapping .map-row b{color:var(--yellow);font-weight:500}

/* ============ PREVIEW ============ */
.preview-wrap{overflow-y:auto;background:var(--bg);padding:24px}
.preview-toolbar{
  position:sticky;top:0;z-index:50;
  background:rgba(43,47,51,.95);backdrop-filter:blur(8px);
  padding:10px 16px;border-radius:12px;margin-bottom:20px;
  display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;
  border-bottom:2px solid var(--yellow);
}
.preview-toolbar .info{color:#cbd5e1;font-size:12px}
.preview-toolbar .info b{color:var(--yellow)}
.preview-toolbar .actions{display:flex;gap:8px;flex-wrap:wrap}
.preview-toolbar .actions .btn{width:auto;padding:9px 16px}

/* Empty state */
.empty{
  display:flex;flex-direction:column;align-items:center;justify-content:center;
  min-height:60vh;color:#9ca3af;text-align:center;padding:40px;
}
.empty .big{font-size:64px;color:var(--yellow);margin-bottom:12px}
.empty h2{color:var(--graphite);margin:0 0 8px}
.empty p{max-width:400px}

/* ============ PDF CONTENT (corporate style) ============ */
#pdf-content{background:var(--bg);border-radius:16px;overflow:hidden;box-shadow:0 10px 40px rgba(0,0,0,.08)}
.corp-header{background:linear-gradient(135deg,var(--graphite-dark),var(--graphite));color:#fff;padding:32px 36px;position:relative;overflow:hidden}
.corp-header::after{content:'';position:absolute;right:-60px;top:-60px;width:220px;height:220px;background:var(--yellow);opacity:.15;border-radius:50%}
.corp-header .top{display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;position:relative;z-index:1}
.corp-logo{font-weight:800;font-size:24px;letter-spacing:1px}
.corp-logo span{color:var(--yellow)}
.corp-badge{background:var(--yellow);color:var(--graphite-dark);padding:6px 14px;border-radius:999px;font-weight:700;font-size:12px;white-space:nowrap}
.corp-header h1{font-size:34px;margin:18px 0 8px;position:relative;z-index:1}
.corp-header .subtitle{color:#d1d5db;max-width:720px;position:relative;z-index:1;font-size:14px}
.corp-header .client{margin-top:14px;font-size:14px;color:#e5e7eb;position:relative;z-index:1}
.corp-header .client b{color:var(--yellow)}

.corp-body{padding:28px 36px}

.hero-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-bottom:28px}
.corp-card{background:#fff;border-radius:14px;padding:18px;border:1px solid var(--border);box-shadow:0 2px 12px rgba(0,0,0,.04)}
.corp-card.dark{background:var(--graphite);color:#fff;border-color:var(--graphite)}
.corp-card .label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.corp-card.dark .label{color:#cbd5e1}
.corp-card .price{font-size:26px;font-weight:800;color:var(--graphite);margin-top:4px}
.corp-card.dark .price{color:var(--yellow)}
.corp-card .hint{font-size:12px;color:var(--muted);margin-top:4px}
.corp-card.dark .hint{color:#9ca3af}

.corp-section-title{font-size:20px;margin:28px 0 12px;color:var(--graphite);display:flex;align-items:center;gap:10px}
.corp-section-title::before{content:'';width:6px;height:22px;background:var(--yellow);border-radius:2px}

.table-wrap{overflow-x:auto;background:#fff;border-radius:14px;box-shadow:0 2px 12px rgba(0,0,0,.04);margin-bottom:20px}
table{width:100%;border-collapse:collapse;min-width:640px}
th,td{padding:10px 12px;text-align:left;border-bottom:1px solid var(--border);font-size:13px;vertical-align:top}
th{background:#f9fafb;color:var(--muted);font-weight:600;text-transform:uppercase;font-size:11px;letter-spacing:.5px}
tr:last-child td{border-bottom:none}
td.num{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
td.center{text-align:center}
tr.total-row td{font-weight:800;background:#fff8e1;font-size:13px}
tr.grand-total td{font-weight:800;background:var(--graphite);color:#fff;font-size:14px}
tr.grand-total td span{color:var(--yellow)}
tr.section-sep td{background:#f3f4f6;font-weight:700;color:var(--graphite);text-transform:uppercase;font-size:11px;letter-spacing:1px}

.corp-footer{background:var(--graphite-dark);color:#fff;padding:28px 36px;margin-top:8px}
.corp-footer .grid{display:grid;grid-template-columns:1fr 1fr;gap:24px}
.corp-footer h3{margin:0 0 8px;color:var(--yellow);font-size:15px}
.corp-footer p{margin:4px 0;font-size:13px;color:#d1d5db}
.corp-footer a{color:var(--yellow);text-decoration:none}
.corp-footer .sign{margin-top:16px;padding-top:16px;border-top:1px solid #454b52;display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap}
.corp-footer .sign div{font-size:13px}
.qr{width:88px;height:88px;background:#fff;border-radius:8px;padding:4px}

/* Loader */
#pdf-loader{position:fixed;inset:0;background:rgba(43,47,51,.9);display:none;align-items:center;justify-content:center;z-index:1000;color:#fff;font-weight:600}
#pdf-loader .box{background:var(--graphite-dark);padding:32px 48px;border-radius:16px;border:2px solid var(--yellow);text-align:center}
#pdf-loader .spinner{width:40px;height:40px;border:4px solid rgba(255,255,255,.2);border-top-color:var(--yellow);border-radius:50%;margin:0 auto 16px;animation:spin 1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

/* ============ PRINT ============ */
@media print{
  @page{size:A4;margin:10mm}
  body{background:#fff}
  .sidebar,.preview-toolbar,#pdf-loader{display:none!important}
  .app{display:block;height:auto}
  .preview-wrap{padding:0;overflow:visible}
  #pdf-content{box-shadow:none;border-radius:0}
  .corp-header,.corp-footer{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .table-wrap{overflow:visible;box-shadow:none}
  table{min-width:0;font-size:10px}
  th,td{padding:5px 7px}
  tr{page-break-inside:avoid}
  thead{display:table-header-group}
  .corp-section-title{page-break-after:avoid}
}
</style>
</head>
<body>

<div class="app">

  <!-- ================= SIDEBAR ================= -->
  <aside class="sidebar">
    <h2>ЛУЧ<span>СТРОЙ</span> · КП-Генератор</h2>
    <div class="tagline">Загрузите расчёт — получите фирменное КП и PDF</div>

    <!-- 1. ЗАГРУЗКА -->
    <div class="panel">
      <h3>1. Загрузка расчёта</h3>
      <div class="dropzone" id="dropzone">
        <div class="icon">📊</div>
        <div class="text">Перетащите файл или нажмите</div>
        <div class="hint">.xlsx · .xls · .csv</div>
      </div>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" hidden>

      <div style="margin-top:12px">
        <div class="field">
          <label>…или вставьте данные из Excel</label>
          <textarea id="pasteArea" rows="4" placeholder="Скопируйте ячейки и вставьте сюда (Ctrl+V)"></textarea>
        </div>
        <div class="btn-group">
          <button class="btn ghost" onclick="useSample()">Пример</button>
          <button class="btn ghost" onclick="parsePaste()">Разобрать</button>
        </div>
      </div>

      <div class="status" id="status"></div>
      <div class="mapping" id="mapping"></div>
    </div>

    <!-- 2. РЕКВИЗИТЫ -->
    <div class="panel">
      <h3>2. Реквизиты КП</h3>
      <div class="field row">
        <div class="field"><label>№ КП</label><input id="kpNumber" value="0810–2026/1"></div>
        <div class="field"><label>Дата</label><input id="kpDate" value="08.10.2026"></div>
      </div>
      <div class="field"><label>Заказчик</label><input id="client" value="ВсеИнструменты.ру"></div>
      <div class="field"><label>Объект / заголовок</label><input id="object" value="Контейнерная площадка под ключ"></div>
      <div class="field row">
        <div class="field"><label>НДС, %</label><input id="vat" type="number" value="22"></div>
        <div class="field"><label>Срок действия, дней</label><input id="validity" type="number" value="15"></div>
      </div>
      <div class="field"><label>Описание (подзаголовок)</label>
        <textarea id="subtitle" rows="2">Производство и установка по эскизному проекту и ТЗ. Прозрачная смета, фиксированные цены.</textarea>
      </div>
    </div>

    <!-- 3. РЕКВИЗИТЫ КОМПАНИИ -->
    <div class="panel">
      <h3>3. Реквизиты компании</h3>
      <div class="field"><label>Юр. лицо</label><input id="company" value="ООО «ЛучСтрой Групп»"></div>
      <div class="field"><label>Адрес</label><input id="address" value="117420, Москва, Профсоюзная ул., д. 57, пом. 1/7"></div>
      <div class="field row">
        <div class="field"><label>Телефон</label><input id="phone" value="8 (985) 889-90-69"></div>
        <div class="field"><label>Email</label><input id="email" value="info@luchsg.ru"></div>
      </div>
      <div class="field"><label>Сайт</label><input id="site" value="https://luchsg.ru"></div>
      <div class="field row">
        <div class="field"><label>Директор</label><input id="director" value="Вилков А.В."></div>
        <div class="field"><label>Исполнитель</label><input id="manager" value="Кочетов В.В."></div>
      </div>
    </div>

    <!-- 4. ЭКСПОРТ -->
    <div class="panel">
      <h3>4. Экспорт</h3>
      <button class="btn primary" id="btnBuild" disabled>👁 Сформировать превью</button>
      <div class="btn-group" style="margin-top:8px">
        <button class="btn ghost" id="btnPdf" disabled>⬇ PDF</button>
        <button class="btn ghost" id="btnPrint" disabled>🖨 Печать</button>
      </div>
    </div>
  </aside>

  <!-- ================= PREVIEW ================= -->
  <main class="preview-wrap">
    <div class="preview-toolbar">
      <div class="info" id="toolbarInfo">Загрузите расчёт, чтобы увидеть превью</div>
      <div class="actions">
        <button class="btn primary" id="btnPdf2" disabled onclick="downloadPDF()">⬇ Скачать PDF</button>
        <button class="btn ghost" id="btnPrint2" disabled onclick="window.print()">🖨 Печать</button>
      </div>
    </div>

    <div id="emptyState" class="empty">
      <div class="big">📄</div>
      <h2>Превью появится здесь</h2>
      <p>Загрузите Excel/CSV или нажмите «Пример» в панели слева, чтобы увидеть, как сервис оформляет КП.</p>
    </div>

    <div id="pdf-content" style="display:none"></div>
  </main>
</div>

<!-- Loader -->
<div id="pdf-loader"><div class="box"><div class="spinner"></div>Формируем PDF…</div></div>

<script>
/* ============================================================
   ЛУЧСТРОЙ КП-ГЕНЕРАТОР
   ============================================================ */

let PARSED = null; // { meta, sections:[{title,rows,total}], grandTotal }

/* ---------- Форматирование чисел ---------- */
function fmt(n){
  if(n===null||n===undefined||isNaN(n)) return '—';
  return Number(n).toLocaleString('ru-RU',{maximumFractionDigits:2});
}
function parseNum(v){
  if(v===null||v===undefined) return null;
  if(typeof v==='number') return v;
  let s=String(v).replace(/[₽руб.\s]/gi,'').replace(/\u00a0/g,'').replace(',','.');
  s=s.replace(/[^0-9.\-]/g,'');
  const n=parseFloat(s);
  return isNaN(n)?null:n;
}

/* ---------- Определение колонок ---------- */
const COL_PATTERNS={
  num:      /^(№|no|номер|п\/п|пп)$/i,
  name:     /(наимен|назван|работа|материал|позиц)/i,
  req:      /(требован|характер|специф|гост)/i,
  qty:      /(кол-?во|колич|объ[её]м)/i,
  unit:     /^(ед|ед\.|единица|ед\.изм)/i,
  price:    /(цена|стоимост[ьи] за|price)/i,
  sum:      /(сумма|итого|всего|стоимост[ьи]$)/i,
  note:     /(примеч|коммент|note)/i
};
function detectColumns(headerRow){
  const map={};
  headerRow.forEach((cell,idx)=>{
    const v=String(cell||'').trim();
    for(const key in COL_PATTERNS){
      if(map[key]===undefined && COL_PATTERNS[key].test(v)){ map[key]=idx; break; }
    }
  });
  return map;
}

/* ---------- Определение раздела ---------- */
const SECTION_RE=/^(материал|работ|оборуд|услуг|монтаж|доставк)/i;
const TOTAL_RE=/^(итого|всего|итог)/i;

/* ---------- Парсинг массива строк ---------- */
function parseRows(rows){
  if(!rows||!rows.length) throw new Error('Пустые данные');

  // Находим строку заголовка (первые 5 строк ищем по ключевым словам)
  let headerIdx=-1, colMap=null;
  for(let i=0;i<Math.min(6,rows.length);i++){
    const m=detectColumns(rows[i]);
    if(m.name!==undefined && (m.qty!==undefined||m.sum!==undefined||m.price!==undefined)){
      headerIdx=i; colMap=m; break;
    }
  }
  if(headerIdx===-1) throw new Error('Не удалось найти строку заголовков. Нужны колонки: Наименование, Кол-во, Цена/Сумма.');

  const sections=[];
  let current={title:'ПОЗИЦИИ', rows:[], total:0};
  let grandTotal=0;
  let pendingTotal=null;

  for(let i=headerIdx+1;i<rows.length;i++){
    const r=rows[i]||[];
    const first=String(r[0]||'').trim();
    const name=String(r[colMap.name]!==undefined?r[colMap.name]:'').trim();

    // Пустая строка — пропускаем
    if(!first && !name) continue;

    // Строка-раздел (например, «МАТЕРИАЛЫ», «РАБОТЫ»)
    if(SECTION_RE.test(first) && !name){
      if(current.rows.length) sections.push(current);
      current={title:first.toUpperCase(), rows:[], total:0};
      continue;
    }

    // Строка-итог
    if(TOTAL_RE.test(first) || TOTAL_RE.test(name)){
      const sumVal = colMap.sum!==undefined ? parseNum(r[colMap.sum]) : null;
      if(sumVal!==null) pendingTotal=sumVal;
      continue;
    }

    // Обычная позиция
    const qty   = colMap.qty   !==undefined ? parseNum(r[colMap.qty])   : null;
    const price = colMap.price !==undefined ? parseNum(r[colMap.price]) : null;
    let   sum   = colMap.sum   !==undefined ? parseNum(r[colMap.sum])   : null;
    if(sum===null && qty!==null && price!==null) sum=qty*price;

    current.rows.push({
      num: colMap.num!==undefined ? String(r[colMap.num]||'').trim() : String(current.rows.length+1),
      name: name || first,
      req:  colMap.req!==undefined ? String(r[colMap.req]||'').trim() : '',
      qty, unit: colMap.unit!==undefined ? String(r[colMap.unit]||'').trim() : '',
      price, sum,
      note: colMap.note!==undefined ? String(r[colMap.note]||'').trim() : ''
    });
  }
  if(current.rows.length) sections.push(current);

  // Считаем итоги по разделам и общий
  sections.forEach(s=>{
    if(!s.total) s.total=s.rows.reduce((a,r)=>a+(r.sum||0),0);
    grandTotal+=s.total;
  });

  return {colMap, sections, grandTotal, headerIdx};
}

/* ---------- Чтение файла ---------- */
function handleFile(file){
  setStatus('warn','Читаем файл…');
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const data=new Uint8Array(e.target.result);
      const wb=XLSX.read(data,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      finishParse(rows);
    }catch(err){
      console.error(err);
      setStatus('err','Ошибка чтения файла: '+err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------- Парсинг вставки ---------- */
function parsePaste(){
  const text=document.getElementById('pasteArea').value.trim();
  if(!text){ setStatus('err','Вставьте данные в поле'); return; }
  const rows=text.split(/\r?\n/).map(line=>line.split('\t'));
  finishParse(rows);
}

/* ---------- Финализация парсинга ---------- */
function finishParse(rows){
  try{
    const res=parseRows(rows);
    PARSED=res;
    // Показываем какие колонки распознаны
    const m=res.colMap;
    const names={num:'№',name:'Наименование',req:'Требования',qty:'Кол-во',unit:'Ед.',price:'Цена',sum:'Сумма',note:'Примечание'};
    let html='';
    for(const k in names){
      html+=`<div class="map-row"><span>${names[k]}</span><b>${m[k]!==undefined?'✓ колонка '+(m[k]+1):'—'}</b></div>`;
    }
    const mp=document.getElementById('mapping');
    mp.innerHTML=html; mp.classList.add('show');

    const total=sectionsCount(res);
    setStatus('ok',`Распознано: ${total.sections} раздел(а), ${total.rows} позиций, итого ${fmt(res.grandTotal)} ₽`);
    document.getElementById('btnBuild').disabled=false;
  }catch(err){
    console.error(err);
    setStatus('err',err.message);
  }
}
function sectionsCount(res){
  let rows=0; res.sections.forEach(s=>rows+=s.rows.length);
  return {sections:res.sections.length, rows};
}

/* ---------- Статус ---------- */
function setStatus(type,msg){
  const el=document.getElementById('status');
  el.className='status '+type; el.textContent=msg;
}

/* ---------- Сборка мета ---------- */
function getMeta(){
  const g=id=>document.getElementById(id).value.trim();
  return {
    kpNumber:g('kpNumber'), date:g('kpDate'), client:g('client'),
    object:g('object'), subtitle:g('subtitle'),
    vat:parseNum(g('vat'))||0, validity:g('validity'),
    company:g('company'), address:g('address'),
    phone:g('phone'), email:g('email'), site:g('site'),
    director:g('director'), manager:g('manager')
  };
}

/* ---------- ГЛАВНАЯ: построение превью ---------- */
function buildPreview(){
  if(!PARSED){ setStatus('err','Сначала загрузите расчёт'); return; }
  const m=getMeta();
  const el=document.getElementById('pdf-content');
  el.innerHTML=renderKP(PARSED,m);
  el.style.display='block';
  document.getElementById('emptyState').style.display='none';
  document.getElementById('toolbarInfo').innerHTML=
    `КП <b>№ ${m.kpNumber}</b> от ${m.date} · <b>${fmt(PARSED.grandTotal)} ₽</b> с НДС ${m.vat}%`;
  document.getElementById('btnPdf').disabled=false;
  document.getElementById('btnPrint').disabled=false;
  document.getElementById('btnPdf2').disabled=false;
  document.getElementById('btnPrint2').disabled=false;
  setStatus('ok','Превью сформировано. Можно скачать PDF.');
}

/* ---------- Рендер КП ---------- */
function renderKP(data,m){
  const {sections,grandTotal}=data;

  // Определяем «варианты»: если разделов МАТЕРИАЛЫ+РАБОТЫ — это один вариант.
  // Разделим на варианты по ключевым словам, если есть два набора.
  const hasMultipleVariants = sections.some(s=>/вариант/i.test(s.title));
  const variants = hasMultipleVariants ? splitVariants(sections) : [{sections, total:grandTotal}];

  // Hero cards: если несколько вариантов — показываем цены, иначе одну общую
  let heroHTML='';
  if(variants.length>1){
    heroHTML=`<div class="hero-cards">`+variants.map((v,i)=>`
      <div class="corp-card ${i===1?'dark':''}">
        <div class="label">${v.title||('Вариант '+(i+1))}</div>
        <div class="price">${fmt(v.total)} ₽</div>
        <div class="hint">с НДС ${m.vat}%</div>
      </div>`).join('')+`
      <div class="corp-card">
        <div class="label">Срок действия</div>
        <div class="price">${m.validity} дн.</div>
        <div class="hint">с даты получения</div>
      </div>
    </div>`;
  }else{
    heroHTML=`
    <div class="hero-cards">
      <div class="corp-card">
        <div class="label">Общая стоимость</div>
        <div class="price">${fmt(grandTotal)} ₽</div>
        <div class="hint">с НДС ${m.vat}%</div>
      </div>
      <div class="corp-card dark">
        <div class="label">Срок действия КП</div>
        <div class="price">${m.validity} дней</div>
        <div class="hint">со дня фактического получения</div>
      </div>
      <div class="corp-card">
        <div class="label">Позиций</div>
        <div class="price">${sections.reduce((a,s)=>a+s.rows.length,0)}</div>
        <div class="hint">в ${sections.length} разделах</div>
      </div>
    </div>`;
  }

  // Основные таблицы
  let tablesHTML='';
  variants.forEach((v,i)=>{
    if(variants.length>1){
      tablesHTML+=`<h2 class="corp-section-title">${v.title||('Вариант '+(i+1))}</h2>`;
    }
    tablesHTML+=renderVariantTables(v.sections,m, v.total);
  });

  // Подвал
  const qrURL=`https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=${encodeURIComponent(m.site)}`;

  return `
  <div class="corp-header">
    <div class="top">
      <div class="corp-logo">ЛУЧ<span>СТРОЙ</span> ГРУПП</div>
      <div class="corp-badge">КП № ${m.kpNumber} от ${m.date}</div>
    </div>
    <h1>${m.object}</h1>
    <div class="subtitle">${m.subtitle}</div>
    <div class="client">Заказчик: <b>${m.client}</b></div>
  </div>

  <div class="corp-body">
    ${heroHTML}
    ${tablesHTML}
  </div>

  <div class="corp-footer">
    <div class="grid">
      <div>
        <h3>${m.company}</h3>
        <p>${m.address}</p>
        <p>Тел.: <a href="tel:${m.phone.replace(/[^\d+]/g,'')}">${m.phone}</a></p>
        <p>Email: <a href="mailto:${m.email}">${m.email}</a></p>
        <p>Сайт: <a href="${m.site}">${m.site}</a></p>
      </div>
      <div style="text-align:right">
        <img class="qr" src="${qrURL}" alt="QR">
      </div>
    </div>
    <div class="sign">
      <div>Генеральный директор: <b>${m.director}</b></div>
      <div>Исполнитель: <b>${m.manager}</b></div>
    </div>
    <p style="margin-top:16px;color:#9ca3af;font-size:12px">
      Срок действия коммерческого предложения — ${m.validity} календарных дней со дня фактического получения.
      Цены указаны с НДС ${m.vat}%.
    </p>
  </div>`;
}

/* ---------- Разбиение на варианты ---------- */
function splitVariants(sections){
  // Группируем: каждый раз, когда встречаем раздел с «Вариант» в названии,
  // начинаем новый вариант. Иначе объединяем пары МАТЕРИАЛЫ+РАБОТЫ.
  const variants=[]; let cur=null;
  sections.forEach(s=>{
    if(/вариант/i.test(s.title)){
      if(cur) variants.push(cur);
      cur={title:s.title, sections:[s], total:0};
    }else{
      if(!cur) cur={title:'', sections:[], total:0};
      cur.sections.push(s);
    }
    if(cur) cur.total=cur.sections.reduce((a,x)=>a+x.total,0);
  });
  if(cur) variants.push(cur);
  // Считаем итог по варианту
  variants.forEach(v=>v.total=v.sections.reduce((a,s)=>a+s.total,0));
  return variants;
}

/* ---------- Таблицы одного варианта ---------- */
function renderVariantTables(sections,m,total){
  let html='';
  sections.forEach((s,si)=>{
    html+=`<div class="table-wrap"><table>
      <thead><tr>
        <th style="width:36px">№</th>
        <th>Наименование</th>
        <th>Требования</th>
        <th class="num" style="width:70px">Кол-во</th>
        <th class="center" style="width:50px">Ед.</th>
        <th class="num" style="width:90px">Цена, ₽</th>
        <th class="num" style="width:100px">Сумма, ₽</th>
        <th>Примечание</th>
      </tr></thead>
      <tbody>`;
    s.rows.forEach(r=>{
      html+=`<tr>
        <td class="center">${r.num||''}</td>
        <td>${r.name||''}</td>
        <td>${r.req||''}</td>
        <td class="num">${r.qty!==null?fmt(r.qty):''}</td>
        <td class="center">${r.unit||''}</td>
        <td class="num">${r.price!==null?fmt(r.price):''}</td>
        <td class="num">${r.sum!==null?fmt(r.sum):''}</td>
        <td>${r.note||''}</td>
      </tr>`;
    });
    html+=`<tr class="total-row"><td colspan="6" style="text-align:right">Итого ${s.title}</td>
      <td class="num">${fmt(s.total)}</td><td></td></tr>`;
    html+=`</tbody></table></div>`;
  });

  html+=`<div class="table-wrap"><table>
    <tbody>
      <tr class="grand-total"><td style="text-align:right">ИТОГО с НДС ${m.vat}%</td>
        <td class="num" style="width:180px"><span>${fmt(total)} ₽</span></td>
        <td style="width:40%"></td></tr>
    </tbody></table></div>`;
  return html;
}

/* ---------- PDF ---------- */
function downloadPDF(){
  if(!PARSED){ setStatus('err','Сначала сформируйте превью'); return; }
  const loader=document.getElementById('pdf-loader');
  loader.style.display='flex';
  const m=getMeta();
  const filename=`КП_${m.company.replace(/[^\wа-яА-Я]/g,'')}_${m.kpNumber.replace(/[^\w\-]/g,'')}.pdf`;
  const element=document.getElementById('pdf-content');

  const opt={
    margin:[8,6,8,6],
    filename,
    image:{type:'jpeg',quality:0.98},
    html2canvas:{scale:2,useCORS:true,backgroundColor:'#ffffff',scrollY:0},
    jsPDF:{unit:'mm',format:'a4',orientation:'portrait'},
    pagebreak:{mode:['css','avoid-all'],avoid:'tr,.corp-card,.corp-section-title'}
  };

  html2pdf().set(opt).from(element).save()
    .then(()=>{loader.style.display='none';setStatus('ok','PDF сохранён');})
    .catch(err=>{
      console.error(err);
      loader.style.display='none';
      setStatus('warn','Автосохранение не удалось, открываю печать');
      setTimeout(()=>window.print(),300);
    });
}

/* ---------- Пример ---------- */
function useSample(){
  const sample=[
    ['№','Наименование','Требования','Кол-во','Ед.','Цена','Сумма','Примечание'],
    ['МАТЕРИАЛЫ'],
    [1,'Профильная труба (стойки)','100×100×3 мм, Ст3',30.8,'м.п.',975,30030,'8 стоек'],
    [2,'Профильная труба (ригели, фермы)','60×40×3 мм',110.0,'м.п.',650,71500,'5 ферм пролетом 7 м'],
    [3,'Профильная труба (обвязка, связи)','40×20×2 мм',123.0,'м.п.',450,55350,'прогоны кровли'],
    [4,'Профильная труба (каркас ворот)','40×40×2 мм',46.0,'м.п.',450,20700,'2 проема, 2 калитки'],
    [5,'Профлист стеновой С-21','0,5 мм, RAL 7024',119.4,'м²',1100,131296,'стены, ворота, калитки'],
    [6,'Профлист кровельный С-21','0,5 мм, RAL 7024',68.6,'м²',1100,75482,'кровля 9×7 м'],
    [7,'Лист гладкий на пол','1,5 мм',63.0,'м²',1450,91350,'742 кг'],
    ['','Итого материалы','','','','',475708,''],
    ['РАБОТЫ'],
    [9,'Изготовление навеса комплектом','резка, сварка, КМД, окраска',1,'компл.',681716,681716,''],
    [10,'Доставка на объект','Барыбино, длинномер',1,'рейс',69908,69908,''],
    [11,'Сварка листов пола','',63,'м²',7000,441000,''],
    [12,'Монтаж каркаса','',1,'компл.',298983,298983,''],
    [13,'Монтаж обшивки и кровли','саморезы, доборные планки',188,'м²',2300,432400,''],
    ['','Итого работы','','','','',1924007,'']
  ];
  finishParse(sample);
  setStatus('ok','Пример загружен. Нажмите «Сформировать превью».');
}

/* ---------- Слушатели ---------- */
document.getElementById('dropzone').addEventListener('click',()=>{
  document.getElementById('fileInput').click();
});
document.getElementById('fileInput').addEventListener('change',e=>{
  if(e.target.files[0]) handleFile(e.target.files[0]);
});
const dz=document.getElementById('dropzone');
['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{
  e.preventDefault(); dz.classList.add('drag');
}));
['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{
  e.preventDefault(); dz.classList.remove('drag');
}));
dz.addEventListener('drop',e=>{
  const f=e.dataTransfer.files[0];
  if(f) handleFile(f);
});
document.getElementById('btnBuild').addEventListener('click',buildPreview);
document.getElementById('btnPdf').addEventListener('click',downloadPDF);
document.getElementById('btnPrint').addEventListener('click',()=>window.print());

// Автоперестройка превью при изменении реквизитов
['kpNumber','kpDate','client','object','subtitle','vat','validity',
 'company','address','phone','email','site','director','manager'].forEach(id=>{
  document.getElementById(id).addEventListener('input',()=>{
    if(PARSED && document.getElementById('pdf-content').style.display!=='none') buildPreview();
  });
});
</script>
</body>
</html>