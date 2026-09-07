const EVENT_APP = Object.freeze({
  SHEETS:{TASKS:'Attività',EXPENSES:'Spese',PARTICIPANTS:'Partecipanti',META:'_META'},
  MASTER_DOCUMENTS:'_DOCUMENTI',
  INPUT_ROW:9,
  LIST_START_ROW:13,
  LIST_END_ROW:500,
  TEMPLATE_SPECS:[
    {key:'CONV_ATLETI',label:'Convocazione atleti',id:'1SpOMrpDwe8aTW9QAyby865WufnD40AsPx6t8zm6Yw3E'},
    {key:'CONV_TECNICO',label:'Convocazione tecnico',id:'1Xb61H5TQ0avz8n5_THn-BSd1Xd4YBtmOD3Lu2uwUuJw'},
    {key:'GOMMONE',label:'Richiesta gommone alla Zona',id:'1uvK4J8WrWekpMrkPTkpUvsdYTggBhwzCDJaJmyIhUko'},
    {key:'OSPITALITA',label:'Richiesta ospitalità circolo',id:'1NuheJqcwtLxp8muTdbnpcZc22Z8HJgA-jbdt_Sq6ruk'},
    {key:'RINGRAZIAMENTO',label:'Ringraziamento circolo',id:'11Fqn8U3Hs8MgJPH0tQTKDEozmWDftPtmELoLjgMFg7c'}
  ]
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Scheda evento')
    .addItem('➕ Importa nuova spesa','eventImportExpense')
    .addItem('🔄 Aggiorna cruscotto spese','eventRefreshExpenseDashboard')
    .addSeparator()
    .addItem('📄 Genera documenti','eventGenerateDocuments')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === EVENT_APP.SHEETS.EXPENSES) {
    if (e.range.getA1Notation() === 'J7' && e.value === 'TRUE') {
      try { eventImportExpense(); }
      catch (err) {
        sheet.getRange('J7').setValue(false);
        SpreadsheetApp.getActive().toast(err.message || String(err),'Spese',6);
      }
      return;
    }
    const row = e.range.getRow();
    if (row >= EVENT_APP.LIST_START_ROW && row <= EVENT_APP.LIST_END_ROW && e.range.getColumn() <= 10) {
      eventRefreshExpenseDashboard();
    }
  }
  if (sheet.getName() === EVENT_APP.SHEETS.PARTICIPANTS) {
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row >= 3 && row <= 42 && col >= 6 && col <= 8) eventRefreshExpenseDashboard();
  }
}

function eventMeta_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.META);
  if (!sheet) throw new Error('Foglio _META non trovato.');
  const values = sheet.getRange(1,1,sheet.getLastRow(),2).getDisplayValues();
  const out = {};
  values.forEach(r=>{ if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}

function eventImportExpense() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  const meta = eventMeta_();
  if (!sheet) throw new Error('Foglio Spese non trovato.');
  const input = sheet.getRange(EVENT_APP.INPUT_ROW,1,1,10).getValues()[0];
  const amount = Number(input[0]||0);
  const description = String(input[1]||'').trim();
  if (!amount && !description) throw new Error('Compila IMPORTO e DESCRIZIONE.');
  if (!description) throw new Error('La DESCRIZIONE è obbligatoria.');
  if (!input[2]) input[2] = 'ALTRO';
  if (!input[4]) input[4] = 'PREVENTIVO';
  if (String(input[4]).toUpperCase() !== 'PREVENTIVO' && !input[5]) input[5] = 'DA PAGARE';
  input[7] = meta.EVENT_COMMITMENT || '';

  let target = EVENT_APP.LIST_START_ROW;
  const descriptions = sheet.getRange(EVENT_APP.LIST_START_ROW,2,EVENT_APP.LIST_END_ROW-EVENT_APP.LIST_START_ROW+1,1).getDisplayValues();
  while (target <= EVENT_APP.LIST_END_ROW && String(descriptions[target-EVENT_APP.LIST_START_ROW][0]||'').trim()) target++;
  if (target > EVENT_APP.LIST_END_ROW) throw new Error('Elenco spese pieno.');

  const now = new Date();
  sheet.getRange(target,1,1,10).setValues([input]);
  sheet.getRange(target,11,1,6).setValues([['SPESA-' + Utilities.getUuid(),'','','',now,now]]);
  sheet.getRange(target,1).setNumberFormat('€ #,##0.00');
  sheet.getRange(target,8).setNumberFormat('@');
  sheet.getRange(EVENT_APP.INPUT_ROW,1,1,10).clearContent();
  sheet.getRange('H9').setValue(meta.EVENT_COMMITMENT || '').setNumberFormat('@');
  sheet.getRange('J7').setValue(false);
  eventRefreshExpenseDashboard();
  ss.toast('Spesa aggiunta all elenco. Usa Aggiorna calendario dalla scheda per sincronizzarla.','Spese',4);
}

function eventRefreshExpenseDashboard() {
  const ss = SpreadsheetApp.getActive();
  const expense = ss.getSheetByName(EVENT_APP.SHEETS.EXPENSES);
  const participants = ss.getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  if (!expense || String(expense.getRange('A1').getDisplayValue()) !== 'RIEPILOGO SPESE EVENTO') return;

  const rows = expense.getRange(EVENT_APP.LIST_START_ROW,1,EVENT_APP.LIST_END_ROW-EVENT_APP.LIST_START_ROW+1,11).getValues();
  const buckets = {VIAGGI:0,VITTO:0,ALLOGGIO:0,NOLEGGI:0,RIMBORSI:0,ALTRO:0};
  let forecast = 0;
  let paid = 0;
  rows.forEach(r=>{
    const id = String(r[10]||'');
    if (id.indexOf('RIMBORSO-AUTO-') === 0) return;
    const amount = Number(r[0]||0);
    if (!amount && !String(r[1]||'').trim()) return;
    const category = String(r[2]||'ALTRO').toUpperCase();
    let bucket = 'ALTRO';
    if (category === 'VIAGGI' || category === 'VITTO' || category === 'ALLOGGIO' || category === 'NOLEGGI') bucket = category;
    forecast += amount;
    buckets[bucket] += amount;
    if (String(r[5]||'').toUpperCase() === 'PAGATO') paid += amount;
  });

  if (participants) {
    const pRows = participants.getRange(3,1,40,10).getValues();
    pRows.forEach(r=>{
      const name = String(r[0]||'').trim();
      const surname = String(r[1]||'').trim();
      if (!name && !surname) return;
      const status = String(r[5]||'').toUpperCase();
      const max = Number(r[6]||0);
      const passed = Number(r[7]||0);
      const current = status === 'ASSENTE' ? 0 : (passed > 0 ? passed : max);
      forecast += current;
      buckets.RIMBORSI += current;
      if (passed > 0) paid += passed;
    });
  }

  expense.getRange('D2').setValue(forecast).setNumberFormat('€ #,##0.00');
  expense.getRange('F2').setValue(paid).setNumberFormat('€ #,##0.00');
  expense.getRange('A5:F5').setValues([[
    buckets.VIAGGI,buckets.VITTO,buckets.ALLOGGIO,buckets.NOLEGGI,buckets.RIMBORSI,buckets.ALTRO
  ]]).setNumberFormat('€ #,##0.00');
}

function eventGenerateDocuments() {
  const ui = SpreadsheetApp.getUi();
  const meta = eventMeta_();
  let prompt = 'Scegli i documenti da creare, separando i numeri con una virgola:\n\n';
  EVENT_APP.TEMPLATE_SPECS.forEach((t,i)=>prompt += (i+1) + ' = ' + t.label + '\n');
  const response = ui.prompt('Genera documenti',prompt + '\nEsempio: 1,2,4',ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const keys = eventParseDocumentSelection_(response.getResponseText());
  if (!keys.length) { ui.alert('Nessun documento valido selezionato.'); return; }

  const folder = DriveApp.getFolderById(meta.EVENT_FOLDER_ID);
  const replacements = eventDocumentReplacements_(meta);
  const created = [];
  keys.forEach(key=>{
    const spec = EVENT_APP.TEMPLATE_SPECS.find(x=>x.key===key);
    if (!spec) return;
    const copy = DriveApp.getFileById(spec.id).makeCopy(eventDocumentName_(key,replacements),folder);
    eventFillDocument_(copy.getId(),replacements,key);
    eventRegisterDocument_(meta,spec,copy);
    created.push(copy.getName());
  });
  ui.alert('Documenti creati',created.map(x=>'• '+x).join('\n'),ui.ButtonSet.OK);
}

function eventParseDocumentSelection_(text) {
  const seen = {};
  return String(text||'').split(/[;,\s]+/).map(x=>parseInt(x,10)).filter(n=>{
    if (!n || n<1 || n>EVENT_APP.TEMPLATE_SPECS.length || seen[n]) return false;
    seen[n]=true; return true;
  }).map(n=>EVENT_APP.TEMPLATE_SPECS[n-1].key);
}

function eventDocumentReplacements_(meta) {
  const participants = SpreadsheetApp.getActive().getSheetByName(EVENT_APP.SHEETS.PARTICIPANTS);
  const rows = participants ? participants.getRange(3,1,40,10).getDisplayValues() : [];
  const athletes = rows.filter(r=>{
    const hasName = String(r[0]||'').trim() || String(r[1]||'').trim();
    const role = String(r[4]||'ATLETA').toUpperCase();
    const status = String(r[5]||'').toUpperCase();
    return hasName && role === 'ATLETA' && status !== 'ASSENTE';
  });
  const athleteLines = athletes.map(r=>{
    const name = [r[0],r[1]].filter(Boolean).join(' ').trim();
    return r[2] ? name + ' – ' + r[2] : name;
  });
  const start = eventParseIsoDate_(meta.EVENT_START);
  const end = eventParseIsoDate_(meta.EVENT_END);
  return {
    'tipo':eventTitleCase_(meta.EVENT_TYPE||''),
    'classe':meta.EVENT_CLASS||'',
    'luogo':meta.EVENT_LOCATION||'',
    'localita':meta.EVENT_LOCATION||'',
    'zona':meta.EVENT_ZONE||'',
    'circolo':meta.EVENT_CLUB||'',
    'tecnico':meta.EVENT_TECHNICIANS||'',
    'lista tecnici':meta.EVENT_TECHNICIANS||'',
    'lista atleti':athleteLines.join('\n'),
    'numero atleti':athletes.length ? String(athletes.length) : '',
    'hotel/struttura':meta.EVENT_LODGING||'',
    'data inizio':eventFormatDate_(start),
    'data fine':eventFormatDate_(end),
    'data in':eventFormatDate_(start),
    'data out':eventFormatDate_(end),
    'date':eventFormatDateRange_(start,end),
    'data di oggi':eventFormatDate_(new Date()),
    'data oggi':eventFormatDate_(new Date()),
    'impegno':meta.EVENT_COMMITMENT||''
  };
}

function eventFillDocument_(docId,replacements,key) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const reps = Object.assign({},replacements);
  if (key === 'GOMMONE') reps.N = replacements.zona || '';
  if (key === 'OSPITALITA') reps.N = replacements['numero atleti'] || '';
  Object.keys(reps).forEach(k=>{
    const v = reps[k];
    if (v === '' || v === null || v === undefined) return;
    body.replaceText('(?i)\\{\\{' + eventEscapeRegex_(k) + '\\}\\}',eventSafeReplacement_(v));
  });
  if (reps.classe) body.replaceText('(?i)\\{classe\\}\\}',eventSafeReplacement_(reps.classe));
  doc.saveAndClose();
}

function eventDocumentName_(key,r) {
  const base = {
    CONV_ATLETI:['Convocazione atleti',r.tipo,r.classe,r.luogo,r.date],
    CONV_TECNICO:['Convocazione tecnico',r.tecnico,r.tipo,r.classe,r.date],
    GOMMONE:['Richiesta gommone',r.zona ? r.zona + ' Zona' : '',r.tipo,r.classe,r.date],
    OSPITALITA:['Richiesta ospitalità',r.circolo,r.tipo,r.classe,r.date],
    RINGRAZIAMENTO:['Ringraziamento',r.circolo,r.tipo,r.classe,r.date]
  }[key] || [r.tipo,r.classe,r.date];
  return base.filter(Boolean).join(' - ');
}

function eventRegisterDocument_(meta,spec,file) {
  if (!meta.MASTER_SPREADSHEET_ID) return;
  const master = SpreadsheetApp.openById(meta.MASTER_SPREADSHEET_ID);
  const sheet = master.getSheetByName(EVENT_APP.MASTER_DOCUMENTS);
  if (!sheet) return;
  sheet.appendRow([
    'DOC-' + Utilities.getUuid(),meta.EVENT_ID,spec.key,spec.id,file.getName(),
    file.getId(),file.getUrl(),'BOZZA MODIFICABILE',new Date(),''
  ]);
}

function eventParseIsoDate_(s) {
  const m = String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0) : null;
}
function eventFormatDate_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const months=['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];
  return d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
}
function eventFormatDateRange_(s,e) {
  if (!(s instanceof Date)) return '';
  if (!(e instanceof Date)) return eventFormatDate_(s);
  if (s.getTime()===e.getTime()) return eventFormatDate_(s);
  const months=['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre'];
  if (s.getFullYear()===e.getFullYear() && s.getMonth()===e.getMonth()) return s.getDate() + '-' + e.getDate() + ' ' + months[s.getMonth()] + ' ' + s.getFullYear();
  return eventFormatDate_(s) + ' - ' + eventFormatDate_(e);
}
function eventTitleCase_(s) { return String(s||'').toLowerCase().replace(/(^|\s|[-/])([a-zà-öø-ÿ])/g,(_,p,c)=>p+c.toUpperCase()); }
function eventEscapeRegex_(s) { return String(s||'').replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function eventSafeReplacement_(s) { return String(s||'').replace(/\\/g,'\\\\').replace(/\$/g,'\\$'); }
