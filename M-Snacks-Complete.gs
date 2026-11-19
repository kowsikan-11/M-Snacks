/**
 * M-Snacks | Single-file Google Apps Script app
 *
 * 1. Create a project at https://script.google.com and paste this entire file
 *    into Code.gs. Save the project.
 * 2. Run setup in the editor and approve Google Sheets access. Open the
 *    spreadsheet URL from the execution log. Keep the spreadsheet private.
 *    The Daily quantities tab has one column per material and product field.
 *    Run setup again after an upgrade to fill those columns from saved entries.
 * 3. Entry and Reports use the web app audience selected in Google.
 * 4. Deploy > New deployment > Web app. Set Execute as to Me, and select the
 *    narrowest access setting that covers your users. Share the /exec URL.
 * 5. Run setup again after deployment to save both web URLs in the Sheet.
 *
 * The setup date zone is Asia/Kuala_Lumpur. Review material names and units
 * in the Default materials tab before use. The historical yields are weighted
 * and the report requires the matching product input to make an estimate.
 * No sign-in or access code: anyone with the link can enter and view reports.
 *
 * 19-Nov-2026: M/C BAGODA and WHITE KACHANG hidden from Raw materials; Curry Leaf moved from
 * Finished products to Raw materials; பொ.கடலை மாவு fixed to kg (no unit drop-down); Finished
 * products "Input" is now "No. of Lot" with a + ADD NEW PRODUCTION button; new Import products
 * section (Product name, Bundle of 80kg, Bundle of 330kg) saved in column H of Daily production.
 * The GitHub Pages front end calls doPost, so set Who has access to Anyone.
 */
const PRODUCTS = [
  ['Plain Murukku','lot'], ['White Mullu Murukku','lot'], ['Red Mullu Murukku','lot'],
  ['Curry Murukku','lot'], ['Wheel Chips','kg'], ['Onion Ring','kg'], ['Star','lot'],
  ['Omam','lot'], ['Seeval','lot'], ['Corn Chips','box']
];
// Products that are no longer on the form. Kept so old entries keep their columns and report rows.
const OLD_PRODUCTS = [['Curry Leaf','kg']];
// Raw materials that are not shown on the form any more (old entries keep them).
const HIDDEN_MATERIALS = ['M/C BAGODA','WHITE KACHANG'];
// Raw materials added to the Default materials tab on the next load, and fixed units (no unit drop-down).
const ADDED_MATERIALS = [['Curry Leaf','kg']];
const FIXED_UNITS = [['பொ.கடலை மாவு','kg']];
const PRODUCT_UNIT = 'lot';     // unit of a product added with + ADD NEW PRODUCTION
const EXTRA_PRODUCTS_TAB = 'Added products';
const LOG_WIDTH = 8;            // Daily production columns (8 = Import products)
const MATERIALS = [
  ['க.மாவு நைஸ்','kg'], ['க.அ.மாவு','kg'],
  ['கி.மாவு 5T54','bag'], ['மிளகாய் திப்பை','kg'], ['எல்லு','kg'], ['சீரகம்','kg'],
  ['ஓமம்','kg'], ['உப்பு','kg'], ['சோடப்பு','kg'], ['(இ)ம.கலர்','kg'], ['சி.கலர்','kg'],
  ['பச்சை','kg'], ['ஆயில்','confirm'], ['மிளகாய் தூள்','kg'], ['கறிவேப்பிலை','kg'],
  ['மிளகு திப்பை','kg'], ['மிளகு','kg'], ['டால்டா','L'], ['முறுக்கு மாவு கம்பி','kg'],
  ['முறுக்கு மாவு முள்ளு','kg'], ['பெருங்காயம்','kg'], ['சவ்வரிசி மாவு','g'],
  ['பொ.கடலை மாவு','kg'], ['உளுந்து மாவு','kg'], ['கார்ன் சிப்ஸ்','box'],
  ['வீல் சிப்ஸ்','kg'], ['Onion Ring','kg'], ['Curry Leaf','kg']
];
const UNITS = ['kg','g','L','mL','bag','box','piece','lot','pack'];
const ZONE = 'Asia/Kuala_Lumpur';

function doGet(event) {
  const report = event && event.parameter && event.parameter.page === 'report';
  const page = HtmlService.createHtmlOutput(PAGE_HTML);
  if (report) page.setContent(page.getContent().replace('<body>', '<body data-page="report">'));
  return page.setTitle(report ? 'M- Snacks Report' : 'M- Snacks')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

// JSON API for the GitHub Pages front end (index.html / report.html).
// The front end POSTs {method, args} as text/plain so the browser sends no CORS preflight.
const API_METHODS = {loadProduction, saveProduction, productionReport, dayTotals};
function doPost(event) {
  let result;
  try {
    const body = JSON.parse(event && event.postData ? event.postData.contents : '{}');
    const handler = API_METHODS[body.method];
    if (!handler || !Object.prototype.hasOwnProperty.call(API_METHODS, body.method)) throw Error('Unknown request.');
    result = {ok:true, data:handler.apply(null, Array.isArray(body.args) ? body.args : [])};
  } catch (error) {
    result = {ok:false, error:String(error && error.message || error)};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// Visible in the Apps Script editor. The web app runs as the deploying owner.
function setup() {
  const active = Session.getActiveUser().getEmail();
  const effective = Session.getEffectiveUser().getEmail();
  if (!active || active.toLowerCase() !== effective.toLowerCase())
    throw Error('Run setup from the Apps Script editor using the deployment owner account.');
  setup_();
}

// Keep provisioning and generated codes on the server.
function setup_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    let id = props.getProperty('SHEET_ID');
    if (!id) {
      id = SpreadsheetApp.create('M- Snacks Production').getId();
      props.setProperty('SHEET_ID',id);
    }
    const book = SpreadsheetApp.openById(id);
    book.setSpreadsheetTimeZone(ZONE);
    const log = ensureTab_(book,'Daily production',[
      'Submission ID','Production date','Submitted at','Role','Total output kg','Raw material usage','Finished products','Import products'
    ]);
    if (!log.getRange(1,8).getValue()) log.getRange(1,8).setValue('Import products').setBackground('#153b32').setFontColor('#ffffff').setFontWeight('bold');
    ensureTab_(book,EXTRA_PRODUCTS_TAB,['Product','Unit']);
    const catalog = ensureTab_(book,'Default materials',['Material','Unit']);
    const history = ensureTab_(book,'WhatsApp production history',[
      'Production date','Product','Input quantity','Input unit','Output kg','WhatsApp source line','Note'
    ]);
    if (catalog.getLastRow() === 1) catalog.getRange(2,1,MATERIALS.length,2).setValues(MATERIALS);
    if (history.getLastRow() === 1) {
      history.getRange(2,1,HISTORY.length,1).setNumberFormat('@');
      history.getRange(2,1,HISTORY.length,HISTORY[0].length).setValues(HISTORY);
    }
    const starter = book.getSheetByName('Sheet1');
    if (starter && !starter.getLastRow()) book.deleteSheet(starter);
    log.setColumnWidth(6,340);
    log.setColumnWidth(7,340);
    log.setColumnWidth(8,300);
    upgradeCatalog_(book);
    props.setProperty('CATALOG_V3','1');
    backfillCatalog_(book,recent_(book));
    props.setProperty('CATALOG_V2','1');
    syncQuantities_(book,recent_(book));
    const settings = ensureTab_(book,'Settings',['Setting','URL']);
    const deploymentUrl = ScriptApp.getService().getUrl();
    if (deploymentUrl) {
      settings.getRange(2,1,2,2).setValues([
        ['Production form',deploymentUrl],['Production report',deploymentUrl+'?page=report']
      ]);
    }
    SpreadsheetApp.flush();
    console.log('Private backend: '+book.getUrl());
  } finally {
    lock.releaseLock();
  }
}

function ensureTab_(book,name,headers) {
  const tab = book.getSheetByName(name) || book.insertSheet(name);
  if (!tab.getLastRow()) {
    tab.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#153b32').setFontColor('#ffffff').setFontWeight('bold');
    tab.setFrozenRows(1);
    tab.setColumnWidths(1,headers.length,145);
  }
  return tab;
}

function book_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (!id) throw Error('The owner needs to run setup first.');
  const book = SpreadsheetApp.openById(id);
  // One-time upgrade: copy materials used in old entries into the Default materials tab.
  if (props.getProperty('CATALOG_V2') !== '1') {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (props.getProperty('CATALOG_V2') !== '1') {
        backfillCatalog_(book,recent_(book));
        SpreadsheetApp.flush();
        props.setProperty('CATALOG_V2','1');
      }
    } finally { lock.releaseLock(); }
  }
  // One-time upgrade (19-Nov-2026): add Curry Leaf as a raw material, fix the unit of பொ.கடலை மாவு,
  // and add the Import products column / Added products tab.
  if (props.getProperty('CATALOG_V3') !== '1') {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (props.getProperty('CATALOG_V3') !== '1') {
        upgradeCatalog_(book);
        SpreadsheetApp.flush();
        props.setProperty('CATALOG_V3','1');
      }
    } finally { lock.releaseLock(); }
  }
  return book;
}

function upgradeCatalog_(book) {
  const tab = book.getSheetByName(CATALOG_TAB);
  const rows = catalog_(book);
  const index = new Map(rows.map((row,i)=>[materialKey_(row[0]),i]));
  for (const [name,unit] of FIXED_UNITS) {
    const i = index.get(materialKey_(name));
    if (i !== undefined && rows[i][1] === 'confirm') tab.getRange(i+2,2).setValue(unit);
  }
  for (const [name,unit] of ADDED_MATERIALS)
    if (!index.has(materialKey_(name))) { tab.appendRow([name,unit]); index.set(materialKey_(name),rows.length); rows.push([name,unit]); }
  const log = book.getSheetByName(LOG_TAB);
  if (log && !log.getRange(1,8).getValue())
    log.getRange(1,8).setValue('Import products').setBackground('#153b32').setFontColor('#ffffff').setFontWeight('bold');
  ensureTab_(book,EXTRA_PRODUCTS_TAB,['Product','Unit']);
}

// Products on the form = PRODUCTS + products added with "+ ADD NEW PRODUCTION" (Added products tab).
function products_(book) {
  const tab = book.getSheetByName(EXTRA_PRODUCTS_TAB);
  const extra = tab ? rows_(tab,2).filter(row=>row[0]).map(([name,unit])=>[String(name),String(unit||PRODUCT_UNIT)]) : [];
  const seen = new Set(PRODUCTS.map(([name])=>materialKey_(name)));
  return [...PRODUCTS,...extra.filter(([name])=>!seen.has(materialKey_(name)) && seen.add(materialKey_(name)))];
}

// Columns of the Daily quantities tab: current products, then old ones (so history keeps its columns).
function sheetProducts_(book) {
  const list = products_(book);
  const seen = new Set(list.map(([name])=>name));
  return [...list,...OLD_PRODUCTS.filter(([name])=>!seen.has(name))];
}

// Raw material list shown on the form (hidden materials left out).
function formMaterials_(catalog) {
  const hidden = new Set(HIDDEN_MATERIALS.map(materialKey_));
  return catalog.filter(([name])=>!hidden.has(materialKey_(name)));
}

function json_(text) {
  if (!text) return [];
  try { const value = JSON.parse(text); return Array.isArray(value) ? value : []; } catch (e) { return []; }
}

function rows_(tab,width,limit) {
  const last = tab.getLastRow();
  if (last < 2) return [];
  const start = limit ? Math.max(2,last-limit+1) : 2;
  return tab.getRange(start,1,last-start+1,width).getValues();
}

function recent_(book,limit) {
  return rows_(book.getSheetByName('Daily production'),LOG_WIDTH,limit);
}

function materialKey_(name) {
  return name.normalize('NFC').trim().replace(/\s+/gu,' ').toLowerCase();
}

const LOG_TAB = 'Daily production';
const CATALOG_TAB = 'Default materials';
const RECENT_ROWS = 400; // rows scanned for duplicate saves and the day's running totals

// Material list = the Default materials tab. New materials and confirmed units are written
// back to that tab on save, so loading the form never has to parse the whole history.
function catalog_(book) {
  return rows_(book.getSheetByName(CATALOG_TAB),2).filter(row=>row[0]).map(([name,unit])=>[String(name),String(unit)]);
}

// One-time backfill (setup): add any material used in old entries to the Default materials tab.
function backfillCatalog_(book,records) {
  const tab = book.getSheetByName(CATALOG_TAB);
  const rows = catalog_(book);
  const index = new Map(rows.map((row,i)=>[materialKey_(row[0]),i]));
  for (const row of records) {
    for (const {name,unit} of JSON.parse(row[5])) {
      const key = materialKey_(name);
      if (!index.has(key)) { index.set(key,rows.length); rows.push([name,unit]); }
      else if (rows[index.get(key)][1] === 'confirm') rows[index.get(key)][1] = unit;
    }
  }
  if (rows.length) tab.getRange(2,1,rows.length,2).setValues(rows);
}

function quantityHeaders_(columns,products) {
  return ['Submission ID','Production date','Submitted at','Role','Total output (kg)',
    ...columns.map(([name,unit])=>'Used · '+name+' ('+(unit==='confirm'?'unit pending':unit)+')'),
    ...products.flatMap(([name,unit])=>['No. of Lot · '+name+' ('+unit+')','Produced · '+name+' (kg)'])];
}

function quantityRow_(row,raw,finished,columns,products) {
  const usage = new Map(raw.map(item=>[JSON.stringify([item.name,item.unit]),item.qty]));
  const output = new Map(finished.map(item=>[item.name,item]));
  return [...row.slice(0,5),...columns.map(column=>usage.get(JSON.stringify(column)) ?? ''),
    ...products.flatMap(([name])=>{
      const item = output.get(name);
      return [item?.input ?? '',item?.output ?? ''];
    })];
}

// Full rebuild of the Daily quantities tab from the production log (setup, or when columns change).
function syncQuantities_(book,records) {
  const materials = new Map();
  for (const [name,unit] of catalog_(book)) materials.set(JSON.stringify([name,unit]),[name,unit]);
  const entries = records.map(row=>{
    const raw = JSON.parse(row[5]);
    for (const {name,unit} of raw) materials.set(JSON.stringify([name,unit]),[name,unit]);
    return {row,raw,finished:JSON.parse(row[6])};
  });
  const columns = [...materials.values()];
  const products = sheetProducts_(book);
  const headers = quantityHeaders_(columns,products);
  const tab = book.getSheetByName('Daily quantities') || book.insertSheet('Daily quantities');
  const previousRows=tab.getLastRow();
  const previousColumns=tab.getLastColumn();
  const values = entries.map(({row,raw,finished})=>quantityRow_(row,raw,finished,columns,products));
  if (headers.length > tab.getMaxColumns())
    tab.insertColumnsAfter(tab.getMaxColumns(),headers.length-tab.getMaxColumns());
  if (records.length+1 > tab.getMaxRows())
    tab.insertRowsAfter(tab.getMaxRows(),records.length+1-tab.getMaxRows());
  if (values.length) tab.getRange(2,1,values.length,4).setNumberFormat('@');
  tab.getRange(1,1,values.length+1,headers.length).setValues([headers,...values]);
  tab.getRange(1,1,1,headers.length).setBackground('#153b32').setFontColor('#ffffff')
    .setFontWeight('bold').setWrap(true);
  tab.setFrozenRows(1);
  if (previousRows > values.length+1)
    tab.getRange(values.length+2,1,previousRows-values.length-1,previousColumns).clearContent();
  if (previousColumns > headers.length)
    tab.getRange(1,headers.length+1,values.length+1,previousColumns-headers.length).clearContent();
}

// Fast path on save: append one row when the column layout is unchanged; otherwise rebuild.
function appendQuantities_(book,cells,raw,finished) {
  const tab = book.getSheetByName('Daily quantities');
  const columns = catalog_(book);
  const products = sheetProducts_(book);
  const headers = quantityHeaders_(columns,products);
  const last = tab ? tab.getLastRow() : 0;
  const same = tab && last >= 1 && tab.getLastColumn() === headers.length &&
    tab.getRange(1,1,1,headers.length).getValues()[0].every((header,i)=>header===headers[i]);
  if (!same) { syncQuantities_(book,recent_(book)); return; }
  tab.getRange(last+1,1,1,4).setNumberFormat('@');
  tab.getRange(last+1,1,1,headers.length).setValues([quantityRow_(cells,raw,finished,columns,products)]);
}

function validDate_(date) {
  return typeof date==='string' && /^\d{4}-\d{2}-\d{2}$/.test(date) &&
    Number.isFinite(Date.parse(date)) && new Date(date).toISOString().slice(0,10)===date;
}

function today_() {
  return Utilities.formatDate(new Date(),ZONE,'yyyy-MM-dd');
}

// Running totals for one production date: every save for that date is added together.
function sumDay_(rows,date) {
  const raw = new Map(), finished = new Map(), imported = new Map();
  let count = 0;
  for (const row of rows) {
    if (String(row[1]) !== date) continue;
    count++;
    for (const item of JSON.parse(row[5])) {
      const key = JSON.stringify([item.name,item.unit]);
      const total = raw.get(key) || {name:item.name,unit:item.unit,qty:0};
      total.qty += item.qty;
      raw.set(key,total);
    }
    for (const item of JSON.parse(row[6])) {
      const total = finished.get(item.name) || {name:item.name,unit:item.unit,input:0,output:0};
      total.output += item.output;
      if (item.input != null) total.input += item.input;
      finished.set(item.name,total);
    }
    for (const item of json_(row[7])) {
      const key = materialKey_(item.name);
      const total = imported.get(key) || {name:item.name,b80:0,b330:0};
      total.b80 += item.b80 || 0;
      total.b330 += item.b330 || 0;
      imported.set(key,total);
    }
  }
  return {date,count,raw:[...raw.values()],finished:[...finished.values()],imported:[...imported.values()]};
}

function dayTotals(date) {
  if (!validDate_(date)) throw Error('Choose a valid production date.');
  const book = book_();
  return sumDay_(recent_(book,RECENT_ROWS),date);
}

function loadProduction(date) {
  const book = book_();
  const today = today_();
  const day = validDate_(date) && date <= today ? date : today;
  return {products:products_(book),units:UNITS,materials:formMaterials_(catalog_(book)),today,
    day:sumDay_(recent_(book,RECENT_ROWS),day)};
}

function quantity_(value,label) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > 1000000)
    throw Error('Check the quantity for '+label+'.');
  return value;
}

function saveProduction(entry) {
  const role = 'production';
  if (!entry || typeof entry.id !== 'string' || !/^[a-f0-9-]{36}$/.test(entry.id))
    throw Error('Invalid submission.');
  if (!validDate_(entry.date) || entry.date > today_()) throw Error('Choose a valid production date.');
  if (!Array.isArray(entry.raw) || entry.raw.length > 200 ||
      !Array.isArray(entry.finished) || entry.finished.length > 200 ||
      (entry.imported != null && (!Array.isArray(entry.imported) || entry.imported.length > 100))) throw Error('Check the entry.');
  const book = book_();                           // opened before the lock (book_ may take it once to migrate)
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const recent = recent_(book,RECENT_ROWS);
    const previous = recent.find(row=>row[0]===entry.id);
    const catalog = catalog_(book);
    if (previous) {
      if (previous[3] !== role) throw Error('Invalid submission.');
      return {duplicate:true,materials:formMaterials_(catalog),products:products_(book),day:sumDay_(recent,entry.date)};
    }
    const existing = new Map(catalog.map((row,i)=>[materialKey_(row[0]),i]));
    const catalogTab = book.getSheetByName(CATALOG_TAB);
    const seen = new Set();
    const raw = entry.raw.map(item=>{
      if (!item || typeof item.name !== 'string' || item.name.length > 100) throw Error('Check the material name.');
      const key = materialKey_(item.name);
      const index = existing.get(key);
      const match = index === undefined ? null : catalog[index];
      if (!key || seen.has(key) || !UNITS.includes(item.unit) ||
          (match && match[1] !== 'confirm' && match[1] !== item.unit)) throw Error('Check the name and unit for '+item.name+'.');
      seen.add(key);
      const name = match ? match[0] : item.name.normalize('NFC').trim().replace(/\s+/gu,' ');
      quantity_(item.qty,item.name);
      if (!match) {                               // new material: add it to the list
        catalog.push([name,item.unit]);
        existing.set(key,catalog.length-1);
        catalogTab.appendRow([name,item.unit]);
      } else if (match[1] === 'confirm') {        // first use fixes the unit
        match[1] = item.unit;
        catalogTab.getRange(index+2,2).setValue(item.unit);
      }
      return {name,unit:item.unit,qty:item.qty};
    });
    const productList = products_(book);
    const productIndex = new Map(productList.map(([name,unit])=>[materialKey_(name),[name,unit]]));
    const productTab = book.getSheetByName(EXTRA_PRODUCTS_TAB);
    seen.clear();
    const finished = entry.finished.map(item=>{
      if (!item || typeof item.name !== 'string' || item.name.length > 100) throw Error('Check the finished products.');
      const key = materialKey_(item.name);
      if (!key || seen.has(key)) throw Error('Check the finished products.');
      seen.add(key);
      let match = productIndex.get(key);
      if (!match) {                               // + ADD NEW PRODUCTION: keep the new product on the form
        match = [item.name.normalize('NFC').trim().replace(/\s+/gu,' '),PRODUCT_UNIT];
        productIndex.set(key,match);
        productTab.appendRow(match);
      }
      return {name:match[0],output:quantity_(item.output,match[0]),
        input:item.input===null ? null : quantity_(item.input,match[0]+' no. of lot'),unit:match[1]};
    });
    seen.clear();
    const imported = (entry.imported || []).map(item=>{
      if (!item || typeof item.name !== 'string' || item.name.length > 100) throw Error('Check the import product name.');
      const name = item.name.normalize('NFC').trim().replace(/\s+/gu,' ');
      const key = materialKey_(name);
      if (!key || seen.has(key)) throw Error('Check the import product '+item.name+'.');
      seen.add(key);
      const b80 = item.b80==null ? 0 : quantity_(item.b80,name+' bundle of 80kg');
      const b330 = item.b330==null ? 0 : quantity_(item.b330,name+' bundle of 330kg');
      if (!b80 && !b330) throw Error('Enter the bundles for '+name+'.');
      return {name,b80,b330};
    });
    if (!raw.length && !finished.length && !imported.length) throw Error('Enter at least one quantity.');
    const cells = [entry.id,entry.date,new Date().toISOString(),role,finished.reduce((sum,item)=>sum+item.output,0),
      JSON.stringify(raw),JSON.stringify(finished),JSON.stringify(imported)];
    const tab = book.getSheetByName(LOG_TAB);
    const next = tab.getLastRow()+1;
    tab.getRange(next,1,1,4).setNumberFormat('@');
    tab.getRange(next,1,1,cells.length).setValues([cells]);
    appendQuantities_(book,cells,raw,finished);
    SpreadsheetApp.flush();
    recent.push(cells);
    return {duplicate:false,materials:formMaterials_(catalog),products:products_(book),day:sumDay_(recent,entry.date)};
  } finally {
    lock.releaseLock();
  }
}

function productionReport(from,to=from) {
  if (!validDate_(from) || !validDate_(to) || from>to) throw Error('Choose a valid production date range.');
  const book = book_();
  const history = rows_(book.getSheetByName('WhatsApp production history'),7);
  const baselines = products_(book).map(([name,unit])=>{
    const matches = history.filter(([,product,input,inputUnit,output])=>
      product===name && inputUnit===unit && typeof input==='number' && input>0 && typeof output==='number' && output>=0);
    const input = matches.reduce((sum,row)=>sum+row[2],0);
    const output = matches.reduce((sum,row)=>sum+row[4],0);
    return {name,unit,count:matches.length,input,output,rate:input ? output/input : null};
  });
  const rates = new Map(baselines.map(item=>[item.name,item.rate]));
  const reportRows = recent_(book).filter(row=>row[1]>=from&&row[1]<=to)
    .sort((a,b)=>String(a[1]).localeCompare(String(b[1]))||String(a[2]).localeCompare(String(b[2])));
  const records = reportRows.reverse().map(([,date,submitted,role,total,raw,finished,imported])=>({
    date,submitted,role,total,raw:JSON.parse(raw),imported:json_(imported),finished:JSON.parse(finished).map(item=>{
      const rate = rates.get(item.name);
      const expected = item.input > 0 && rate != null ? item.input * rate : null;
      return {...item,expected,variance:expected===null ? null : item.output-expected};
    })
  }));
  return {baselines,records};
}

// Source: WhatsApp Chat with STAR PRODUCTION.txt. Source lines are one-based.
const HISTORY = [
  [
    "2026-10-07",
    "Plain Murukku",
    6,
    "lot",
    57,
    110,
    ""
  ],
  [
    "2026-10-08",
    "Plain Murukku",
    20,
    "lot",
    200,
    172,
    ""
  ],
  [
    "2026-10-09",
    "Plain Murukku",
    30,
    "lot",
    300,
    230,
    ""
  ],
  [
    "2026-10-10",
    "Plain Murukku",
    11,
    "lot",
    109,
    293,
    ""
  ],
  [
    "2026-10-10",
    "White Mullu Murukku",
    19,
    "lot",
    187,
    294,
    ""
  ],
  [
    "2026-10-11",
    "Red Mullu Murukku",
    14,
    "lot",
    139,
    358,
    ""
  ],
  [
    "2026-10-11",
    "Wheel Chips",
    60,
    "kg",
    89,
    359,
    ""
  ],
  [
    "2026-10-11",
    "Onion Ring",
    40,
    "kg",
    55,
    360,
    ""
  ],
  [
    "2026-10-11",
    "Star",
    6,
    "lot",
    82,
    361,
    ""
  ],
  [
    "2026-10-11",
    "Omam",
    6,
    "lot",
    78,
    362,
    ""
  ],
  [
    "2026-10-12",
    "Seeval",
    20,
    "lot",
    260,
    374,
    ""
  ],
  [
    "2026-10-12",
    "Corn Chips",
    3,
    "box",
    90,
    375,
    ""
  ],
  [
    "2026-10-13",
    "Seeval",
    14,
    "lot",
    179,
    438,
    ""
  ],
  [
    "2026-10-13",
    "Corn Chips",
    6,
    "box",
    185,
    439,
    ""
  ],
  [
    "2026-10-14",
    "Plain Murukku",
    25,
    "lot",
    248,
    506,
    ""
  ],
  [
    "2026-10-15",
    "Plain Murukku",
    24,
    "lot",
    237,
    570,
    ""
  ],
  [
    "2026-10-15",
    "White Mullu Murukku",
    3,
    "lot",
    29,
    571,
    ""
  ],
  [
    "2026-10-16",
    "White Mullu Murukku",
    27,
    "lot",
    268,
    635,
    ""
  ],
  [
    "2026-10-17",
    "Red Mullu Murukku",
    27,
    "lot",
    265,
    698,
    ""
  ],
  [
    "2026-10-17",
    "Wheel Chips",
    60,
    "kg",
    83,
    699,
    ""
  ],
  [
    "2026-10-17",
    "Onion Ring",
    40,
    "kg",
    52,
    700,
    ""
  ],
  [
    "2026-10-19",
    "Curry Murukku",
    4,
    "lot",
    54,
    759,
    ""
  ],
  [
    "2026-10-19",
    "Seeval",
    20,
    "lot",
    259,
    760,
    ""
  ],
  [
    "2026-10-19",
    "Corn Chips",
    5,
    "box",
    150,
    761,
    ""
  ],
  [
    "2026-10-19",
    "Star",
    4,
    "lot",
    50,
    762,
    ""
  ],
  [
    "2026-10-19",
    "Omam",
    4,
    "lot",
    60,
    763,
    ""
  ],
  [
    "2026-10-21",
    "Plain Murukku",
    4,
    "lot",
    35,
    832,
    ""
  ],
  [
    "2026-10-21",
    "White Mullu Murukku",
    9,
    "lot",
    87,
    833,
    ""
  ],
  [
    "2026-10-22",
    "Plain Murukku",
    26,
    "lot",
    253,
    895,
    ""
  ],
  [
    "2026-10-23",
    "Plain Murukku",
    12,
    "lot",
    115,
    954,
    ""
  ],
  [
    "2026-10-23",
    "White Mullu Murukku",
    18,
    "lot",
    175,
    955,
    ""
  ],
  [
    "2026-10-24",
    "White Mullu Murukku",
    18,
    "lot",
    178,
    1014,
    ""
  ],
  [
    "2026-10-24",
    "Red Mullu Murukku",
    14,
    "lot",
    135,
    1015,
    ""
  ],
  [
    "2026-10-25",
    "Red Mullu Murukku",
    28,
    "lot",
    276,
    1057,
    ""
  ],
  [
    "2026-10-26",
    "Wheel Chips",
    40,
    "kg",
    55,
    1116,
    ""
  ],
  [
    "2026-10-26",
    "Onion Ring",
    40,
    "kg",
    55,
    1117,
    ""
  ],
  [
    "2026-10-26",
    "Plain Murukku",
    2,
    "lot",
    14,
    1118,
    ""
  ],
  [
    "2026-10-27",
    "White Mullu Murukku",
    8,
    "lot",
    78,
    1184,
    ""
  ],
  [
    "2026-10-27",
    "Corn Chips",
    7,
    "box",
    210,
    1185,
    ""
  ],
  [
    "2026-10-28",
    "Plain Murukku",
    4,
    "lot",
    35,
    1197,
    ""
  ],
  [
    "2026-10-28",
    "Seeval",
    18,
    "lot",
    230,
    1198,
    ""
  ],
  [
    "2026-10-28",
    "Corn Chips",
    3,
    "box",
    90,
    1199,
    ""
  ],
  [
    "2026-10-29",
    "Plain Murukku",
    16,
    "lot",
    154,
    1265,
    ""
  ],
  [
    "2026-10-30",
    "Plain Murukku",
    23,
    "lot",
    217,
    1326,
    ""
  ],
  [
    "2026-10-31",
    "Plain Murukku",
    4,
    "lot",
    36,
    1387,
    ""
  ],
  [
    "2026-10-31",
    "White Mullu Murukku",
    16,
    "lot",
    156,
    1388,
    ""
  ],
  [
    "2026-10-31",
    "Red Mullu Murukku",
    12,
    "lot",
    115,
    1389,
    "Output kg inferred from the report total"
  ],
  [
    "2026-11-02",
    "Plain Murukku",
    4,
    "lot",
    36,
    1464,
    ""
  ],
  [
    "2026-11-03",
    "Wheel Chips",
    45,
    "kg",
    93,
    1527,
    ""
  ],
  [
    "2026-11-03",
    "Onion Ring",
    40,
    "kg",
    55,
    1528,
    ""
  ],
  [
    "2026-11-03",
    "Plain Murukku",
    24,
    "lot",
    229,
    1529,
    ""
  ],
  [
    "2026-11-04",
    "Plain Murukku",
    6,
    "lot",
    57,
    1595,
    ""
  ],
  [
    "2026-11-04",
    "Corn Chips",
    10,
    "box",
    300,
    1596,
    ""
  ],
  [
    "2026-11-05",
    "Seeval",
    30,
    "lot",
    371,
    1661,
    ""
  ],
  [
    "2026-11-05",
    "Star",
    4,
    "lot",
    50,
    1662,
    ""
  ],
  [
    "2026-11-05",
    "Omam",
    4,
    "lot",
    57,
    1663,
    ""
  ],
  [
    "2026-11-10",
    "Plain Murukku",
    20,
    "lot",
    197,
    1747,
    ""
  ],
  [
    "2026-11-10",
    "Corn Chips",
    6,
    "box",
    178,
    1748,
    ""
  ],
  [
    "2026-11-11",
    "Plain Murukku",
    24,
    "lot",
    233,
    1810,
    ""
  ],
  [
    "2026-11-12",
    "Curry Leaf",
    20,
    "kg",
    7.6,
    1884,
    ""
  ],
  [
    "2026-11-13",
    "Plain Murukku",
    22,
    "lot",
    213,
    1937,
    ""
  ],
  [
    "2026-11-13",
    "White Mullu Murukku",
    14,
    "lot",
    137,
    1938,
    ""
  ],
  [
    "2026-11-14",
    "Red Mullu Murukku",
    8,
    "lot",
    78,
    2005,
    ""
  ],
  [
    "2026-11-14",
    "Wheel Chips",
    120,
    "kg",
    154,
    2006,
    ""
  ],
  [
    "2026-11-14",
    "Onion Ring",
    40,
    "kg",
    55,
    2007,
    ""
  ],
  [
    "2026-11-15",
    "Seeval",
    14,
    "lot",
    174,
    2074,
    ""
  ],
  [
    "2026-11-15",
    "Star",
    4,
    "lot",
    53,
    2075,
    ""
  ],
  [
    "2026-11-15",
    "Omam",
    4,
    "lot",
    56,
    2076,
    ""
  ],
  [
    "2026-11-16",
    "Curry Murukku",
    8,
    "lot",
    105,
    2139,
    ""
  ],
  [
    "2026-11-16",
    "Seeval",
    14,
    "lot",
    173,
    2140,
    ""
  ],
  [
    "2026-11-16",
    "Corn Chips",
    2,
    "box",
    59,
    2141,
    ""
  ],
  [
    "2026-11-17",
    "Red Mullu Murukku",
    1,
    "lot",
    8.5,
    2207,
    ""
  ],
  [
    "2026-11-17",
    "White Mullu Murukku",
    4,
    "lot",
    36,
    2208,
    ""
  ]
];

const PAGE_HTML = "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>M- Snacks</title>\n<style>\n:root{font:16px/1.5 system-ui,sans-serif;color:#192d25;background:#f3f5ef;--green:#174c38;--line:#d9e2d8}\n*{box-sizing:border-box}body{margin:0}header{background:#153b32;color:white;padding:18px max(20px,calc((100% - 1000px)/2));display:flex;align-items:center;justify-content:space-between;gap:15px}\nheader strong{font-size:1.45rem;white-space:nowrap}main{max-width:1040px;margin:auto;padding:28px 20px 56px}\nh1,h2,h3{line-height:1.2;margin:0 0 12px}h1{font-size:clamp(1.85rem,4vw,2.65rem);letter-spacing:-.035em}h2{font-size:1.25rem}\np{margin:0 0 16px;color:#496054}.card{background:#fff;border:1px solid var(--line);border-radius:16px;padding:clamp(16px,3vw,26px);margin-bottom:18px}\n.field{display:grid;gap:6px}input,select,button{font:inherit}input,select{padding:9px 10px;border:1px solid #9aafa1;border-radius:8px;background:#fff;color:#192d25;min-width:0;width:100%}\ninput:focus-visible,select:focus-visible,button:focus-visible,summary:focus-visible{outline:3px solid #7baf75;outline-offset:2px}button{padding:9px 15px;border-radius:8px;border:1px solid var(--line);background:white;color:var(--green);cursor:pointer;font-weight:650}\nbutton.primary{background:var(--green);color:white;border-color:var(--green)}button:disabled{opacity:.6;cursor:wait}.actions{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-top:14px}\n.split{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;align-items:start}.split>.card{margin:0}\n.row{display:grid;grid-template-columns:minmax(0,1fr) 122px;gap:10px;align-items:center;border-top:1px solid var(--line);padding:9px 0}\n.product{grid-template-columns:minmax(0,1fr) repeat(2,minmax(78px,100px))}.caption{font-size:.8rem;color:#53665a;display:block}.coltitle{font-size:.78rem;font-weight:650;color:#53665a}\n.custom{display:grid;grid-template-columns:minmax(0,1fr) 105px 85px auto;gap:6px;padding:7px 0}\n.top{display:flex;flex-wrap:wrap;align-items:end;justify-content:space-between;gap:12px;margin-bottom:18px}\n.notice{padding:9px 12px;background:#e7f4e7;border-radius:8px}.error{color:#842d20;background:#fce9e2}.scroll{overflow:auto}\ntable{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}td,th{text-align:left;border-bottom:1px solid var(--line);padding:8px;white-space:nowrap}th{font-size:.85rem;color:#38574a}\ndetails{border-top:1px solid var(--line);padding:12px 0}summary{cursor:pointer;font-weight:650}details p{margin:10px 0 4px}[hidden]{display:none!important}\n@media(max-width:760px){.split{grid-template-columns:1fr}.split>.card:first-child{grid-row:2}.product{grid-template-columns:minmax(0,1fr) repeat(2,82px)}}\n.report-date{max-width:230px;margin-bottom:14px}#usageTotals{margin-bottom:20px}#settingsStatus{margin-top:12px}\n.yield-row{border-top:1px solid var(--line);padding:8px 0}.yield-line{display:grid;grid-template-columns:75px minmax(60px,1fr) 90px;align-items:center;gap:8px;font-size:.875rem}.yield-line progress{width:100%;height:.8rem;accent-color:#17533a}.yield-line progress.estimated{accent-color:#997032}#reportFilter .actions{margin-top:0;align-items:end}#finishedTotals{margin-bottom:18px}\n\n.sofar{display:block;font-size:.8rem;font-weight:650;color:#17533a}.sofar:empty{display:none}\n.has-total{background:#f3f9f2}#dayStatus{margin-bottom:14px}\n.stack{display:grid;gap:18px}.stack>.card{margin:0}\n.import{grid-template-columns:minmax(0,1fr) repeat(2,minmax(70px,92px)) 34px}.import button{padding:6px 0}\n.newproduct input:first-child{font-weight:650}\n@media(max-width:760px){.split>.stack{grid-row:1}.import{grid-template-columns:minmax(0,1fr) repeat(2,58px) 30px;gap:6px}}\n.newname{display:grid;grid-template-columns:minmax(0,1fr) 34px;gap:6px}.newname button{padding:6px 0}#newProducts .row{border-top:1px dashed var(--line)}#addProduct{margin-top:8px}\n</style>\n</head>\n<body>\n<header><strong>M- Snacks</strong><button id=\"settingsToggle\" type=\"button\" aria-controls=\"settings\" aria-expanded=\"false\" hidden>Settings</button></header>\n<main>\n  <div id=\"startup\" class=\"card\" hidden><p id=\"startupStatus\" role=\"status\"></p><button id=\"retryStartup\" type=\"button\" hidden>Retry</button></div>\n  <form id=\"settings\" class=\"card\" hidden>\n    <h2>Google Apps Script connection</h2>\n    <label class=\"field\">Deployment URL<input id=\"deploymentUrl\" type=\"url\" value=\"https://script.google.com/macros/s/AKfycbzv9F0lr8IhHqZ2fsZVGRYQ6sde9fv0tUDNl0UM1bJpO0i7raZZxOLvxAWO1aHGv65J/exec\" placeholder=\"https://script.google.com/macros/s/…/exec\" required></label>\n    <div class=\"actions\"><button class=\"primary\">Save URL</button></div>\n    <p id=\"settingsStatus\" role=\"status\" hidden></p>\n  </form>\n  <section id=\"app\" hidden>\n    <div id=\"nav\" class=\"actions\" hidden>\n      <button id=\"entryTab\" type=\"button\" aria-pressed=\"true\">Entry</button>\n      <button id=\"reviewTab\" type=\"button\" aria-pressed=\"false\">Reports</button>\n    </div>\n    <form id=\"entry\">\n      <div class=\"top\">\n        <div><h1>Daily production</h1><p>Enter only the quantity you are adding now (morning, evening…). Each save is added to the day's total. Leave unused fields blank.</p></div>\n        <label class=\"field\">Production date<input id=\"date\" type=\"date\" required></label>\n      </div>\n      <p id=\"dayStatus\" role=\"status\" hidden></p>\n      <div class=\"split\">\n        <section class=\"card\"><h2>Raw materials used</h2>\n          <div id=\"materials\"></div><div id=\"custom\"></div>\n          <button id=\"add\" type=\"button\">+ Add raw material</button>\n        </section>\n        <div class=\"stack\">\n          <section class=\"card\"><h2>Finished products</h2>\n            <p>Total quantity of product produced from the raw material used in lots. One lot is 4kg of Finished Product.</p>\n            <div class=\"row product coltitle\"><span>Product</span><span>No. of Lot</span><span>Output kg</span></div>\n            <div id=\"products\"></div><div id=\"newProducts\"></div>\n            <button id=\"addProduct\" type=\"button\">+ ADD NEW PRODUCTION</button>\n          </section>\n          <section class=\"card\"><h2>Import products</h2>\n            <div class=\"row import coltitle\"><span>Product name</span><span>Bundle of 80kg</span><span>Bundle of 330kg</span><span></span></div>\n            <div id=\"imports\"></div>\n            <button id=\"addImport\" type=\"button\">+ Add import product</button>\n          </section>\n        </div>\n      </div>\n      <div class=\"actions\"><button class=\"primary\">Save production</button><p id=\"entryStatus\" role=\"status\" hidden></p></div>\n    </form>\n    <section id=\"review\" hidden>\n      <h1>Production reports</h1>\n      <p id=\"reviewStatus\" role=\"status\" hidden></p>\n      <form id=\"reportFilter\" class=\"card\">\n        <div class=\"actions\">\n          <label class=\"field\">From<input id=\"reportDate\" type=\"date\" required></label>\n          <label class=\"field\">To<input id=\"reportEnd\" type=\"date\" required></label>\n          <button class=\"primary\">Run report</button>\n        </div>\n      </form>\n      <section class=\"card\"><h2>Finished products</h2>\n        <div id=\"finishedTotals\" class=\"scroll\"></div>\n        <h3>Actual vs estimated output</h3>\n        <p>Chart totals use entries with a recorded product input. Estimates use the historical average yield below.</p>\n        <div id=\"yieldChart\"></div>\n      </section>\n      <section class=\"card\"><h2>Historical average yield</h2>\n        <p>Average yield is total output divided by total recorded input. Lot-based yields describe batch output. The chat gives no ingredient weights for those lots. Check the number of reports before using a yield as a benchmark.</p>\n        <div class=\"scroll\" id=\"baseline\"></div>\n      </section>\n      <section class=\"card\"><h2>Import products</h2>\n        <div id=\"importTotals\" class=\"scroll\"></div>\n      </section>\n      <section class=\"card\"><h2>Raw materials used</h2>\n        <div id=\"usageTotals\" class=\"scroll\"></div>\n        <h3>Production submissions</h3>\n        <p>Estimated output uses the matching product input and historical yield. Differences are shown in kg for review.</p>\n        <div id=\"records\"></div>\n      </section>\n    </section>\n  </section>\n</main>\n<script>\nconst $=id=>document.getElementById(id);\nconst reportOnly=document.body.dataset.page==='report';\nlet session=null, requestId='', busy=false, version=0;\nconst number=new Intl.NumberFormat('en',{maximumFractionDigits:3});\nfunction el(tag,text,cls){\n  const node=document.createElement(tag);\n  if(text!=null)node.textContent=text;\n  if(cls)node.className=cls;\n  return node;\n}\nfunction status(id,message,error=false){\n  const node=$(id);\n  node.textContent=message;\n  node.className=error?'notice error':'notice';\n  node.hidden=!message;\n}\nconst hosted=Boolean(window.google && google.script && google.script.run);\nlet apiUrl='';\nasync function rpc(method,...args){\n  if(hosted) return new Promise((resolve,reject)=>google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[method](...args));\n  if(!apiUrl) throw Error('Set the deployment URL in Settings.');\n  let response;\n  try{\n    response=await fetch(apiUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({method,args})});\n  }catch{throw Error('Could not reach the server. Check your internet connection and that the web app access is set to Anyone.')}\n  let result;\n  try{result=await response.json()}\n  catch{throw Error('The server did not return data. Redeploy the Apps Script web app with access set to Anyone.')}\n  if(!result.ok) throw Error(result.error||'Request failed.');\n  return result.data;\n}\nfunction numeric(label){\n  const input=el('input');\n  Object.assign(input,{type:'number',min:'0',max:'1000000',step:'any',inputMode:'decimal',placeholder:'—'});\n  input.setAttribute('aria-label',label);\n  return input;\n}\nfunction units(label){\n  const select=el('select');\n  select.setAttribute('aria-label',label);\n  select.append(new Option('Unit',''));\n  session.units.forEach(unit=>select.append(new Option(unit,unit)));\n  return select;\n}\nfunction materialRow(name,unit){\n  const row=el('div',null,'row');\n  row.dataset.name=name;\n  row.dataset.unit=unit;\n  const label=el('label',name);\n  const input=numeric(name+' quantity used'+(unit==='confirm'?'':' in '+unit));\n  input.id='material-'+$('materials').children.length;\n  label.htmlFor=input.id;\n  label.append(el('small',unit==='confirm'?'Choose unit when used':unit,'caption'));\n  label.append(el('small','','sofar'));\n  row.append(label,input);\n  if(unit==='confirm') row.append(units(name+' unit'));\n  return row;\n}\nfunction renderEntry(){\n  $('materials').replaceChildren();\n  session.materials.forEach(row=>$('materials').append(materialRow(...row)));\n  $('custom').replaceChildren();\n  $('products').replaceChildren(...session.products.map(([name,unit])=>{\n    const row=el('div',null,'row product');\n    row.dataset.name=name;\n    const label=el('span',name);\n    label.append(el('small',unit,'caption'));\n    label.append(el('small','','sofar'));\n    row.append(label,numeric(name+' no. of lot in '+unit),numeric(name+' output in kg'));\n    return row;\n  }));\n  $('newProducts').replaceChildren();\n  $('imports').replaceChildren();\n  addImport(false);\n  requestId=globalThis.crypto?.randomUUID?.() || 'preview';\n}\n// Show what is already saved for the chosen date. New entries are added on top of these totals.\nfunction applyDay(day){\n  session.day=day;\n  const raw=new Map((day?.raw||[]).map(r=>[r.name,r]));\n  const fin=new Map((day?.finished||[]).map(r=>[r.name,r]));\n  for(const row of $('materials').children){\n    const t=raw.get(row.dataset.name);\n    const node=row.querySelector('.sofar');\n    node.textContent=t?'Saved: '+format(t.qty)+' '+t.unit:'';\n    row.classList.toggle('has-total',Boolean(t));\n  }\n  for(const row of $('products').children){\n    const t=fin.get(row.dataset.name);\n    const node=row.querySelector('.sofar');\n    node.textContent=t?'Saved: '+(t.input?format(t.input)+' '+t.unit+' → ':'')+format(t.output)+' kg':'';\n    row.classList.toggle('has-total',Boolean(t));\n  }\n  const imp=day?.imported||[];\n  $('importSaved')?.remove();\n  if(imp.length){\n    const saved=el('p','Saved today: '+imp.map(r=>r.name+' ('+format(r.b80)+' × 80kg, '+format(r.b330)+' × 330kg)').join(', '),'sofar');\n    saved.id='importSaved';\n    $('imports').before(saved);\n  }\n  const n=day?.count||0;\n  status('dayStatus',n?n+(n===1?' entry':' entries')+' already saved for '+day.date+'. Enter only the new quantity — it is added to the total.':'');\n}\nlet dayVersion=0;\nasync function loadDay(){\n  const date=$('date').value;\n  if(!date||!session)return;\n  const current=++dayVersion;\n  status('dayStatus','Loading saved quantities for '+date+'…');\n  try{\n    const day=await rpc('dayTotals',date);\n    if(current===dayVersion) applyDay(day);\n  }catch(error){if(current===dayVersion)status('dayStatus',error.message,true)}\n}\nfunction addMaterial(){\n  const row=el('div',null,'custom');\n  const name=el('input');\n  name.placeholder='Material name'; name.maxLength=100;\n  name.setAttribute('aria-label','New material name');\n  const remove=el('button','×');\n  remove.type='button'; remove.setAttribute('aria-label','Remove new material');\n  remove.onclick=()=>row.remove();\n  row.append(name,units('New material unit'),numeric('New material quantity used'),remove);\n  $('custom').append(row);\n  name.focus();\n}\n// + ADD NEW PRODUCTION: a product that is not in the list yet (saved to the list on save).\nfunction addProduct(){\n  const row=el('div',null,'row product newproduct');\n  const name=el('input');\n  name.placeholder='New product name'; name.maxLength=100;\n  name.setAttribute('aria-label','New product name');\n  const lot=numeric('New product no. of lot'), out=numeric('New product output in kg');\n  const remove=el('button','×');\n  remove.type='button'; remove.setAttribute('aria-label','Remove new product');\n  remove.onclick=()=>row.remove();\n  const cell=el('div',null,'newname'); cell.append(name,remove);\n  row.append(cell,lot,out);\n  $('newProducts').append(row);\n  name.focus();\n}\nfunction addImport(focus=true){\n  const row=el('div',null,'row import');\n  const name=el('input');\n  name.placeholder='Product name'; name.maxLength=100;\n  name.setAttribute('aria-label','Import product name');\n  const remove=el('button','×');\n  remove.type='button'; remove.setAttribute('aria-label','Remove import product');\n  remove.onclick=()=>row.remove();\n  row.append(name,numeric('Bundle of 80kg'),numeric('Bundle of 330kg'),remove);\n  $('imports').append(row);\n  if(focus)name.focus();\n}\nfunction collect(){\n  const raw=[];\n  for(const row of $('materials').children){\n    const quantity=row.querySelector('input');\n    if(quantity.value==='')continue;\n    const {name,unit}=row.dataset;\n    const selected=unit==='confirm'?row.querySelector('select').value:unit;\n    if(!selected)throw Error('Choose the unit for '+name+'.');\n    raw.push({name,unit:selected,qty:quantity.valueAsNumber});\n  }\n  for(const row of $('custom').children){\n    const [name,unit,qty]=row.children;\n    if(!name.value.trim()&&!unit.value&&qty.value==='')continue;\n    if(!name.value.trim()||!unit.value||qty.value==='')throw Error('Complete the name, unit and quantity for each new material.');\n    raw.push({name:name.value,unit:unit.value,qty:qty.valueAsNumber});\n  }\n  const finished=[];\n  for(const row of $('products').children){\n    const [input,output]=row.querySelectorAll('input');\n    if(input.value!==''&&output.value==='')throw Error('Enter the output for '+row.dataset.name+'.');\n    if(output.value!=='') finished.push({name:row.dataset.name,output:output.valueAsNumber,input:input.value===''?null:input.valueAsNumber});\n  }\n  for(const row of $('newProducts').children){\n    const [name,input,output]=row.querySelectorAll('input');\n    if(!name.value.trim()&&input.value===''&&output.value==='')continue;\n    if(!name.value.trim()||output.value==='')throw Error('Enter the name and output kg for each new product.');\n    finished.push({name:name.value,output:output.valueAsNumber,input:input.value===''?null:input.valueAsNumber});\n  }\n  const imported=[];\n  for(const row of $('imports').children){\n    const [name,b80,b330]=row.querySelectorAll('input');\n    if(!name.value.trim()&&b80.value===''&&b330.value==='')continue;\n    if(!name.value.trim())throw Error('Enter the import product name.');\n    if(b80.value===''&&b330.value==='')throw Error('Enter the bundles for '+name.value.trim()+'.');\n    imported.push({name:name.value,b80:b80.value===''?null:b80.valueAsNumber,b330:b330.value===''?null:b330.valueAsNumber});\n  }\n  if(!raw.length&&!finished.length&&!imported.length)throw Error('Enter at least one quantity.');\n  return {id:requestId,date:$('date').value,raw,finished,imported};\n}\nfunction table(headers,rows){\n  const result=el('table');\n  const head=result.createTHead().insertRow();\n  headers.forEach(label=>{const cell=el('th',label);cell.scope='col';head.append(cell)});\n  const body=result.createTBody();\n  rows.forEach(row=>{const tr=body.insertRow();row.forEach(value=>tr.insertCell().textContent=value)});\n  return result;\n}\nfunction format(value){return value==null?'—':number.format(value)}\nfunction renderFinished(data){\n  const totals=new Map();\n  for(const record of data.records){\n    for(const item of record.finished){\n      const total=totals.get(item.name)||{name:item.name,unit:item.unit,output:0,input:0,matched:0,expected:0};\n      total.output+=item.output;\n      if(item.expected!=null){\n        total.input+=item.input;\n        total.matched+=item.output;\n        total.expected+=item.expected;\n      }\n      totals.set(item.name,total);\n    }\n  }\n  const products=[...totals.values()].sort((a,b)=>a.name.localeCompare(b.name));\n  $('finishedTotals').replaceChildren(products.length\n    ? table(['Product','Total output kg','Paired input','Average yield in range'],\n      products.map(r=>[r.name,format(r.output),r.input?format(r.input)+' '+r.unit:'—',\n        r.input?format(r.matched/r.input)+' kg / '+r.unit:'—']))\n    : el('p','No finished products recorded in this date range.'));\n  const paired=products.filter(r=>r.input>0);\n  const max=Math.max(1,...paired.flatMap(r=>[r.matched,r.expected]));\n  $('yieldChart').replaceChildren(...(paired.length?paired.map(r=>{\n    const section=el('section',null,'yield-row');\n    section.append(el('strong',r.name));\n    for(const [label,value] of [['Actual',r.matched],['Estimated',r.expected]]){\n      const line=el('div',null,'yield-line');\n      const bar=el('progress',null,label.toLowerCase());\n      bar.value=value;bar.max=max;\n      bar.setAttribute('aria-label',r.name+' '+label+' output in kg');\n      line.append(el('span',label),bar,el('span',format(value)+' kg'));\n      section.append(line);\n    }\n    return section;\n  }):[el('p','Enter product inputs on production submissions to plot a comparison.')]));\n}\nfunction renderReview(data){\n  renderFinished(data);\n  $('baseline').replaceChildren(table(['Product','Reports','Recorded input','Output kg','Average yield'],\n    data.baselines.map(r=>[r.name,r.count,format(r.input)+' '+r.unit,format(r.output),\n      r.rate==null?'—':format(r.rate)+' kg / '+r.unit])));\n  const totals=new Map();\n  for(const record of data.records){\n    for(const item of record.raw){\n      const key=JSON.stringify([item.name,item.unit]);\n      const total=totals.get(key)||{name:item.name,unit:item.unit,qty:0};\n      total.qty+=item.qty;\n      totals.set(key,total);\n    }\n  }\n  const imports=new Map();\n  for(const record of data.records){\n    for(const item of record.imported||[]){\n      const key=item.name.toLowerCase();\n      const total=imports.get(key)||{name:item.name,b80:0,b330:0};\n      total.b80+=item.b80||0; total.b330+=item.b330||0;\n      imports.set(key,total);\n    }\n  }\n  const imported=[...imports.values()].sort((a,b)=>a.name.localeCompare(b.name));\n  $('importTotals').replaceChildren(imported.length\n    ? table(['Product name','Bundle of 80kg','Bundle of 330kg','Total kg'],imported.map(r=>[r.name,format(r.b80),format(r.b330),format(r.b80*80+r.b330*330)]))\n    : el('p','No import products recorded in this date range.'));\n  const usage=[...totals.values()].sort((a,b)=>a.name.localeCompare(b.name));\n  $('usageTotals').replaceChildren(usage.length\n    ? table(['Raw material','Total used','Unit'],usage.map(r=>[r.name,format(r.qty),r.unit]))\n    : el('p','No raw material usage recorded in this date range.'));\n  const nodes=data.records.map(record=>{\n    const details=el('details');\n    details.append(el('summary',record.date+' · '+format(record.total)+' kg · '+record.role));\n    details.append(el('p','Saved '+new Date(record.submitted).toLocaleString()));\n    details.append(el('h3','Raw materials used'));\n    details.append(record.raw.length?table(['Material','Quantity','Unit'],record.raw.map(r=>[r.name,format(r.qty),r.unit])):el('p','No raw material quantities entered.'));\n    details.append(el('h3','Finished products'));\n    const scroll=el('div',null,'scroll');\n    scroll.append(record.finished.length?table(['Product','No. of Lot','Actual kg','Estimated kg','Difference kg'],\n      record.finished.map(r=>[r.name,r.input==null?'—':format(r.input)+' '+r.unit,format(r.output),format(r.expected),\n        r.variance>0?'+'+format(r.variance):format(r.variance)])):el('p','No finished products entered.'));\n    details.append(scroll);\n    if((record.imported||[]).length){\n      details.append(el('h3','Import products'));\n      details.append(table(['Product name','Bundle of 80kg','Bundle of 330kg'],record.imported.map(r=>[r.name,format(r.b80),format(r.b330)])));\n    }\n    return details;\n  });\n  $('records').replaceChildren(...(nodes.length?nodes:[el('p','No production submissions in this date range.')]));\n}\nfunction openProduction(data){\n  session=data;\n  for(const id of ['date','reportDate','reportEnd']){\n    $(id).value=session.today;\n    $(id).max=session.today;\n  }\n  renderEntry();\n  applyDay(data.day);\n  $('app').hidden=false;\n  $('nav').hidden=false;\n  if(reportOnly) $('reviewTab').click();\n}\nasync function startProduction(){\n  $('startup').hidden=false;\n  $('retryStartup').hidden=true;\n  status('startupStatus','Loading production form…');\n  try{\n    openProduction(await rpc('loadProduction',$('date').value||null));\n    $('startup').hidden=true;\n  }catch(error){\n    status('startupStatus',error.message,true);\n    $('retryStartup').hidden=false;\n  }\n}\n$('retryStartup').onclick=startProduction;\n$('entry').onsubmit=async event=>{\n  event.preventDefault();\n  if(busy)return;\n  try{\n    const data=collect();\n    busy=true;\n    for(const control of $('entry').querySelectorAll('input,select,button'))control.disabled=true;\n    status('entryStatus','Saving…');\n    const result=await rpc('saveProduction',data);\n    session.materials=result.materials;\n    if(result.products) session.products=result.products;\n    renderEntry();\n    applyDay(result.day);\n    status('entryStatus',result.duplicate?'This entry was already saved.':'Added to '+data.date+'. Totals updated.');\n  }catch(error){status('entryStatus',error.message,true)}\n  finally{\n    busy=false;\n    for(const control of $('entry').querySelectorAll('input,select,button'))control.disabled=false;\n  }\n};\n$('add').onclick=addMaterial;\n$('addProduct').onclick=addProduct;\n$('addImport').onclick=()=>addImport();\n$('date').onchange=()=>{status('entryStatus','');loadDay()};\n$('entryTab').onclick=()=>{\n  version++;\n  $('entry').hidden=false;$('review').hidden=true;\n  $('entryTab').setAttribute('aria-pressed','true');$('reviewTab').setAttribute('aria-pressed','false');\n};\n$('reviewTab').onclick=()=>{\n  $('entry').hidden=true;$('review').hidden=false;\n  $('entryTab').setAttribute('aria-pressed','false');$('reviewTab').setAttribute('aria-pressed','true');\n  loadReport();\n};\n$('reportFilter').onsubmit=event=>{event.preventDefault();loadReport()};\nfunction clearReport(){\n  for(const id of ['usageTotals','records','finishedTotals','yieldChart','baseline','importTotals']) $(id).replaceChildren();\n}\nasync function loadReport(){\n  const current=++version;\n  const from=$('reportDate'),to=$('reportEnd');\n  clearReport();\n  if(!from.checkValidity()||!to.checkValidity()||from.value>to.value){\n    status('reviewStatus','Choose a valid date range. From must be on or before To.',true);\n    return;\n  }\n  status('reviewStatus','Loading review…');\n  try{\n    const data=await rpc('productionReport',from.value,to.value);\n    if(current!==version)return;\n    renderReview(data);\n    status('reviewStatus','');\n  }catch(error){if(current===version)status('reviewStatus',error.message,true)}\n}\n\nfunction deploymentUrl(value){\n  const url=new URL(value.trim());\n  if(url.origin!=='https://script.google.com'||url.username||url.password||url.search||url.hash||\n    !/^\\/macros\\/(?:a\\/[^/]+\\/)?s\\/[\\w-]+\\/exec$/.test(url.pathname))\n    throw Error('Invalid deployment URL.');\n  return url.href;\n}\nfunction showDeployment(url){\n  $('deploymentUrl').value=url;\n  apiUrl=url;\n}\n{\n  $('settingsToggle').hidden=false;\n  $('settingsToggle').onclick=()=>{\n    $('settings').hidden=!$('settings').hidden;\n    $('settingsToggle').setAttribute('aria-expanded',String(!$('settings').hidden));\n    if(!$('settings').hidden) $('deploymentUrl').focus();\n  };\n  $('settings').onsubmit=event=>{\n    event.preventDefault();\n    let url;\n    try{url=deploymentUrl($('deploymentUrl').value)}\n    catch{status('settingsStatus','Enter the Google Apps Script web app URL ending in /exec.',true);return}\n    showDeployment(url);\n    try{\n      localStorage.setItem('m-snacks-deployment-url',url);\n      status('settingsStatus','URL saved in this browser.');\n    }catch{\n      status('settingsStatus','URL ready for this visit. This browser blocked saving the setting.');\n    }\n    if(!session) startProduction();\n  };\n  try{\n    let saved=null;\n    try{saved=localStorage.getItem('m-snacks-deployment-url')}catch{}\n    const configured=saved||$('deploymentUrl').value;\n    if(configured) showDeployment(deploymentUrl(configured));\n  }catch{}\n}\nstartProduction();\n</script>\n</body>\n</html>\n";
