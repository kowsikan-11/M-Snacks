const $=id=>document.getElementById(id);
const reportOnly=document.body.dataset.page==='report';
let session=null, requestId='', busy=false, version=0;
const number=new Intl.NumberFormat('en',{maximumFractionDigits:3});
function el(tag,text,cls){
  const node=document.createElement(tag);
  if(text!=null)node.textContent=text;
  if(cls)node.className=cls;
  return node;
}
function status(id,message,error=false){
  const node=$(id);
  node.textContent=message;
  node.className=error?'notice error':'notice';
  node.hidden=!message;
}
const hosted=Boolean(window.google && google.script && google.script.run);
let apiUrl='';
async function rpc(method,...args){
  if(hosted) return new Promise((resolve,reject)=>google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[method](...args));
  if(!apiUrl) throw Error('Set the deployment URL in Settings.');
  let response;
  try{
    response=await fetch(apiUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({method,args})});
  }catch{throw Error('Could not reach the server. Check your internet connection and that the web app access is set to Anyone.')}
  let result;
  try{result=await response.json()}
  catch{throw Error('The server did not return data. Redeploy the Apps Script web app with access set to Anyone.')}
  if(!result.ok) throw Error(result.error||'Request failed.');
  return result.data;
}
function numeric(label){
  const input=el('input');
  Object.assign(input,{type:'number',min:'0',max:'1000000',step:'any',inputMode:'decimal',placeholder:'—'});
  input.setAttribute('aria-label',label);
  return input;
}
function units(label){
  const select=el('select');
  select.setAttribute('aria-label',label);
  select.append(new Option('Unit',''));
  session.units.forEach(unit=>select.append(new Option(unit,unit)));
  return select;
}
function materialRow(name,unit){
  const row=el('div',null,'row');
  row.dataset.name=name;
  row.dataset.unit=unit;
  const label=el('label',name);
  const input=numeric(name+' quantity used'+(unit==='confirm'?'':' in '+unit));
  input.id='material-'+$('materials').children.length;
  label.htmlFor=input.id;
  label.append(el('small',unit==='confirm'?'Choose unit when used':unit,'caption'));
  label.append(el('small','','sofar'));
  row.append(label,input);
  if(unit==='confirm') row.append(units(name+' unit'));
  return row;
}
function renderEntry(){
  $('materials').replaceChildren();
  session.materials.forEach(row=>$('materials').append(materialRow(...row)));
  $('custom').replaceChildren();
  $('products').replaceChildren(...session.products.map(([name,unit])=>{
    const row=el('div',null,'row product');
    row.dataset.name=name;
    const label=el('span',name);
    label.append(el('small','Input in '+unit,'caption'));
    label.append(el('small','','sofar'));
    row.append(label,numeric(name+' input in '+unit),numeric(name+' output in kg'));
    return row;
  }));
  requestId=globalThis.crypto?.randomUUID?.() || 'preview';
}
// Show what is already saved for the chosen date. New entries are added on top of these totals.
function applyDay(day){
  session.day=day;
  const raw=new Map((day?.raw||[]).map(r=>[r.name,r]));
  const fin=new Map((day?.finished||[]).map(r=>[r.name,r]));
  for(const row of $('materials').children){
    const t=raw.get(row.dataset.name);
    const node=row.querySelector('.sofar');
    node.textContent=t?'Saved: '+format(t.qty)+' '+t.unit:'';
    row.classList.toggle('has-total',Boolean(t));
  }
  for(const row of $('products').children){
    const t=fin.get(row.dataset.name);
    const node=row.querySelector('.sofar');
    node.textContent=t?'Saved: '+(t.input?format(t.input)+' '+t.unit+' → ':'')+format(t.output)+' kg':'';
    row.classList.toggle('has-total',Boolean(t));
  }
  const n=day?.count||0;
  status('dayStatus',n?n+(n===1?' entry':' entries')+' already saved for '+day.date+'. Enter only the new quantity — it is added to the total.':'');
}
let dayVersion=0;
async function loadDay(){
  const date=$('date').value;
  if(!date||!session)return;
  const current=++dayVersion;
  status('dayStatus','Loading saved quantities for '+date+'…');
  try{
    const day=await rpc('dayTotals',date);
    if(current===dayVersion) applyDay(day);
  }catch(error){if(current===dayVersion)status('dayStatus',error.message,true)}
}
function addMaterial(){
  const row=el('div',null,'custom');
  const name=el('input');
  name.placeholder='Material name'; name.maxLength=100;
  name.setAttribute('aria-label','New material name');
  const remove=el('button','×');
  remove.type='button'; remove.setAttribute('aria-label','Remove new material');
  remove.onclick=()=>row.remove();
  row.append(name,units('New material unit'),numeric('New material quantity used'),remove);
  $('custom').append(row);
  name.focus();
}
function collect(){
  const raw=[];
  for(const row of $('materials').children){
    const quantity=row.querySelector('input');
    if(quantity.value==='')continue;
    const {name,unit}=row.dataset;
    const selected=unit==='confirm'?row.querySelector('select').value:unit;
    if(!selected)throw Error('Choose the unit for '+name+'.');
    raw.push({name,unit:selected,qty:quantity.valueAsNumber});
  }
  for(const row of $('custom').children){
    const [name,unit,qty]=row.children;
    if(!name.value.trim()&&!unit.value&&qty.value==='')continue;
    if(!name.value.trim()||!unit.value||qty.value==='')throw Error('Complete the name, unit and quantity for each new material.');
    raw.push({name:name.value,unit:unit.value,qty:qty.valueAsNumber});
  }
  const finished=[];
  for(const row of $('products').children){
    const [input,output]=row.querySelectorAll('input');
    if(input.value!==''&&output.value==='')throw Error('Enter the output for '+row.dataset.name+'.');
    if(output.value!=='') finished.push({name:row.dataset.name,output:output.valueAsNumber,input:input.value===''?null:input.valueAsNumber});
  }
  if(!raw.length&&!finished.length)throw Error('Enter at least one quantity.');
  return {id:requestId,date:$('date').value,raw,finished};
}
function table(headers,rows){
  const result=el('table');
  const head=result.createTHead().insertRow();
  headers.forEach(label=>{const cell=el('th',label);cell.scope='col';head.append(cell)});
  const body=result.createTBody();
  rows.forEach(row=>{const tr=body.insertRow();row.forEach(value=>tr.insertCell().textContent=value)});
  return result;
}
function format(value){return value==null?'—':number.format(value)}
function renderFinished(data){
  const totals=new Map();
  for(const record of data.records){
    for(const item of record.finished){
      const total=totals.get(item.name)||{name:item.name,unit:item.unit,output:0,input:0,matched:0,expected:0};
      total.output+=item.output;
      if(item.expected!=null){
        total.input+=item.input;
        total.matched+=item.output;
        total.expected+=item.expected;
      }
      totals.set(item.name,total);
    }
  }
  const products=[...totals.values()].sort((a,b)=>a.name.localeCompare(b.name));
  $('finishedTotals').replaceChildren(products.length
    ? table(['Product','Total output kg','Paired input','Average yield in range'],
      products.map(r=>[r.name,format(r.output),r.input?format(r.input)+' '+r.unit:'—',
        r.input?format(r.matched/r.input)+' kg / '+r.unit:'—']))
    : el('p','No finished products recorded in this date range.'));
  const paired=products.filter(r=>r.input>0);
  const max=Math.max(1,...paired.flatMap(r=>[r.matched,r.expected]));
  $('yieldChart').replaceChildren(...(paired.length?paired.map(r=>{
    const section=el('section',null,'yield-row');
    section.append(el('strong',r.name));
    for(const [label,value] of [['Actual',r.matched],['Estimated',r.expected]]){
      const line=el('div',null,'yield-line');
      const bar=el('progress',null,label.toLowerCase());
      bar.value=value;bar.max=max;
      bar.setAttribute('aria-label',r.name+' '+label+' output in kg');
      line.append(el('span',label),bar,el('span',format(value)+' kg'));
      section.append(line);
    }
    return section;
  }):[el('p','Enter product inputs on production submissions to plot a comparison.')]));
}
function renderReview(data){
  renderFinished(data);
  $('baseline').replaceChildren(table(['Product','Reports','Recorded input','Output kg','Average yield'],
    data.baselines.map(r=>[r.name,r.count,format(r.input)+' '+r.unit,format(r.output),
      r.rate==null?'—':format(r.rate)+' kg / '+r.unit])));
  const totals=new Map();
  for(const record of data.records){
    for(const item of record.raw){
      const key=JSON.stringify([item.name,item.unit]);
      const total=totals.get(key)||{name:item.name,unit:item.unit,qty:0};
      total.qty+=item.qty;
      totals.set(key,total);
    }
  }
  const usage=[...totals.values()].sort((a,b)=>a.name.localeCompare(b.name));
  $('usageTotals').replaceChildren(usage.length
    ? table(['Raw material','Total used','Unit'],usage.map(r=>[r.name,format(r.qty),r.unit]))
    : el('p','No raw material usage recorded in this date range.'));
  const nodes=data.records.map(record=>{
    const details=el('details');
    details.append(el('summary',record.date+' · '+format(record.total)+' kg · '+record.role));
    details.append(el('p','Saved '+new Date(record.submitted).toLocaleString()));
    details.append(el('h3','Raw materials used'));
    details.append(record.raw.length?table(['Material','Quantity','Unit'],record.raw.map(r=>[r.name,format(r.qty),r.unit])):el('p','No raw material quantities entered.'));
    details.append(el('h3','Finished products'));
    const scroll=el('div',null,'scroll');
    scroll.append(record.finished.length?table(['Product','Input','Actual kg','Estimated kg','Difference kg'],
      record.finished.map(r=>[r.name,r.input==null?'—':format(r.input)+' '+r.unit,format(r.output),format(r.expected),
        r.variance>0?'+'+format(r.variance):format(r.variance)])):el('p','No finished products entered.'));
    details.append(scroll);
    return details;
  });
  $('records').replaceChildren(...(nodes.length?nodes:[el('p','No production submissions in this date range.')]));
}
function openProduction(data){
  session=data;
  for(const id of ['date','reportDate','reportEnd']){
    $(id).value=session.today;
    $(id).max=session.today;
  }
  renderEntry();
  applyDay(data.day);
  $('app').hidden=false;
  $('nav').hidden=false;
  if(reportOnly) $('reviewTab').click();
}
async function startProduction(){
  $('startup').hidden=false;
  $('retryStartup').hidden=true;
  status('startupStatus','Loading production form…');
  try{
    openProduction(await rpc('loadProduction',$('date').value||null));
    $('startup').hidden=true;
  }catch(error){
    status('startupStatus',error.message,true);
    $('retryStartup').hidden=false;
  }
}
$('retryStartup').onclick=startProduction;
$('entry').onsubmit=async event=>{
  event.preventDefault();
  if(busy)return;
  try{
    const data=collect();
    busy=true;
    for(const control of $('entry').querySelectorAll('input,select,button'))control.disabled=true;
    status('entryStatus','Saving…');
    const result=await rpc('saveProduction',data);
    session.materials=result.materials;
    renderEntry();
    applyDay(result.day);
    status('entryStatus',result.duplicate?'This entry was already saved.':'Added to '+data.date+'. Totals updated.');
  }catch(error){status('entryStatus',error.message,true)}
  finally{
    busy=false;
    for(const control of $('entry').querySelectorAll('input,select,button'))control.disabled=false;
  }
};
$('add').onclick=addMaterial;
$('date').onchange=()=>{status('entryStatus','');loadDay()};
$('entryTab').onclick=()=>{
  version++;
  $('entry').hidden=false;$('review').hidden=true;
  $('entryTab').setAttribute('aria-pressed','true');$('reviewTab').setAttribute('aria-pressed','false');
};
$('reviewTab').onclick=()=>{
  $('entry').hidden=true;$('review').hidden=false;
  $('entryTab').setAttribute('aria-pressed','false');$('reviewTab').setAttribute('aria-pressed','true');
  loadReport();
};
$('reportFilter').onsubmit=event=>{event.preventDefault();loadReport()};
function clearReport(){
  for(const id of ['usageTotals','records','finishedTotals','yieldChart','baseline']) $(id).replaceChildren();
}
async function loadReport(){
  const current=++version;
  const from=$('reportDate'),to=$('reportEnd');
  clearReport();
  if(!from.checkValidity()||!to.checkValidity()||from.value>to.value){
    status('reviewStatus','Choose a valid date range. From must be on or before To.',true);
    return;
  }
  status('reviewStatus','Loading review…');
  try{
    const data=await rpc('productionReport',from.value,to.value);
    if(current!==version)return;
    renderReview(data);
    status('reviewStatus','');
  }catch(error){if(current===version)status('reviewStatus',error.message,true)}
}

function deploymentUrl(value){
  const url=new URL(value.trim());
  if(url.origin!=='https://script.google.com'||url.username||url.password||url.search||url.hash||
    !/^\/macros\/(?:a\/[^/]+\/)?s\/[\w-]+\/exec$/.test(url.pathname))
    throw Error('Invalid deployment URL.');
  return url.href;
}
function showDeployment(url){
  $('deploymentUrl').value=url;
  apiUrl=url;
}
{
  $('settingsToggle').hidden=false;
  $('settingsToggle').onclick=()=>{
    $('settings').hidden=!$('settings').hidden;
    $('settingsToggle').setAttribute('aria-expanded',String(!$('settings').hidden));
    if(!$('settings').hidden) $('deploymentUrl').focus();
  };
  $('settings').onsubmit=event=>{
    event.preventDefault();
    let url;
    try{url=deploymentUrl($('deploymentUrl').value)}
    catch{status('settingsStatus','Enter the Google Apps Script web app URL ending in /exec.',true);return}
    showDeployment(url);
    try{
      localStorage.setItem('m-snacks-deployment-url',url);
      status('settingsStatus','URL saved in this browser.');
    }catch{
      status('settingsStatus','URL ready for this visit. This browser blocked saving the setting.');
    }
    if(!session) startProduction();
  };
  try{
    let saved=null;
    try{saved=localStorage.getItem('m-snacks-deployment-url')}catch{}
    const configured=saved||$('deploymentUrl').value;
    if(configured) showDeployment(deploymentUrl(configured));
  }catch{}
}
startProduction();
