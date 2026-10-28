'use strict';
const $ = id => document.getElementById(id);
let workbook = null;
let report = null;
const fields = ['shop','sales_basis','as_of','production_days','shipping_days','target_days','safety_days','plan_store','coverage_mode'];
$('as_of').value = new Intl.DateTimeFormat('sv-SE', {timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());

function message(value, error=false) {
  $('message').textContent = value;
  $('message').className = error ? 'error' : '';
}
function invalidate() {
  report = null;
  $('output').hidden = true;
}
async function api(path, payload) {
  const response = await fetch(path, {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const body = await response.json();
  if (!response.ok) throw new Error(body.error || '请求失败');
  return body;
}
function table(target, rows, columns) {
  const container = $(target);
  container.replaceChildren();
  if (!rows.length) {container.textContent = '暂无记录'; return;}
  const keys = columns || [...new Set(rows.flatMap(Object.keys))];
  const element = document.createElement('table');
  const head = document.createElement('thead');
  const tr = document.createElement('tr');
  keys.forEach(k => {const th = document.createElement('th'); th.textContent=k; tr.append(th);});
  head.append(tr); element.append(head);
  const body = document.createElement('tbody');
  rows.forEach(row => {
    const tr = document.createElement('tr');
    keys.forEach(k => {const td=document.createElement('td');const v=row[k];td.textContent=v===null||v===undefined?'—':typeof v==='number'?v.toLocaleString('zh-CN',{maximumFractionDigits:2}):String(v);tr.append(td);});
    body.append(tr);
  });
  element.append(body);container.append(element);
}
function showResults() {
  if (!report) return;
  const query = $('search').value.trim().toLowerCase();
  const rows=report.results.filter(r=>['SKU','FNSKU','产品名称'].some(k=>String(r[k]||'').toLowerCase().includes(query)));
  table('results',rows,['FNSKU','SKU','产品名称','状态','采用日销','采用覆盖天数','FBA可售','FBA可售天数','AWD可用','FBA补货需求量','建议FBA发货量','国内现货不足量','建议采购量','采购缺口试算量','FBA断货风险']);
}
$('file').addEventListener('change',async()=>{
  invalidate();workbook=null;$('calculate').disabled=true;$('sheets').replaceChildren();
  const file=$('file').files[0];
  if(!file)return;
  if(!file.name.toLowerCase().endsWith('.xlsx')||file.size>20*1024*1024){message('请选择不超过20 MB的.xlsx文件',true);return;}
  $('file-info').textContent='正在识别工作表…';message('');
  try {
    const b64=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result.split(',')[1]);r.onerror=reject;r.readAsDataURL(file);});
    const info=await api('/api/inspect',{file:b64});
    workbook=b64;
    $('file-info').textContent=`${file.name} · ${info.sheets.length} 个工作表已识别`;
    $('shop').replaceChildren();
    info.shops.forEach(s=>{const o=document.createElement('option');o.value=s;o.textContent=s;$('shop').append(o);});
    info.sheets.forEach(s=>{const span=document.createElement('span');span.className='sheet-chip';span.textContent=`${s.name} · ${s.rows} 行`;$('sheets').append(span);});
    $('calculate').disabled=false;
  }catch(error){$('file-info').textContent='文件未加载';message(error.message,true);}
});
fields.concat(['data_confirmed']).forEach(id=>$(id).addEventListener('change',()=>{$('target_days').disabled=$('coverage_mode').value!=='manual';invalidate();message('参数已更新，请重新计算。');}));
$('calculate').addEventListener('click',async()=>{
  if(!workbook)return;
  const options={};
  for(const id of fields){if(!$(id).checkValidity()){$(id).reportValidity();return;}options[id]=$(id).value;}
  options.data_confirmed=$('data_confirmed').checked;
  $('calculate').disabled=true;invalidate();message('正在核算库存、订单与装箱数据…');
  try{
    report=await api('/api/calculate',{file:workbook,options});
    $('count').textContent=report.results.length;
    $('pending').textContent=report.results.filter(r=>r['状态']!=='可复核建议').length;
    const sum=key=>{const valid=report.results.filter(r=>r[key]!==null&&r[key]!==undefined);return valid.length?valid.reduce((s,r)=>s+r[key],0).toLocaleString():'待确认';};
    $('ship-count').textContent=sum('建议FBA发货量');$('buy-count').textContent=sum('建议采购量');
    $('search').value='';showResults();table('issues',report.issues);table('shipments',report.shipments);
    $('assumptions').replaceChildren();report.assumptions.forEach(x=>{const li=document.createElement('li');li.textContent=x;$('assumptions').append(li);});
    $('output').hidden=false;
    message(`核算完成：${report.results.length} 个商品。请先处理待确认事项，再采用正式建议。`);
  }catch(error){message(error.message,true);}finally{$('calculate').disabled=false;}
});
$('search').addEventListener('input',showResults);
$('download').addEventListener('click',()=>{
  if(!report)return;
  const raw=atob(report.excel),bytes=new Uint8Array(raw.length);
  for(let i=0;i<raw.length;i++)bytes[i]=raw.charCodeAt(i);
  const url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  const a=document.createElement('a');a.href=url;a.download=`补货与采购核算-${report.settings.as_of}.xlsx`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
});
