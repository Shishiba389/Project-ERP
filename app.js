const definitions = [
  {key:'assortment',title:'Assortment',desc:'Any source file containing Brand, Product name and EAN',need:'Brand, Product name, EAN'},
  {key:'master',title:'Master Data',desc:'Status reference for each product EAN',need:'BAR CODE, STATUS'},
  {key:'template',title:'Output template',desc:'The workbook structure to preserve',need:'Product Package Card sheet'}
];
const state={files:{},data:null,mode:'auto'};
const $=s=>document.querySelector(s);
const grid=$('#uploadGrid');
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const ean=v=>String(v??'').replace(/\.0$/,'').replace(/[^0-9]/g,'');
const norm=v=>String(v??'').toLowerCase().replace(/[\s_-]+/g,'');
const metric=(label,value)=>`<div class="metric"><small>${label}</small><strong>${value}</strong></div>`;
function show(id){document.querySelectorAll('.view').forEach(v=>v.classList.add('hidden'));$(id).classList.remove('hidden')}
function step(n){document.querySelectorAll('.step').forEach(x=>{let i=+x.dataset.step;x.classList.toggle('current',i===n);x.classList.toggle('done',i<n)})}
function renderInputs(){
  grid.innerHTML=definitions.map((f,i)=>`<article class="upload-card" data-key="${f.key}"><h3>${i+1} &nbsp; ${f.title}</h3><p>${f.desc}</p><label class="drop" tabindex="0"><input hidden type="file" accept=".xlsx,.xls" data-key="${f.key}"><div><div class="file-icon">X</div><strong>Drag and drop Excel file here</strong><em>or click to browse</em><small>${f.need}</small></div></label></article>`).join('');
  grid.querySelectorAll('input').forEach(input=>input.addEventListener('change',e=>e.target.files[0]&&selectFile(e.target.dataset.key,e.target.files[0])));
  grid.querySelectorAll('.drop').forEach(drop=>{
    ['dragenter','dragover'].forEach(t=>drop.addEventListener(t,e=>{e.preventDefault();drop.classList.add('dragover')}));
    ['dragleave','drop'].forEach(t=>drop.addEventListener(t,e=>{e.preventDefault();drop.classList.remove('dragover');if(t==='drop'&&e.dataTransfer.files[0]){let f=e.dataTransfer.files[0];if(!/\.xlsx?$/i.test(f.name))return alert('Please choose an .xlsx or .xls file.');selectFile(drop.querySelector('input').dataset.key,f)}}));
  });
}
function selectFile(key,file){
  state.files[key]=file;
  const card=grid.querySelector(`[data-key="${key}"]`);
  card.querySelector('.drop strong').textContent=file.name;
  const sub=card.querySelector('.drop em');sub.textContent=`${Math.ceil(file.size/1024)} KB selected`;sub.style.color='#15945d';
  if(key==='template'&&$('#templateStatus'))$('#templateStatus').textContent=`Selected for this session: ${file.name}`;
  $('#continueBtn').disabled=definitions.some(d=>!state.files[d.key]);
}
function sheetRows(book){return XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1,raw:false,defval:''})}
function headers(rows,terms){
  for(let r=0;r<Math.min(rows.length,6);r++){let cols=rows[r].map(norm),idx=terms.map(t=>cols.findIndex(c=>c===norm(t)||c.includes(norm(t))));if(idx.every(i=>i>=0))return{r,idx}}
  return null;
}
function assortmentModel(book){
  const rows=sheetRows(book),aliases={brand:['brand','brands','brandname','manufacturer'],name:['productname','produktnavn','product','itemname','name'],ean:['ean','barcode','barcodeean','itembarcode','gtin']};
  let best={row:0,score:-1};
  for(let row=0;row<Math.min(rows.length,20);row++){const cells=rows[row].map(norm),score=Object.values(aliases).reduce((n,list)=>n+(cells.some(c=>list.some(a=>c===a||c.includes(a)))?1:0),0);if(score>best.score)best={row,score};}
  const model={rows,headerRow:best.row,aliases,columns:[],mapping:{brand:-1,name:-1,ean:-1}};
  setHeaderRow(model,best.row);return model;
}
function setHeaderRow(model,row){model.headerRow=row;model.columns=(model.rows[row]||[]).map((value,index)=>String(value||`Column ${index+1}`));const find=key=>model.columns.findIndex(value=>model.aliases[key].some(alias=>norm(value)===alias||norm(value).includes(alias)));model.mapping={brand:find('brand'),name:find('name'),ean:find('ean')};}
function renderMapping(){
  const source=state.source;const header=$('#mapHeader');header.innerHTML=source.rows.slice(0,Math.min(20,source.rows.length)).map((row,index)=>`<option value="${index}">Row ${index+1}: ${esc(row.filter(value=>value!==''&&value!=null).slice(0,4).join(' | ')||'empty')}</option>`).join('');header.value=source.headerRow;header.onchange=()=>{setHeaderRow(source,+header.value);renderMapping();refreshValidation()};
  const options='<option value="">Select a column</option>'+source.columns.map((label,index)=>`<option value="${index}">${esc(label)}</option>`).join('');
  [['mapBrand','brand'],['mapName','name'],['mapEan','ean']].forEach(([id,key])=>{const select=$('#'+id);select.innerHTML=options;select.value=source.mapping[key]??'';select.addEventListener('change',()=>{source.mapping[key]=select.value===''?-1:+select.value;refreshValidation()});});
}
function computeData(){
  const source=state.source,map=source.mapping;if([map.brand,map.name,map.ean].some(value=>value<0)||new Set([map.brand,map.name,map.ean]).size!==3)return null;
  const assortment=source.rows.slice(source.headerRow+1).map(row=>({ean:ean(row[map.ean]),brand:String(row[map.brand]||'').trim(),name:String(row[map.name]||'').trim()})).filter(item=>item.ean&&item.brand&&item.name);
  const all=assortment.map(item=>({...item,status:state.master.get(item.ean)||'',matched:state.master.has(item.ean)}));
  const ready=all.filter(item=>item.matched&&!discontinued(item.status)),removed=all.filter(item=>item.matched&&discontinued(item.status)),unmatched=all.filter(item=>!item.matched);
  const brands=ready.reduce((result,item)=>{if(!result.has(item.brand))result.set(item.brand,[]);result.get(item.brand).push(item);return result},new Map());
  state.data={all,ready,removed,unmatched,brands};return state.data;
}
function refreshValidation(){
  const data=computeData();if(!data){$('#validationSummary').innerHTML=metric('Required fields','Choose 3 columns')+metric('Ready products','-')+metric('Discontinued removed','-')+metric('EAN not matched','-');$('#brandList').innerHTML='';return;}
  $('#validationSummary').innerHTML=metric('Assortment products',data.all.length)+metric('Ready products',data.ready.length)+metric('Discontinued removed',data.removed.length)+metric('EAN not matched',data.unmatched.length);
  $('#brandList').innerHTML=[...data.brands.entries()].sort(([a],[b])=>a.localeCompare(b)).map(([brand,items])=>`<label class="brand-check"><input type="checkbox" value="${esc(brand)}" checked> ${esc(brand)} <small>(${items.length})</small></label>`).join('');
}
function parseMaster(book){
  let rows=sheetRows(book),h=headers(rows,['BAR CODE','STATUS']);
  if(!h)throw new Error('Master Data must contain BAR CODE and STATUS columns.');
  let [ec,sc]=h.idx,map=new Map();
  rows.slice(h.r+1).forEach(r=>{let code=ean(r[ec]);if(code&&!map.has(code))map.set(code,String(r[sc]||'').trim())});
  return map;
}
function discontinued(status){return['discontinued','discontined'].includes(norm(status))}
async function configure(){
  try{
    if(!window.XLSX||!window.JSZip)throw new Error('Excel libraries could not load. Check your internet connection and reload.');
    $('#continueBtn').disabled=true;$('#continueBtn').textContent='Reading files...';
    const [a,m]=await Promise.all([state.files.assortment.arrayBuffer(),state.files.master.arrayBuffer()]);
    state.source=assortmentModel(XLSX.read(a,{type:'array',cellText:true}));
    state.master=parseMaster(XLSX.read(m,{type:'array',cellText:true}));
    renderMapping();refreshValidation();
    step(2);show('#configureView');
  }catch(err){alert(err.message)}
  finally{$('#continueBtn').textContent='Continue to configure ->';$('#continueBtn').disabled=definitions.some(d=>!state.files[d.key])}
}
function chosen(){return state.mode==='auto'?[...state.data.brands.keys()]:[...$('#brandList').querySelectorAll('input:checked')].map(x=>x.value)}
function results(){
  if(!computeData())return alert('Choose a different column for Brand, Product name, and EAN before continuing.');
  const names=chosen();if(!names.length)return alert('Select at least one brand.');
  const d=state.data,rows=names.map(brand=>({brand,products:d.brands.get(brand)||[]}));
  $('#resultIntro').textContent=`${rows.length} output file${rows.length===1?'':'s'} will be generated from your original template.`;
  $('#resultStats').innerHTML=metric('Files to create',rows.length)+metric('Products included',rows.reduce((n,r)=>n+r.products.length,0))+metric('Discontinued excluded',d.removed.length)+metric('EAN unmatched',d.unmatched.length);
  $('#resultRows').innerHTML=rows.map(r=>`<tr><td><strong>${esc(r.brand)}</strong></td><td>${r.products.length}</td><td>${d.removed.length}</td><td>${d.unmatched.length}</td><td><button class="download-one" data-brand="${esc(r.brand)}">Download</button></td></tr>`).join('');
  let warning=$('#warningBox');warning.classList.toggle('hidden',!d.unmatched.length);warning.innerHTML=d.unmatched.length?`<span>i</span><p><strong>${d.unmatched.length} EAN value${d.unmatched.length===1?'':'s'} could not be found in Master Data.</strong> These products are not included in any export.</p>`:'';
  $('#resultRows').querySelectorAll('.download-one').forEach(b=>b.addEventListener('click',()=>downloadBrand(b.dataset.brand)));
  addHistory({brands:rows.length,products:rows.reduce((n,row)=>n+row.products.length,0)});step(4);show('#resultView');
}
function targetSheet(workbook,references){
  const rels=new Map([...references.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/g)].map(m=>[m[1],m[2]]));
  const sheets=[...workbook.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*>/g)];
  const target=sheets.find(m=>/product package card/i.test(m[1]))||sheets[0];
  if(!target||!rels.get(target[2]))throw new Error('The output worksheet could not be resolved.');
  return 'xl/'+rels.get(target[2]).replace(/^\//,'').replace(/^\.\//,'');
}
function writeSheet(xml,products){
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  if(doc.querySelector('parsererror'))throw new Error('The output worksheet could not be read.');
  const data=doc.getElementsByTagName('sheetData')[0],rows=[...data.getElementsByTagName('row')],sample=rows.find(r=>+r.getAttribute('r')>1);
  if(!sample)throw new Error('The output template needs a formatted data row beneath its header.');
  rows.filter(r=>+r.getAttribute('r')>1).forEach(r=>r.remove());
  products.forEach((product,index)=>{
    const rowNo=index+2,row=sample.cloneNode(true);row.setAttribute('r',rowNo);
    [...row.getElementsByTagName('c')].forEach(cell=>{let col=cell.getAttribute('r').replace(/\d+/g,'');cell.setAttribute('r',col+rowNo);while(cell.firstChild)cell.removeChild(cell.firstChild);cell.removeAttribute('t')});
    const values={A:index+1,B:product.ean,D:product.name};
    for(const [col,value] of Object.entries(values)){
      let cell=[...row.getElementsByTagName('c')].find(c=>c.getAttribute('r')===col+rowNo);
      if(!cell){cell=doc.createElementNS(doc.documentElement.namespaceURI,'c');cell.setAttribute('r',col+rowNo);row.appendChild(cell)}
      cell.setAttribute('t','inlineStr');let is=doc.createElementNS(doc.documentElement.namespaceURI,'is'),text=doc.createElementNS(doc.documentElement.namespaceURI,'t');text.textContent=String(value);is.appendChild(text);cell.appendChild(is);
    }
    data.appendChild(row);
  });
  const dimension=doc.getElementsByTagName('dimension')[0];if(dimension)dimension.setAttribute('ref',`A1:Y${Math.max(2,products.length+1)}`);
  return new XMLSerializer().serializeToString(doc);
}
async function makeWorkbook(products){
  const zip=await JSZip.loadAsync(await state.files.template.arrayBuffer());
  const wb=await zip.file('xl/workbook.xml').async('string'),rels=await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const path=targetSheet(wb,rels),xml=await zip.file(path).async('string');
  zip.file(path,writeSheet(xml,products));
  zip.remove('xl/calcChain.xml');
  zip.file('xl/_rels/workbook.xml.rels',rels.replace(/<Relationship\b(?=[^>]*calcChain)[^>]*\/>/g,''));
  return zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
}
function save(blob,name){let url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500)}
async function downloadBrand(brand){try{let blob=await makeWorkbook(state.data.brands.get(brand));save(blob,`Mass_Product_Packaging_EN_${brand.replace(/[\\\\/:*?"<>|]/g,'-')}.xlsx`)}catch(err){alert(`Could not create ${brand}'s export: ${err.message}`)}}
async function downloadAll(){let b=$('#downloadAllBtn');b.disabled=true;b.textContent='Preparing downloads...';for(const brand of chosen())await downloadBrand(brand);b.disabled=false;b.textContent='Download all output files'}
function getHistory(){try{return JSON.parse(localStorage.getItem('project-erp-history')||'[]')}catch{return []}}
function addHistory(entry){const rows=[{...entry,at:new Date().toISOString()},...getHistory()].slice(0,20);localStorage.setItem('project-erp-history',JSON.stringify(rows));}
function renderHistory(){const rows=getHistory(),host=$('#historyList');host.innerHTML=rows.length?rows.map(row=>`<div class="history-row"><div><strong>${row.products} validated product${row.products===1?'':'s'}</strong><small>Output prepared for ${row.brands} brand${row.brands===1?'':'s'}</small></div><div><strong>${new Date(row.at).toLocaleDateString()}</strong><small>${new Date(row.at).toLocaleTimeString()}</small></div><div><strong>Local session</strong><small>Source files were not retained</small></div></div>`).join(''):'<p class="empty-message">No processing activity has been recorded in this browser yet.</p>'}
function route(name){document.querySelectorAll('#process,.route').forEach(section=>section.classList.add('hidden'));const target=$('#'+name);target.classList.remove('hidden','route-enter');requestAnimationFrame(()=>target.classList.add('route-enter'));document.querySelectorAll('.nav-link').forEach(link=>link.classList.toggle('active',link.dataset.route===name));if(name==='history')renderHistory()}

renderInputs();
$('#continueBtn').addEventListener('click',configure);
$('#processBtn').addEventListener('click',results);
$('#downloadAllBtn').addEventListener('click',downloadAll);
$('#resetBtn').addEventListener('click',()=>{state.files={};state.data=null;renderInputs();$('#continueBtn').disabled=true});
document.querySelectorAll('.back').forEach(b=>b.addEventListener('click',()=>{step(1);show('#uploadView')}));
document.querySelectorAll('.back-config').forEach(b=>b.addEventListener('click',()=>{step(2);show('#configureView')}));
document.querySelectorAll('input[name="mode"]').forEach(i=>i.addEventListener('change',e=>{state.mode=e.target.value;$('#brandPicker').classList.toggle('hidden',state.mode!=='manual')}));
document.querySelectorAll('.nav-link').forEach(link=>link.addEventListener('click',event=>{event.preventDefault();route(link.dataset.route)}));
$('#templatePicker').addEventListener('change',event=>{if(event.target.files[0]){selectFile('template',event.target.files[0]);route('process')}});
const clearHistory=()=>{localStorage.removeItem('project-erp-history');renderHistory()};
$('#historyClearBtn').addEventListener('click',clearHistory);$('#settingsClearHistory').addEventListener('click',clearHistory);
$('#requirementsBtn').addEventListener('click',()=>alert('Assortment: choose any Excel file containing Brand, Product name and EAN. You will confirm the column mapping in the next step.\\nMaster Data: BAR CODE and STATUS.\\nOutput template: a workbook containing Product Package Card.'));
