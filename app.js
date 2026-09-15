const definitions = [
  {key:'assortment',title:'Matas assortment',desc:'Source product list - columns A, D and E',need:'EAN, Brand, Product name'},
  {key:'master',title:'Master Data',desc:'Status reference for each product EAN',need:'BAR CODE, STATUS'},
  {key:'template',title:'Kopia template',desc:'The workbook structure to preserve',need:'Product Package Card sheet'}
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
  $('#continueBtn').disabled=definitions.some(d=>!state.files[d.key]);
}
function sheetRows(book){return XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1,raw:false,defval:''})}
function headers(rows,terms){
  for(let r=0;r<Math.min(rows.length,6);r++){let cols=rows[r].map(norm),idx=terms.map(t=>cols.findIndex(c=>c===norm(t)||c.includes(norm(t))));if(idx.every(i=>i>=0))return{r,idx}}
  return null;
}
function parseAssortment(book){
  let rows=sheetRows(book),h=headers(rows,['EAN','Brand','Produktnavn']);
  if(!h)throw new Error('Matas assortment must contain EAN, Brand, and Produktnavn/Product name columns.');
  let [ec,bc,nc]=h.idx;
  return rows.slice(h.r+1).map(r=>({ean:ean(r[ec]),brand:String(r[bc]||'').trim(),name:String(r[nc]||'').trim()})).filter(x=>x.ean&&x.brand&&x.name);
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
    const assortment=parseAssortment(XLSX.read(a,{type:'array',cellText:true}));
    const master=parseMaster(XLSX.read(m,{type:'array',cellText:true}));
    const all=assortment.map(x=>({...x,status:master.get(x.ean)||'',matched:master.has(x.ean)}));
    const ready=all.filter(x=>x.matched&&!discontinued(x.status));
    const removed=all.filter(x=>x.matched&&discontinued(x.status));
    const unmatched=all.filter(x=>!x.matched);
    const brands=ready.reduce((map,x)=>{if(!map.has(x.brand))map.set(x.brand,[]);map.get(x.brand).push(x);return map},new Map());
    state.data={all,ready,removed,unmatched,brands};
    $('#validationSummary').innerHTML=metric('Assortment products',all.length)+metric('Ready products',ready.length)+metric('Discontinued removed',removed.length)+metric('EAN not matched',unmatched.length);
    $('#brandList').innerHTML=[...brands.entries()].sort(([a],[b])=>a.localeCompare(b)).map(([brand,items])=>`<label class="brand-check"><input type="checkbox" value="${esc(brand)}" checked> ${esc(brand)} <small>(${items.length})</small></label>`).join('');
    step(2);show('#configureView');
  }catch(err){alert(err.message)}
  finally{$('#continueBtn').textContent='Continue to configure ->';$('#continueBtn').disabled=definitions.some(d=>!state.files[d.key])}
}
function chosen(){return state.mode==='auto'?[...state.data.brands.keys()]:[...$('#brandList').querySelectorAll('input:checked')].map(x=>x.value)}
function results(){
  const names=chosen();if(!names.length)return alert('Select at least one brand.');
  const d=state.data,rows=names.map(brand=>({brand,products:d.brands.get(brand)||[]}));
  $('#resultIntro').textContent=`${rows.length} Kopia file${rows.length===1?'':'s'} will be generated from your original template.`;
  $('#resultStats').innerHTML=metric('Files to create',rows.length)+metric('Products included',rows.reduce((n,r)=>n+r.products.length,0))+metric('Discontinued excluded',d.removed.length)+metric('EAN unmatched',d.unmatched.length);
  $('#resultRows').innerHTML=rows.map(r=>`<tr><td><strong>${esc(r.brand)}</strong></td><td>${r.products.length}</td><td>${d.removed.length}</td><td>${d.unmatched.length}</td><td><button class="download-one" data-brand="${esc(r.brand)}">Download</button></td></tr>`).join('');
  let warning=$('#warningBox');warning.classList.toggle('hidden',!d.unmatched.length);warning.innerHTML=d.unmatched.length?`<span>i</span><p><strong>${d.unmatched.length} EAN value${d.unmatched.length===1?'':'s'} could not be found in Master Data.</strong> These products are not included in any export.</p>`:'';
  $('#resultRows').querySelectorAll('.download-one').forEach(b=>b.addEventListener('click',()=>downloadBrand(b.dataset.brand)));
  step(4);show('#resultView');
}
function targetSheet(workbook,references){
  const rels=new Map([...references.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/g)].map(m=>[m[1],m[2]]));
  const sheets=[...workbook.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*>/g)];
  const target=sheets.find(m=>/product package card/i.test(m[1]))||sheets[0];
  if(!target||!rels.get(target[2]))throw new Error('The Kopia worksheet could not be resolved.');
  return 'xl/'+rels.get(target[2]).replace(/^\//,'').replace(/^\.\//,'');
}
function writeSheet(xml,products){
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  if(doc.querySelector('parsererror'))throw new Error('The Kopia worksheet could not be read.');
  const data=doc.getElementsByTagName('sheetData')[0],rows=[...data.getElementsByTagName('row')],sample=rows.find(r=>+r.getAttribute('r')>1);
  if(!sample)throw new Error('The Kopia template needs a formatted data row beneath its header.');
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
async function downloadBrand(brand){try{let blob=await makeWorkbook(state.data.brands.get(brand));save(blob,`Kopia_Mass_Product_Packaging_EN_${brand.replace(/[\\\\/:*?"<>|]/g,'-')}.xlsx`)}catch(err){alert(`Could not create ${brand}'s export: ${err.message}`)}}
async function downloadAll(){let b=$('#downloadAllBtn');b.disabled=true;b.textContent='Preparing downloads...';for(const brand of chosen())await downloadBrand(brand);b.disabled=false;b.textContent='Download all Kopia files'}

renderInputs();
$('#continueBtn').addEventListener('click',configure);
$('#processBtn').addEventListener('click',results);
$('#downloadAllBtn').addEventListener('click',downloadAll);
$('#resetBtn').addEventListener('click',()=>{state.files={};state.data=null;renderInputs();$('#continueBtn').disabled=true});
document.querySelectorAll('.back').forEach(b=>b.addEventListener('click',()=>{step(1);show('#uploadView')}));
document.querySelectorAll('.back-config').forEach(b=>b.addEventListener('click',()=>{step(2);show('#configureView')}));
document.querySelectorAll('input[name="mode"]').forEach(i=>i.addEventListener('change',e=>{state.mode=e.target.value;$('#brandPicker').classList.toggle('hidden',state.mode!=='manual')}));
$('#requirementsBtn').addEventListener('click',()=>alert('Matas assortment: EAN, Brand and Produktnavn/Product name.\\nMaster Data: BAR CODE and STATUS.\\nKopia: a workbook containing Product Package Card.'));
