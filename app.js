const files = [
  { key: 'assortment', title: 'Matas assortment', desc: 'Active product list from Matas' },
  { key: 'master', title: 'Master Data', desc: 'EAN status reference file' },
  { key: 'template', title: 'Kopia template', desc: 'Brand packaging template' }
];
const state = {};
const grid = document.querySelector('#uploadGrid');
grid.innerHTML = files.map((f, i) => `<article class="upload-card"><h3>${i + 1} &nbsp; ${f.title}</h3><p>${f.desc}</p><label class="drop" tabindex="0"><input hidden type="file" accept=".xlsx,.xls" data-key="${f.key}"><div><div class="file-icon">▧</div><strong>Drag & drop Excel file here</strong><em>or click to browse</em><small>Supports .xlsx, .xls · Max 50MB</small></div></label></article>`).join('');
grid.addEventListener('change', e => { const input = e.target; if (input.matches('input[type=file]')) { state[input.dataset.key] = input.files[0]; const drop = input.closest('.drop'); drop.querySelector('strong').textContent = input.files[0]?.name || 'Drag & drop Excel file here'; drop.querySelector('em').textContent = input.files[0] ? 'File selected' : 'or click to browse'; drop.querySelector('em').style.color = input.files[0] ? '#15945d' : ''; document.querySelector('#continueBtn').disabled = files.some(f => !state[f.key]); } });
