/* Project ERP runs entirely in the browser. Source workbooks are never uploaded. */
const state = { matas: null, master: null, template: null, products: [], excluded: [], brands: new Map(), templateSheets: [] };
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

function normalize(value) { return String(value ?? '').trim().toLocaleLowerCase(); }
function ean(value) { const v = String(value ?? '').trim(); return v.endsWith('.0') ? v.slice(0, -2) : v; }
function escapeXml(value) { return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;'); }
function safeName(value) { return String(value).trim().replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_'); }
function toast(message, error = false) { const box = $('#toast'); box.textContent = message; box.classList.toggle('error', error); box.classList.add('show'); window.setTimeout(() => box.classList.remove('show'), 3500); }
function logActivity(title, message) { const entry = document.createElement('div'); entry.innerHTML = `<time>Now</time><span class="timeline-dot orange"></span><p><strong>${title}</strong><small>${message}</small></p>`; $('#activity-log').prepend(entry); }

function openView(name) {
  $$('.view').forEach((view) => view.classList.toggle('active', view.id === `view-${name}`));
  $$('.nav-item').forEach((item) => item.classList.toggle('active', item.dataset.view === name));
  $('#breadcrumb-name').textContent = name === 'workspace' ? 'Packaging data' : name.replace('-', ' ');
  $('#main').focus({ preventScroll: true });
}

$$('[data-view]').forEach((button) => button.addEventListener('click', () => openView(button.dataset.view)));
$('#top-create').addEventListener('click', () => $('#export-dialog').showModal());
$('#start-export').addEventListener('click', () => $('#export-dialog').showModal());
$('#products-export').addEventListener('click', () => $('#export-dialog').showModal());
$('#files-start').addEventListener('click', () => $('#export-dialog').showModal());

const commandDialog = $('#command-dialog');
function showCommands() { commandDialog.showModal(); window.setTimeout(() => $('#command-input').focus(), 0); }
$('#command-trigger').addEventListener('click', showCommands);
window.addEventListener('keydown', (event) => { if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') { event.preventDefault(); showCommands(); } if (event.key === 'Escape') { commandDialog.close(); } });
$$('[data-command]').forEach((button) => button.addEventListener('click', () => { commandDialog.close(); button.dataset.command === 'export' ? $('#export-dialog').showModal() : openView(button.dataset.command); }));

const savedTheme = localStorage.getItem('erp-theme') || 'light';
document.body.dataset.theme = savedTheme;
function syncThemeToggle() {
  if (!$('#theme-toggle')) return;
  const dark = document.body.dataset.theme === 'dark';
  $('#theme-toggle').innerHTML = `<i data-lucide="${dark ? 'moon-star' : 'sun-medium'}"></i><span>${dark ? 'Dark deck' : 'Light studio'}</span>`;
  lucide.createIcons();
}
if ($('#theme-toggle')) $('#theme-toggle').addEventListener('click', () => { document.body.dataset.theme = document.body.dataset.theme === 'dark' ? 'light' : 'dark'; localStorage.setItem('erp-theme', document.body.dataset.theme); syncThemeToggle(); });
$('#stage-export').addEventListener('click', () => $('#export-dialog').showModal());
$('.output-core').addEventListener('click', () => $('#export-dialog').showModal());
syncThemeToggle();

function fileNameOutput(input, output) { input.addEventListener('change', () => { $(output).textContent = input.files[0]?.name || 'No file selected'; }); }
fileNameOutput($('#matas-file'), '#matas-name'); fileNameOutput($('#master-file'), '#master-name'); fileNameOutput($('#template-file'), '#template-name');

async function readWorkbook(file) { return XLSX.read(await file.arrayBuffer(), { type: 'array', cellText: true, raw: false }); }
function rowsFrom(workbook, sheetName) { return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false }); }
function findColumn(headers, label) { return headers.findIndex((header) => normalize(header) === normalize(label)); }

async function prepareData() {
  const matasFile = $('#matas-file').files[0]; const masterFile = $('#master-file').files[0]; const templateFile = $('#template-file').files[0];
  if (!matasFile || !masterFile || !templateFile) return;
  try {
    const [matasBook, masterBook, templateBook] = await Promise.all([readWorkbook(matasFile), readWorkbook(masterFile), readWorkbook(templateFile)]);
    const matasRows = rowsFrom(matasBook, matasBook.SheetNames[0]);
    const masterSheet = masterBook.SheetNames.find((name) => normalize(name) === 'artikkel') || masterBook.SheetNames[0];
    const masterRows = rowsFrom(masterBook, masterSheet);
    const masterHeaders = masterRows[0] || [];
    const barcodeColumn = findColumn(masterHeaders, 'BAR CODE'); const statusColumn = findColumn(masterHeaders, 'STATUS');
    if (barcodeColumn < 0 || statusColumn < 0) throw new Error('Master Data needs BAR CODE and STATUS columns.');
    const master = new Map();
    masterRows.slice(1).forEach((row) => { const key = ean(row[barcodeColumn]); if (key) master.set(key, { status: String(row[statusColumn] ?? '').trim() }); });
    const products = []; const excluded = []; const unmatched = [];
    matasRows.slice(2).forEach((row) => {
      const product = { ean: ean(row[0]), brand: String(row[3] ?? '').trim(), name: String(row[4] ?? '').trim() };
      if (!product.ean) return;
      const record = master.get(product.ean);
      if (!record) { unmatched.push(product); return; }
      product.status = record.status;
      if (normalize(record.status) === 'discontinued') excluded.push(product); else products.push(product);
    });
    if (unmatched.length) throw new Error(`${unmatched.length} EAN value(s) could not be matched in Master Data.`);
    if (!products.length) throw new Error('No eligible products remain after the Discontinued rule.');
    state.matas = matasBook; state.master = masterBook; state.template = templateFile; state.products = products; state.excluded = excluded;
    state.brands = products.reduce((map, product) => { map.set(product.brand, [...(map.get(product.brand) || []), product]); return map; }, new Map());
    state.templateSheets = templateBook.SheetNames;
    const select = $('#sheet-select'); select.innerHTML = state.templateSheets.map((name) => `<option value="${escapeXml(name)}">${escapeXml(name)}</option>`).join(''); select.disabled = false;
    const preferred = state.templateSheets.find((name) => normalize(name).startsWith('product package card'));
    if (preferred) select.value = preferred;
    renderBrands(); renderProducts(); renderDashboard(); renderAnalytics();
    $('#validation-result').textContent = `${products.length} eligible products in ${state.brands.size} brands. ${excluded.length} Discontinued product(s) will be excluded.`;
    $('#validation-result').className = 'validation-result ready'; $('#generate-button').disabled = false;
    logActivity('Source files validated', `${products.length} eligible products matched to Master Data.`); toast('Files validated. Choose an export mode.');
  } catch (error) { $('#validation-result').textContent = error.message; $('#validation-result').className = 'validation-result error'; $('#generate-button').disabled = true; toast(error.message, true); }
}

['#matas-file', '#master-file', '#template-file'].forEach((selector) => $(selector).addEventListener('change', prepareData));
function renderBrands() { const selector = $('#brand-selector'); selector.innerHTML = [...state.brands.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([brand, products]) => `<label><input type="checkbox" value="${escapeXml(brand)}" checked> ${escapeXml(brand)} <span class="mono">${products.length}</span></label>`).join(''); }
$$('input[name="mode"]').forEach((radio) => radio.addEventListener('change', () => { $('#brand-selector').hidden = $('input[name="mode"]:checked').value !== 'choose'; }));
function renderProducts() {
  const target = $('#products-table'); if (!target) return; const items = state.products;
  target.innerHTML = items.map((product, index) => `<tr><td class="row-signal"></td><td><strong>${escapeXml(product.name)}</strong></td><td class="mono">${escapeXml(product.ean)}</td><td>${escapeXml(product.brand)}</td><td><span class="status active">${escapeXml(product.status)}</span></td><td><span class="mono">READY</span></td><td><button class="icon-button" aria-label="Inspect ${escapeXml(product.name)}"><i data-lucide="ellipsis"></i></button></td></tr>`).join(''); lucide.createIcons();
}
function renderDashboard() {
  if ($('#brand-count')) $('#brand-count').textContent = state.brands.size; if ($('#product-count')) $('#product-count').textContent = state.products.length; if ($('#excluded-count')) $('#excluded-count').textContent = state.excluded.length;
  if (!$('#batch-preview')) return;
  $('#batch-preview').innerHTML = [...state.brands.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([brand, products]) => `<button class="batch-tile" data-batch="${escapeXml(brand)}"><strong>${escapeXml(brand)}</strong><span>${products.length} PRODUCTS · READY</span></button>`).join('');
  $$('#batch-preview [data-batch]').forEach((button) => button.addEventListener('click', () => { $('#export-dialog').showModal(); $('input[name="mode"][value="choose"]').checked = true; $('#brand-selector').hidden = false; $$('#brand-selector input').forEach((input) => { input.checked = input.value === button.dataset.batch; }); }));
}
function renderAnalytics() { if (!$('#brand-bars')) return; $('#brand-bars').innerHTML = [...state.brands.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([brand, items]) => `<div class="bar-row"><span>${escapeXml(brand)}</span><div class="bar"><i style="width:${Math.max(12, Math.round(items.length / state.products.length * 100))}%"></i></div><strong>${items.length}</strong></div>`).join(''); }
if ($('#product-search')) $('#product-search').addEventListener('input', (event) => { const term = normalize(event.target.value); $$('#products-table tr').forEach((row) => { row.hidden = Boolean(term) && !normalize(row.textContent).includes(term); }); });

function getSheetPath(zip, sheetName) {
  const parser = new DOMParser(); const workbook = parser.parseFromString(zip.file('xl/workbook.xml').async('string'), 'application/xml');
  return workbook;
}
async function resolveSheetPart(zip, sheetName) {
  const parser = new DOMParser(); const workbook = parser.parseFromString(await zip.file('xl/workbook.xml').async('string'), 'application/xml');
  const sheet = [...workbook.getElementsByTagName('*')].find((node) => node.localName === 'sheet' && node.getAttribute('name') === sheetName);
  if (!sheet) throw new Error(`The selected worksheet "${sheetName}" was not found in the template.`);
  const relationshipId = sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
  const relations = parser.parseFromString(await zip.file('xl/_rels/workbook.xml.rels').async('string'), 'application/xml');
  const relation = [...relations.getElementsByTagName('*')].find((node) => node.localName === 'Relationship' && node.getAttribute('Id') === relationshipId);
  if (!relation) throw new Error('Could not locate the selected worksheet in the template package.');
  const target = relation.getAttribute('Target').replace(/^\//, '');
  return target.startsWith('xl/') ? target : `xl/${target.replace(/^\.\//, '')}`;
}
function cellStyle(existing) { return existing?.match(/\ss="[^"]*"/)?.[0] || ''; }
function cell(reference, value, existing) {
  if (value === null) return existing ? `<c r="${reference}"${cellStyle(existing)}/>` : '';
  // Keep indexes and standard numeric EANs as numeric Excel values. This matches the Kopia data columns while still protecting non-numeric identifiers.
  if ((reference.startsWith('A') || reference.startsWith('B')) && /^\d{1,15}$/.test(String(value))) return `<c r="${reference}"${cellStyle(existing)}><v>${value}</v></c>`;
  return `<c r="${reference}"${cellStyle(existing)} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
}
function replaceCell(row, reference, value) {
  const matcher = new RegExp(`<c\\b[^>]*\\br="${reference}"[^>]*(?:/>|>[\\s\\S]*?<\\/c>)`, 'i'); const existing = row.match(matcher)?.[0]; const replacement = cell(reference, value, existing);
  if (existing) return row.replace(matcher, replacement);
  return value === null ? row : row.replace(/<\/row>$/i, `${replacement}</row>`);
}
function replaceRow(xml, rowNumber, values) {
  const matcher = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`, 'i'); const current = xml.match(matcher)?.[0];
  if (!current) throw new Error(`Template row ${rowNumber} does not exist. Use a template with enough prepared rows.`);
  let next = current; Object.entries(values).forEach(([column, value]) => { next = replaceCell(next, `${column}${rowNumber}`, value); }); return xml.replace(matcher, next);
}
async function createExport(brand, products, sheetName) {
  const zip = await JSZip.loadAsync(await state.template.arrayBuffer()); const sheetPath = await resolveSheetPart(zip, sheetName); const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error('The selected worksheet file could not be opened.'); let sheetXml = await sheetFile.async('string');
  const rowMatches = [...sheetXml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>/gi)].map((match) => Number(match[1])); const lastRow = Math.max(...rowMatches);
  if (products.length + 1 > lastRow) throw new Error(`${brand} has ${products.length} products but the template has only ${lastRow - 1} data rows.`);
  for (let row = 2; row <= lastRow; row += 1) { const product = products[row - 2]; sheetXml = replaceRow(sheetXml, row, { A: product ? row - 1 : null, B: product ? product.ean : null, D: product ? product.name : null }); }
  zip.file(sheetPath, sheetXml); const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  const anchor = document.createElement('a'); anchor.href = URL.createObjectURL(blob); anchor.download = `Kopia_Mass_Product_Packaging_EN_${safeName(brand)}.xlsx`; anchor.click(); window.setTimeout(() => URL.revokeObjectURL(anchor.href), 1200);
}
$('#generate-button').addEventListener('click', async () => {
  const mode = $('input[name="mode"]:checked').value; const selected = mode === 'auto' ? [...state.brands.keys()] : $$('#brand-selector input:checked').map((input) => input.value);
  if (!selected.length) return toast('Select at least one brand to export.', true);
  const button = $('#generate-button'); button.disabled = true; button.innerHTML = '<i data-lucide="loader-circle"></i>Preparing…'; lucide.createIcons();
  try { for (const brand of selected) await createExport(brand, state.brands.get(brand), $('#sheet-select').value); $('#export-dialog').close(); logActivity('Brand files generated', `${selected.length} Kopia file(s) prepared for download.`); toast(`${selected.length} Kopia file(s) generated.`); }
  catch (error) { toast(error.message, true); }
  finally { button.disabled = false; button.innerHTML = '<i data-lucide="file-output"></i>Generate files'; lucide.createIcons(); }
});

lucide.createIcons();
const citrusNote = document.querySelector('.sidebar-note');
if (citrusNote && !citrusNote.querySelector('.citrus-art')) {
  const citrus = new Image(); citrus.className = 'citrus-art'; citrus.alt = 'Orange and fresh green leaves'; citrus.src = 'assets/citrus-corner.png'; citrus.style.cssText = 'position:absolute;left:-24px;bottom:-88px;width:310px;max-width:none;mix-blend-mode:screen;opacity:.94;pointer-events:none;z-index:0;'; citrusNote.prepend(citrus);
}
