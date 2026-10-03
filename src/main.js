let JSZip;

function loadZipLibrary() {
  if (window.JSZip) return Promise.resolve(window.JSZip);
  if (window.__zipLibrary) return window.__zipLibrary;
  window.__zipLibrary = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
    script.onload = () => resolve(window.JSZip);
    script.onerror = () => reject(new Error('Не удалось загрузить библиотеку обработки DOCX.'));
    document.head.append(script);
  });
  return window.__zipLibrary;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const state = { title: null, content: null };
const fileSize = bytes => `${(bytes / 1024 / 1024).toFixed(1)} МБ`;
const safeName = name => name.replace(/\.docx$/i, '').replace(/[^а-яёa-z0-9 _-]/gi, '').trim() || 'лабораторная';

function wordElements(element, name) { return [...element.getElementsByTagNameNS(W, name)]; }
function directWordChild(element, name) { return [...element.children].find(child => child.namespaceURI === W && child.localName === name); }
function hasAncestor(element, name) { for (let node = element.parentElement; node; node = node.parentElement) if (node.namespaceURI === W && node.localName === name) return true; return false; }
function setAttribute(element, name, value) { element.setAttributeNS(W, `w:${name}`, value); }
function createWordElement(document, name) { return document.createElementNS(W, `w:${name}`); }
function paragraphText(paragraph) { return wordElements(paragraph, 't').map(node => node.textContent).join('').trim(); }

function setParagraphFormatting(paragraph) {
  const document = paragraph.ownerDocument;
  let properties = directWordChild(paragraph, 'pPr');
  if (!properties) { properties = createWordElement(document, 'pPr'); paragraph.insertBefore(properties, paragraph.firstChild); }
  ['jc', 'spacing', 'ind'].forEach(name => directWordChild(properties, name)?.remove());

  const containsImage = wordElements(paragraph, 'drawing').length > 0 || wordElements(paragraph, 'pict').length > 0;
  const isTableText = hasAncestor(paragraph, 'tbl');
  const isList = Boolean(directWordChild(properties, 'numPr'));
  const isCaption = /^(рисунок|рис\.)\s*\d+/i.test(paragraphText(paragraph));
  const centered = containsImage || isCaption;

  const alignment = createWordElement(document, 'jc');
  setAttribute(alignment, 'val', centered ? 'center' : 'both');
  properties.append(alignment);

  const spacing = createWordElement(document, 'spacing');
  setAttribute(spacing, 'line', '360');
  setAttribute(spacing, 'lineRule', 'auto');
  properties.append(spacing);

  const indent = createWordElement(document, 'ind');
  if (isTableText || centered) {
    // Explicit zero values override a first-line indent inherited from a Word style.
    setAttribute(indent, 'firstLine', '0');
    setAttribute(indent, 'hanging', '0');
    setAttribute(indent, 'left', '0');
  } else if (!isList) {
    setAttribute(indent, 'firstLine', '709');
  }
  if (!isList || isTableText || centered) properties.append(indent);
}

function setRunFormatting(run) {
  const document = run.ownerDocument;
  let properties = directWordChild(run, 'rPr');
  if (!properties) { properties = createWordElement(document, 'rPr'); run.insertBefore(properties, run.firstChild); }
  directWordChild(properties, 'rFonts')?.remove();
  ['sz', 'szCs'].forEach(name => directWordChild(properties, name)?.remove());
  const fonts = createWordElement(document, 'rFonts');
  ['ascii', 'hAnsi', 'cs', 'eastAsia'].forEach(name => setAttribute(fonts, name, 'Times New Roman'));
  properties.append(fonts);
  ['sz', 'szCs'].forEach(name => { const size = createWordElement(document, name); setAttribute(size, 'val', '28'); properties.append(size); });
}

function normaliseContent(contentXml) {
  const xml = new DOMParser().parseFromString(contentXml, 'application/xml');
  if (xml.querySelector('parsererror')) throw new Error('В DOCX не удалось прочитать документ.');
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  wordElements(body, 'p').forEach(setParagraphFormatting);
  wordElements(body, 'r').forEach(setRunFormatting);
  wordElements(body, 'tbl').forEach(setTableBorders);
  const section = directWordChild(body, 'sectPr');
  section?.remove();
  removeLeadingPageBreaks(body);
  return [...body.childNodes].map(node => new XMLSerializer().serializeToString(node)).join('');
}

function setTableBorders(table) {
  const document = table.ownerDocument;
  let properties = directWordChild(table, 'tblPr');
  if (!properties) { properties = createWordElement(document, 'tblPr'); table.insertBefore(properties, table.firstChild); }
  directWordChild(properties, 'tblBorders')?.remove();

  const borders = createWordElement(document, 'tblBorders');
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(side => {
    const border = createWordElement(document, side);
    setAttribute(border, 'val', 'single');
    setAttribute(border, 'sz', '4');
    setAttribute(border, 'space', '0');
    setAttribute(border, 'color', '000000');
    borders.append(border);
  });
  properties.append(borders);
}

function removeLeadingPageBreaks(body) {
  for (const paragraph of [...body.children]) {
    if (paragraph.namespaceURI !== W || paragraph.localName !== 'p') break;
    const hasContent = wordElements(paragraph, 't').some(node => node.textContent.trim()) || wordElements(paragraph, 'drawing').length || wordElements(paragraph, 'pict').length;
    wordElements(paragraph, 'br').filter(node => node.getAttributeNS(W, 'type') === 'page').forEach(node => node.remove());
    directWordChild(paragraph, 'pPr')?.getElementsByTagNameNS(W, 'pageBreakBefore')[0]?.remove();
    if (hasContent) break;
    if (!paragraph.children.length || !wordElements(paragraph, 'r').length) paragraph.remove();
  }
}

async function mergeNumbering(titleZip, contentZip, contentXml) {
  const sourceFile = contentZip.file('word/numbering.xml');
  if (!sourceFile) return contentXml;
  const titleFile = titleZip.file('word/numbering.xml');
  if (!titleFile) {
    titleZip.file('word/numbering.xml', await sourceFile.async('uint8array'));
    await addNumberingRelationship(titleZip);
    await addNumberingContentType(titleZip, contentZip);
    return contentXml;
  }

  const [sourceXml, titleXml] = await Promise.all([sourceFile.async('string'), titleFile.async('string')]);
  const parser = new DOMParser();
  const source = parser.parseFromString(sourceXml, 'application/xml');
  const title = parser.parseFromString(titleXml, 'application/xml');
  const sourceRoot = source.documentElement;
  const titleRoot = title.documentElement;
  const directChildren = (root, name) => [...root.children].filter(node => node.namespaceURI === W && node.localName === name);
  const maxId = (nodes, attribute) => Math.max(-1, ...nodes.map(node => Number(node.getAttributeNS(W, attribute) || node.getAttribute(`w:${attribute}`) || node.getAttribute(attribute)) || 0));
  let nextAbstractId = maxId(directChildren(titleRoot, 'abstractNum'), 'abstractNumId') + 1;
  let nextNumId = maxId(directChildren(titleRoot, 'num'), 'numId') + 1;
  const abstractIds = new Map();
  const numIds = new Map();

  for (const abstractNum of directChildren(sourceRoot, 'abstractNum')) {
    const oldId = abstractNum.getAttributeNS(W, 'abstractNumId') || abstractNum.getAttribute('w:abstractNumId');
    const newId = String(nextAbstractId++);
    const copy = title.importNode(abstractNum, true);
    setAttribute(copy, 'abstractNumId', newId);
    titleRoot.append(copy);
    abstractIds.set(oldId, newId);
  }
  for (const num of directChildren(sourceRoot, 'num')) {
    const oldId = num.getAttributeNS(W, 'numId') || num.getAttribute('w:numId');
    const newId = String(nextNumId++);
    const copy = title.importNode(num, true);
    setAttribute(copy, 'numId', newId);
    const abstractNumId = copy.getElementsByTagNameNS(W, 'abstractNumId')[0];
    const mappedAbstractId = abstractIds.get(abstractNumId?.getAttributeNS(W, 'val'));
    if (abstractNumId && mappedAbstractId) setAttribute(abstractNumId, 'val', mappedAbstractId);
    titleRoot.append(copy);
    numIds.set(oldId, newId);
  }
  titleZip.file('word/numbering.xml', new XMLSerializer().serializeToString(title));
  for (const [oldId, newId] of numIds) contentXml = contentXml.replace(new RegExp(`(<w:numId\\b[^>]*\\bw:val=")${oldId}("[^>]*>)`, 'g'), `$1${newId}$2`);
  return contentXml;
}

async function addNumberingRelationship(titleZip) {
  const file = titleZip.file('word/_rels/document.xml.rels');
  if (!file) return;
  const xml = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
  const root = xml.documentElement;
  if ([...root.getElementsByTagName('Relationship')].some(node => /\/numbering$/.test(node.getAttribute('Type') || ''))) return;
  const used = new Set([...root.getElementsByTagName('Relationship')].map(node => node.getAttribute('Id')));
  let index = 1; while (used.has(`rId${index}`)) index += 1;
  const relationship = xml.createElement('Relationship');
  relationship.setAttribute('Id', `rId${index}`);
  relationship.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering');
  relationship.setAttribute('Target', 'numbering.xml');
  root.append(relationship);
  titleZip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(xml));
}

async function addNumberingContentType(titleZip, contentZip) {
  const [titleFile, sourceFile] = [titleZip.file('[Content_Types].xml'), contentZip.file('[Content_Types].xml')];
  if (!titleFile || !sourceFile) return;
  const [titleXml, sourceXml] = await Promise.all([titleFile.async('string'), sourceFile.async('string')]);
  const parser = new DOMParser();
  const title = parser.parseFromString(titleXml, 'application/xml');
  const source = parser.parseFromString(sourceXml, 'application/xml');
  if ([...title.documentElement.getElementsByTagName('Override')].some(node => node.getAttribute('PartName') === '/word/numbering.xml')) return;
  const override = [...source.documentElement.getElementsByTagName('Override')].find(node => node.getAttribute('PartName') === '/word/numbering.xml');
  if (override) title.documentElement.append(title.importNode(override, true));
  titleZip.file('[Content_Types].xml', new XMLSerializer().serializeToString(title));
}

async function mergeImageRelationships(titleZip, contentZip, contentXml) {
  const sourceRelationshipsFile = contentZip.file('word/_rels/document.xml.rels');
  const titleRelationshipsFile = titleZip.file('word/_rels/document.xml.rels');
  if (!sourceRelationshipsFile || !titleRelationshipsFile) return contentXml;
  const [sourceXml, titleXml] = await Promise.all([sourceRelationshipsFile.async('string'), titleRelationshipsFile.async('string')]);
  const parser = new DOMParser();
  const source = parser.parseFromString(sourceXml, 'application/xml');
  const title = parser.parseFromString(titleXml, 'application/xml');
  const root = title.documentElement;
  const usedIds = new Set([...root.getElementsByTagName('Relationship')].map(node => node.getAttribute('Id')));
  let nextId = 1;
  let nextImage = 1;
  const remap = new Map();

  for (const relationship of [...source.documentElement.getElementsByTagName('Relationship')]) {
    if (!/\/image$/.test(relationship.getAttribute('Type') || '')) continue;
    while (usedIds.has(`rId${nextId}`)) nextId += 1;
    const id = `rId${nextId++}`;
    const sourceTarget = relationship.getAttribute('Target');
    const extension = (sourceTarget.match(/(\.[^./]+)$/)?.[1] || '.bin').replace(/[^.a-z0-9]/gi, '');
    const target = `media/imported-image-${nextImage++}${extension}`;
    const image = contentZip.file(`word/${sourceTarget.replace(/^\.\//, '')}`);
    if (!image) continue;
    titleZip.file(`word/${target}`, await image.async('uint8array'));
    const copied = title.createElement('Relationship');
    copied.setAttribute('Id', id);
    copied.setAttribute('Type', relationship.getAttribute('Type'));
    copied.setAttribute('Target', target);
    root.append(copied);
    usedIds.add(id);
    remap.set(relationship.getAttribute('Id'), id);
  }

  titleZip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(title));
  for (const [oldId, newId] of remap) contentXml = contentXml.replace(new RegExp(`(r:(?:embed|link|id)=")${oldId}("|')`, 'g'), `$1${newId}$2`);
  return contentXml;
}

function titleBodyWithoutTrailingPageBreaks(titleXml) {
  const xml = new DOMParser().parseFromString(titleXml, 'application/xml');
  if (xml.querySelector('parsererror')) throw new Error('В титульном листе не удалось прочитать документ.');
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  const section = directWordChild(body, 'sectPr');
  const titleSection = section ? new XMLSerializer().serializeToString(section) : '';
  section?.remove();

  // Title templates often end with a manual page break. Keeping it and adding our
  // own separator produces an unwanted blank page before the lab text.
  const paragraphs = [...body.children].filter(node => node.namespaceURI === W && node.localName === 'p');
  const trailingEmptyParagraphs = [];
  for (let index = paragraphs.length - 1; index >= 0; index -= 1) {
    const paragraph = paragraphs[index];
    const pageBreaks = wordElements(paragraph, 'br').filter(node => node.getAttributeNS(W, 'type') === 'page');
    const hasContent = wordElements(paragraph, 't').some(node => node.textContent.trim()) || wordElements(paragraph, 'drawing').length || wordElements(paragraph, 'pict').length;
    if (!hasContent && !pageBreaks.length) { trailingEmptyParagraphs.push(paragraph); continue; }
    if (!pageBreaks.length) break;
    trailingEmptyParagraphs.forEach(node => node.remove());
    pageBreaks.forEach(node => node.remove());
    if (!hasContent) paragraph.remove();
    // Once real title content is reached, only its terminal manual page break is removed.
    break;
  }
  return { titleSection, content: [...body.childNodes].map(node => new XMLSerializer().serializeToString(node)).join('') };
}

async function makeDocument(titleFile, contentFile) {
  JSZip = JSZip || await loadZipLibrary();
  const [titleZip, contentZip] = await Promise.all([JSZip.loadAsync(titleFile), JSZip.loadAsync(contentFile)]);
  const [titleXml, rawContentXml] = await Promise.all([titleZip.file('word/document.xml').async('string'), contentZip.file('word/document.xml').async('string')]);
  const numberedContentXml = await mergeNumbering(titleZip, contentZip, rawContentXml);
  const contentXml = await mergeImageRelationships(titleZip, contentZip, numberedContentXml);
  const { titleSection, content: titleContent } = titleBodyWithoutTrailingPageBreaks(titleXml);
  const titleBody = titleXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/)?.[1] || '';
  // A page-break-before paragraph avoids the additional empty page caused by a
  // standalone manual break in some title-page templates.
  const pageBreak = '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>';
  titleZip.file('word/document.xml', titleXml.replace(titleBody, `${titleContent}${pageBreak}${normaliseContent(contentXml)}${titleSection}`));
  return titleZip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}

const rules = [['T', 'Times New Roman', '14 pt'], ['↔', 'Выравнивание', 'по ширине'], ['↕', 'Интервал', '1,5 строки'], ['¶', 'Красная строка', '1,25 см']];
document.querySelector('#root').innerHTML = `<main><nav><a class="brand" href="#top"><i></i>ФОРМАТикс</a><div class="navlinks"><a href="#how">Как это работает</a><a href="#rules">Требования ГОСТ</a></div><button class="help">?</button></nav><section class="hero" id="top"><div class="eyebrow"><span></span> Лабораторные без мучений</div><h1>Оформление <em>по ГОСТу</em><br/>за пару минут.</h1><p>Загрузите титульный лист и текст работы — мы бережно соберём готовый документ с нужным форматированием.</p><a class="start" href="#upload">Начать оформление <b>↓</b></a><div class="paper paper-one"></div><div class="paper paper-two"><div class="paper-logo">L</div><div class="paper-lines"><b></b><b></b><b></b><b></b><b></b></div><div class="paper-stamp">ГОСТ</div></div></section><section class="workspace" id="upload"><div class="section-label">01 — Загрузка файлов</div><div class="workspace-head"><h2>Соберём вашу<br/><em>лабораторную.</em></h2><p>Сначала титульный лист, затем текст работы.<br/>Титульник останется в точности как у вас.</p></div><div class="steps"><article class="upload-card" data-kind="title"><input type="file" accept=".docx"/><div class="number">01</div><div class="file-icon">↑</div><h3>Титульный лист</h3><p>Загрузите готовый титульник</p><small>DOCX · до 20 МБ</small></article><div class="connector"><span>→</span></div><article class="upload-card locked" data-kind="content"><input type="file" accept=".docx"/><div class="number">02</div><div class="file-icon">•</div><h3>Текст работы</h3><p>Станет доступно после титульника</p><small>DOCX · до 20 МБ</small></article></div><div class="action-row"><span>Загрузите два файла в формате DOCX</span><button class="format-button" disabled>Оформить и скачать <b>→</b></button></div></section><section class="rules" id="rules"><div><div class="section-label">02 — Уже настроено</div><h2>Все детали<br/><em>под контролем.</em></h2><p>Текст, списки, таблицы и рисунки сохраняются аккуратными. Мы меняем только то, что требуется для оформления.</p></div><div class="rule-grid">${rules.map(([icon, title, text]) => `<div class="rule"><span>${icon}</span><div><b>${title}</b><small>${text}</small></div></div>`).join('')}</div></section><section class="note" id="how"><span>Важно</span><p>Титульный лист не изменяется. Форматирование начинается с новой страницы после него.</p><b>✦</b></section><footer><a class="brand" href="#top"><i></i>ФОРМАТикс</a><span>Ваши файлы обрабатываются прямо в браузере</span><span>© 2026</span></footer></main>`;
function render() { for (const kind of ['title', 'content']) { const file = state[kind], card = document.querySelector(`[data-kind=${kind}]`); if (kind === 'content') card.classList.toggle('locked', !state.title); card.classList.toggle('filled', !!file); card.querySelector('.file-icon').textContent = file ? '✓' : kind === 'content' && !state.title ? '•' : '↑'; card.querySelector('p').textContent = file ? file.name : kind === 'title' ? 'Загрузите готовый титульник' : state.title ? 'Загрузите текст лабораторной' : 'Станет доступно после титульника'; card.querySelector('small').textContent = file ? fileSize(file.size) : 'DOCX · до 20 МБ'; } const ready = state.title && state.content, button = document.querySelector('.format-button'); button.disabled = !ready; document.querySelector('.action-row span').textContent = ready ? 'Файлы готовы к оформлению' : 'Загрузите два файла в формате DOCX'; }
function accept(file, kind) { if (!file || !/\.docx$/i.test(file.name)) return; state[kind] = file; render(); }
document.querySelectorAll('.upload-card').forEach(card => { const kind = card.dataset.kind, input = card.querySelector('input'); card.addEventListener('click', () => { if (kind === 'title' || state.title) input.click(); }); input.addEventListener('change', event => accept(event.target.files[0], kind)); for (const eventName of ['dragover', 'drop']) card.addEventListener(eventName, event => { if (kind === 'title' || state.title) event.preventDefault(); }); card.addEventListener('drop', event => accept(event.dataTransfer.files[0], kind)); });
document.querySelector('.format-button').addEventListener('click', async event => { const button = event.currentTarget; button.disabled = true; button.firstChild.textContent = 'Оформляем… '; try { const blob = await makeDocument(state.title, state.content), anchor = document.createElement('a'); anchor.href = URL.createObjectURL(blob); anchor.download = `${safeName(state.content.name)}_ГОСТ.docx`; anchor.click(); URL.revokeObjectURL(anchor.href); button.firstChild.textContent = 'Скачать ещё раз '; } catch (error) { alert(error.message || 'Не удалось обработать документ. Проверьте, что оба файла — корректные DOCX.'); button.firstChild.textContent = 'Оформить и скачать '; } finally { button.disabled = false; } });
