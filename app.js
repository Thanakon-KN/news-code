const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PLACEHOLDER = '<figure class="text-center"><img class="img-fluid" src="xxx"></figure>';
const MAX_FILE_SIZE = 20 * 1024 * 1024;
const decoder = new TextDecoder('utf-8');

const fileInput = document.querySelector('#docx-files');
const dropzone = document.querySelector('#dropzone');
const fileList = document.querySelector('#file-list');
const fileCount = document.querySelector('#file-count');
const convertButton = document.querySelector('#convert');
const clearButton = document.querySelector('#clear');
const statusNode = document.querySelector('#status');
const resultSection = document.querySelector('#results');
const resultCards = document.querySelector('#result-cards');
const resultFileTabs = document.querySelector('#result-file-tabs');
const resultLanguageTabs = document.querySelector('#result-language-tabs');
const outputWorkspace = document.querySelector('#output-workspace');
const publishEmpty = document.querySelector('#publish-empty');
const publishControls = document.querySelector('#publish-controls');
const publishFileSelect = document.querySelector('#publish-file');
const publishLanguages = document.querySelector('#publish-languages');
const publishSummary = document.querySelector('#publish-summary');
const publishCopyButton = document.querySelector('#publish-copy');
const publishDownloadButton = document.querySelector('#publish-download');
const headlineControl = document.querySelector('#skip-headline');
const workflowSteps = Array.from(document.querySelectorAll('.workflow-step'));
const workflowRules = Array.from(document.querySelectorAll('.workflow-rule'));

let selectedFiles = [];
let completedResults = [];
let publishLocale = '';

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]);
}

function keepSpaces(value) {
  return escapeHtml(value).replace(/ {2,}/g, spaces => '&nbsp;'.repeat(spaces.length - 1) + ' ');
}

function directChildren(node, localName) {
  return Array.from(node?.children || []).filter(child => child.localName === localName);
}

function firstChild(node, localName) {
  return directChildren(node, localName)[0] || null;
}

function descendants(node, localName) {
  return Array.from(node.getElementsByTagNameNS(WORD_NS, localName));
}

function wordValue(node) {
  return node?.getAttributeNS(WORD_NS, 'val') || node?.getAttribute('w:val') || '';
}

function xmlDocument(source, name) {
  if (!source) throw new Error(`DOCX ไม่มี ${name}`);
  const xml = new DOMParser().parseFromString(source, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error(`อ่าน ${name} ไม่ได้`);
  return xml;
}

async function inflateEntry(bytes, method) {
  if (method === 0) return bytes;
  if (method !== 8) throw new Error(`ไฟล์ DOCX ใช้การบีบอัดที่ไม่รองรับ (${method})`);
  if (!('DecompressionStream' in window)) throw new Error('เบราว์เซอร์นี้ยังไม่รองรับการอ่าน DOCX แบบบีบอัด');
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function readDocxEntries(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const searchStart = Math.max(0, bytes.length - 65_557);
  let endOffset = -1;
  for (let offset = bytes.length - 22; offset >= searchStart; offset--) {
    if (view.getUint32(offset, true) === 0x06054b50) { endOffset = offset; break; }
  }
  if (endOffset < 0) throw new Error('ไฟล์นี้ไม่ใช่ DOCX ที่รองรับ');
  const entryCount = view.getUint16(endOffset + 10, true);
  let cursor = view.getUint32(endOffset + 16, true);
  if (entryCount > 2000) throw new Error('ไฟล์มีรายการภายในมากเกินกว่าจะประมวลผลได้');

  const required = new Set([
    'word/document.xml',
    'word/_rels/document.xml.rels',
    'word/numbering.xml',
    'word/styles.xml'
  ]);
  const contents = new Map();

  for (let index = 0; index < entryCount; index++) {
    if (view.getUint32(cursor, true) !== 0x02014b50) throw new Error('อ่านสารบัญ ZIP ของไฟล์ DOCX ไม่ได้');
    const method = view.getUint16(cursor + 10, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const uncompressedSize = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const localOffset = view.getUint32(cursor + 42, true);
    const name = decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength));

    if (required.has(name)) {
      if (uncompressedSize > 25 * 1024 * 1024) throw new Error('ส่วนข้อความในไฟล์มีขนาดเกินกำหนด');
      if (view.getUint32(localOffset, true) !== 0x04034b50) throw new Error('อ่านรายการภายใน DOCX ไม่ได้');
      const localNameLength = view.getUint16(localOffset + 26, true);
      const localExtraLength = view.getUint16(localOffset + 28, true);
      const dataStart = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = bytes.subarray(dataStart, dataStart + compressedSize);
      const content = await inflateEntry(compressed, method);
      contents.set(name, decoder.decode(content));
    }
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  return contents;
}

function parseRelationships(xmlSource) {
  if (!xmlSource) return new Map();
  const xml = xmlDocument(xmlSource, 'document relationships');
  return new Map(Array.from(xml.getElementsByTagName('Relationship')).map(relationship => [
    relationship.getAttribute('Id'),
    { target: relationship.getAttribute('Target') || '', external: relationship.getAttribute('TargetMode') === 'External' }
  ]));
}

function parseStyles(xmlSource) {
  if (!xmlSource) return { byId: new Map(), docDefaults: {}, defaultParagraph: {}, defaultCharacter: {} };
  const xml = xmlDocument(xmlSource, 'styles');
  const styles = new Map();
  for (const style of Array.from(xml.getElementsByTagNameNS(WORD_NS, 'style'))) {
    const id = style.getAttributeNS(WORD_NS, 'styleId') || style.getAttribute('w:styleId');
    const name = firstChild(style, 'name');
    if (id) styles.set(id, {
      name: name?.getAttributeNS(WORD_NS, 'val') || name?.getAttribute('w:val') || id,
      type: style.getAttributeNS(WORD_NS, 'type') || style.getAttribute('w:type') || '',
      basedOn: wordValue(firstChild(style, 'basedOn')),
      formatting: formatProperties(firstChild(style, 'rPr')),
      isDefault: ['1', 'true', 'on'].includes((style.getAttributeNS(WORD_NS, 'default') || style.getAttribute('w:default') || '').toLowerCase())
    });
  }
  const defaults = xml.getElementsByTagNameNS(WORD_NS, 'docDefaults')[0];
  const defaultRun = firstChild(firstChild(defaults, 'rPrDefault'), 'rPr');
  const styleContext = { byId: styles };
  const defaultParagraphId = Array.from(styles.entries()).find(([, style]) => style.isDefault && style.type === 'paragraph')?.[0];
  const defaultCharacterId = Array.from(styles.entries()).find(([, style]) => style.isDefault && style.type === 'character')?.[0];
  const defaultParagraph = resolveStyleFormatting(defaultParagraphId, styleContext);
  const defaultCharacter = resolveStyleFormatting(defaultCharacterId, styleContext);
  return { byId: styles, docDefaults: formatProperties(defaultRun), defaultParagraph, defaultCharacter };
}

function parseNumbering(xmlSource) {
  const numMap = new Map();
  const formatMap = new Map();
  if (!xmlSource) return { numMap, formatMap };
  const xml = xmlDocument(xmlSource, 'numbering');
  for (const num of Array.from(xml.getElementsByTagNameNS(WORD_NS, 'num'))) {
    const numId = num.getAttributeNS(WORD_NS, 'numId') || num.getAttribute('w:numId');
    const abstract = firstChild(num, 'abstractNumId');
    if (numId && abstract) numMap.set(numId, wordValue(abstract));
  }
  for (const abstract of Array.from(xml.getElementsByTagNameNS(WORD_NS, 'abstractNum'))) {
    const abstractId = abstract.getAttributeNS(WORD_NS, 'abstractNumId') || abstract.getAttribute('w:abstractNumId');
    for (const level of directChildren(abstract, 'lvl')) {
      const levelIndex = level.getAttributeNS(WORD_NS, 'ilvl') || level.getAttribute('w:ilvl') || '0';
      const format = firstChild(level, 'numFmt');
      const text = firstChild(level, 'lvlText');
      formatMap.set(`${abstractId}:${levelIndex}`, {
        format: wordValue(format) || 'bullet',
        marker: wordValue(text)
      });
    }
  }
  return { numMap, formatMap };
}

function xmlTextOf(node) {
  let text = '';
  const visit = child => {
    if (child.nodeType !== 1) return;
    if (child.localName === 't') text += child.textContent || '';
    else if (child.localName === 'tab') text += '\t';
    else if (child.localName === 'br' || child.localName === 'cr') text += '\n';
    else for (const nested of child.children) visit(nested);
  };
  visit(node);
  return text;
}

function autoLinkText(value) {
  const pattern = /(https?:\/\/[^\s<>"']+|www\.[^\s<>"']+|[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/gi;
  let output = '';
  let cursor = 0;
  for (const match of value.matchAll(pattern)) {
    const raw = match[0];
    const index = match.index ?? 0;
    let link = raw;
    let trailing = '';
    while (/[.,;:!?)]$/.test(link)) { trailing = link.slice(-1) + trailing; link = link.slice(0, -1); }
    output += keepSpaces(value.slice(cursor, index));
    const href = /^www\./i.test(link) ? `https://${link}` : /^[A-Z0-9._%+-]+@/i.test(link) ? `mailto:${link}` : link;
    output += `<a href="${escapeHtml(href)}" target="_blank">${keepSpaces(link)}</a>${escapeHtml(trailing)}`;
    cursor = index + raw.length;
  }
  output += keepSpaces(value.slice(cursor));
  return output;
}

function propertyState(properties, names) {
  const states = [];
  for (const name of names) {
    const element = firstChild(properties, name);
    if (!element) continue;
    const value = wordValue(element).toLowerCase();
    states.push(!['0', 'false', 'off', 'none'].includes(value));
  }
  if (!states.length) return undefined;
  return states.includes(true);
}

function formatProperties(properties) {
  if (!properties) return {};
  const formatting = {};
  const bold = propertyState(properties, ['b', 'bCs']);
  const italic = propertyState(properties, ['i', 'iCs']);
  const underline = propertyState(properties, ['u']);
  if (bold !== undefined) formatting.bold = bold;
  if (italic !== undefined) formatting.italic = italic;
  if (underline !== undefined) formatting.underline = underline;
  return formatting;
}

function resolveStyleFormatting(styleId, styles, visited = new Set()) {
  if (!styleId || visited.has(styleId)) return {};
  visited.add(styleId);
  const style = styles.byId.get(styleId);
  if (!style) return {};
  return { ...resolveStyleFormatting(style.basedOn, styles, visited), ...style.formatting };
}

function runFormatting(run, context, paragraphStyleId) {
  const properties = firstChild(run, 'rPr');
  const characterStyleId = wordValue(firstChild(properties, 'rStyle'));
  const effective = {
    ...context.styles.docDefaults,
    ...context.styles.defaultParagraph,
    ...context.styles.defaultCharacter,
    ...resolveStyleFormatting(paragraphStyleId, context.styles),
    ...resolveStyleFormatting(characterStyleId, context.styles),
    ...formatProperties(properties)
  };
  return { bold: effective.bold === true, italic: effective.italic === true, underline: effective.underline === true };
}

function wrapFormatting(html, format) {
  if (format.underline) html = `<u>${html}</u>`;
  if (format.italic) html = `<em>${html}</em>`;
  if (format.bold) html = `<strong>${html}</strong>`;
  return html;
}

function mergeAdjacentFormatting(html) {
  return html.replace(/<\/(strong|em|u)><\1>/g, '');
}

function safeHref(target) {
  const value = String(target || '').trim();
  if (/^(https?:\/\/|mailto:|tel:)/i.test(value)) return value;
  if (/^www\./i.test(value)) return `https://${value}`;
  return '';
}

function renderInline(node, relationships, context, paragraphStyleId) {
  if (node.nodeType !== 1) return '';
  switch (node.localName) {
    case 'pPr': case 'rPr': return '';
    case 't': return autoLinkText(node.textContent || '');
    case 'tab': return '&nbsp;&nbsp;&nbsp;';
    case 'br': case 'cr': return '<br>';
    case 'noBreakHyphen': return '&#8209;';
    case 'softHyphen': return '&shy;';
    case 'r': {
      const inner = Array.from(node.children).map(child => renderInline(child, relationships, context, paragraphStyleId)).join('');
      return wrapFormatting(inner, runFormatting(node, context, paragraphStyleId));
    }
    case 'hyperlink': {
      const id = node.getAttributeNS(REL_NS, 'id') || node.getAttribute('r:id') || '';
      const relationship = relationships.get(id);
      const inner = Array.from(node.children).map(child => renderInline(child, relationships, context, paragraphStyleId)).join('');
      const href = relationship?.external ? safeHref(relationship.target) : '';
      return href ? `<a href="${escapeHtml(href)}" target="_blank">${inner}</a>` : inner;
    }
    default: return Array.from(node.children || []).map(child => renderInline(child, relationships, context, paragraphStyleId)).join('');
  }
}

function countStyles(node, context, paragraphStyleId) {
  const runs = descendants(node, 'r').filter(run => xmlTextOf(run).trim());
  let bold = 0;
  let maxSize = 0;
  for (const run of runs) {
    const format = runFormatting(run, context, paragraphStyleId);
    if (format.bold) bold++;
    const size = firstChild(firstChild(run, 'rPr'), 'sz');
    maxSize = Math.max(maxSize, Number(wordValue(size)) || 0);
  }
  return { runCount: runs.length, boldCount: bold, maxSize };
}

function listProperties(paragraph, numbering) {
  const paragraphProps = firstChild(paragraph, 'pPr');
  const numPr = firstChild(paragraphProps, 'numPr');
  if (!numPr) return null;
  const numId = wordValue(firstChild(numPr, 'numId'));
  const level = wordValue(firstChild(numPr, 'ilvl')) || '0';
  if (!numId) return null;
  const abstractId = numbering.numMap.get(numId);
  const definition = numbering.formatMap.get(`${abstractId}:${level}`) || { format: 'bullet', marker: '' };
  const ordered = !['bullet', 'none'].includes(definition.format);
  const type = ({ upperRoman: 'I', lowerRoman: 'i', upperLetter: 'A', lowerLetter: 'a' })[definition.format];
  return { id: numId, level: Number(level), ordered, type: type || '', format: definition.format };
}

function paragraphNode(paragraph, context, isFirstContentParagraph = false) {
  const text = xmlTextOf(paragraph);
  const paragraphProps = firstChild(paragraph, 'pPr');
  const styleId = wordValue(firstChild(paragraphProps, 'pStyle'));
  const styleName = context.styles.byId.get(styleId)?.name || styleId;
  const alignment = wordValue(firstChild(paragraphProps, 'jc'));
  const runs = countStyles(paragraph, context, styleId);
  const body = mergeAdjacentFormatting(Array.from(paragraph.children).map(child => renderInline(child, context.relationships, context, styleId)).join(''));
  const size = runs.maxSize;
  const isHeadingStyle = /^(title|subtitle|heading\s*[1-3])$/i.test(styleName || '');
  const hasHeadlineLabel = /^\s*headline\s*:/i.test(text);
  const fullyBoldHeadline = text.trim().length >= 24 && runs.runCount > 0 && runs.boldCount === runs.runCount;
  const candidate = hasHeadlineLabel || isHeadingStyle || (fullyBoldHeadline && (alignment === 'center' || size >= 48 || (isFirstContentParagraph && size >= 32)));
  return {
    kind: 'paragraph', text, html: `<p>${body}</p>`, locale: detectLocale(text),
    alignment, styleName, candidate, hasHeadlineLabel, list: listProperties(paragraph, context.numbering)
  };
}

function parseTable(table, context) {
  const rows = directChildren(table, 'tr').map(row => directChildren(row, 'tc').map(cell => {
    const blocks = directChildren(cell, 'p').map(paragraph => paragraphNode(paragraph, context)).filter(block => block.text.trim());
    return { text: blocks.map(block => block.text).join('\n'), html: blocks.map(block => block.html).join('\n'), blocks };
  }));
  const bodyLabels = /^(เนื้อข่าว|เนื้อหา|body|news body|article body|news content|content)$/i;
  const labelText = rows.flatMap(row => row[0] ? row[0].text.split('\n') : []).map(value => value.trim()).filter(Boolean);
  const contentRows = rows.filter(row => row[0] && row[0].text.split('\n').some(line => bodyLabels.test(line.trim())));
  if (contentRows.length) {
    const blocks = [];
    for (const row of contentRows) for (const cell of row.slice(1)) blocks.push(...cell.blocks);
    return { kind: 'formBody', blocks, ignoredMetadata: true };
  }
  const tableText = rows.flat().map(cell => cell.text).join(' ');
  const tableHtml = `<table>\n<tbody>\n${rows.map(row => `<tr>${row.map(cell => `<td>${cell.html}</td>`).join('')}</tr>`).join('\n')}\n</tbody>\n</table>`;
  return { kind: 'table', text: tableText, html: tableHtml, locale: detectLocale(tableText), candidate: false, labels: labelText };
}

function detectLocale(text) {
  const thai = (text.match(/[\u0E00-\u0E7F]/g) || []).length;
  const latin = (text.match(/[A-Za-z]/g) || []).length;
  if (thai && thai >= Math.max(2, latin * 0.12)) return 'TH';
  if (latin || text.trim()) return 'EN';
  return '';
}

function isLanguageLabel(text) {
  return /^(ภาษาไทย|ภาษาอังกฤษ|ไทย|อังกฤษ|thai|english|th|en)$/i.test(text.trim());
}

function processHeadline(blocks, skipHeadline) {
  if (!skipHeadline) return { blocks, skipped: [] };
  const skipped = [];
  const output = blocks.filter(block => {
    if (block.kind !== 'paragraph') return true;
    const plain = block.text.trim();
    if (!plain || isLanguageLabel(plain)) return false;
    if (block.candidate) { skipped.push(plain.replace(/^\s*headline\s*:\s*/i, '')); return false; }
    return true;
  });
  return { blocks: output, skipped };
}

function renderListBlocks(blocks) {
  const rendered = [];
  let list = null;
  const closeList = () => {
    if (list) rendered.push(`</${list.tag}>`);
    list = null;
  };
  for (const block of blocks) {
    if (block.kind === 'paragraph' && block.list) {
      const tag = block.list.ordered ? 'ol' : 'ul';
      const key = `${block.list.id}:${block.list.level}:${tag}:${block.list.type}`;
      if (!list || list.key !== key) {
        closeList();
        const typeAttribute = tag === 'ol' && block.list.type ? ` type="${block.list.type}"` : '';
        list = { key, tag };
        rendered.push(`<${tag}${typeAttribute}>`);
      }
      rendered.push(`<li>${block.html.replace(/^<p>|<\/p>$/g, '')}</li>`);
    } else {
      closeList();
      rendered.push(block.html);
    }
  }
  closeList();
  return rendered.join('\n');
}

function groupByLanguage(blocks) {
  const groups = { TH: [], EN: [] };
  for (const block of blocks) {
    const locale = block.locale || 'EN';
    if (!isLanguageLabel(block.text || '')) groups[locale].push(block);
  }
  return Object.entries(groups).filter(([, group]) => group.length).map(([locale, group]) => ({
    locale,
    html: `${PLACEHOLDER}\n${renderListBlocks(group)}`
  }));
}

async function readDocument(file, skipHeadline) {
  if (!file.name.toLowerCase().endsWith('.docx')) throw new Error('รองรับเฉพาะไฟล์ .docx');
  if (file.size > MAX_FILE_SIZE) throw new Error('ไฟล์มีขนาดเกิน 20 MB');
  const entries = await readDocxEntries(file);
  const documentXml = xmlDocument(entries.get('word/document.xml'), 'เอกสาร Word');
  const relationships = parseRelationships(entries.get('word/_rels/document.xml.rels'));
  const styles = parseStyles(entries.get('word/styles.xml'));
  const numbering = parseNumbering(entries.get('word/numbering.xml'));
  const context = { relationships, styles, numbering };
  const body = documentXml.getElementsByTagNameNS(WORD_NS, 'body')[0];
  if (!body) throw new Error('ไม่พบเนื้อหาในเอกสาร');

  const blocks = [];
  let ignoredMetadata = false;
  let seenContentParagraph = false;
  for (const child of Array.from(body.children)) {
    if (child.localName === 'p') {
      const paragraph = paragraphNode(child, context, !seenContentParagraph);
      if (paragraph.text.trim()) {
        blocks.push(paragraph);
        seenContentParagraph = true;
      }
    } else if (child.localName === 'tbl') {
      const table = parseTable(child, context);
      if (table.kind === 'formBody') { blocks.push(...table.blocks); ignoredMetadata = true; }
      else if (table.text.trim()) blocks.push(table);
    }
  }
  if (!blocks.length) throw new Error('ไม่พบข้อความข่าวใน DOCX');
  const result = processHeadline(blocks, skipHeadline);
  const outputs = groupByLanguage(result.blocks);
  if (!outputs.length) throw new Error('ไม่พบข้อความที่สามารถแยกภาษาได้');
  return {
    outputs,
    skipped: result.skipped,
    paragraphCount: blocks.filter(block => block.kind === 'paragraph').length,
    ignoredMetadata
  };
}

function formatBytes(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function setStatus(message, error = false) {
  statusNode.textContent = message;
  statusNode.classList.toggle('error', error);
}

function updateWorkflow(stage) {
  workflowSteps.forEach(step => {
    const stepNumber = Number(step.dataset.step);
    step.classList.toggle('active', stepNumber === stage);
    step.classList.toggle('completed', stepNumber < stage);
    if (stepNumber === stage) step.setAttribute('aria-current', 'step');
    else step.removeAttribute('aria-current');
  });
  workflowRules.forEach((rule, index) => rule.classList.toggle('completed', stage > index + 1));
}

function resetResults() {
  resultCards.replaceChildren();
  resultFileTabs.replaceChildren();
  resultLanguageTabs.replaceChildren();
  resultSection.hidden = true;
  outputWorkspace.classList.remove('has-results');
  completedResults = [];
  publishLocale = '';
  publishFileSelect.replaceChildren();
  publishLanguages.replaceChildren();
  publishEmpty.hidden = false;
  publishControls.hidden = true;
}

function updateFileView() {
  fileList.replaceChildren();
  selectedFiles.forEach((file, index) => {
    const row = document.createElement('div');
    row.className = 'file-row';
    const icon = document.createElement('span');
    icon.className = 'doc-icon';
    icon.textContent = 'DOCX';
    const info = document.createElement('span');
    info.className = 'file-info';
    const name = document.createElement('span');
    name.className = 'file-name';
    name.textContent = file.name;
    const meta = document.createElement('span');
    meta.className = 'file-meta';
    meta.textContent = formatBytes(file.size);
    info.append(name, meta);
    const remove = document.createElement('button');
    remove.className = 'remove-file';
    remove.type = 'button';
    remove.setAttribute('aria-label', `ลบไฟล์ ${file.name}`);
    remove.title = 'ลบไฟล์';
    remove.textContent = '×';
    remove.addEventListener('click', () => {
      selectedFiles.splice(index, 1);
      resetResults();
      updateFileView();
      setStatus('');
    });
    row.append(icon, info, remove);
    fileList.append(row);
  });
  fileCount.textContent = `${selectedFiles.length} ไฟล์`;
  convertButton.disabled = selectedFiles.length === 0;
  clearButton.disabled = selectedFiles.length === 0;
  updateWorkflow(selectedFiles.length ? 2 : 1);
}

function addFiles(files) {
  const additions = [];
  for (const file of files) {
    if (!file.name.toLowerCase().endsWith('.docx')) { setStatus(`ข้าม ${file.name}: รองรับเฉพาะ DOCX`, true); continue; }
    if (file.size > MAX_FILE_SIZE) { setStatus(`ข้าม ${file.name}: ไฟล์มีขนาดเกิน 20 MB`, true); continue; }
    if (selectedFiles.some(existing => existing.name === file.name && existing.size === file.size && existing.lastModified === file.lastModified)) continue;
    additions.push(file);
  }
  selectedFiles.push(...additions);
  if (additions.length) resetResults();
  updateFileView();
  if (additions.length) setStatus(`${additions.length} ไฟล์พร้อมแปลง`);
}

function renderResults(results) {
  resultCards.replaceChildren();
  resultFileTabs.replaceChildren();
  completedResults = results;
  publishFileSelect.replaceChildren();
  for (const result of results) {
    const option = document.createElement('option');
    option.value = String(publishFileSelect.options.length);
    option.textContent = result.file.name;
    publishFileSelect.append(option);
    const fileTab = document.createElement('button');
    fileTab.type = 'button';
    fileTab.className = 'result-file-tab';
    fileTab.textContent = result.file.name;
    fileTab.title = result.file.name;
    const fileIndex = publishFileSelect.options.length - 1;
    fileTab.addEventListener('click', () => {
      publishFileSelect.value = String(fileIndex);
      updatePublishPanel();
    });
    resultFileTabs.append(fileTab);
    const card = document.createElement('article');
    card.className = 'result-card';
    const head = document.createElement('div');
    head.className = 'result-card-head';
    const title = document.createElement('div');
    title.className = 'result-title';
    const filename = document.createElement('strong');
    filename.textContent = result.file.name;
    const subtitle = document.createElement('span');
    subtitle.textContent = `${result.data.paragraphCount} ย่อหน้า · ${result.data.outputs.map(item => item.locale === 'TH' ? 'ภาษาไทย' : 'English').join(' / ')}`;
    title.append(filename, subtitle);
    head.append(title);
    const body = document.createElement('div');
    body.className = 'result-body';
    for (const output of result.data.outputs) {
      const section = document.createElement('section');
      section.className = 'language-output';
      section.dataset.locale = output.locale;
      const sectionHead = document.createElement('div');
      sectionHead.className = 'language-output-head';
      const label = document.createElement('div');
      label.className = 'language-label';
      const languageName = output.locale === 'TH' ? 'ภาษาไทย' : 'English';
      label.append(languageName);
      const tag = document.createElement('span');
      tag.className = 'language-tag';
      tag.textContent = output.locale;
      label.append(tag);
      const tools = document.createElement('div');
      tools.className = 'output-actions';
      const copy = document.createElement('button');
      copy.className = 'mini-button';
      copy.type = 'button';
      copy.textContent = 'คัดลอก HTML';
      const download = document.createElement('button');
      download.className = 'mini-button';
      download.type = 'button';
      download.textContent = 'ดาวน์โหลด';
      const textarea = document.createElement('textarea');
      textarea.className = 'code-output';
      textarea.readOnly = true;
      textarea.spellcheck = false;
      textarea.setAttribute('aria-label', `HTML ${languageName} สำหรับ ${result.file.name}`);
      textarea.value = output.html;
      const previewLabel = document.createElement('div');
      previewLabel.className = 'preview-label';
      previewLabel.textContent = 'ตัวอย่างการแสดงผล';
      const preview = document.createElement('div');
      preview.className = 'output-preview';
      preview.setAttribute('aria-label', `ตัวอย่างเนื้อหา ${languageName} สำหรับ ${result.file.name}`);
      preview.innerHTML = output.html;
      for (const image of preview.querySelectorAll('figure img[src="xxx"]')) {
        const placeholder = document.createElement('span');
        placeholder.textContent = 'พื้นที่สำหรับรูปภาพข่าว';
        image.replaceWith(placeholder);
      }
      const codeLabel = document.createElement('div');
      codeLabel.className = 'code-label';
      codeLabel.textContent = 'HTML';
      copy.addEventListener('click', async () => {
        try { await navigator.clipboard.writeText(textarea.value); copy.textContent = 'คัดลอกแล้ว'; }
        catch { textarea.focus(); textarea.select(); document.execCommand('copy'); copy.textContent = 'คัดลอกแล้ว'; }
        window.setTimeout(() => { copy.textContent = 'คัดลอก HTML'; }, 1600);
      });
      download.addEventListener('click', () => {
        const suffix = output.locale.toLowerCase();
        downloadText(`${result.file.name.replace(/\.docx$/i, '')}-${suffix}.html`, textarea.value);
      });
      tools.append(copy, download);
      sectionHead.append(label, tools);
      const previewColumn = document.createElement('div');
      previewColumn.className = 'output-column preview-column';
      previewColumn.append(previewLabel, preview);
      const codeColumn = document.createElement('div');
      codeColumn.className = 'output-column code-column';
      codeColumn.append(codeLabel, textarea);
      const content = document.createElement('div');
      content.className = 'output-content';
      content.append(previewColumn, codeColumn);
      section.append(sectionHead, content);
      body.append(section);
    }
    if (result.data.skipped.length || result.data.ignoredMetadata) {
      const notes = document.createElement('details');
      notes.className = 'result-notes';
      const summary = document.createElement('summary');
      summary.textContent = 'รายละเอียดการแปลง';
      const list = document.createElement('ul');
      if (result.data.skipped.length) {
        const item = document.createElement('li');
        item.textContent = `ข้ามพาดหัว ${result.data.skipped.length} บรรทัดตามตัวเลือก: ${result.data.skipped.join(' · ')}`;
        list.append(item);
      }
      if (result.data.ignoredMetadata) {
        const item = document.createElement('li');
        item.textContent = 'พบตารางแบบฟอร์มข่าว: แปลงเฉพาะส่วน “เนื้อข่าว” และละช่องกำกับงาน เช่น หัวข้อข่าว ภาพข่าว และผู้ประสานงาน';
        list.append(item);
      }
      notes.append(summary, list);
      body.append(notes);
    }
    card.append(head, body);
    resultCards.append(card);
  }
  resultSection.hidden = false;
  outputWorkspace.classList.add('has-results');
  publishLocale = results[0]?.data.outputs[0]?.locale || '';
  publishEmpty.hidden = results.length > 0;
  publishControls.hidden = results.length === 0;
  updatePublishPanel();
  updateWorkflow(3);
}

function getPublishOutput() {
  const result = completedResults[Number(publishFileSelect.value)];
  const output = result?.data.outputs.find(item => item.locale === publishLocale);
  return result && output ? { result, output } : null;
}

function updatePublishPanel() {
  const selectedIndex = Number(publishFileSelect.value);
  const result = completedResults[selectedIndex];
  if (!result) return;
  if (!result.data.outputs.some(item => item.locale === publishLocale)) {
    publishLocale = result.data.outputs[0]?.locale || '';
  }
  publishLanguages.replaceChildren();
  resultLanguageTabs.replaceChildren();
  Array.from(resultCards.children).forEach((card, index) => { card.hidden = index !== selectedIndex; });
  Array.from(resultFileTabs.children).forEach((tab, index) => {
    tab.classList.toggle('active', index === selectedIndex);
    tab.setAttribute('aria-pressed', String(index === selectedIndex));
  });
  const selectedCard = resultCards.children[selectedIndex];
  selectedCard?.querySelectorAll('.language-output').forEach(section => { section.hidden = section.dataset.locale !== publishLocale; });
  for (const output of result.data.outputs) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `publish-language${output.locale === publishLocale ? ' active' : ''}`;
    button.setAttribute('aria-pressed', String(output.locale === publishLocale));
    button.textContent = output.locale === 'TH' ? 'ไทย' : 'English';
    button.addEventListener('click', () => {
      publishLocale = output.locale;
      updatePublishPanel();
    });
    publishLanguages.append(button);
    const resultTab = document.createElement('button');
    resultTab.type = 'button';
    resultTab.className = `result-language-tab${output.locale === publishLocale ? ' active' : ''}`;
    resultTab.setAttribute('aria-pressed', String(output.locale === publishLocale));
    resultTab.textContent = output.locale === 'TH' ? 'ภาษาไทย' : 'English';
    resultTab.addEventListener('click', () => { publishLocale = output.locale; updatePublishPanel(); });
    resultLanguageTabs.append(resultTab);
  }
  publishSummary.textContent = `${result.data.paragraphCount} ย่อหน้า · ${publishLocale === 'TH' ? 'ภาษาไทย' : 'English'}`;
}

publishFileSelect.addEventListener('change', updatePublishPanel);
publishCopyButton.addEventListener('click', async () => {
  const selected = getPublishOutput();
  if (!selected) return;
  try {
    await navigator.clipboard.writeText(selected.output.html);
  } catch {
    const temporary = document.createElement('textarea');
    temporary.value = selected.output.html;
    temporary.style.position = 'fixed';
    temporary.style.opacity = '0';
    document.body.append(temporary);
    temporary.select();
    document.execCommand('copy');
    temporary.remove();
  }
  publishCopyButton.textContent = 'คัดลอกแล้ว';
  window.setTimeout(() => { publishCopyButton.textContent = 'คัดลอก HTML'; }, 1600);
});

publishDownloadButton.addEventListener('click', () => {
  const selected = getPublishOutput();
  if (!selected) return;
  const suffix = selected.output.locale.toLowerCase();
  downloadText(`${selected.result.file.name.replace(/\.docx$/i, '')}-${suffix}.html`, selected.output.html);
});

function downloadText(filename, value) {
  const blob = new Blob([value], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

fileInput.addEventListener('change', event => {
  addFiles(event.target.files || []);
  fileInput.value = '';
});

for (const eventName of ['dragenter', 'dragover']) {
  dropzone.addEventListener(eventName, event => { event.preventDefault(); dropzone.classList.add('dragging'); });
}
for (const eventName of ['dragleave', 'drop']) {
  dropzone.addEventListener(eventName, event => { event.preventDefault(); dropzone.classList.remove('dragging'); });
}
dropzone.addEventListener('drop', event => addFiles(event.dataTransfer?.files || []));
dropzone.addEventListener('keydown', event => {
  if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); fileInput.click(); }
});

convertButton.addEventListener('click', async () => {
  if (!selectedFiles.length) return;
  convertButton.disabled = true;
  setStatus('กำลังอ่านไฟล์ในเบราว์เซอร์…');
  const completed = [];
  const failures = [];
  for (let index = 0; index < selectedFiles.length; index++) {
    const file = selectedFiles[index];
    setStatus(`กำลังอ่าน ${index + 1}/${selectedFiles.length}: ${file.name}`);
    try {
      const data = await readDocument(file, headlineControl.checked);
      completed.push({ file, data });
    } catch (error) {
      failures.push(`${file.name}: ${error.message || 'อ่านไฟล์ไม่สำเร็จ'}`);
    }
  }
  convertButton.disabled = selectedFiles.length === 0;
  if (completed.length) renderResults(completed);
  if (failures.length) {
    setStatus(`${completed.length} ไฟล์แปลงแล้ว · ${failures.join(' | ')}`, true);
  } else if (completed.length) {
    setStatus(`แปลงเสร็จ ${completed.length} ไฟล์ · ตรวจผลลัพธ์ก่อนนำไปใช้`);
  } else {
    setStatus(failures.join(' | ') || 'ไม่พบข้อความข่าวในไฟล์', true);
    resultSection.hidden = true;
  }
});

clearButton.addEventListener('click', () => {
  selectedFiles = [];
  updateFileView();
  resetResults();
  setStatus('ล้างไฟล์และผลลัพธ์แล้ว');
});
