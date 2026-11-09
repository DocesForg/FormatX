// Markdown-to-WordprocessingML converter. Runs entirely in the browser.
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function xmlEscape(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function markdownInlineRuns(value) {
  const parts = [];
  const pattern = /(\*\*|__)(.+?)\1|(\*|_)(.+?)\3|~~(.+?)~~|`([^`]+)`/g;
  let last = 0;
  let match;
  while ((match = pattern.exec(value))) {
    if (match.index > last) parts.push({ text: value.slice(last, match.index) });
    if (match[2]) parts.push({ text: match[2], bold: true });
    else if (match[4]) parts.push({ text: match[4], italic: true });
    else if (match[5]) parts.push({ text: match[5], strike: true });
    else if (match[6]) parts.push({ text: match[6] });
    last = pattern.lastIndex;
  }
  if (last < value.length) parts.push({ text: value.slice(last) });
  if (!parts.length) parts.push({ text: '' });
  return parts.map(part => {
    const properties = [
      part.bold ? '<w:b/>' : '',
      part.italic ? '<w:i/>' : '',
      part.strike ? '<w:strike/>' : ''
    ].join('');
    return '<w:r>' + (properties ? '<w:rPr>' + properties + '</w:rPr>' : '') +
      '<w:t xml:space="preserve">' + xmlEscape(part.text) + '</w:t></w:r>';
  }).join('');
}

function markdownParagraph(text, options = {}) {
  const pPr = [];
  if (options.style) pPr.push('<w:pStyle w:val="' + options.style + '"/>');
  if (options.center) pPr.push('<w:jc w:val="center"/>');
  if (options.keepNext) pPr.push('<w:keepNext/>');
  const content = options.plain ? '<w:r><w:t xml:space="preserve">' + xmlEscape(text) + '</w:t></w:r>' : markdownInlineRuns(text);
  return '<w:p>' + (pPr.length ? '<w:pPr>' + pPr.join('') + '</w:pPr>' : '') + content + '</w:p>';
}

function markdownTableRow(line) {
  return line.trim().replace(/^\|/, '').replace(/\|$/, '').split(/(?<!\\)\|/).map(cell => cell.replace(/\\\|/g, '|').trim());
}

function markdownTableXml(rows) {
  const columnCount = Math.max(1, ...rows.map(row => row.length));
  const borders = '<w:tblBorders>' + ['top','left','bottom','right','insideH','insideV'].map(side =>
    '<w:' + side + ' w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
  ).join('') + '</w:tblBorders>';
  const grid = '<w:tblGrid>' + Array.from({length: columnCount}, () => '<w:gridCol w:w="' + Math.floor(9000 / columnCount) + '"/>').join('') + '</w:tblGrid>';
  const body = rows.map((row, rowIndex) => '<w:tr>' + Array.from({length: columnCount}, (_, i) => {
    const value = row[i] || '';
    const header = rowIndex === 0;
    return '<w:tc><w:tcPr><w:tcW w:w="' + Math.floor(9000 / columnCount) + '" w:type="dxa"/></w:tcPr>' +
      '<w:p>' + (header ? '<w:pPr><w:jc w:val="center"/></w:pPr>' : '') +
      (header ? '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + xmlEscape(value) + '</w:t></w:r>' : markdownInlineRuns(value)) +
      '</w:p></w:tc>';
  }).join('') + '</w:tr>').join('');
  return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + borders + '</w:tblPr>' + grid + body + '</w:tbl>';
}

function markdownToWordXml(markdown) {
  const lines = String(markdown || '').replace(/\r\n?/g, '\n').split('\n');
  const output = [];
  let paragraph = [];
  let listMode = null;
  const flushParagraph = () => {
    if (paragraph.length) output.push(markdownParagraph(paragraph.join(' ')));
    paragraph = [];
  };
  const flushList = () => { listMode = null; };
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    if (/^\s*```/.test(line) || /^\s*~~~/.test(line)) {
      flushParagraph(); flushList();
      const fence = line.trim().slice(0, 3);
      const code = [];
      i += 1;
      while (i < lines.length && !new RegExp('^\\s*' + fence + '+\\s*$').test(lines[i])) {
        code.push(lines[i]);
        i += 1;
      }
      // Code fences intentionally become regular paragraphs, not shaded/code-styled blocks.
      if (code.length) code.forEach(codeLine => output.push(markdownParagraph(codeLine, { plain: true })));
      else output.push(markdownParagraph('', { plain: true }));
      continue;
    }
    if (!line.trim()) { flushParagraph(); flushList(); continue; }
    if (/^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(line)) continue;
    if (line.includes('|') && i + 1 < lines.length && /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(lines[i + 1])) {
      flushParagraph(); flushList();
      const rows = [markdownTableRow(line)];
      i += 2;
      while (i < lines.length && lines[i].includes('|')) { rows.push(markdownTableRow(lines[i])); i += 1; }
      i -= 1;
      output.push(markdownTableXml(rows));
      continue;
    }
    const heading = line.match(/^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/);
    if (heading) {
      flushParagraph(); flushList();
      output.push(markdownParagraph(heading[2], { style: 'Heading' + heading[1].length, keepNext: true }));
      continue;
    }
    if (/^\s*(---+|\*\*\*+|___+)\s*$/.test(line)) {
      flushParagraph(); flushList();
      output.push(markdownParagraph(''));
      continue;
    }
    const unordered = line.match(/^\s*[-+*]\s+(.+)$/);
    const ordered = line.match(/^\s*\d+[.)]\s+(.+)$/);
    if (unordered || ordered) {
      flushParagraph();
      const mode = unordered ? 'unordered' : 'ordered';
      if (listMode !== mode) listMode = mode;
      output.push(markdownParagraph((unordered ? '• ' : (line.match(/^\s*(\d+)[.)]/)[1] + '. ')) + (unordered || ordered)[1]));
      continue;
    }
    if (/^\s*>/.test(line)) {
      flushParagraph(); flushList();
      output.push(markdownParagraph(line.replace(/^\s*>\s?/, '')));
      continue;
    }
    flushList();
    paragraph.push(line.trim());
  }
  flushParagraph();
  return '<w:document xmlns:w="' + WORD_NS + '"><w:body>' + output.join('') + '</w:body></w:document>';
}

window.markdownToWordXml = markdownToWordXml;
