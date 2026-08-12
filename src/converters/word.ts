import JSZip from 'jszip';
import { jsPDF } from 'jspdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const all = (node: Document | Element, name: string) => [...node.getElementsByTagNameNS(W, name)];
const value = (node: Element | undefined, name = 'val') => node?.getAttributeNS(W, name) ?? node?.getAttribute(`w:${name}`) ?? '';
const twips = (input: string, fallback = 0) => input ? Number(input) / 15 : fallback;

function styleRun(run: Element, span: HTMLSpanElement) {
  const properties = all(run, 'rPr')[0];
  if (!properties) return;
  if (all(properties, 'b').length) span.style.fontWeight = '700';
  if (all(properties, 'i').length) span.style.fontStyle = 'italic';
  if (all(properties, 'u').length) span.style.textDecoration = 'underline';
  const color = value(all(properties, 'color')[0]);
  if (color && color !== 'auto') span.style.color = `#${color}`;
  const size = Number(value(all(properties, 'sz')[0]));
  if (size) span.style.fontSize = `${size / 2}pt`;
  const fonts = all(properties, 'rFonts')[0];
  const font = fonts && (value(fonts, 'eastAsia') || value(fonts, 'ascii'));
  if (font) span.style.fontFamily = `'${font}', sans-serif`;
  const highlight = value(all(properties, 'highlight')[0]);
  if (highlight && highlight !== 'none') span.style.backgroundColor = highlight;
}

async function renderRun(run: Element, zip: JSZip, relations: Map<string, string>, urls: string[]) {
  const span = document.createElement('span');
  styleRun(run, span);
  for (const child of [...run.children]) {
    if (child.localName === 't') span.append(document.createTextNode(child.textContent ?? ''));
    if (child.localName === 'tab') span.append(document.createTextNode('\t'));
    if (child.localName === 'br') span.append(document.createElement('br'));
  }
  const blip = run.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')[0];
  const id = blip?.getAttributeNS(R, 'embed') ?? blip?.getAttribute('r:embed');
  const target = id && relations.get(id);
  const imageFile = target && zip.file(`word/${target.replace(/^\.\//, '')}`);
  if (imageFile) {
    const url = URL.createObjectURL(await imageFile.async('blob'));
    urls.push(url);
    const image = document.createElement('img'); image.src = url;
    const extent = run.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent')[0];
    if (extent) { image.style.width = `${Number(extent.getAttribute('cx')) / 9525}px`; image.style.height = `${Number(extent.getAttribute('cy')) / 9525}px`; }
    span.append(image);
  }
  return span;
}

async function renderParagraph(node: Element, zip: JSZip, relations: Map<string, string>, urls: string[]) {
  const paragraph = document.createElement('p');
  const properties = all(node, 'pPr')[0];
  const alignment = value(properties && all(properties, 'jc')[0]);
  paragraph.style.textAlign = ({ center: 'center', right: 'right', both: 'justify' } as Record<string, string>)[alignment] ?? 'left';
  const spacing = properties && all(properties, 'spacing')[0];
  paragraph.style.marginTop = `${twips(value(spacing, 'before'))}px`;
  paragraph.style.marginBottom = `${twips(value(spacing, 'after'), 8)}px`;
  if (value(spacing, 'line')) paragraph.style.lineHeight = String(Number(value(spacing, 'line')) / 240);
  const indent = properties && all(properties, 'ind')[0];
  paragraph.style.marginLeft = `${twips(value(indent, 'left'))}px`;
  paragraph.style.marginRight = `${twips(value(indent, 'right'))}px`;
  paragraph.style.textIndent = `${twips(value(indent, 'firstLine')) - twips(value(indent, 'hanging'))}px`;
  if (properties && all(properties, 'pageBreakBefore').length) paragraph.style.breakBefore = 'page';
  const styleName = value(properties && all(properties, 'pStyle')[0]);
  if (/heading|見出し/i.test(styleName)) { const level = Number(styleName.match(/\d+/)?.[0] ?? 1); paragraph.style.fontSize = `${Math.max(14, 24 - level * 2)}pt`; paragraph.style.fontWeight = '700'; }
  for (const child of [...node.children]) {
    if (child.localName === 'r') paragraph.append(await renderRun(child, zip, relations, urls));
    else if (child.localName === 'hyperlink') for (const run of [...child.children].filter(item => item.localName === 'r')) paragraph.append(await renderRun(run, zip, relations, urls));
  }
  if (!paragraph.textContent && !paragraph.querySelector('img')) paragraph.append(document.createElement('br'));
  return paragraph;
}

async function renderTable(node: Element, zip: JSZip, relations: Map<string, string>, urls: string[]) {
  const table = document.createElement('table'); table.style.width = '100%'; table.style.borderCollapse = 'collapse';
  for (const rowNode of [...node.children].filter(child => child.localName === 'tr')) {
    const row = table.insertRow();
    for (const cellNode of [...rowNode.children].filter(child => child.localName === 'tc')) {
      const cell = row.insertCell(); cell.style.border = '1px solid #777'; cell.style.padding = '4px'; cell.style.verticalAlign = 'top';
      const span = Number(value(all(cellNode, 'gridSpan')[0])); if (span) cell.colSpan = span;
      const fill = value(all(cellNode, 'shd')[0], 'fill'); if (fill && fill !== 'auto') cell.style.backgroundColor = `#${fill}`;
      for (const paragraph of [...cellNode.children].filter(child => child.localName === 'p')) cell.append(await renderParagraph(paragraph, zip, relations, urls));
    }
  }
  return table;
}

export async function wordToPdf(file: File, onStatus?: (message: string) => void) {
  const urls: string[] = []; let host: HTMLDivElement | undefined;
  try {
    onStatus?.('Word文書を解析しています…');
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const xml = await zip.file('word/document.xml')?.async('text');
    if (!xml) throw new Error('DOCX文書データが見つかりません');
    const parse = (text: string) => new DOMParser().parseFromString(text, 'application/xml');
    const doc = parse(xml); if (doc.querySelector('parsererror')) throw new Error('文書データが壊れています');
    const relations = new Map<string, string>();
    const relationXml = await zip.file('word/_rels/document.xml.rels')?.async('text');
    if (relationXml) for (const relation of [...parse(relationXml).getElementsByTagName('Relationship')]) relations.set(relation.getAttribute('Id')!, relation.getAttribute('Target')!);
    const section = all(doc, 'sectPr').at(-1); const size = section && all(section, 'pgSz')[0]; const margin = section && all(section, 'pgMar')[0];
    let width = twips(value(size, 'w'), 794), height = twips(value(size, 'h'), 1123);
    if (value(size, 'orient') === 'landscape' && width < height) [width, height] = [height, width];
    host = document.createElement('div'); host.className = 'word-pdf-render';
    Object.assign(host.style, { position: 'fixed', left: '-20000px', top: '0', width: `${width}px`, padding: `${twips(value(margin, 'top'), 76)}px ${twips(value(margin, 'right'), 76)}px ${twips(value(margin, 'bottom'), 76)}px ${twips(value(margin, 'left'), 76)}px`, background: '#fff', color: '#000', fontFamily: "'Calibri','Noto Sans JP',sans-serif", fontSize: '11pt', lineHeight: '1.25' });
    const body = all(doc, 'body')[0];
    for (const child of [...body.children]) { if (child.localName === 'p') host.append(await renderParagraph(child, zip, relations, urls)); else if (child.localName === 'tbl') host.append(await renderTable(child, zip, relations, urls)); }
    document.body.append(host); await document.fonts?.ready; await Promise.all([...host.querySelectorAll('img')].map(image => image.decode().catch(() => undefined)));
    const { default: html2canvas } = await import('html2canvas'); const canvas = await html2canvas(host, { scale: 1.5, backgroundColor: '#fff', logging: false });
    const pageHeight = Math.round(height * 1.5), pages = Math.max(1, Math.ceil(canvas.height / pageHeight)); const orientation = width > height ? 'landscape' : 'portrait';
    const pdf = new jsPDF({ unit: 'px', format: [width, height], orientation, compress: true, hotfixes: ['px_scaling'] });
    for (let index = 0; index < pages; index++) { onStatus?.(`ページを変換しています… ${index + 1}/${pages}`); if (index) pdf.addPage([width, height], orientation); const slice = document.createElement('canvas'); slice.width = canvas.width; slice.height = pageHeight; const context = slice.getContext('2d')!; context.fillStyle = '#fff'; context.fillRect(0, 0, slice.width, slice.height); context.drawImage(canvas, 0, index * pageHeight, canvas.width, pageHeight, 0, 0, canvas.width, pageHeight); pdf.addImage(slice.toDataURL('image/jpeg', .92), 'JPEG', 0, 0, width, height); }
    canvas.width = canvas.height = 0; return pdf.output('blob');
  } catch (error) { const message = error instanceof Error ? error.message : ''; throw new Error(`Wordファイルを変換できませんでした${message ? `（${message}）` : ''}`); }
  finally { host?.remove(); urls.forEach(URL.revokeObjectURL); }
}
