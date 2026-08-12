import JSZip from 'jszip';
import { jsPDF } from 'jspdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const all = (node: Document | Element, name: string) => [...node.getElementsByTagNameNS(W, name)];
const value = (node: Element | undefined, name = 'val') => node?.getAttributeNS(W, name) ?? node?.getAttribute(`w:${name}`) ?? '';
const twips = (input: string, fallback = 0) => input ? Number(input) / 15 : fallback;
type WordStyles = Map<string, Element>;
type Numbering = { formats: Map<string, string>; counters: Map<string, number> };

function resolveStyle(id: string, styles: WordStyles, seen = new Set<string>()): Element[] {
  if (!id || seen.has(id)) return [];
  seen.add(id);
  const style = styles.get(id);
  if (!style) return [];
  const parent = value(all(style, 'basedOn')[0]);
  return [...resolveStyle(parent, styles, seen), style];
}

function applyRunProperties(properties: Element | undefined, span: HTMLElement) {
  if (!properties) return;
  const toggle = (name: string) => {
    const node = all(properties, name)[0];
    return Boolean(node && !['0', 'false', 'off', 'none'].includes(value(node).toLowerCase()));
  };
  if (toggle('b')) span.style.fontWeight = '700';
  if (toggle('i')) span.style.fontStyle = 'italic';
  if (toggle('u')) span.style.textDecoration = 'underline';
  if (toggle('strike')) span.style.textDecoration = `${span.style.textDecoration} line-through`.trim();
  const color = value(all(properties, 'color')[0]);
  if (color && color !== 'auto') span.style.color = `#${color}`;
  const size = Number(value(all(properties, 'sz')[0]));
  if (size) span.style.fontSize = `${size / 2}pt`;
  const fonts = all(properties, 'rFonts')[0];
  const font = fonts && (value(fonts, 'eastAsia') || value(fonts, 'ascii') || value(fonts, 'hAnsi'));
  if (font) span.style.fontFamily = `'${font}', 'Noto Sans JP', sans-serif`;
  const highlight = value(all(properties, 'highlight')[0]);
  if (highlight && highlight !== 'none') span.style.backgroundColor = highlight;
  const shading = value(all(properties, 'shd')[0], 'fill');
  if (shading && shading !== 'auto') span.style.backgroundColor = `#${shading}`;
}

function styleRun(run: Element, span: HTMLSpanElement, styles: WordStyles) {
  const properties = all(run, 'rPr')[0];
  const styleId = value(properties && all(properties, 'rStyle')[0]);
  for (const style of resolveStyle(styleId, styles)) applyRunProperties(all(style, 'rPr')[0], span);
  applyRunProperties(properties, span);
}

async function renderRun(run: Element, zip: JSZip, relations: Map<string, string>, urls: string[], styles: WordStyles) {
  const span = document.createElement('span');
  styleRun(run, span, styles);
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

async function renderParagraph(node: Element, zip: JSZip, relations: Map<string, string>, urls: string[], styles: WordStyles, numbering: Numbering) {
  const paragraph = document.createElement('p');
  const properties = all(node, 'pPr')[0];
  const styleName = value(properties && all(properties, 'pStyle')[0]);
  const inherited = resolveStyle(styleName, styles);
  for (const style of inherited) {
    const runProperties = all(style, 'rPr')[0];
    applyRunProperties(runProperties, paragraph);
    const paragraphProperties = all(style, 'pPr')[0];
    const styleSpacing = paragraphProperties && all(paragraphProperties, 'spacing')[0];
    if (value(styleSpacing, 'before')) paragraph.style.marginTop = `${twips(value(styleSpacing, 'before'))}px`;
    if (value(styleSpacing, 'after')) paragraph.style.marginBottom = `${twips(value(styleSpacing, 'after'))}px`;
    if (value(styleSpacing, 'line')) paragraph.style.lineHeight = String(Number(value(styleSpacing, 'line')) / 240);
    const outline = Number(value(paragraphProperties && all(paragraphProperties, 'outlineLvl')[0]));
    if (Number.isFinite(outline) && outline >= 0) { paragraph.style.fontWeight = '700'; paragraph.style.breakAfter = 'avoid'; }
  }
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
  const numProperties = properties && all(properties, 'numPr')[0];
  if (numProperties) {
    const numId = value(all(numProperties, 'numId')[0]);
    const level = value(all(numProperties, 'ilvl')[0]) || '0';
    const key = `${numId}:${level}`;
    const format = numbering.formats.get(key) ?? 'decimal';
    const count = (numbering.counters.get(key) ?? 0) + 1;
    numbering.counters.set(key, count);
    const marker = format === 'bullet' ? '•' : format === 'lowerLetter' ? `${String.fromCharCode(96 + count)}.` : `${count}.`;
    const label = document.createElement('span');
    label.textContent = `${marker} `; label.style.display = 'inline-block'; label.style.minWidth = '2em'; label.style.textIndent = '0';
    paragraph.append(label);
  }
  if (/heading|見出し/i.test(styleName)) { const level = Number(styleName.match(/\d+/)?.[0] ?? 1); paragraph.style.fontSize = `${Math.max(14, 24 - level * 2)}pt`; paragraph.style.fontWeight = '700'; }
  for (const child of [...node.children]) {
    if (child.localName === 'r') paragraph.append(await renderRun(child, zip, relations, urls, styles));
    else if (child.localName === 'hyperlink') for (const run of [...child.children].filter(item => item.localName === 'r')) paragraph.append(await renderRun(run, zip, relations, urls, styles));
  }
  if (!paragraph.textContent && !paragraph.querySelector('img')) paragraph.append(document.createElement('br'));
  return paragraph;
}

async function renderTable(node: Element, zip: JSZip, relations: Map<string, string>, urls: string[], styles: WordStyles, numbering: Numbering) {
  const table = document.createElement('table'); table.style.width = '100%'; table.style.borderCollapse = 'collapse';
  for (const rowNode of [...node.children].filter(child => child.localName === 'tr')) {
    const row = table.insertRow();
    for (const cellNode of [...rowNode.children].filter(child => child.localName === 'tc')) {
      const cell = row.insertCell(); cell.style.border = '1px solid #777'; cell.style.padding = '4px'; cell.style.verticalAlign = 'top';
      const span = Number(value(all(cellNode, 'gridSpan')[0])); if (span) cell.colSpan = span;
      const fill = value(all(cellNode, 'shd')[0], 'fill'); if (fill && fill !== 'auto') cell.style.backgroundColor = `#${fill}`;
      for (const paragraph of [...cellNode.children].filter(child => child.localName === 'p')) cell.append(await renderParagraph(paragraph, zip, relations, urls, styles, numbering));
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
    const styles: WordStyles = new Map();
    const stylesXml = await zip.file('word/styles.xml')?.async('text');
    if (stylesXml) for (const style of all(parse(stylesXml), 'style')) styles.set(value(style, 'styleId'), style);
    const numbering: Numbering = { formats: new Map(), counters: new Map() };
    const numberingXml = await zip.file('word/numbering.xml')?.async('text');
    if (numberingXml) {
      const numberingDoc = parse(numberingXml);
      const abstractFormats = new Map<string, Map<string, string>>();
      for (const abstract of all(numberingDoc, 'abstractNum')) {
        const formats = new Map<string, string>();
        for (const level of all(abstract, 'lvl')) formats.set(value(level, 'ilvl'), value(all(level, 'numFmt')[0]));
        abstractFormats.set(value(abstract, 'abstractNumId'), formats);
      }
      for (const num of all(numberingDoc, 'num')) {
        const numId = value(num, 'numId');
        const formats = abstractFormats.get(value(all(num, 'abstractNumId')[0]));
        formats?.forEach((format, level) => numbering.formats.set(`${numId}:${level}`, format));
      }
    }
    const section = all(doc, 'sectPr').at(-1); const size = section && all(section, 'pgSz')[0]; const margin = section && all(section, 'pgMar')[0];
    let width = twips(value(size, 'w'), 794), height = twips(value(size, 'h'), 1123);
    if (value(size, 'orient') === 'landscape' && width < height) [width, height] = [height, width];
    host = document.createElement('div'); host.className = 'word-pdf-render';
    Object.assign(host.style, { boxSizing: 'border-box', position: 'fixed', left: '-20000px', top: '0', width: `${width}px`, padding: `${twips(value(margin, 'top'), 76)}px ${twips(value(margin, 'right'), 76)}px ${twips(value(margin, 'bottom'), 76)}px ${twips(value(margin, 'left'), 76)}px`, background: '#fff', color: '#000', fontFamily: "'Calibri','Noto Sans JP',sans-serif", fontSize: '11pt', lineHeight: '1.25', overflowWrap: 'break-word' });
    const body = all(doc, 'body')[0];
    for (const child of [...body.children]) { if (child.localName === 'p') host.append(await renderParagraph(child, zip, relations, urls, styles, numbering)); else if (child.localName === 'tbl') host.append(await renderTable(child, zip, relations, urls, styles, numbering)); }
    document.body.append(host); await document.fonts?.ready; await Promise.all([...host.querySelectorAll('img')].map(image => image.decode().catch(() => undefined)));
    const { default: html2canvas } = await import('html2canvas'); const canvas = await html2canvas(host, { scale: 1.5, backgroundColor: '#fff', logging: false });
    const pageHeight = Math.round(height * 1.5), pages = Math.max(1, Math.ceil(canvas.height / pageHeight)); const orientation = width > height ? 'landscape' : 'portrait';
    const pdf = new jsPDF({ unit: 'px', format: [width, height], orientation, compress: true, hotfixes: ['px_scaling'] });
    for (let index = 0; index < pages; index++) { onStatus?.(`ページを変換しています… ${index + 1}/${pages}`); if (index) pdf.addPage([width, height], orientation); const slice = document.createElement('canvas'); slice.width = canvas.width; slice.height = pageHeight; const context = slice.getContext('2d')!; context.fillStyle = '#fff'; context.fillRect(0, 0, slice.width, slice.height); context.drawImage(canvas, 0, index * pageHeight, canvas.width, pageHeight, 0, 0, canvas.width, pageHeight); pdf.addImage(slice.toDataURL('image/jpeg', .92), 'JPEG', 0, 0, width, height); }
    canvas.width = canvas.height = 0; return pdf.output('blob');
  } catch (error) { const message = error instanceof Error ? error.message : ''; throw new Error(`Wordファイルを変換できませんでした${message ? `（${message}）` : ''}`); }
  finally { host?.remove(); urls.forEach(URL.revokeObjectURL); }
}
