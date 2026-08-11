import * as XLSX from 'xlsx';
import JSZip from 'jszip';

// Excelブックを「描画に必要な情報だけ」に落とし込む。
// .xlsx / .xlsm は ExcelJS で読み、書式・セル結合・画像・印刷設定まで取得する。
// .xls（旧形式）は ExcelJS が扱えないため SheetJS で値だけを読む。
export type Edge = { width: number; color: string; dash?: number[]; double?: boolean };
export type Borders = { top?: Edge; right?: Edge; bottom?: Edge; left?: Edge };
export type CellStyle = {
  fill?: string;
  color: string;
  fontSize: number;
  fontName: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  hAlign: 'left' | 'center' | 'right' | 'general';
  vAlign: 'top' | 'middle' | 'bottom';
  wrap: boolean;
  rotation: number | 'vertical';
  indent: number;
  borders: Borders;
};
export type CellAnchor = { kind: 'anchor'; text: string; numeric: boolean; rowSpan: number; colSpan: number; style: CellStyle };
export type CellCovered = { kind: 'covered'; row: number; col: number };
export type Cell = CellAnchor | CellCovered;
export type Picture = { source: ImageBitmap; col: number; row: number; toCol?: number; toRow?: number; width?: number; height?: number };
export type ShapeText = { text: string; fontSize: number; bold: boolean; color: string; hAlign: 'left' | 'center' | 'right'; vAlign: 'top' | 'middle' | 'bottom' };
export type Shape = {
  kind: 'rect' | 'line';
  col: number; row: number; toCol: number; toRow: number;
  flipH?: boolean; flipV?: boolean;
  fill?: string;
  stroke?: string;
  strokeWidth: number;
  dash?: number[];
  text?: ShapeText;
};
export type PrintSetup = {
  landscape?: boolean;
  gridLines: boolean;
  margins?: { left: number; right: number; top: number; bottom: number }; // インチ
  scale?: number; // 100 = 等倍
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  horizontalCentered?: boolean;
  titleRows?: [number, number]; // 印刷タイトルの行範囲（1始まり）
  titleCols?: [number, number]; // 印刷タイトルの列範囲（1始まり）
  rowBreaks: number[]; // この行の後で改ページする（シート上の行番号・1始まり）
  colBreaks: number[]; // この列の後で改ページする（シート上の列番号・1始まり）
};
export type Sheet = {
  name: string;
  firstRow: number; // シート上の行番号（1始まり）。改ページ位置の対応に使う
  firstCol: number;
  cols: number[];
  rows: number[];
  cells: Cell[][];
  pictures: Picture[];
  shapes: Shape[];
  print: PrintSetup;
};

const DEFAULT_ROW_PT = 15;
export const DEFAULT_COL_PX = 64;
export const DEFAULT_ROW_PX = 20;
const PT_TO_PX = 4 / 3;

// Officeの既定テーマ色（tintは無視する）
const THEME_COLORS = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];

const charsToPx = (chars: number) => Math.round(chars * 7 + 5);
const pointsToPx = (points: number) => Math.round(points * PT_TO_PX);

// 古いExcelで作られたファイルは色をパレット番号で持っている（Excel既定の56色）
const INDEXED_COLORS = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
  '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
  '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
  '#000000', '#ffffff',
];

const argbToCss = (argb: string) => (argb.length === 8 ? `#${argb.slice(2)}` : argb.startsWith('#') ? argb : `#${argb}`);

// テーマ色の濃淡（tint）を適用する
function applyTint(css: string, tint: number | undefined) {
  if (!tint) return css;
  const channels = [1, 3, 5].map(index => parseInt(css.slice(index, index + 2), 16));
  const shifted = channels.map(value => Math.round(tint > 0 ? value + (255 - value) * tint : value * (1 + tint)));
  return `#${shifted.map(value => Math.max(0, Math.min(255, value)).toString(16).padStart(2, '0')).join('')}`;
}

type ExcelColor = { argb?: string; theme?: number; indexed?: number; tint?: number } | undefined;
const colorOf = (color: ExcelColor, fallback: string) => {
  if (!color) return fallback;
  if (color.argb) return applyTint(argbToCss(color.argb), color.tint);
  if (typeof color.theme === 'number') return applyTint(THEME_COLORS[color.theme] ?? fallback, color.tint);
  if (typeof color.indexed === 'number') return applyTint(INDEXED_COLORS[color.indexed] ?? fallback, color.tint);
  return fallback;
};

const BORDER_WIDTH: Record<string, number> = { hair: 0.5, thin: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, slantDashDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, thick: 3, double: 3 };
const BORDER_DASH: Record<string, number[]> = { dotted: [1, 2], dashed: [4, 3], dashDot: [6, 2, 2, 2], dashDotDot: [6, 2, 2, 2, 2, 2], mediumDashed: [5, 3], mediumDashDot: [7, 3, 3, 3], mediumDashDotDot: [7, 3, 3, 3, 3, 3], slantDashDot: [5, 3] };

type ExcelFill = { type?: string; pattern?: string; fgColor?: ExcelColor; bgColor?: ExcelColor; stops?: { position: number; color?: ExcelColor }[] } | undefined;
const fillColorOf = (fill: ExcelFill) => {
  if (!fill) return undefined;
  if (fill.type === 'gradient') {
    const stop = fill.stops?.[0]?.color ?? fill.stops?.[1]?.color;
    return stop ? colorOf(stop, '#ffffff') : undefined;
  }
  if (!fill.pattern || fill.pattern === 'none') return undefined;
  const color = fill.fgColor ?? fill.bgColor;
  return color ? colorOf(color, '#ffffff') : undefined;
};

type ExcelBorderEdge = { style?: string; color?: ExcelColor } | undefined;
const edgeOf = (border: ExcelBorderEdge): Edge | undefined => {
  if (!border?.style) return undefined;
  return { width: BORDER_WIDTH[border.style] ?? 1, color: colorOf(border.color, '#000000'), dash: BORDER_DASH[border.style], double: border.style === 'double' };
};

const dateToSerial = (date: Date) => date.getTime() / 86_400_000 + 25_569;

function formatNumber(value: number, numFmt: string | undefined) {
  if (!numFmt || numFmt === 'General') return String(value);
  try { return XLSX.SSF.format(numFmt, value); } catch { return String(value); }
}

type ExcelValue = unknown;
function valueToText(value: ExcelValue, numFmt: string | undefined): { text: string; numeric: boolean } {
  if (value === null || value === undefined) return { text: '', numeric: false };
  if (typeof value === 'number') return { text: formatNumber(value, numFmt), numeric: true };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', numeric: false };
  if (typeof value === 'string') return { text: value, numeric: false };
  if (value instanceof Date) return { text: formatNumber(dateToSerial(value), numFmt ?? 'yyyy/mm/dd'), numeric: true };
  const record = value as Record<string, unknown>;
  if (Array.isArray(record.richText)) return { text: (record.richText as { text?: string }[]).map(run => run.text ?? '').join(''), numeric: false };
  if ('result' in record) return valueToText(record.result, numFmt);
  if ('text' in record) return { text: String(record.text ?? ''), numeric: false };
  if ('error' in record) return { text: String(record.error), numeric: false };
  if ('formula' in record || 'sharedFormula' in record) return { text: '', numeric: false };
  return { text: String(value), numeric: false };
}

// 'Sheet1'!$A$1:$J$40 や A1:J40 から範囲だけ取り出す
function parseRange(area: string) {
  const body = area.split('!').pop()?.split(',')[0]?.replace(/\$/g, '').trim();
  if (!body) return undefined;
  try {
    const range = XLSX.utils.decode_range(body.includes(':') ? body : `${body}:${body}`);
    return { top: range.s.r + 1, left: range.s.c + 1, bottom: range.e.r + 1, right: range.e.c + 1 };
  } catch { return undefined; }
}

// ExcelJSは手動の改ページと図形を返さないため、シートのXMLから直接読む
type SheetExtras = { rows: number[]; cols: number[]; drawing?: string; titleRows?: [number, number]; titleCols?: [number, number] };

// 印刷タイトル（'Sheet1'!$1:$4,'Sheet1'!$A:$B）を行範囲・列範囲へ分解する
function parseTitles(value: string) {
  let titleRows: [number, number] | undefined;
  let titleCols: [number, number] | undefined;
  for (const part of value.split(',')) {
    const body = part.split('!').pop()?.replace(/\$/g, '').trim();
    if (!body) continue;
    const rows = body.match(/^(\d+):(\d+)$/);
    if (rows) { titleRows = [Number(rows[1]), Number(rows[2])]; continue; }
    const cols = body.match(/^([A-Z]+):([A-Z]+)$/i);
    if (cols) {
      const toIndex = (letters: string) => [...letters.toUpperCase()].reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
      titleCols = [toIndex(cols[1]), toIndex(cols[2])];
    }
  }
  return { titleRows, titleCols };
}
const unescapeXml = (value: string) => value.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)));
const relativeTo = (base: string, target: string) => {
  if (target.startsWith('/')) return target.replace(/^\//, '');
  const parts = base.split('/').slice(0, -1);
  for (const piece of target.split('/')) {
    if (piece === '..') parts.pop();
    else if (piece !== '.') parts.push(piece);
  }
  return parts.join('/');
};

async function readSheetExtras(buffer: ArrayBuffer): Promise<Map<string, SheetExtras>> {
  const result = new Map<string, SheetExtras>();
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
    if (!workbookXml || !relsXml) return result;
    const targets = new Map<string, string>();
    for (const match of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
      targets.set(match[1], relativeTo('xl/workbook.xml', match[2]));
    }
    // 印刷タイトルはブック側の定義名にシート番号付きで入っている
    const titles = new Map<number, { titleRows?: [number, number]; titleCols?: [number, number] }>();
    for (const match of workbookXml.matchAll(/<definedName name="_xlnm\.Print_Titles" localSheetId="(\d+)"[^>]*>([^<]*)<\/definedName>/g)) {
      titles.set(Number(match[1]), parseTitles(unescapeXml(match[2])));
    }
    let sheetIndex = -1;
    for (const match of workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/>/g)) {
      sheetIndex++;
      const path = targets.get(match[2]);
      if (!path) continue;
      const sheetXml = await zip.file(path)?.async('string');
      if (!sheetXml) continue;
      const collect = (tag: string) => {
        const section = sheetXml.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`));
        if (!section) return [];
        return [...section[1].matchAll(/<brk[^>]*id="(\d+)"[^>]*\/>/g)].map(brk => Number(brk[1])).filter(id => id > 0);
      };
      let drawing: string | undefined;
      const drawingId = sheetXml.match(/<drawing[^>]*r:id="([^"]+)"/)?.[1];
      if (drawingId) {
        const sheetRels = await zip.file(path.replace(/([^/]+)$/, '_rels/$1.rels'))?.async('string');
        const drawingTarget = sheetRels?.match(new RegExp(`<Relationship[^>]*Id="${drawingId}"[^>]*Target="([^"]+)"`))?.[1];
        if (drawingTarget) drawing = await zip.file(relativeTo(path, drawingTarget))?.async('string') ?? undefined;
      }
      result.set(unescapeXml(match[1]), { rows: collect('rowBreaks'), cols: collect('colBreaks'), drawing, ...titles.get(sheetIndex) });
    }
  } catch { /* 読めなければ改ページ・図形なしとして扱う */ }
  return result;
}

const EMU_PER_PX = 9525;
const PRESET_DASH: Record<string, number[]> = { dot:[1,2], sysDot:[1,1], dash:[4,3], sysDash:[3,1], dashDot:[4,3,1,3], sysDashDot:[3,1,1,1], lgDash:[8,3], lgDashDot:[8,3,1,3], lgDashDotDot:[8,3,1,3,1,3], sysDashDotDot:[3,1,1,1,1,1] };

// 図形（四角・直線）を読む。ボタンなどの非表示図形は飛ばす
function readShapes(drawingXml: string, cols: number[], rows: number[], firstRow: number, firstCol: number): Shape[] {
  const shapes: Shape[] = [];
  const position = (sizes: number[], index: number, offsetEmu: number) => {
    if (index < 0) return 0;
    const size = sizes[index] ?? 0;
    return index + (size > 0 ? Math.min(1, offsetEmu / EMU_PER_PX / size) : 0);
  };
  for (const anchor of drawingXml.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)[^>]*>([\s\S]*?)<\/xdr:\1>/g)) {
    const body = anchor[2];
    if (!/<xdr:(sp|cxnSp)[ >]/.test(body)) continue; // 画像はExcelJS側で扱う
    if (/hidden="1"/.test(body)) continue;
    const from = body.match(/<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>\s*<xdr:colOff>(-?\d+)<\/xdr:colOff>\s*<xdr:row>(\d+)<\/xdr:row>\s*<xdr:rowOff>(-?\d+)<\/xdr:rowOff>/);
    if (!from) continue;
    const to = body.match(/<xdr:to>\s*<xdr:col>(\d+)<\/xdr:col>\s*<xdr:colOff>(-?\d+)<\/xdr:colOff>\s*<xdr:row>(\d+)<\/xdr:row>\s*<xdr:rowOff>(-?\d+)<\/xdr:rowOff>/);
    const extent = body.match(/<xdr:ext cx="(\d+)" cy="(\d+)"\/>/);
    const col = position(cols, Number(from[1]) - (firstCol - 1), Number(from[2]));
    const row = position(rows, Number(from[3]) - (firstRow - 1), Number(from[4]));
    let toCol: number;
    let toRow: number;
    if (to) {
      toCol = position(cols, Number(to[1]) - (firstCol - 1), Number(to[2]));
      toRow = position(rows, Number(to[3]) - (firstRow - 1), Number(to[4]));
    } else if (extent) {
      const widthPx = Number(extent[1]) / EMU_PER_PX;
      const heightPx = Number(extent[2]) / EMU_PER_PX;
      let remaining = widthPx;
      let cursor = Math.floor(col);
      while (remaining > 0 && cursor < cols.length) { remaining -= cols[cursor]; cursor++; }
      toCol = cursor;
      remaining = heightPx;
      cursor = Math.floor(row);
      while (remaining > 0 && cursor < rows.length) { remaining -= rows[cursor]; cursor++; }
      toRow = cursor;
    } else continue;
    if (Math.abs(toCol - col) < 0.01 && Math.abs(toRow - row) < 0.01) continue; // 大きさのない図形は飛ばす

    const geometry = body.match(/<a:prstGeom prst="([^"]+)"/)?.[1] ?? 'rect';
    const kind: Shape['kind'] = geometry === 'line' || geometry === 'straightConnector1' || /<xdr:cxnSp/.test(body) ? 'line' : 'rect';
    const spPr = body.match(/<xdr:spPr[\s\S]*?<\/xdr:spPr>/)?.[0] ?? '';
    const lineSection = spPr.match(/<a:ln[ >][\s\S]*?<\/a:ln>/)?.[0] ?? spPr.match(/<a:ln[^>]*\/>/)?.[0] ?? '';
    const fillMatch = spPr.replace(lineSection, '').match(/<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);
    const noFill = /<a:noFill\/>/.test(spPr.replace(lineSection, ''));
    const strokeColor = lineSection.match(/<a:srgbClr val="([0-9A-Fa-f]{6})"/)?.[1];
    const strokeWidthEmu = Number(lineSection.match(/<a:ln[^>]*w="(\d+)"/)?.[1] ?? spPr.match(/<a:ln[^>]*w="(\d+)"/)?.[1] ?? 9525);
    const dash = PRESET_DASH[lineSection.match(/<a:prstDash val="([^"]+)"/)?.[1] ?? ''];
    const strokeless = /<a:ln[^>]*>\s*<a:noFill\/>/.test(lineSection) || /<a:ln[^>]*w="0"/.test(lineSection);

    const runs = [...body.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(run => unescapeXml(run[1])).join('');
    const textStyle = runs.trim() ? {
      text: runs,
      fontSize: Number(body.match(/sz="(\d+)"/)?.[1] ?? 1000) / 100,
      bold: /b="1"/.test(body),
      color: `#${body.match(/<a:rPr[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/)?.[1] ?? '000000'}`,
      hAlign: (body.match(/algn="(ctr|l|r)"/)?.[1] === 'ctr' ? 'center' : body.match(/algn="(ctr|l|r)"/)?.[1] === 'r' ? 'right' : 'left') as ShapeText['hAlign'],
      vAlign: (body.match(/anchor="(t|ctr|b)"/)?.[1] === 't' ? 'top' : body.match(/anchor="(t|ctr|b)"/)?.[1] === 'b' ? 'bottom' : 'middle') as ShapeText['vAlign'],
    } : undefined;

    shapes.push({
      kind,
      col, row, toCol, toRow,
      flipH: /<a:xfrm[^>]*flipH="1"/.test(spPr),
      flipV: /<a:xfrm[^>]*flipV="1"/.test(spPr),
      fill: noFill || !fillMatch ? undefined : `#${fillMatch[1]}`,
      stroke: strokeless ? undefined : `#${strokeColor ?? '000000'}`,
      strokeWidth: Math.max(0.75, strokeWidthEmu / EMU_PER_PX),
      dash,
      text: textStyle,
    });
  }
  return shapes;
}

type ExcelJSModule = typeof import('exceljs');
type ExcelWorksheet = ReturnType<InstanceType<ExcelJSModule['Workbook']>['addWorksheet']>;

async function decodePictures(workbook: InstanceType<ExcelJSModule['Workbook']>, worksheet: ExcelWorksheet, offsetRow: number, offsetCol: number) {
  const pictures: Picture[] = [];
  for (const image of worksheet.getImages()) {
    const media = workbook.model.media?.[Number(image.imageId)];
    if (!media?.buffer) continue;
    try {
      const bytes = media.buffer as unknown as Uint8Array;
      const blob = new Blob([new Uint8Array(bytes)], { type: `image/${media.extension}` });
      const source = await createImageBitmap(blob); // EMF/WMFなどブラウザが解釈できない形式はここで失敗するので飛ばす
      const anchor = image.range as { tl?: { col: number; row: number }; br?: { col: number; row: number }; ext?: { width: number; height: number } };
      if (!anchor.tl) { source.close(); continue; }
      pictures.push({
        source,
        col: anchor.tl.col - offsetCol,
        row: anchor.tl.row - offsetRow,
        toCol: anchor.br ? anchor.br.col - offsetCol : undefined,
        toRow: anchor.br ? anchor.br.row - offsetRow : undefined,
        width: anchor.ext?.width,
        height: anchor.ext?.height,
      });
    } catch { /* 描画できない画像は飛ばす */ }
  }
  return pictures;
}

function printSetupOf(worksheet: ExcelWorksheet, extras: SheetExtras | undefined): PrintSetup {
  const setup = worksheet.pageSetup ?? {};
  const margins = setup.margins;
  return {
    horizontalCentered: !!setup.horizontalCentered,
    rowBreaks: extras?.rows ?? [],
    colBreaks: extras?.cols ?? [],
    titleRows: extras?.titleRows,
    titleCols: extras?.titleCols,
    landscape: setup.orientation === 'landscape' ? true : setup.orientation === 'portrait' ? false : undefined,
    gridLines: !!setup.showGridLines,
    margins: margins ? { left: margins.left, right: margins.right, top: margins.top, bottom: margins.bottom } : undefined,
    scale: typeof setup.scale === 'number' ? setup.scale : undefined,
    fitToPage: !!setup.fitToPage,
    fitToWidth: typeof setup.fitToWidth === 'number' ? setup.fitToWidth : undefined,
    fitToHeight: typeof setup.fitToHeight === 'number' ? setup.fitToHeight : undefined,
  };
}

async function readWithExcelJS(file: File): Promise<Sheet[]> {
  const ExcelJS = (await import('exceljs')).default;
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const extras = await readSheetExtras(buffer);
  const sheets: Sheet[] = [];

  for (const worksheet of workbook.worksheets) {
    if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') continue;
    const dimensions = worksheet.dimensions as unknown as { top: number; left: number; bottom: number; right: number } | undefined;
    // 印刷範囲が設定されていればそれを優先する
    const area = worksheet.pageSetup?.printArea ? parseRange(worksheet.pageSetup.printArea) : undefined;
    const bounds = area ?? (dimensions && dimensions.bottom >= dimensions.top ? dimensions : undefined);
    if (!bounds) continue;
    const top = Math.max(1, bounds.top);
    const left = Math.max(1, bounds.left);
    const bottom = Math.max(top, bounds.bottom);
    const right = Math.max(left, bounds.right);

    const defaultColPx = worksheet.properties?.defaultColWidth ? charsToPx(worksheet.properties.defaultColWidth) : DEFAULT_COL_PX;
    const defaultRowPx = worksheet.properties?.defaultRowHeight ? pointsToPx(worksheet.properties.defaultRowHeight) : pointsToPx(DEFAULT_ROW_PT);

    const cols: number[] = [];
    for (let c = left; c <= right; c++) {
      const column = worksheet.getColumn(c);
      cols.push(column?.hidden ? 0 : column?.width ? charsToPx(column.width) : defaultColPx);
    }
    const rows: number[] = [];
    for (let r = top; r <= bottom; r++) {
      const row = worksheet.getRow(r);
      rows.push(row?.hidden ? 0 : row?.height ? pointsToPx(row.height) : defaultRowPx);
    }

    // セル結合：外周の罫線は右下セル側に入っていることがあるので、上下左右で参照元を変える
    const spans = new Map<string, { rowSpan: number; colSpan: number; lastRow: number; lastCol: number }>();
    const covered = new Map<string, { row: number; col: number }>();
    for (const merge of worksheet.model.merges ?? []) {
      const range = parseRange(merge);
      if (!range) continue;
      if (range.bottom < top || range.top > bottom || range.right < left || range.left > right) continue;
      const mergeTop = Math.max(range.top, top);
      const mergeLeft = Math.max(range.left, left);
      const mergeBottom = Math.min(range.bottom, bottom);
      const mergeRight = Math.min(range.right, right);
      spans.set(`${mergeTop},${mergeLeft}`, { rowSpan: mergeBottom - mergeTop + 1, colSpan: mergeRight - mergeLeft + 1, lastRow: mergeBottom, lastCol: mergeRight });
      for (let r = mergeTop; r <= mergeBottom; r++) for (let c = mergeLeft; c <= mergeRight; c++) {
        if (r !== mergeTop || c !== mergeLeft) covered.set(`${r},${c}`, { row: mergeTop - top, col: mergeLeft - left });
      }
    }

    const drawing = extras.get(worksheet.name)?.drawing;
    const shapes = drawing ? readShapes(drawing, cols, rows, top, left) : [];

    const cells: Cell[][] = [];
    for (let r = top; r <= bottom; r++) {
      const line: Cell[] = [];
      for (let c = left; c <= right; c++) {
        const from = covered.get(`${r},${c}`);
        if (from) { line.push({ kind:'covered', row: from.row, col: from.col }); continue; }
        const cell = worksheet.getCell(r, c);
        const span = spans.get(`${r},${c}`);
        const font = (cell.font ?? {}) as { size?: number; name?: string; bold?: boolean; italic?: boolean; underline?: boolean | string; color?: ExcelColor };
        const alignment = (cell.alignment ?? {}) as { horizontal?: string; vertical?: string; wrapText?: boolean; textRotation?: number | 'vertical'; indent?: number };
        const fill = cell.fill as ExcelFill;
        const border = (cell.border ?? {}) as { top?: ExcelBorderEdge; left?: ExcelBorderEdge; bottom?: ExcelBorderEdge; right?: ExcelBorderEdge };
        const lastCell = span ? worksheet.getCell(span.lastRow, span.lastCol) : cell;
        const lastBorder = (lastCell.border ?? {}) as { bottom?: ExcelBorderEdge; right?: ExcelBorderEdge };
        const { text, numeric } = valueToText(cell.value, cell.numFmt);
        line.push({
          kind:'anchor',
          text,
          numeric,
          rowSpan: span?.rowSpan ?? 1,
          colSpan: span?.colSpan ?? 1,
          style: {
            fill: fillColorOf(fill),
            color: colorOf(font.color, '#000000'),
            fontSize: font.size ?? 11,
            fontName: font.name ?? '',
            bold: !!font.bold,
            italic: !!font.italic,
            underline: !!font.underline,
            hAlign: alignment.horizontal === 'center' || alignment.horizontal === 'centerContinuous' ? 'center'
              : alignment.horizontal === 'right' ? 'right'
              : alignment.horizontal === 'left' || alignment.horizontal === 'justify' || alignment.horizontal === 'distributed' ? 'left'
              : 'general',
            vAlign: alignment.vertical === 'top' ? 'top'
              : alignment.vertical === 'middle' || alignment.vertical === 'distributed' || alignment.vertical === 'justify' ? 'middle'
              : 'bottom',
            wrap: !!alignment.wrapText,
            rotation: alignment.textRotation ?? 0,
            indent: alignment.indent ?? 0,
            borders: { top: edgeOf(border.top), left: edgeOf(border.left), bottom: edgeOf(lastBorder.bottom ?? border.bottom), right: edgeOf(lastBorder.right ?? border.right) },
          },
        });
      }
      cells.push(line);
    }

    if (!cells.some(line => line.some(cell => cell.kind === 'anchor' && (cell.text || cell.style.fill || Object.values(cell.style.borders).some(Boolean))))) {
      const pictures = await decodePictures(workbook, worksheet, top - 1, left - 1);
      if (!pictures.length) continue; // 何も無いシートは飛ばす
      sheets.push({ name: worksheet.name, firstRow: top, firstCol: left, cols, rows, cells, pictures, shapes, print: printSetupOf(worksheet, extras.get(worksheet.name)) });
      continue;
    }

    sheets.push({
      name: worksheet.name,
      firstRow: top,
      firstCol: left,
      cols,
      rows,
      cells,
      pictures: await decodePictures(workbook, worksheet, top - 1, left - 1),
      shapes,
      print: printSetupOf(worksheet, extras.get(worksheet.name)),
    });
  }
  return sheets;
}

const plainStyle = (numeric: boolean): CellStyle => ({
  color: '#000000', fontSize: 11, fontName: '', bold: false, italic: false, underline: false,
  hAlign: numeric ? 'right' : 'left', vAlign: 'bottom', wrap: false, rotation: 0, indent: 0, borders: {},
});

// .xls（旧形式）はExcelJSが読めないため、値・列幅・行高・セル結合だけをSheetJSで読む
function readWithSheetJS(buffer: ArrayBuffer): Sheet[] {
  const workbook = XLSX.read(buffer, { type:'array', cellDates:true, cellText:true, cellStyles:true });
  const hidden = new Set((workbook.Workbook?.Sheets ?? []).filter(sheet => sheet.Hidden).map(sheet => sheet.name));
  const sheets: Sheet[] = [];
  for (const name of workbook.SheetNames) {
    if (hidden.has(name)) continue;
    const ws = workbook.Sheets[name];
    if (!ws?.['!ref']) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    let lastRow = -1;
    let lastCol = -1;
    for (const key of Object.keys(ws)) {
      if (key.startsWith('!')) continue;
      const cell = ws[key] as XLSX.CellObject;
      if (cell?.v === undefined || cell.v === '') continue;
      const address = XLSX.utils.decode_cell(key);
      if (address.r > lastRow) lastRow = address.r;
      if (address.c > lastCol) lastCol = address.c;
    }
    if (lastRow < 0 || lastCol < 0) continue;
    const firstRow = Math.max(0, Math.min(range.s.r, lastRow));
    const firstCol = Math.max(0, Math.min(range.s.c, lastCol));

    const spans = new Map<string, { rowSpan: number; colSpan: number }>();
    const covered = new Map<string, { row: number; col: number }>();
    for (const merge of ws['!merges'] ?? []) {
      if (merge.s.r > lastRow || merge.s.c > lastCol) continue;
      const endRow = Math.min(merge.e.r, lastRow);
      const endCol = Math.min(merge.e.c, lastCol);
      spans.set(`${merge.s.r},${merge.s.c}`, { rowSpan: endRow - merge.s.r + 1, colSpan: endCol - merge.s.c + 1 });
      for (let r = merge.s.r; r <= endRow; r++) for (let c = merge.s.c; c <= endCol; c++) {
        if (r !== merge.s.r || c !== merge.s.c) covered.set(`${r},${c}`, { row: merge.s.r - firstRow, col: merge.s.c - firstCol });
      }
    }

    const cols: number[] = [];
    for (let c = firstCol; c <= lastCol; c++) {
      const info = ws['!cols']?.[c];
      cols.push(Math.max(8, Math.round(info?.wpx ?? (info?.wch ? charsToPx(info.wch) : DEFAULT_COL_PX))));
    }
    const rows: number[] = [];
    const cells: Cell[][] = [];
    for (let r = firstRow; r <= lastRow; r++) {
      const info = ws['!rows']?.[r];
      rows.push(Math.max(8, Math.round(info?.hpx ?? (info?.hpt ? pointsToPx(info.hpt) : DEFAULT_ROW_PX))));
      const line: Cell[] = [];
      for (let c = firstCol; c <= lastCol; c++) {
        const from = covered.get(`${r},${c}`);
        if (from) { line.push({ kind:'covered', row: from.row, col: from.col }); continue; }
        const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
        const span = spans.get(`${r},${c}`);
        const numeric = cell?.t === 'n' || cell?.t === 'd';
        line.push({
          kind:'anchor',
          text: cell === undefined || cell.v === undefined ? '' : (cell.w ?? String(cell.v)),
          numeric,
          rowSpan: span?.rowSpan ?? 1,
          colSpan: span?.colSpan ?? 1,
          style: plainStyle(numeric),
        });
      }
      cells.push(line);
    }
    sheets.push({ name, firstRow: firstRow + 1, firstCol: firstCol + 1, cols, rows, cells, pictures: [], shapes: [], print: { gridLines: true, rowBreaks: [], colBreaks: [] } });
  }
  return sheets;
}

export async function readWorkbook(file: File): Promise<Sheet[]> {
  const isOldFormat = file.name.toLowerCase().endsWith('.xls');
  if (isOldFormat) return readWithSheetJS(await file.arrayBuffer());
  return readWithExcelJS(file);
}
