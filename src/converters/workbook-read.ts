import * as XLSX from 'xlsx';

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
export type Sheet = {
  name: string;
  cols: number[];
  rows: number[];
  cells: Cell[][];
  pictures: Picture[];
  gridLines: boolean;
  landscape?: boolean;
};

const DEFAULT_ROW_PT = 15;
export const DEFAULT_COL_PX = 64;
export const DEFAULT_ROW_PX = 20;
const PT_TO_PX = 4 / 3;

// Officeの既定テーマ色（tintは無視する）
const THEME_COLORS = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];

const charsToPx = (chars: number) => Math.round(chars * 7 + 5);
const pointsToPx = (points: number) => Math.round(points * PT_TO_PX);

const argbToCss = (argb: string) => (argb.length === 8 ? `#${argb.slice(2)}` : argb.startsWith('#') ? argb : `#${argb}`);

type ExcelColor = { argb?: string; theme?: number; indexed?: number } | undefined;
const colorOf = (color: ExcelColor, fallback: string) => {
  if (!color) return fallback;
  if (color.argb) return argbToCss(color.argb);
  if (typeof color.theme === 'number') return THEME_COLORS[color.theme] ?? fallback;
  return fallback;
};

const BORDER_WIDTH: Record<string, number> = { hair: 0.5, thin: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, slantDashDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, thick: 3, double: 3 };
const BORDER_DASH: Record<string, number[]> = { dotted: [1, 2], dashed: [4, 3], dashDot: [6, 2, 2, 2], dashDotDot: [6, 2, 2, 2, 2, 2], mediumDashed: [5, 3], mediumDashDot: [7, 3, 3, 3], mediumDashDotDot: [7, 3, 3, 3, 3, 3], slantDashDot: [5, 3] };

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

async function readWithExcelJS(file: File): Promise<Sheet[]> {
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
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
        const fill = cell.fill as { type?: string; pattern?: string; fgColor?: ExcelColor } | undefined;
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
            fill: fill?.pattern && fill.pattern !== 'none' ? colorOf(fill.fgColor, '#ffffff') : undefined,
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
            vAlign: alignment.vertical === 'top' ? 'top' : alignment.vertical === 'middle' ? 'middle' : 'bottom',
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
      sheets.push({ name: worksheet.name, cols, rows, cells, pictures, gridLines: !!worksheet.pageSetup?.showGridLines, landscape: worksheet.pageSetup?.orientation === 'landscape' ? true : worksheet.pageSetup?.orientation === 'portrait' ? false : undefined });
      continue;
    }

    sheets.push({
      name: worksheet.name,
      cols,
      rows,
      cells,
      pictures: await decodePictures(workbook, worksheet, top - 1, left - 1),
      gridLines: !!worksheet.pageSetup?.showGridLines,
      landscape: worksheet.pageSetup?.orientation === 'landscape' ? true : worksheet.pageSetup?.orientation === 'portrait' ? false : undefined,
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
    sheets.push({ name, cols, rows, cells, pictures: [], gridLines: true });
  }
  return sheets;
}

export async function readWorkbook(file: File): Promise<Sheet[]> {
  const isOldFormat = file.name.toLowerCase().endsWith('.xls');
  if (isOldFormat) return readWithSheetJS(await file.arrayBuffer());
  return readWithExcelJS(file);
}
