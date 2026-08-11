import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';

// ブック全体をPDFにする。シートごとに用紙の向きと倍率を決め、
// 幅が収まらなければ列を分割、高さが収まらなければ行で改ページする（Excelの印刷に近い挙動）。
export type PaperSize = 'a4' | 'a3';
export type PageOrientation = 'auto' | 'portrait' | 'landscape';
export type FitMode = 'width' | 'page' | 'none';
export type ExcelPdfOptions = { paper: PaperSize; orientation: PageOrientation; fit: FitMode; repeatHeader: boolean };

export const defaultExcelPdfOptions: ExcelPdfOptions = { paper:'a4', orientation:'auto', fit:'width', repeatHeader:true };

const PAPER_MM: Record<PaperSize, { width: number; height: number }> = { a4:{ width:210, height:297 }, a3:{ width:297, height:420 } };
const MARGIN_MM = 10;
const PX_PER_MM = 96 / 25.4;
const DEFAULT_COL_PX = 64; // Excel標準の列幅（8.43文字）
const DEFAULT_ROW_PX = 20; // Excel標準の行の高さ（15pt）
const MIN_WIDTH_SCALE = 0.35; // 幅を1ページに収めるときの下限。これ以下になる場合は列を分割する
const TARGET_DPI = 150;
const CAPTION_PX = 16;
const MAX_PAGES = 500;

class ConversionError extends Error {}

type Anchor = { kind:'anchor'; text: string; numeric: boolean; rowSpan: number; colSpan: number };
type Covered = { kind:'covered'; row: number; col: number };
type Grid = { name: string; rows: number[]; cols: number[]; cells: (Anchor | Covered)[][] };
type Block = { start: number; end: number; width: number };

const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value));
const sum = (values: number[]) => values.reduce((total, value) => total + value, 0);

function buildGrid(name: string, ws: XLSX.WorkSheet): Grid | undefined {
  if (!ws['!ref']) return undefined;
  const range = XLSX.utils.decode_range(ws['!ref']);
  // 壊れた!refで巨大な範囲を走査しないよう、実在するセルから使用範囲を求める
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
  if (lastRow < 0 || lastCol < 0) return undefined;
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
      if (r !== merge.s.r || c !== merge.s.c) covered.set(`${r},${c}`, { row: merge.s.r, col: merge.s.c });
    }
  }

  const cols: number[] = [];
  for (let c = firstCol; c <= lastCol; c++) {
    const info = ws['!cols']?.[c];
    cols.push(Math.max(8, Math.round(info?.wpx ?? (info?.wch ? info.wch * 7 + 5 : DEFAULT_COL_PX))));
  }
  const rows: number[] = [];
  const cells: (Anchor | Covered)[][] = [];
  for (let r = firstRow; r <= lastRow; r++) {
    const info = ws['!rows']?.[r];
    rows.push(Math.max(8, Math.round(info?.hpx ?? (info?.hpt ? info.hpt * 4 / 3 : DEFAULT_ROW_PX))));
    const line: (Anchor | Covered)[] = [];
    for (let c = firstCol; c <= lastCol; c++) {
      const from = covered.get(`${r},${c}`);
      if (from) { line.push({ kind:'covered', row: from.row - firstRow, col: from.col - firstCol }); continue; }
      const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
      const span = spans.get(`${r},${c}`);
      line.push({
        kind:'anchor',
        text: cell === undefined || cell.v === undefined ? '' : (cell.w ?? String(cell.v)),
        numeric: cell?.t === 'n' || cell?.t === 'd',
        rowSpan: span?.rowSpan ?? 1,
        colSpan: span?.colSpan ?? 1,
      });
    }
    cells.push(line);
  }
  return { name, rows, cols, cells };
}

// ページに入る列の範囲へ分割する（1列が上限より広い場合はその列だけで1ブロック）
function splitColumns(cols: number[], maxWidth: number): Block[] {
  const blocks: Block[] = [];
  let start = 0;
  let width = 0;
  for (let index = 0; index < cols.length; index++) {
    if (width > 0 && width + cols[index] > maxWidth) { blocks.push({ start, end: index - 1, width }); start = index; width = 0; }
    width += cols[index];
  }
  blocks.push({ start, end: cols.length - 1, width });
  return blocks;
}

// ページに入る行の範囲へ分割する（2ページ目以降は繰り返す見出し行の高さを差し引く）
function splitRows(rows: number[], maxHeight: number, headerRow: number | undefined) {
  const pages: number[][] = [];
  const headerHeight = headerRow === undefined ? 0 : rows[headerRow];
  let current: number[] = [];
  let height = 0;
  for (let index = 0; index < rows.length; index++) {
    if (current.length > 0 && height + rows[index] > maxHeight) { pages.push(current); current = []; height = headerHeight; }
    current.push(index);
    height += rows[index];
  }
  if (current.length) pages.push(current);
  return pages;
}

type PageLayout = { grid: Grid; columns: number[]; lines: number[]; caption: string; captionHeight: number; scale: number; rasterScale: number };

const FONT_STACK = '"Yu Gothic","Meiryo","Noto Sans JP","Hiragino Sans",sans-serif';
const CELL_FONT_PX = 12;
const GRID_COLOR = '#c8c8c8';
const PAD_PX = 3;

// このページに連続して並んでいる要素数（結合セルをページ内に収めるために使う）
const runLength = (list: number[], position: number) => {
  let length = 1;
  while (list[position + length] === list[position] + length) length++;
  return length;
};

// レイアウトはこちらで確定しているので、DOMを介さずCanvasへ直接描く（描画が速く、寸法もぶれない）
function drawPage({ grid, columns, lines, caption, captionHeight, scale, rasterScale }: PageLayout) {
  const widths = columns.map(c => grid.cols[c]);
  const heights = lines.map(r => grid.rows[r]);
  const width = sum(widths);
  const height = captionHeight + sum(heights);

  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(width * rasterScale);
  canvas.height = Math.ceil(height * rasterScale);
  const ctx = canvas.getContext('2d', { alpha: false });
  if (!ctx) throw new Error('キャンバスを準備できませんでした');
  ctx.scale(rasterScale, rasterScale);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  // シート名とページ番号
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, 0, width, captionHeight);
  ctx.clip();
  ctx.fillStyle = '#333333';
  ctx.font = `${Math.round(11 / clamp(scale, 0.3, 1))}px ${FONT_STACK}`;
  ctx.textBaseline = 'middle';
  ctx.fillText(caption, 0, captionHeight / 2);
  ctx.restore();

  const x = [0];
  for (const columnWidth of widths) x.push(x[x.length - 1] + columnWidth);
  const y = [captionHeight];
  for (const rowHeight of heights) y.push(y[y.length - 1] + rowHeight);
  const rowsOnPage = new Set(lines);
  const columnsOnPage = new Set(columns);

  ctx.lineWidth = 1;
  ctx.strokeStyle = GRID_COLOR;
  ctx.font = `${CELL_FONT_PX}px ${FONT_STACK}`;
  for (const [row, r] of lines.entries()) {
    for (let column = 0; column < columns.length; column++) {
      const c = columns[column];
      const cell = grid.cells[r][c];
      if (cell.kind === 'covered') {
        // 結合元がこのページにあるなら結合セルが覆う。ページや列ブロックで切れた場合は空セルとして枠だけ描く
        if (columnsOnPage.has(cell.col) && rowsOnPage.has(cell.row)) continue;
        ctx.strokeRect(x[column] + 0.5, y[row] + 0.5, widths[column], heights[row]);
        continue;
      }
      const colSpan = Math.min(cell.colSpan, runLength(columns, column));
      const rowSpan = Math.min(cell.rowSpan, runLength(lines, row));
      const cellWidth = x[column + colSpan] - x[column];
      const cellHeight = y[row + rowSpan] - y[row];
      ctx.strokeRect(x[column] + 0.5, y[row] + 0.5, cellWidth, cellHeight);
      if (cell.text) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(x[column] + 1, y[row], cellWidth - 2, cellHeight);
        ctx.clip();
        ctx.fillStyle = '#000000';
        ctx.textBaseline = 'alphabetic';
        ctx.textAlign = cell.numeric ? 'right' : 'left';
        // Excelの既定と同じく下揃え
        const baseline = y[row] + cellHeight - Math.max(3, Math.round(cellHeight * 0.2));
        ctx.fillText(cell.text, cell.numeric ? x[column] + cellWidth - PAD_PX : x[column] + PAD_PX, baseline);
        ctx.restore();
      }
      column += colSpan - 1;
    }
  }
  return canvas;
}

const toPngBytes = (canvas: HTMLCanvasElement) => new Promise<ArrayBuffer>((resolve, reject) => {
  canvas.toBlob(blob => {
    if (!blob) { reject(new Error('ページを画像にできませんでした')); return; }
    blob.arrayBuffer().then(resolve, reject);
  }, 'image/png');
});

export async function excelToPdf(file: File, options: ExcelPdfOptions, onStatus?: (message: string) => void) {
  onStatus?.('ブックを解析しています…');
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:true, cellText:true, cellStyles:true });
    const hidden = new Set((workbook.Workbook?.Sheets ?? []).filter(sheet => sheet.Hidden).map(sheet => sheet.name));
    const grids = workbook.SheetNames
      .filter(name => !hidden.has(name))
      .map(name => buildGrid(name, workbook.Sheets[name]))
      .filter((grid): grid is Grid => grid !== undefined);
    if (!grids.length) throw new ConversionError('PDFにできるシートが見つかりませんでした');

    const paper = PAPER_MM[options.paper];
    let pdf: jsPDF | undefined;
    let pageCount = 0;

    for (const [sheetIndex, grid] of grids.entries()) {
      onStatus?.(`シートを変換しています… ${sheetIndex + 1}/${grids.length}（${grid.name}）`);
      const tableWidth = sum(grid.cols);
      const tableHeight = sum(grid.rows);

      // 用紙の向き：指定がなければ縮小率の小さくて済む向きを選ぶ
      const fitFor = (landscape: boolean) => {
        const width = (landscape ? paper.height : paper.width) - MARGIN_MM * 2;
        const height = (landscape ? paper.width : paper.height) - MARGIN_MM * 2;
        const byWidth = Math.min(1, width / (tableWidth / PX_PER_MM));
        return options.fit === 'page' ? Math.min(byWidth, height / ((tableHeight + CAPTION_PX) / PX_PER_MM)) : byWidth;
      };
      const landscape = options.orientation === 'auto' ? fitFor(true) > fitFor(false) : options.orientation === 'landscape';
      const contentWidth = (landscape ? paper.height : paper.width) - MARGIN_MM * 2;
      const contentHeight = (landscape ? paper.width : paper.height) - MARGIN_MM * 2;

      const maxBlockWidth = options.fit === 'page' ? Infinity : options.fit === 'none' ? contentWidth * PX_PER_MM : contentWidth * PX_PER_MM / MIN_WIDTH_SCALE;
      const blocks = splitColumns(grid.cols, maxBlockWidth);
      // 見出しの繰り返し：先頭行は改ページのたび、先頭列は列を分けたページのたびに付ける
      const headerRow = options.repeatHeader && grid.rows.length > 1 ? 0 : undefined;
      const headerColumn = options.repeatHeader && blocks.length > 1 ? 0 : undefined;

      for (const block of blocks) {
        const range = Array.from({ length: block.end - block.start + 1 }, (_, index) => block.start + index);
        const columns = headerColumn !== undefined && range[0] !== headerColumn ? [headerColumn, ...range] : range;
        const pageWidth = sum(columns.map(c => grid.cols[c]));
        const byWidth = Math.min(1, contentWidth / (pageWidth / PX_PER_MM));
        const scale = options.fit === 'none' ? 1
          : options.fit === 'page' ? Math.min(byWidth, contentHeight / ((tableHeight + CAPTION_PX) / PX_PER_MM))
          : byWidth;
        const captionHeight = Math.round(CAPTION_PX / clamp(scale, 0.3, 1));
        const usableHeight = options.fit === 'page' ? Infinity : contentHeight * PX_PER_MM / scale - captionHeight;
        const rowPages = splitRows(grid.rows, usableHeight, headerRow);
        const rasterScale = clamp(scale * TARGET_DPI / 96, 0.6, 2);

        for (const [pageIndex, rowIndexes] of rowPages.entries()) {
          if (++pageCount > MAX_PAGES) throw new ConversionError(`ページ数が${MAX_PAGES}を超えました。用紙サイズを大きくするか、シートを分けてください`);
          const columnLabel = blocks.length > 1 ? `・列${block.start + 1}〜${block.end + 1}` : '';
          const caption = rowPages.length > 1 || blocks.length > 1 ? `${grid.name}（${pageIndex + 1}/${rowPages.length}${columnLabel}）` : grid.name;
          const lines = headerRow !== undefined && rowIndexes[0] !== headerRow ? [headerRow, ...rowIndexes] : rowIndexes;
          const canvas = drawPage({ grid, columns, lines, caption, captionHeight, scale, rasterScale });
          // 罫線と文字が主体なのでPNG（可逆）の方がJPEGより軽く鮮明。
          // エンコード中に画面が固まらないよう、同期のtoDataURLではなくtoBlobを使う
          const image = new Uint8Array(await toPngBytes(canvas));
          let width = canvas.width / rasterScale * scale / PX_PER_MM;
          let height = canvas.height / rasterScale * scale / PX_PER_MM;
          canvas.width = canvas.height = 0; // 大きなキャンバスを早めに解放する
          if (height > contentHeight) { width *= contentHeight / height; height = contentHeight; }
          if (!pdf) pdf = new jsPDF({ unit:'mm', format:[paper.width, paper.height], orientation: landscape ? 'landscape' : 'portrait', compress:true });
          else pdf.addPage([paper.width, paper.height], landscape ? 'landscape' : 'portrait');
          pdf.addImage(image, 'PNG', MARGIN_MM, MARGIN_MM, width, height);
        }
      }
    }
    return pdf!.output('blob');
  } catch (error) {
    if (error instanceof ConversionError) throw error;
    const message = error instanceof Error ? error.message : '';
    throw new Error(`ExcelファイルをPDFに変換できませんでした${message ? `（${message}）` : ''}`);
  }
}
