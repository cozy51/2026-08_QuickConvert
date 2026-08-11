import { jsPDF } from 'jspdf';
import { readWorkbook } from './workbook-read';
import type { Cell, CellAnchor, CellStyle, Edge, Sheet, Shape } from './workbook-read';

// ブック全体をPDFにする。シートごとに用紙の向きと倍率を決め、
// 幅が収まらなければ列を分割、高さが収まらなければ行で改ページする（Excelの印刷に近い挙動）。
// セルの塗り・罫線・フォント・配置・画像はCanvasへ直接描く。
export type PaperSize = 'a4' | 'a3';
export type PageOrientation = 'auto' | 'portrait' | 'landscape';
export type FitMode = 'file' | 'width' | 'page' | 'none';
export type GridLineMode = 'auto' | 'on' | 'off';
export type RepeatMode = 'auto' | 'on' | 'off';
export type ExcelPdfOptions = { paper: PaperSize; orientation: PageOrientation; fit: FitMode; repeatHeader: RepeatMode; gridLines: GridLineMode; sheetLabel: boolean };

export const defaultExcelPdfOptions: ExcelPdfOptions = { paper:'a4', orientation:'auto', fit:'width', repeatHeader:'auto', gridLines:'auto', sheetLabel:false };

const PAPER_MM: Record<PaperSize, { width: number; height: number }> = { a4:{ width:210, height:297 }, a3:{ width:297, height:420 } };
const MARGIN_MM = 10;
const MIN_MARGIN_MM = 5;
const PX_PER_MM = 96 / 25.4;
const MIN_WIDTH_SCALE = 0.35; // 幅を1ページに収めるときの下限。これ以下になる場合は列を分割する
const TARGET_DPI = 150;
const CAPTION_PX = 16;
const MAX_PAGES = 500;
const FONT_STACK = '"Yu Gothic","Meiryo","Noto Sans JP","Hiragino Sans",sans-serif';
const GRID_COLOR = '#d0d0d0';
const PAD_PX = 3;
const PT_TO_PX = 4 / 3;

class ConversionError extends Error {}

type Block = { start: number; end: number; width: number };

const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value));
const sum = (values: number[]) => values.reduce((total, value) => total + value, 0);

// このページに連続して並んでいる要素数（結合セルをページ内に収めるために使う）
const runLength = (list: number[], position: number) => {
  let length = 1;
  while (list[position + length] === list[position] + length) length++;
  return length;
};

// ページに入る列の範囲へ分割する（1列が上限より広い場合はその列だけで1ブロック）
// forced に入っている位置（その列の後）では必ず改ページする
function splitColumns(cols: number[], maxWidth: number, forced: Set<number>): Block[] {
  const blocks: Block[] = [];
  let start = 0;
  let width = 0;
  for (let index = 0; index < cols.length; index++) {
    if (width > 0 && width + cols[index] > maxWidth) { blocks.push({ start, end: index - 1, width }); start = index; width = 0; }
    width += cols[index];
    if (forced.has(index) && index < cols.length - 1) { blocks.push({ start, end: index, width }); start = index + 1; width = 0; }
  }
  if (start < cols.length) blocks.push({ start, end: cols.length - 1, width });
  return blocks;
}

// ページに入る行の範囲へ分割する（2ページ目以降は繰り返す見出し行の高さを差し引く）
// forced に入っている位置（その行の後）では必ず改ページする
function splitRows(rows: number[], maxHeight: number, titleRows: number[], forced: Set<number>) {
  const pages: number[][] = [];
  const headerHeight = sum(titleRows.map(index => rows[index] ?? 0));
  let current: number[] = [];
  let height = 0;
  for (let index = 0; index < rows.length; index++) {
    if (current.length > 0 && height + rows[index] > maxHeight) { pages.push(current); current = []; height = headerHeight; }
    current.push(index);
    height += rows[index];
    if (forced.has(index) && index < rows.length - 1) { pages.push(current); current = []; height = headerHeight; }
  }
  if (current.length) pages.push(current);
  return pages;
}

const fontOf = (style: CellStyle) => {
  const family = style.fontName ? `"${style.fontName}",${FONT_STACK}` : FONT_STACK;
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${Math.max(6, Math.round(style.fontSize * PT_TO_PX))}px ${family}`;
};

function strokeEdge(ctx: CanvasRenderingContext2D, edge: Edge, x1: number, y1: number, x2: number, y2: number) {
  ctx.save();
  ctx.strokeStyle = edge.color;
  ctx.lineWidth = edge.width;
  if (edge.dash) ctx.setLineDash(edge.dash);
  const offset = edge.width % 2 === 1 ? 0.5 : 0;
  const draw = (shift: number) => {
    ctx.beginPath();
    ctx.moveTo(x1 + (y1 === y2 ? 0 : offset + shift), y1 + (y1 === y2 ? offset + shift : 0));
    ctx.lineTo(x2 + (y1 === y2 ? 0 : offset + shift), y2 + (y1 === y2 ? offset + shift : 0));
    ctx.stroke();
  };
  if (edge.double) { ctx.lineWidth = 1; draw(-1); draw(1); } else draw(0);
  ctx.restore();
}

// 折り返し表示のセル用に、幅に収まる行へ分ける
function wrapLines(ctx: CanvasRenderingContext2D, text: string, maxWidth: number) {
  const lines: string[] = [];
  for (const paragraph of text.split('\n')) {
    let current = '';
    for (const char of paragraph) {
      if (current && ctx.measureText(current + char).width > maxWidth) {
        const breakAt = /[ \t]/.test(char) ? current.length : current.lastIndexOf(' ');
        if (breakAt > 0 && current.length - breakAt < 12) { lines.push(current.slice(0, breakAt)); current = current.slice(breakAt + 1); }
        else { lines.push(current); current = ''; }
      }
      current += char;
    }
    lines.push(current);
  }
  return lines;
}

type TextBox = { left: number; top: number; width: number; height: number };

function drawCellText(ctx: CanvasRenderingContext2D, cell: CellAnchor, box: TextBox, clip: TextBox) {
  const style = cell.style;
  ctx.save();
  ctx.beginPath();
  ctx.rect(clip.left, clip.top, clip.width, clip.height);
  ctx.clip();
  ctx.font = fontOf(style);
  ctx.fillStyle = style.color;
  const lineHeight = Math.round(style.fontSize * PT_TO_PX * 1.25);
  const horizontal = style.hAlign === 'general' ? (cell.numeric ? 'right' : 'left') : style.hAlign;

  if (style.rotation === 'vertical') {
    // 縦書き：1文字ずつ縦に積む
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    const characters = [...cell.text];
    const total = characters.length * lineHeight;
    let y = box.top + (style.vAlign === 'top' ? 0 : style.vAlign === 'middle' ? Math.max(0, (box.height - total) / 2) : Math.max(0, box.height - total));
    for (const char of characters) { ctx.fillText(char, box.left + box.width / 2, y); y += lineHeight; }
    ctx.restore();
    return;
  }

  const rotation = typeof style.rotation === 'number' ? style.rotation : 0;
  if (rotation !== 0) {
    ctx.translate(box.left + box.width / 2, box.top + box.height / 2);
    ctx.rotate(-rotation * Math.PI / 180);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(cell.text, 0, 0);
    ctx.restore();
    return;
  }

  const indent = style.indent * 8;
  const lines = style.wrap ? wrapLines(ctx, cell.text, Math.max(8, box.width - PAD_PX * 2 - indent)) : cell.text.split('\n');
  const total = lines.length * lineHeight;
  const startY = box.top + (style.vAlign === 'top' ? 0 : style.vAlign === 'middle' ? Math.max(0, (box.height - total) / 2) : Math.max(0, box.height - total));
  ctx.textBaseline = 'top';
  ctx.textAlign = horizontal;
  const x = horizontal === 'center' ? box.left + box.width / 2
    : horizontal === 'right' ? box.left + box.width - PAD_PX
    : box.left + PAD_PX + indent;
  lines.forEach((line, index) => ctx.fillText(line, x, startY + index * lineHeight + Math.round(lineHeight * 0.12)));
  ctx.restore();
}

type PageLayout = { sheet: Sheet; columns: number[]; lines: number[]; caption: string; captionHeight: number; scale: number; rasterScale: number; gridLines: boolean };

// レイアウトはこちらで確定しているので、DOMを介さずCanvasへ直接描く（描画が速く、寸法もぶれない）
function drawPage({ sheet, columns, lines, caption, captionHeight, scale, rasterScale, gridLines }: PageLayout) {
  const widths = columns.map(c => sheet.cols[c]);
  const heights = lines.map(r => sheet.rows[r]);
  const width = Math.max(1, sum(widths));
  const height = captionHeight + sum(heights);

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.ceil(width * rasterScale));
  canvas.height = Math.max(1, Math.ceil(height * rasterScale));
  const ctx = canvas.getContext('2d', { alpha: false });
  if (!ctx) throw new ConversionError('キャンバスを準備できませんでした');
  ctx.scale(rasterScale, rasterScale);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  // シート名とページ番号
  if (caption && captionHeight > 0) {
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, 0, width, captionHeight);
  ctx.clip();
  ctx.fillStyle = '#333333';
  ctx.font = `${Math.round(11 / clamp(scale, 0.3, 1))}px ${FONT_STACK}`;
  ctx.textBaseline = 'middle';
  ctx.fillText(caption, 0, captionHeight / 2);
  ctx.restore();
  }

  const x = [0];
  for (const columnWidth of widths) x.push(x[x.length - 1] + columnWidth);
  const y = [captionHeight];
  for (const rowHeight of heights) y.push(y[y.length - 1] + rowHeight);
  const rowsOnPage = new Map(lines.map((r, position) => [r, position]));
  const columnsOnPage = new Map(columns.map((c, position) => [c, position]));

  type Placed = { cell: CellAnchor; left: number; top: number; width: number; height: number; row: number; column: number; colSpan: number };
  const placed: Placed[] = [];
  for (const [row, r] of lines.entries()) {
    for (let column = 0; column < columns.length; column++) {
      const c = columns[column];
      const cell: Cell = sheet.cells[r][c];
      if (cell.kind === 'covered') continue; // 結合セルは結合元をまとめて描く
      const colSpan = Math.min(cell.colSpan, runLength(columns, column));
      const rowSpan = Math.min(cell.rowSpan, runLength(lines, row));
      placed.push({ cell, left: x[column], top: y[row], width: x[column + colSpan] - x[column], height: y[row + rowSpan] - y[row], row, column, colSpan });
      column += colSpan - 1;
    }
  }

  // 1) 罫線（グリッド線）→ 2) 塗り → 3) 実際の罫線 → 4) 文字 → 5) 画像 の順に重ねる
  if (gridLines) {
    ctx.save();
    ctx.strokeStyle = GRID_COLOR;
    ctx.lineWidth = 1;
    for (const item of placed) ctx.strokeRect(item.left + 0.5, item.top + 0.5, item.width, item.height);
    ctx.restore();
  }
  for (const item of placed) {
    if (!item.cell.style.fill) continue;
    ctx.fillStyle = item.cell.style.fill;
    ctx.fillRect(item.left, item.top, item.width, item.height);
  }
  for (const item of placed) {
    const { borders } = item.cell.style;
    const right = item.left + item.width;
    const bottom = item.top + item.height;
    if (borders.top) strokeEdge(ctx, borders.top, item.left, item.top, right, item.top);
    if (borders.bottom) strokeEdge(ctx, borders.bottom, item.left, bottom, right, bottom);
    if (borders.left) strokeEdge(ctx, borders.left, item.left, item.top, item.left, bottom);
    if (borders.right) strokeEdge(ctx, borders.right, right, item.top, right, bottom);
  }
  for (const item of placed) {
    if (!item.cell.text) continue;
    const box = { left: item.left, top: item.top, width: item.width, height: item.height };
    // 折り返しでも回転でもない文字は、Excelと同じく隣の空セルへはみ出して表示する
    // （左揃えは右へ、右揃えは左へ、中央揃えは両方向へ）
    const style = item.cell.style;
    const align = style.hAlign === 'general' ? (item.cell.numeric ? 'right' : 'left') : style.hAlign;
    const spills = !style.wrap && style.rotation === 0;
    let clipLeft = item.left;
    let clipRight = item.left + item.width;
    const isEmpty = (row: number, column: number) => {
      const neighbour = sheet.cells[lines[row]][columns[column]];
      return neighbour.kind === 'anchor' && !neighbour.text;
    };
    if (spills && (align === 'left' || align === 'center')) {
      for (let next = item.column + item.colSpan; next < columns.length && isEmpty(item.row, next); next++) clipRight += sheet.cols[columns[next]];
    }
    if (spills && (align === 'right' || align === 'center')) {
      for (let previous = item.column - 1; previous >= 0 && isEmpty(item.row, previous); previous--) clipLeft -= sheet.cols[columns[previous]];
    }
    drawCellText(ctx, item.cell, box, { left: clipLeft, top: item.top, width: clipRight - clipLeft, height: item.height });
  }

  // グリッド上の位置（列・行の小数指定）をページ上のpxへ変換する
  const locate = (gridCol: number, gridRow: number) => {
    const columnIndex = Math.floor(gridCol);
    const rowIndex = Math.floor(gridRow);
    const column = columnsOnPage.get(columnIndex);
    const row = rowsOnPage.get(rowIndex);
    if (column === undefined || row === undefined) return undefined;
    return {
      x: x[column] + (gridCol - columnIndex) * (sheet.cols[columnIndex] ?? 0),
      y: y[row] + (gridRow - rowIndex) * (sheet.rows[rowIndex] ?? 0),
    };
  };
  const spanSize = (sizes: number[], from: number, to: number) => {
    const start = Math.floor(from);
    const end = Math.floor(to);
    let total = -(from - start) * (sizes[start] ?? 0);
    for (let index = start; index < end; index++) total += sizes[index] ?? 0;
    total += (to - end) * (sizes[end] ?? 0);
    return total;
  };

  for (const picture of sheet.pictures) {
    const spot = locate(picture.col, picture.row);
    if (!spot) continue; // このページに載らない画像は描かない
    let pictureWidth = picture.toCol !== undefined ? spanSize(sheet.cols, picture.col, picture.toCol) : picture.width;
    let pictureHeight = picture.toRow !== undefined ? spanSize(sheet.rows, picture.row, picture.toRow) : picture.height;
    if (!pictureWidth || !pictureHeight) { pictureWidth = picture.width ?? picture.source.width; pictureHeight = picture.height ?? picture.source.height; }
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, captionHeight, width, height - captionHeight);
    ctx.clip();
    ctx.drawImage(picture.source, spot.x, spot.y, pictureWidth, pictureHeight);
    ctx.restore();
  }

  for (const shape of sheet.shapes) drawShape(ctx, sheet, shape, locate, spanSize, { top: captionHeight, width, height });

  return canvas;
}

type Locator = (col: number, row: number) => { x: number; y: number } | undefined;
type Spanner = (sizes: number[], from: number, to: number) => number;

// 罫線や画像のあとに、Excelの図形（四角・直線）を重ねる
function drawShape(ctx: CanvasRenderingContext2D, sheet: Sheet, shape: Shape, locate: Locator, spanSize: Spanner, page: { top: number; width: number; height: number }) {
  const start = locate(Math.min(shape.col, shape.toCol), Math.min(shape.row, shape.toRow));
  if (!start) return; // このページに載らない図形は描かない
  const shapeWidth = Math.abs(spanSize(sheet.cols, shape.col, shape.toCol));
  const shapeHeight = Math.abs(spanSize(sheet.rows, shape.row, shape.toRow));
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, page.top, page.width, page.height - page.top);
  ctx.clip();
  if (shape.dash) ctx.setLineDash(shape.dash);
  ctx.lineWidth = shape.strokeWidth;
  if (shape.kind === 'line') {
    if (shape.stroke) {
      ctx.strokeStyle = shape.stroke;
      ctx.beginPath();
      ctx.moveTo(shape.flipH ? start.x + shapeWidth : start.x, shape.flipV ? start.y + shapeHeight : start.y);
      ctx.lineTo(shape.flipH ? start.x : start.x + shapeWidth, shape.flipV ? start.y : start.y + shapeHeight);
      ctx.stroke();
    }
    ctx.restore();
    return;
  }
  if (shape.fill) { ctx.fillStyle = shape.fill; ctx.fillRect(start.x, start.y, shapeWidth, shapeHeight); }
  if (shape.stroke) { ctx.strokeStyle = shape.stroke; ctx.strokeRect(start.x + 0.5, start.y + 0.5, shapeWidth, shapeHeight); }
  if (shape.text) {
    ctx.setLineDash([]);
    ctx.beginPath();
    ctx.rect(start.x, start.y, shapeWidth, shapeHeight);
    ctx.clip();
    ctx.fillStyle = shape.text.color;
    ctx.font = `${shape.text.bold ? 'bold ' : ''}${Math.max(6, Math.round(shape.text.fontSize * PT_TO_PX))}px ${FONT_STACK}`;
    ctx.textAlign = shape.text.hAlign;
    ctx.textBaseline = 'middle';
    const lines = shape.text.text.split('\n');
    const lineHeight = Math.round(shape.text.fontSize * PT_TO_PX * 1.25);
    const total = lines.length * lineHeight;
    const startY = shape.text.vAlign === 'top' ? start.y + lineHeight / 2
      : shape.text.vAlign === 'bottom' ? start.y + shapeHeight - total + lineHeight / 2
      : start.y + (shapeHeight - total) / 2 + lineHeight / 2;
    const textX = shape.text.hAlign === 'center' ? start.x + shapeWidth / 2 : shape.text.hAlign === 'right' ? start.x + shapeWidth - PAD_PX : start.x + PAD_PX;
    lines.forEach((line, index) => ctx.fillText(line, textX, startY + index * lineHeight));
  }
  ctx.restore();
}

const toImageBytes = (canvas: HTMLCanvasElement, type: string, quality?: number) => new Promise<ArrayBuffer>((resolve, reject) => {
  canvas.toBlob(blob => {
    if (!blob) { reject(new ConversionError('ページを画像にできませんでした')); return; }
    blob.arrayBuffer().then(resolve, reject);
  }, type, quality);
});

export async function excelToPdf(file: File, options: ExcelPdfOptions, onStatus?: (message: string) => void) {
  onStatus?.('ブックを解析しています…');
  let sheets: Sheet[] = [];
  try {
    sheets = await readWorkbook(file);
    if (!sheets.length) throw new ConversionError('PDFにできるシートが見つかりませんでした');

    const paper = PAPER_MM[options.paper];
    let pdf: jsPDF | undefined;
    let pageCount = 0;

    for (const [sheetIndex, sheet] of sheets.entries()) {
      onStatus?.(`シートを変換しています… ${sheetIndex + 1}/${sheets.length}（${sheet.name}）`);
      const tableWidth = sum(sheet.cols);
      const tableHeight = sum(sheet.rows);
      const gridLines = options.gridLines === 'auto' ? sheet.print.gridLines : options.gridLines === 'on';
      // ファイルの余白設定（インチ）を使う。極端に狭いものは最低値まで広げる
      const margin = sheet.print.margins
        ? {
            left: clamp(sheet.print.margins.left * 25.4, 0, 40),
            right: clamp(sheet.print.margins.right * 25.4, 0, 40),
            top: clamp(sheet.print.margins.top * 25.4, MIN_MARGIN_MM, 40),
            bottom: clamp(sheet.print.margins.bottom * 25.4, 0, 40),
          }
        : { left: MARGIN_MM, right: MARGIN_MM, top: MARGIN_MM, bottom: MARGIN_MM };

      // 用紙の向き：指定がなければファイルの印刷設定を優先し、無ければ縮小率が小さくて済む向きを選ぶ
      const fitFor = (landscape: boolean) => {
        const width = (landscape ? paper.height : paper.width) - margin.left - margin.right;
        const height = (landscape ? paper.width : paper.height) - margin.top - margin.bottom;
        const byWidth = Math.min(1, width / (tableWidth / PX_PER_MM));
        return options.fit === 'page' ? Math.min(byWidth, height / ((tableHeight + (options.sheetLabel ? CAPTION_PX : 0)) / PX_PER_MM)) : byWidth;
      };
      const landscape = options.orientation === 'auto'
        ? sheet.print.landscape ?? fitFor(true) > fitFor(false)
        : options.orientation === 'landscape';
      const contentWidth = (landscape ? paper.height : paper.width) - margin.left - margin.right;
      const contentHeight = (landscape ? paper.width : paper.height) - margin.top - margin.bottom;

      // Excelの印刷設定に従うときの倍率（「N ページ幅に収める」または「倍率 x%」）
      const fileScale = () => {
        if (sheet.print.fitToPage) {
          const acrossWidth = sheet.print.fitToWidth ?? 1;
          const acrossHeight = sheet.print.fitToHeight ?? 0;
          const byWidth = acrossWidth > 0 ? contentWidth * acrossWidth / (tableWidth / PX_PER_MM) : Infinity;
          const byHeight = acrossHeight > 0 ? contentHeight * acrossHeight / ((tallestSegment + (options.sheetLabel ? CAPTION_PX : 0)) / PX_PER_MM) : Infinity;
          const scale = Math.min(byWidth, byHeight);
          return Number.isFinite(scale) ? Math.min(1, scale) : 1;
        }
        return clamp((sheet.print.scale ?? 100) / 100, 0.1, 4);
      };
      const sheetScale = options.fit === 'file' ? fileScale() : undefined;

      // シート上の行番号・列番号を、この範囲の添字へ変換する
      const forcedRows = new Set(sheet.print.rowBreaks.map(row => row - sheet.firstRow).filter(index => index >= 0 && index < sheet.rows.length - 1));
      const forcedCols = new Set(sheet.print.colBreaks.map(col => col - sheet.firstCol).filter(index => index >= 0 && index < sheet.cols.length - 1));
      // 改ページで区切られた区間のうち、一番高いものを基準に「1ページに収める」倍率を決める
      const segmentHeights = (() => {
        const heights: number[] = [];
        let height = 0;
        sheet.rows.forEach((rowHeight, index) => {
          height += rowHeight;
          if (forcedRows.has(index)) { heights.push(height); height = 0; }
        });
        heights.push(height);
        return heights.filter(value => value > 0);
      })();
      const tallestSegment = segmentHeights.length ? Math.max(...segmentHeights) : tableHeight;

      const maxBlockWidth = options.fit === 'page' ? Infinity
        : options.fit === 'file' ? contentWidth * PX_PER_MM / Math.max(sheetScale ?? 1, MIN_WIDTH_SCALE)
        : options.fit === 'none' ? contentWidth * PX_PER_MM
        : contentWidth * PX_PER_MM / MIN_WIDTH_SCALE;
      const blocks = splitColumns(sheet.cols, maxBlockWidth, forcedCols);
      // 見出しの繰り返し：既定はExcelの印刷タイトル設定に従う（設定が無ければ繰り返さない）
      const titleRows = options.repeatHeader === 'off' ? []
        : options.repeatHeader === 'on' ? (sheet.rows.length > 1 ? [0] : [])
        : sheet.print.titleRows
          ? Array.from({ length: sheet.print.titleRows[1] - sheet.print.titleRows[0] + 1 }, (_, index) => sheet.print.titleRows![0] + index - sheet.firstRow)
              .filter(index => index >= 0 && index < sheet.rows.length)
          : [];
      const titleColumns = options.repeatHeader === 'off' ? []
        : options.repeatHeader === 'on' ? (blocks.length > 1 ? [0] : [])
        : sheet.print.titleCols
          ? Array.from({ length: sheet.print.titleCols[1] - sheet.print.titleCols[0] + 1 }, (_, index) => sheet.print.titleCols![0] + index - sheet.firstCol)
              .filter(index => index >= 0 && index < sheet.cols.length)
          : [];

      for (const block of blocks) {
        const range = Array.from({ length: block.end - block.start + 1 }, (_, index) => block.start + index);
        const repeatedColumns = titleColumns.filter(index => index < range[0]);
        const columns = repeatedColumns.length ? [...repeatedColumns, ...range] : range;
        const pageWidth = sum(columns.map(c => sheet.cols[c]));
        const byWidth = Math.min(1, contentWidth / (pageWidth / PX_PER_MM));
        const scale = options.fit === 'none' ? 1
          : options.fit === 'file' ? Math.min(sheetScale ?? 1, byWidth === 1 ? Infinity : byWidth)
          : options.fit === 'page' ? Math.min(byWidth, contentHeight / ((tallestSegment + (options.sheetLabel ? CAPTION_PX : 0)) / PX_PER_MM))
          : byWidth;
        const captionHeight = options.sheetLabel ? Math.round(CAPTION_PX / clamp(scale, 0.3, 1)) : 0;
        const usableHeight = options.fit === 'page' ? Infinity : contentHeight * PX_PER_MM / scale - captionHeight;
        const rowPages = splitRows(sheet.rows, usableHeight, titleRows, forcedRows);
        const rasterScale = clamp(scale * TARGET_DPI / 96, 0.6, 2);

        for (const [pageIndex, rowIndexes] of rowPages.entries()) {
          if (++pageCount > MAX_PAGES) throw new ConversionError(`ページ数が${MAX_PAGES}を超えました。用紙サイズを大きくするか、シートを分けてください`);
          const columnLabel = blocks.length > 1 ? `・列${block.start + 1}〜${block.end + 1}` : '';
          const caption = !options.sheetLabel ? ''
            : rowPages.length > 1 || blocks.length > 1 ? `${sheet.name}（${pageIndex + 1}/${rowPages.length}${columnLabel}）`
            : sheet.name;
          const repeatedRows = titleRows.filter(index => index < rowIndexes[0]);
          const lines = repeatedRows.length ? [...repeatedRows, ...rowIndexes] : rowIndexes;
          const canvas = drawPage({ sheet, columns, lines, caption, captionHeight, scale, rasterScale, gridLines });
          // 写真を含むページはJPEG、罫線と文字だけのページはPNGの方が軽く鮮明
          const image = sheet.pictures.length
            ? new Uint8Array(await toImageBytes(canvas, 'image/jpeg', 0.92))
            : new Uint8Array(await toImageBytes(canvas, 'image/png'));
          let width = canvas.width / rasterScale * scale / PX_PER_MM;
          let height = canvas.height / rasterScale * scale / PX_PER_MM;
          canvas.width = canvas.height = 0; // 大きなキャンバスを早めに解放する
          if (height > contentHeight) { width *= contentHeight / height; height = contentHeight; }
          if (!pdf) pdf = new jsPDF({ unit:'mm', format:[paper.width, paper.height], orientation: landscape ? 'landscape' : 'portrait', compress:true });
          else pdf.addPage([paper.width, paper.height], landscape ? 'landscape' : 'portrait');
          const left = sheet.print.horizontalCentered ? Math.max(margin.left, (paper[landscape ? 'height' : 'width'] - width) / 2) : margin.left;
          pdf.addImage(image, sheet.pictures.length ? 'JPEG' : 'PNG', left, margin.top, width, height);
        }
      }
    }
    return pdf!.output('blob');
  } catch (error) {
    if (error instanceof ConversionError) throw error;
    const message = error instanceof Error ? error.message : '';
    throw new Error(`ExcelファイルをPDFに変換できませんでした${message ? `（${message}）` : ''}`);
  } finally {
    for (const sheet of sheets) for (const picture of sheet.pictures) picture.source.close();
  }
}
