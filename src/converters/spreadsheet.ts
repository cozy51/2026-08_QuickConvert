import * as XLSX from 'xlsx';
import Encoding from 'encoding-japanese';

export async function sheetNames(file: File) { const wb = XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:true, cellText:true }); return wb.SheetNames; }
export async function previewSheet(file: File, sheet?: string) { const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,cellText:true}); return XLSX.utils.sheet_to_json<string[]>(wb.Sheets[sheet || wb.SheetNames[0]],{header:1,raw:false,defval:'',blankrows:true}).slice(0,20); }
export async function excelToCsv(file: File, sheet: string | undefined, delimiter: string, encoding: string) {
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,cellText:true}); const ws=wb.Sheets[sheet || wb.SheetNames[0]];
  const csv=XLSX.utils.sheet_to_csv(ws,{FS:delimiter,blankrows:true,forceQuotes:false}); let bytes:Uint8Array;
  if(encoding==='sjis') bytes=new Uint8Array(Encoding.convert(Encoding.stringToCode(csv),{to:'SJIS',from:'UNICODE'}));
  else { const prefix=encoding==='bom'?'\ufeff':''; bytes=new TextEncoder().encode(prefix+csv); }
  return new Blob([new Uint8Array(bytes).buffer],{type:'text/csv;charset='+encoding});
}
export async function csvToExcel(file: File, delimiter: string) {
  const text=await file.text(); const wb=XLSX.read(text,{type:'string',FS:delimiter,raw:true}); return new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

const escapeMarkdownCell = (value: string) => value.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>');

export function rowsToMarkdown(rows: string[][], firstRowAsHeader = true) {
  if (rows.length === 0) return '';
  const columnCount = Math.max(...rows.map(row => row.length));
  const normalized = rows.map(row => Array.from({ length: columnCount }, (_, index) => escapeMarkdownCell(String(row[index] ?? ''))));
  const header = firstRowAsHeader ? normalized[0] : Array.from({ length: columnCount }, (_, index) => `列${index + 1}`);
  const body = firstRowAsHeader ? normalized.slice(1) : normalized;
  const line = (row: string[]) => `| ${row.join(' | ')} |`;
  return [line(header), line(header.map(() => '---')), ...body.map(line)].join('\n') + '\n';
}

export async function excelToMarkdown(file: File, sheet?: string, firstRowAsHeader = true) {
  const wb = XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:true, cellText:true });
  const ws = wb.Sheets[sheet || wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header:1, raw:false, defval:'', blankrows:true });
  return new Blob([rowsToMarkdown(rows, firstRowAsHeader)], { type:'text/markdown;charset=utf-8' });
}
