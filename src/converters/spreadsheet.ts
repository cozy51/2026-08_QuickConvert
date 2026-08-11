import * as XLSX from 'xlsx';
import Encoding from 'encoding-japanese';

export async function sheetNames(file: File) { const wb = XLSX.read(await file.arrayBuffer(), { type:'array', cellDates:true, cellText:true }); return wb.SheetNames; }
export async function previewSheet(file: File, sheet?: string) { const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,cellText:true}); return XLSX.utils.sheet_to_json<string[]>(wb.Sheets[sheet || wb.SheetNames[0]],{header:1,raw:false,defval:'',blankrows:true}).slice(0,20); }
export async function excelToCsv(file: File, sheet: string | undefined, delimiter: string, encoding: string) {
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,cellText:true}); const ws=wb.Sheets[sheet || wb.SheetNames[0]];
  const csv=XLSX.utils.sheet_to_csv(ws,{FS:delimiter,blankrows:true,forceQuotes:false}); let bytes:Uint8Array;
  if(encoding==='sjis') bytes=new Uint8Array(Encoding.convert(Encoding.stringToCode(csv),{to:'SJIS',from:'UNICODE'}));
  else { const prefix=encoding==='bom'?'\ufeff':''; bytes=new TextEncoder().encode(prefix+csv); }
  return new Blob([bytes],{type:'text/csv;charset='+encoding});
}
export async function csvToExcel(file: File, delimiter: string) {
  const text=await file.text(); const wb=XLSX.read(text,{type:'string',FS:delimiter,raw:true}); return new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
