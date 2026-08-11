import { FileSpreadsheet, FileImage, Images, Sheet, TableProperties } from 'lucide-react';
import type { ToolId } from '../types';
export const tools: { id: ToolId; title: string; short: string; description: string; accept: string; multiple: boolean; icon: typeof Sheet; tone: string }[] = [
  { id:'excel-csv', title:'Excel → CSV', short:'Excel → CSV', description:'シートを選んでCSVへ変換', accept:'.xlsx,.xls', multiple:true, icon:FileSpreadsheet, tone:'mint' },
  { id:'csv-excel', title:'CSV → Excel', short:'CSV → Excel', description:'CSVを.xlsx形式へ変換', accept:'.csv,.tsv', multiple:true, icon:Sheet, tone:'blue' },
  { id:'excel-markdown', title:'Excel → Markdown', short:'Excel → Markdown', description:'シートをMarkdownテーブルへ変換', accept:'.xlsx,.xls', multiple:true, icon:TableProperties, tone:'blue' },
  { id:'images-pdf', title:'画像 → PDF', short:'画像 → PDF', description:'複数の画像を1つのPDFへ', accept:'image/png,image/jpeg', multiple:true, icon:Images, tone:'violet' },
  { id:'image-format', title:'画像形式変換', short:'PNG・JPEG・WebP', description:'PNG・JPEG・WebPを相互変換', accept:'image/png,image/jpeg,image/webp', multiple:true, icon:FileImage, tone:'orange' },
];

export type ImageFormat = 'png' | 'jpeg' | 'webp';
export type ConversionSource = 'excel' | 'csv' | 'image';

export const conversionGroups: {
  id: ConversionSource;
  label: string;
  icon: typeof Sheet;
  outputs: { label: string; tool: ToolId; format?: ImageFormat }[];
}[] = [
  { id:'excel', label:'Excel', icon:FileSpreadsheet, outputs:[
    { label:'CSV', tool:'excel-csv' },
    { label:'Markdown', tool:'excel-markdown' },
  ] },
  { id:'csv', label:'CSV', icon:Sheet, outputs:[{ label:'Excel', tool:'csv-excel' }] },
  { id:'image', label:'画像', icon:FileImage, outputs:[
    { label:'PDF', tool:'images-pdf' },
    { label:'PNG', tool:'image-format', format:'png' },
    { label:'JPEG', tool:'image-format', format:'jpeg' },
    { label:'WebP', tool:'image-format', format:'webp' },
  ] },
];
