import { FileSpreadsheet, FileImage, Images, Sheet } from 'lucide-react';
import type { ToolId } from '../types';
export const tools: { id: ToolId; title: string; short: string; description: string; accept: string; multiple: boolean; icon: typeof Sheet; tone: string }[] = [
  { id:'excel-csv', title:'Excel → CSV', short:'Excel → CSV', description:'シートを選んでCSVへ変換', accept:'.xlsx,.xls', multiple:true, icon:FileSpreadsheet, tone:'mint' },
  { id:'csv-excel', title:'CSV → Excel', short:'CSV → Excel', description:'CSVを.xlsx形式へ変換', accept:'.csv,.tsv', multiple:true, icon:Sheet, tone:'blue' },
  { id:'images-pdf', title:'画像 → PDF', short:'画像 → PDF', description:'複数の画像を1つのPDFへ', accept:'image/png,image/jpeg', multiple:true, icon:Images, tone:'violet' },
  { id:'image-format', title:'画像形式変換', short:'PNG・JPEG・WebP', description:'PNG・JPEG・WebPを相互変換', accept:'image/png,image/jpeg,image/webp', multiple:true, icon:FileImage, tone:'orange' },
];
