import { jsPDF } from 'jspdf';
const isHeic=(file:File)=>/^image\/hei[cf]/i.test(file.type)||/\.(heic|heif)$/i.test(file.name);
const loadImage=(src:string)=>new Promise<HTMLImageElement>((ok,no)=>{const i=new Image();i.onload=()=>ok(i);i.onerror=()=>no(new Error('画像を読み込めませんでした。対応していない形式か、ファイルが壊れている可能性があります。'));i.src=src});
const loadBlob=async(blob:Blob)=>{const url=URL.createObjectURL(blob);try{return await loadImage(url)}finally{URL.revokeObjectURL(url)}};
// HEIC/HEIF を表示できるブラウザは Safari 系のみ。まず標準デコードを試し、失敗したら libheif(wasm) で展開する。
const decodeHeic=async(file:File)=>{const { heicTo }=await import('heic-to/csp');return heicTo({ blob:file, type:'image/png' })};
const loadSource=async(file:File,onProgress?:(message:string)=>void)=>{try{return await loadBlob(file)}catch(e){if(!isHeic(file))throw e}
 onProgress?.('HEIC画像を展開しています…');
 try{return await loadBlob(await decodeHeic(file))}catch{throw new Error('HEIC画像を読み込めませんでした。ファイルが壊れているか、対応していない形式の可能性があります。')}};
const drawImage=async(file:File,background=false,onProgress?:(message:string)=>void)=>{const img=await loadSource(file,onProgress);const c=document.createElement('canvas');c.width=img.naturalWidth||img.width;c.height=img.naturalHeight||img.height;if(!c.width||!c.height)throw new Error('画像のサイズを取得できませんでした');const ctx=c.getContext('2d')!;if(background){ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height)}ctx.drawImage(img,0,0,c.width,c.height);return c};
const canvasBlob=(canvas:HTMLCanvasElement,type:string,quality?:number)=>new Promise<Blob>((ok,no)=>canvas.toBlob(b=>b?ok(b):no(new Error('画像を変換できませんでした')),type,quality));
export async function imagesToPdf(files:File[],onProgress?:(message:string)=>void){let pdf:jsPDF|undefined; for(const [index,file] of files.entries()){const canvas=await drawImage(file,true,message=>onProgress?.(files.length>1?`${message}（${index+1}/${files.length}）`:message));const landscape=canvas.width>canvas.height;const w=landscape?297:210,h=landscape?210:297;if(!pdf)pdf=new jsPDF({orientation:landscape?'landscape':'portrait',unit:'mm',format:'a4'});else pdf.addPage('a4',landscape?'landscape':'portrait');const scale=Math.min(w/canvas.width,h/canvas.height);const iw=canvas.width*scale,ih=canvas.height*scale;pdf.addImage(canvas.toDataURL('image/jpeg',.92),'JPEG',(w-iw)/2,(h-ih)/2,iw,ih)}return pdf!.output('blob')}
export async function convertImage(file:File,format:'png'|'jpeg'|'webp',quality=.9,onProgress?:(message:string)=>void){return canvasBlob(await drawImage(file,format==='jpeg',onProgress),`image/${format}`,quality)}
