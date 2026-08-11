export type ToolId = 'excel-csv' | 'csv-excel' | 'excel-markdown' | 'images-pdf' | 'image-format' | 'video-audio';
export type OutputFile = { name: string; blob: Blob; url: string };
export type Status = 'idle' | 'processing' | 'done' | 'error';
