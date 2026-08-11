export type ToolId = 'excel-csv' | 'csv-excel' | 'images-pdf' | 'image-format';
export type OutputFile = { name: string; blob: Blob; url: string };
export type Status = 'idle' | 'processing' | 'done' | 'error';
