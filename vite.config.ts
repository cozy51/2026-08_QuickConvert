import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
// worker.format:'es' — ffmpeg.wasm は変換エンジンのワーカーを type:'module' で起動するため。
export default defineConfig({ plugins: [react()], worker: { format: 'es' }, optimizeDeps: { exclude: ['@ffmpeg/ffmpeg', '@ffmpeg/util'] } });
