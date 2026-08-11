import { FFmpeg } from '@ffmpeg/ffmpeg';
import { fetchFile, toBlobURL } from '@ffmpeg/util';

const coreBaseUrl = 'https://cdn.jsdelivr.net/npm/@ffmpeg/core@0.12.6/dist/esm';
let ffmpeg: FFmpeg | undefined;
let loading: Promise<FFmpeg> | undefined;

async function getFFmpeg() {
  if (ffmpeg?.loaded) return ffmpeg;
  if (!loading) {
    loading = (async () => {
      const instance = new FFmpeg();
      await instance.load({
        coreURL: await toBlobURL(`${coreBaseUrl}/ffmpeg-core.js`, 'text/javascript'),
        wasmURL: await toBlobURL(`${coreBaseUrl}/ffmpeg-core.wasm`, 'application/wasm'),
      });
      ffmpeg = instance;
      return instance;
    })();
  }
  try {
    return await loading;
  } catch {
    loading = undefined;
    throw new Error('音声変換エンジンを読み込めませんでした。ネットワーク接続を確認してください');
  }
}

export async function videoToMp3(file: File) {
  const engine = await getFFmpeg();
  const token = crypto.randomUUID();
  const inputName = `input-${token}.mp4`;
  const outputName = `output-${token}.mp3`;
  try {
    await engine.writeFile(inputName, await fetchFile(file));
    const exitCode = await engine.exec(['-i', inputName, '-vn', '-codec:a', 'libmp3lame', '-q:a', '2', outputName]);
    if (exitCode !== 0) throw new Error('動画から音声を抽出できませんでした');
    const data = await engine.readFile(outputName);
    if (typeof data === 'string') throw new Error('音声データを読み込めませんでした');
    return new Blob([new Uint8Array(data).buffer], { type:'audio/mpeg' });
  } finally {
    await engine.deleteFile(inputName).catch(() => undefined);
    await engine.deleteFile(outputName).catch(() => undefined);
  }
}
