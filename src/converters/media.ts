import { FFmpeg } from '@ffmpeg/ffmpeg';
import { fetchFile } from '@ffmpeg/util';
// 変換エンジン（ffmpeg.wasm）は外部CDNに依存せず、アプリと同じオリジンから配信する。
import coreURL from '@ffmpeg/core?url';
import wasmURL from '@ffmpeg/core/wasm?url';
import classWorkerURL from '@ffmpeg/ffmpeg/worker?worker&url';

let ffmpeg: FFmpeg | undefined;
let loading: Promise<FFmpeg> | undefined;

const absolute = (url: string) => new URL(url, window.location.href).href;

const withTimeout = <T>(promise: Promise<T>, milliseconds: number, message: string) => new Promise<T>((resolve, reject) => {
  const timer = window.setTimeout(() => reject(new Error(message)), milliseconds);
  promise.then(
    value => { window.clearTimeout(timer); resolve(value); },
    error => { window.clearTimeout(timer); reject(error); },
  );
});

async function getFFmpeg(onStatus?: (message: string) => void) {
  if (ffmpeg?.loaded) return ffmpeg;
  if (!loading) {
    loading = (async () => {
      const instance = new FFmpeg();
      try {
        onStatus?.('変換エンジンを読み込んでいます…');
        await withTimeout(instance.load({
          classWorkerURL: absolute(classWorkerURL),
          coreURL: absolute(coreURL),
          wasmURL: absolute(wasmURL),
        }), 120_000, '変換エンジンの読み込みがタイムアウトしました');
        ffmpeg = instance;
        return instance;
      } catch (error) {
        instance.terminate();
        throw error;
      }
    })();
  }
  try {
    return await loading;
  } catch (error) {
    loading = undefined;
    const detail = error instanceof Error && error.message ? `（${error.message}）` : '';
    throw new Error(`変換エンジンを読み込めませんでした${detail}`);
  }
}

export async function videoToMp3(file: File, onStatus?: (message: string) => void) {
  const engine = await getFFmpeg(onStatus);
  const token = crypto.randomUUID();
  const inputName = `input-${token}.mp4`;
  const outputName = `output-${token}.mp3`;
  const progressHandler = ({ progress }: { progress: number }) => onStatus?.(`音声を変換しています… ${Math.min(100, Math.max(0, Math.round(progress * 100)))}%`);
  engine.on('progress', progressHandler);
  try {
    onStatus?.('動画を準備しています…');
    await engine.writeFile(inputName, await fetchFile(file));
    onStatus?.('音声を変換しています… 0%');
    const exitCode = await withTimeout(engine.exec(['-nostdin', '-y', '-i', inputName, '-map', '0:a:0', '-vn', '-codec:a', 'libmp3lame', '-q:a', '2', outputName]), 5 * 60_000, '動画変換がタイムアウトしました');
    if (exitCode !== 0) throw new Error('動画から音声を抽出できませんでした');
    const data = await engine.readFile(outputName);
    if (typeof data === 'string') throw new Error('音声データを読み込めませんでした');
    return new Blob([new Uint8Array(data).buffer], { type:'audio/mpeg' });
  } catch (error) {
    if (error instanceof Error && error.message.includes('タイムアウト')) {
      engine.terminate();
      ffmpeg = undefined;
      loading = undefined;
    }
    throw error;
  } finally {
    engine.off('progress', progressHandler);
    if (engine.loaded) {
      await engine.deleteFile(inputName).catch(() => undefined);
      await engine.deleteFile(outputName).catch(() => undefined);
    }
  }
}
