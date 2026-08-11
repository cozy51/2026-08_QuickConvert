# QuickConvert

Excel・CSV・画像をドラッグ＆ドロップで変換する、React + TypeScript 製の Web アプリです。ファイルはサーバーへアップロードせず、変換処理はブラウザ内で完結します。

## 対応する変換

- Excel（`.xlsx` / `.xls`）→ CSV（シート・文字コード・区切り文字を選択可能）
- CSV → Excel（`.xlsx`）
- Excel（`.xlsx` / `.xls`）→ Markdownテーブル（シート・見出し行を選択可能）
- 複数の PNG / JPEG → 1つの PDF
- PNG / JPEG / WebP の相互変換
- MP4からMP3音声を抽出
- 複数ファイルの個別ダウンロード / ZIP 一括ダウンロード

## セットアップ

Node.js 18 以降を用意してください。

```bash
npm install
npm run dev
```

表示されたローカル URL（通常は `http://localhost:5173`）を開きます。

## ビルドと確認

```bash
npm run lint
npm run build
npm run preview
```

成果物は `dist/` に生成されます。

## Vercel への公開

1. このリポジトリを GitHub / GitLab / Bitbucket に push します。
2. Vercel の **Add New Project** からリポジトリを選択します。
3. Framework Preset は **Vite**、Build Command は `npm run build`、Output Directory は `dist` のままデプロイします。

Vercel CLI を使う場合は、プロジェクトルートで `npx vercel` を実行してください。バックエンドや環境変数は不要です。

## 構成

- `src/config/tools.ts` — 変換ツール一覧（ツール追加時の設定データ）
- `src/converters/` — 種類ごとに分離した変換ロジック
- `src/App.tsx` — UI と変換フロー

## MP4 → MP3 の変換エンジンについて

MP4からの音声抽出には ffmpeg.wasm（`@ffmpeg/ffmpeg` + `@ffmpeg/core`）を使います。変換エンジン（`ffmpeg-core.js` / `ffmpeg-core.wasm` / ワーカー）は npm 依存としてバンドルし、`dist/assets/` へ出力してアプリと同じオリジンから配信します。外部CDNへは取得しにいかないため、CDNがブロックされている環境やオフラインでも読み込めます。

`ffmpeg-core.wasm` は約31MBあり、MP4変換を実行したときに初めて読み込まれます（他の変換では読み込まれません）。エンジンはブラウザキャッシュに載るため、2回目以降は読み込みが不要です。ワーカーは `type: 'module'` で起動するため、`vite.config.ts` で `worker.format: 'es'` を指定しています。

## プライバシー

SheetJS、jsPDF、Canvas API、JSZip、ffmpeg.wasm を使用し、すべての処理をクライアント側で行います。選択したファイルや変換結果が QuickConvert のサーバーへ送信されることはありません。MP4変換の変換エンジンもアプリと同じオリジンから配信し、動画自体はブラウザの外へ送信されません。変換結果は画面上の「変換ファイルを破棄」からオブジェクト URL を解放できます。
