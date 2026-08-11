import { jsPDF } from 'jspdf';

// PPTXはブラウザ内でDOMに描画し、スライドごとに画像化してPDFへ貼る。
// PowerPointの完全再現ではなく「見た目を保った配布用PDF」を目的とする。
// 変換エンジン（レンダラ・html2canvas）はPPTX変換時のみ読み込む。
const DPI_TARGET = 1920; // 横幅の目標ピクセル数（16:9で約144dpi）

const waitForImages = (root: HTMLElement) => Promise.all(
  [...root.querySelectorAll('img')].map(img => img.complete ? undefined : img.decode().catch(() => undefined)),
);

export async function pptxToPdf(file: File, onStatus?: (message: string) => void) {
  onStatus?.('変換エンジンを読み込んでいます…');
  const [{ PptxViewer }, { default: html2canvas }] = await Promise.all([
    import('@aiden0z/pptx-renderer/browser'),
    import('html2canvas'),
  ]);

  const host = document.createElement('div');
  host.setAttribute('aria-hidden', 'true');
  Object.assign(host.style, { position: 'fixed', top: '0', left: '-20000px', width: 'max-content', background: '#fff', pointerEvents: 'none' });
  document.body.appendChild(host);

  const rendered = new Map<number, HTMLElement>();
  let viewer: InstanceType<typeof PptxViewer> | undefined;
  try {
    onStatus?.('スライドを解析しています…');
    viewer = new PptxViewer(host, {
      fitMode: 'none',
      pdfjs: false,
      onSlideRendered: (index, element) => rendered.set(index, element),
    });
    await viewer.open(await file.arrayBuffer(), { renderMode: 'list', listOptions: { windowed: false } });

    const slides = rendered.size
      ? [...rendered.entries()].sort((a, b) => a[0] - b[0]).map(entry => entry[1])
      : [...host.querySelectorAll<HTMLElement>('[data-slide-index]')];
    if (!slides.length) throw new Error('スライドを読み込めませんでした');

    const pageWidth = viewer.slideWidth * 0.75;
    const pageHeight = viewer.slideHeight * 0.75;
    const orientation = pageWidth >= pageHeight ? 'landscape' : 'portrait';
    const scale = Math.min(2, Math.max(1, DPI_TARGET / viewer.slideWidth));
    await document.fonts?.ready;

    let pdf: jsPDF | undefined;
    for (const [index, element] of slides.entries()) {
      onStatus?.(`スライドを変換しています… ${index + 1}/${slides.length}`);
      await waitForImages(element);
      const canvas = await html2canvas(element, { scale, backgroundColor: '#ffffff', logging: false, useCORS: true });
      const image = canvas.toDataURL('image/jpeg', 0.92);
      canvas.width = canvas.height = 0; // 大きなキャンバスを早めに解放する
      if (!pdf) pdf = new jsPDF({ unit: 'pt', format: [pageWidth, pageHeight], orientation, compress: true });
      else pdf.addPage([pageWidth, pageHeight], orientation);
      pdf.addImage(image, 'JPEG', 0, 0, pageWidth, pageHeight);
    }
    return pdf!.output('blob');
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (/zip|central directory/i.test(message)) throw new Error('PowerPointファイルを読み込めませんでした（ファイルが壊れているか、.pptx形式ではない可能性があります）');
    throw new Error(`PowerPointファイルを変換できませんでした${message ? `（${message}）` : ''}`);
  } finally {
    viewer?.destroy();
    host.remove();
  }
}
