import { EMU_PER_CSS_PIXEL, graphemeBoundaries, OfficeError, assert } from '@office-ex/contracts';
import type { FontMeasurement, FontMetrics, FontRequest } from '@office-ex/render';
/** Chromium Canvas supplies both measure and draw. Full mixed-direction shaping is M2. */
export class CanvasFontMetrics implements FontMetrics {
  readonly environmentId = `chromium-canvas:${navigator.userAgent}:${devicePixelRatio}`;
  private readonly context:CanvasRenderingContext2D;
  constructor(canvas:HTMLCanvasElement = document.createElement('canvas')) {
    const context = canvas.getContext('2d');
    if (!context) throw new OfficeError('unsupported','Canvas text backend unavailable');
    this.context = context;
  }
  private font(request:FontRequest,family:string):string {
    return `${request.italic ? 'italic ' : ''}${request.bold ? 'bold ' : ''}${request.size / EMU_PER_CSS_PIXEL}px ${family}`;
  }
  private resolvedFamily(request:FontRequest):string {
    if (/^(sans-serif|serif|monospace|system-ui)$/.test(request.family)) return request.family;
    // Font Loading API alone cannot prove a local family is installed. Compare two fallbacks.
    const quoted = JSON.stringify(request.family);
    const probe = 'mmmmmmWWWWiii漢字000';
    return ['serif','monospace'].some(fallback => {
      this.context.font = this.font(request,fallback); const reference = this.context.measureText(probe).width;
      this.context.font = this.font(request,`${quoted},${fallback}`); return this.context.measureText(probe).width !== reference;
    }) ? request.family : 'sans-serif';
  }
  measure(text:string,request:FontRequest):FontMeasurement {
    assert(typeof request.family === 'string' && request.family.length > 0 && request.family.length <= 512 && !/[\x00-\x1f\x7f]/.test(request.family) && Number.isSafeInteger(request.size) && request.size > 0 && request.size <= 12_700 * 1638 && typeof request.bold === 'boolean' && typeof request.italic === 'boolean' && (request.direction === 'ltr' || request.direction === 'rtl'), 'invalid', 'Invalid font request');
    if (text.length > 16_384) throw new OfficeError('limit','Split long text into shaping runs');
    const actualFamily = this.resolvedFamily(request);
    this.context.font = this.font(request,/^(sans-serif|serif|monospace|system-ui)$/.test(actualFamily) ? actualFamily : JSON.stringify(actualFamily));
    this.context.direction = request.direction;
    // All layout coordinates use the left edge and alphabetic baseline, even in RTL runs.
    this.context.textAlign = 'left';
    this.context.textBaseline = 'alphabetic';
    const measured = this.context.measureText(text), advance = measured.width * EMU_PER_CSS_PIXEL;
    const boundaries = graphemeBoundaries(text); const clusters = [];
    let previous = 0;
    for (let i=0;i<boundaries.length-1;i++) {
      const end = boundaries[i+1]!;
      const next = this.context.measureText(text.slice(0,end)).width * EMU_PER_CSS_PIXEL;
      const a = request.direction === 'rtl' ? advance-next : previous;
      const b = request.direction === 'rtl' ? advance-previous : next;
      clusters.push({start:boundaries[i]!,end,left:Math.min(a,b),right:Math.max(a,b)}); previous = next;
    }
    return {advance,ascent:Math.max(0,measured.actualBoundingBoxAscent*EMU_PER_CSS_PIXEL),descent:Math.max(0,measured.actualBoundingBoxDescent*EMU_PER_CSS_PIXEL),actualFamily,clusters,clusterPrecision:'estimated'};
  }
  draw(text:string,request:FontRequest,x:number,y:number):void {
    assert(Number.isFinite(x) && Number.isFinite(y),'invalid','Invalid text drawing coordinates');
    this.measure(text,request); this.context.fillText(text,x/EMU_PER_CSS_PIXEL,y/EMU_PER_CSS_PIXEL);
  }
}
