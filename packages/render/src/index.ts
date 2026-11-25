import { EMU_PER_CSS_PIXEL, ROTATION_UNITS_PER_DEGREE, assert, graphemeBoundaries, type Diagnostic } from '@office-ex/contracts';
import { freezeData, validateGeometry, validateTextStyle, type DeepReadonly, type Geometry, type Shape, type TextStyle } from '@office-ex/core';

export interface Point { readonly x: number; readonly y: number }
/** Affine matrix [a,b,c,d,e,f], mapping x'=ax+cy+e, y'=bx+dy+f. */
export type Matrix = readonly [number, number, number, number, number, number];
export const IDENTITY: Matrix = Object.freeze([1, 0, 0, 1, 0, 0]);
function finite(values: readonly number[]): void { assert(values.every(Number.isFinite), 'invalid', 'Non-finite render geometry'); }
export function transform(matrix: Matrix, point: Point): Point {
  finite([...matrix, point.x, point.y]); const [a, b, c, d, e, f] = matrix; return { x: a * point.x + c * point.y + e, y: b * point.x + d * point.y + f };
}
export function multiply(left: Matrix, right: Matrix): Matrix {
  finite([...left, ...right]); const [a, b, c, d, e, f] = left; const [g, h, i, j, k, l] = right;
  const result: Matrix = [a * g + c * h, b * g + d * h, a * i + c * j, b * i + d * j, a * k + c * l + e, b * k + d * l + f]; finite(result); return result;
}
export function inverse(matrix: Matrix): Matrix {
  finite(matrix); const [a, b, c, d, e, f] = matrix; const determinant = a * d - b * c;
  assert(Number.isFinite(determinant) && Math.abs(determinant) > 1e-12, 'invalid', 'Singular render transform');
  const result: Matrix = [d / determinant, -b / determinant, -c / determinant, a / determinant, (c * f - d * e) / determinant, (b * e - a * f) / determinant]; finite(result); return result;
}
export function shapeTransform(geometry: Geometry | DeepReadonly<Geometry>): Matrix {
  validateGeometry(geometry);
  const radians = geometry.rotation / ROTATION_UNITS_PER_DEGREE * Math.PI / 180;
  const cosine = Math.cos(radians); const sine = Math.sin(radians); const cx = geometry.width / 2; const cy = geometry.height / 2;
  return [cosine, sine, -sine, cosine, geometry.x + cx - cosine * cx + sine * cy, geometry.y + cy - sine * cx - cosine * cy];
}
/** Group child coordinate system including scaling, flips and rotation. */
export function groupChildTransform(group:DeepReadonly<Extract<Shape,{type:'group'}>>):Matrix{
 const g=group.geometry,c=group.childSpace,fx=group.flipH?-1:1,fy=group.flipV?-1:1,sx=fx*g.width/c.width,sy=fy*g.height/c.height;
 return multiply(shapeTransform(g),[sx,0,0,sy,(group.flipH?g.width:0)-sx*c.x,(group.flipV?g.height:0)-sy*c.y]);
}
export function slideShapeParentTransform(shapes:readonly DeepReadonly<Shape>[],id:string,parent:Matrix=IDENTITY):Matrix|undefined{
 for(const shape of shapes){
  if(shape.id===id)return parent;
  if(shape.type==='group'){const found=slideShapeParentTransform(shape.children,id,multiply(parent,groupChildTransform(shape)));if(found)return found;}
 }
 return undefined;
}
export interface Viewport { readonly zoom: number; readonly pan: Point; readonly origin: Point; readonly devicePixelRatio: number }
export function viewportTransform(viewport: Viewport, devicePixels = false): Matrix {
  finite([viewport.zoom, viewport.devicePixelRatio, viewport.pan.x, viewport.pan.y, viewport.origin.x, viewport.origin.y]);
  assert(viewport.zoom > 0 && viewport.zoom <= 64 && viewport.devicePixelRatio > 0 && viewport.devicePixelRatio <= 16, 'invalid', 'Invalid viewport scale');
  const device = devicePixels ? viewport.devicePixelRatio : 1; const scale = viewport.zoom / EMU_PER_CSS_PIXEL * device;
  return [scale, 0, 0, scale, (viewport.pan.x - viewport.origin.x * viewport.zoom / EMU_PER_CSS_PIXEL) * device, (viewport.pan.y - viewport.origin.y * viewport.zoom / EMU_PER_CSS_PIXEL) * device];
}
export function containsPoint(shape: DeepReadonly<Shape>, point: Point, tolerance = 0): boolean {
  assert(Number.isFinite(tolerance) && tolerance >= 0, 'invalid', 'Invalid hit tolerance');
  let local = transform(inverse(shapeTransform(shape.geometry)), point); const { width, height } = shape.geometry;
  local={x:shape.flipH?width-local.x:local.x,y:shape.flipV?height-local.y:local.y};
  if (shape.type === 'line') {
    const lengthSquared = width * width + height * height;
    const parameter = lengthSquared === 0 ? 0 : Math.max(0, Math.min(1, (local.x * width + local.y * height) / lengthSquared));
    return Math.hypot(local.x - width * parameter, local.y - height * parameter) <= tolerance + (shape.stroke?.width ?? 0) / 2;
  }
  if (shape.type === 'ellipse') {
    const rx = width / 2 + tolerance; const ry = height / 2 + tolerance;
    return rx > 0 && ry > 0 && ((local.x - width / 2) / rx) ** 2 + ((local.y - height / 2) / ry) ** 2 <= 1;
  }
  if(shape.type==='triangle'||shape.type==='diamond'||shape.type==='parallelogram'){
    if(width<=0||height<=0)return false;
    const u=local.x/width,v=local.y/height,eps=tolerance/Math.max(1,Math.min(width,height));
    if(shape.type==='triangle')return v>=-eps&&v<=1+eps&&Math.abs(u-.5)<=v/2+eps;
    if(shape.type==='diamond')return Math.abs(u-.5)+Math.abs(v-.5)<=.5+eps;
    const inset=.25*Math.min(width,height)/width;return v>=-eps&&v<=1+eps&&u>=inset*(1-v)-eps&&u<=1-inset*v+eps;
  }
  return local.x >= -tolerance && local.y >= -tolerance && local.x <= width + tolerance && local.y <= height + tolerance;
}
/** Shapes are in paint order; selection resolves the topmost hit. */
export function hitTest(shapes: readonly DeepReadonly<Shape>[], point: Point, tolerance = 0): string | undefined {
  for (let index = shapes.length - 1; index >= 0; index--) if (containsPoint(shapes[index]!, point, tolerance)) return shapes[index]!.id;
  return undefined;
}

export interface FontRequest { readonly family: string; readonly size: number; readonly bold: boolean; readonly italic: boolean; readonly direction: 'ltr' | 'rtl' }
export interface TextCluster { readonly start: number; readonly end: number; readonly left: number; readonly right: number }
export interface FontMeasurement {
  readonly advance: number; readonly ascent: number; readonly descent: number;
  readonly actualFamily: string; readonly clusters: readonly TextCluster[];
  /** Canvas prefix advances are estimates; a native shaping engine can supply exact cluster positions. */
  readonly clusterPrecision?: 'exact' | 'estimated';
}
/** The same native shaping backend must measure and draw; EMU values and UTF-16 indices. */
export interface FontMetrics {
  readonly environmentId: string;
  measure(text: string, request: FontRequest): FontMeasurement;
}
export interface MeasuredText { readonly text: string; readonly request: FontRequest; readonly metrics: FontMeasurement; readonly diagnostics: readonly Diagnostic[] }
export function measureText(metrics: FontMetrics, text: string, style: TextStyle, direction: 'ltr' | 'rtl' = 'ltr'): MeasuredText {
  validateTextStyle(style); assert(text.length <= 1_000_000, 'limit', 'Text measurement budget exceeded');
  const boundaries = new Set(graphemeBoundaries(text));
  const request: FontRequest = { family: style.fontFamily ?? 'sans-serif', size: style.fontSize ?? 152_400, bold: style.bold ?? false, italic: style.italic ?? false, direction };
  const measured = metrics.measure(text, request);
  finite([measured.advance, measured.ascent, measured.descent]);
  assert(measured.advance >= 0 && measured.ascent >= 0 && measured.descent >= 0 && typeof measured.actualFamily === 'string' && measured.actualFamily.length > 0, 'invalid', 'Invalid font measurement');
  const sorted = [...measured.clusters].sort((a, b) => a.start - b.start); let end = 0;
  for (const cluster of sorted) {
    finite([cluster.left, cluster.right]);
    assert(cluster.start === end && cluster.end > cluster.start && boundaries.has(cluster.start) && boundaries.has(cluster.end) && cluster.left <= cluster.right, 'invalid', 'Font clusters split or omit text'); end = cluster.end;
  }
  assert(end === text.length, 'invalid', 'Font measurement does not cover text');
  const diagnostics: Diagnostic[] = measured.actualFamily === request.family ? [] : [{ code: 'font.substituted', severity: 'warning', featureId: 'WRT-02', messageKey: 'font.substituted', safeAction: 'none' }];
  if (measured.clusterPrecision === 'estimated') diagnostics.push({ code: 'font.cluster-positions-estimated', severity: 'info', featureId: 'WRT-02', messageKey: 'font.cluster-positions-estimated', safeAction: 'none' });
  return freezeData({ text, request, metrics: { ...measured, clusters: measured.clusters.map((cluster) => ({ ...cluster })) }, diagnostics });
}
/** Logical UTF-16 selection stays separate from visual positions, including RTL clusters. */
export function hitTestText(measured: MeasuredText, x: number): number {
  assert(Number.isFinite(x), 'invalid', 'Invalid text hit coordinate');
  const clusters = [...measured.metrics.clusters].sort((a, b) => a.left - b.left);
  if (clusters.length === 0) return 0;
  const leftOffset = (cluster: TextCluster): number => measured.request.direction === 'rtl' ? cluster.end : cluster.start;
  const rightOffset = (cluster: TextCluster): number => measured.request.direction === 'rtl' ? cluster.start : cluster.end;
  for (const cluster of clusters) { if (x < (cluster.left + cluster.right) / 2) return leftOffset(cluster); if (x <= cluster.right) return rightOffset(cluster); }
  return rightOffset(clusters[clusters.length - 1]!);
}

export class FontMeasurementCache {
  #entries = new Map<string, MeasuredText>();
  constructor(readonly maxEntries = 1000) { assert(Number.isSafeInteger(maxEntries) && maxEntries > 0 && maxEntries <= 10_000, 'invalid', 'Invalid font cache budget'); }
  measure(metrics: FontMetrics, text: string, style: TextStyle, direction: 'ltr' | 'rtl' = 'ltr'): MeasuredText {
    const key = JSON.stringify([metrics.environmentId, text, style, direction]);
    const existing = this.#entries.get(key); if (existing) { this.#entries.delete(key); this.#entries.set(key, existing); return existing; }
    const result = measureText(metrics, text, style, direction);
    // Oversized runs bypass the cache; they still use the same validated shaping path.
    if (text.length <= 16_384) {
      this.#entries.set(key, result); while (this.#entries.size > this.maxEntries) this.#entries.delete(this.#entries.keys().next().value!);
    }
    return result;
  }
  clear(): void { this.#entries.clear(); }
  get size(): number { return this.#entries.size; }
}

export * from './writer-layout.js';

export * from './writer-pagination.js';

export * from './slides-scene.js';
