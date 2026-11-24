import test from 'node:test';
import assert from 'node:assert/strict';
import { EMU_PER_CSS_PIXEL, OfficeError, degreesToRotation, emuToCssPixels, toEmu } from '@office-ex/contracts';
import { FontMeasurementCache, containsPoint, hitTest, hitTestText, inverse, measureText, multiply, shapeTransform, transform, viewportTransform, type FontMetrics } from '@office-ex/render';
import type { Shape } from '@office-ex/core';

const near = (actual: number, expected: number) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ~= ${expected}`);
test('MOD-02: geometry conversions use exact base units and explicit integer rounding', () => {
  assert.equal(toEmu(1, 'inch'), 914400); assert.equal(toEmu(72, 'point'), 914400); assert.equal(toEmu(1440, 'twip'), 914400); assert.equal(toEmu(96, 'css-pixel'), 914400);
  assert.equal(emuToCssPixels(EMU_PER_CSS_PIXEL), 1); assert.equal(degreesToRotation(90), 5400000);
  assert.throws(() => toEmu(Infinity, 'point'), OfficeError); assert.throws(() => toEmu(Number.MAX_SAFE_INTEGER, 'point'), OfficeError);
});
test('M1.9: local rotation and inverse hit-testing share the same geometry', () => {
  const shape: Shape = { id: 'rotated', type: 'rect', geometry: { x: 100, y: 200, width: 100, height: 20, rotation: degreesToRotation(90) } };
  const matrix = shapeTransform(shape.geometry); const point = { x: 25, y: 5 }; const world = transform(matrix, point); const roundTrip = transform(inverse(matrix), world);
  near(roundTrip.x, point.x); near(roundTrip.y, point.y); assert.equal(containsPoint(shape, world), true);
  assert.equal(containsPoint(shape, { x: 100, y: 200 }), false);
  const combined = multiply(inverse(matrix), matrix); near(combined[0], 1); near(combined[3], 1); near(combined[4], 0); near(combined[5], 0);
  assert.throws(() => inverse([0, 0, 0, 0, 0, 0]), OfficeError);
});
test('M1.9: viewport zoom and device scaling never alter document EMU', () => {
  const viewport = { zoom: 2, pan: { x: 10, y: 20 }, origin: { x: EMU_PER_CSS_PIXEL, y: 2 * EMU_PER_CSS_PIXEL }, devicePixelRatio: 2 };
  const point = { x: 3 * EMU_PER_CSS_PIXEL, y: 4 * EMU_PER_CSS_PIXEL };
  assert.deepEqual(transform(viewportTransform(viewport), point), { x: 14, y: 24 });
  assert.deepEqual(transform(viewportTransform(viewport, true), point), { x: 28, y: 48 });
  const restored = transform(inverse(viewportTransform(viewport)), { x: 14, y: 24 }); near(restored.x, point.x); near(restored.y, point.y);
});
test('M1.9: hit testing respects paint order, ellipse edges and line distance', () => {
  const geometry = { x: 0, y: 0, width: 100, height: 100, rotation: 0 };
  const shapes: Shape[] = [{ id: 'back', type: 'rect', geometry }, { id: 'front', type: 'ellipse', geometry }];
  assert.equal(hitTest(shapes, { x: 50, y: 50 }), 'front'); assert.equal(hitTest(shapes, { x: 1, y: 1 }), 'back'); assert.equal(hitTest(shapes, { x: 120, y: 50 }), undefined);
  assert.equal(containsPoint({ id: 'line', type: 'line', geometry }, { x: 50, y: 51 }, 1), true);
  assert.equal(containsPoint({ id: 'line', type: 'line', geometry }, { x: 50, y: 60 }, 1), false);
});

// Injectable synthetic shaping data tests contracts; this is not a native font fidelity claim.
const metrics: FontMetrics = { environmentId: 'synthetic-metrics-v1', measure: (_text, request) => ({ advance: 30, ascent: 8, descent: 2, actualFamily: 'Test Font', clusters: request.direction === 'rtl' ? [{ start: 0, end: 1, left: 20, right: 30 }, { start: 1, end: 3, left: 0, right: 20 }] : [{ start: 0, end: 1, left: 0, right: 10 }, { start: 1, end: 3, left: 10, right: 30 }] }) };
test('M1.9: font shaping contract preserves grapheme indices, RTL visual order and fallback diagnostics', () => {
  const measured = measureText(metrics, 'A🙂', { fontFamily: 'Missing Font', fontSize: 127000 });
  assert.equal(measured.diagnostics[0]?.code, 'font.substituted'); assert.equal(hitTestText(measured, 4), 0); assert.equal(hitTestText(measured, 18), 1); assert.equal(hitTestText(measured, 28), 3);
  const rtl = measureText(metrics, 'A🙂', { fontFamily: 'Test Font' }, 'rtl'); assert.equal(hitTestText(rtl, 1), 3); assert.equal(hitTestText(rtl, 28), 0);
  assert.throws(() => measureText({ environmentId: 'broken', measure: () => ({ advance: 30, ascent: 8, descent: 2, actualFamily: 'Test', clusters: [{ start: 0, end: 2, left: 0, right: 30 }] }) }, 'A🙂', {}), OfficeError);
});
test('M1.9: font measurement caching is bounded and keys include the environment', () => {
  const cache = new FontMeasurementCache(1); const first = cache.measure(metrics, 'A🙂', {});
  assert.equal(cache.measure(metrics, 'A🙂', {}), first);
  const second = cache.measure({ ...metrics, environmentId: 'new-fonts' }, 'A🙂', {}); assert.notEqual(second, first); assert.equal(cache.size, 1);
  cache.clear(); assert.equal(cache.size, 0);
});
