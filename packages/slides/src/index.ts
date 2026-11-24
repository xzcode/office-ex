import { assert } from '@office-ex/contracts';
import { patchCommand, type DocumentCommand, type DocumentSnapshot, type Geometry, type Shape } from '@office-ex/core';
export type { SlidesDocument, Slide, Shape, Geometry } from '@office-ex/core';
function slideIndex(snapshot: DocumentSnapshot, slideId: string): number {
  assert(snapshot.document.kind === 'slides', 'invalid', 'Expected slides document');
  const index = snapshot.document.slides.findIndex((slide) => slide.id === slideId); assert(index >= 0, 'invalid', 'Slide does not exist'); return index;
}
export function insertShapes(snapshot: DocumentSnapshot, slideId: string, index: number, shapes: readonly Shape[]): DocumentCommand {
  return patchCommand(snapshot, 'SLD-02', [{ op: 'splice', path: ['slides', slideIndex(snapshot, slideId), 'shapes'], index, deleteCount: 0, items: shapes }]);
}
/** The UI previews a drag separately and calls this once when the gesture commits. */
export function setShapeGeometry(snapshot: DocumentSnapshot, slideId: string, shapeId: string, geometry: Geometry): DocumentCommand {
  const index = slideIndex(snapshot, slideId); assert(snapshot.document.kind === 'slides', 'invalid', 'Expected slides document');
  const shapeIndex = snapshot.document.slides[index]!.shapes.findIndex((shape) => shape.id === shapeId); assert(shapeIndex >= 0, 'invalid', 'Shape does not exist');
  return patchCommand(snapshot, 'SLD-02', [{ op: 'set', path: ['slides', index, 'shapes', shapeIndex, 'geometry'], value: geometry }]);
}
