// SPDX-License-Identifier: MIT
import type {Shape} from './model.js';
import type {DeepReadonly} from './session.js';
/** Includes every node exactly once, in source paint order. */
export function* walkSlideShapes<T extends Shape|DeepReadonly<Shape>>(shapes:readonly T[]):Generator<T>{
 for(const shape of shapes){yield shape;if(shape.type==='group')yield* walkSlideShapes(shape.children as readonly T[]);}
}
