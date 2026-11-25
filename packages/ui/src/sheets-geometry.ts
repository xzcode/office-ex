/** Sparse row/column geometry; bounded viewport without dense arrays. */
export class GridAxis {
  constructor(count:number, defaultSize:number, overrides:Readonly<Record<string,{readonly height?:number;readonly width?:number;readonly hidden?:boolean}>>, size:"height"|"width") {
    this.count = count;
    this.defaultSize = defaultSize;
    let delta = 0;
    for (const [index, dimension] of Object.entries(overrides).map(([key, value]) => [Number(key), value] as const).sort((a, b) => a[0] - b[0])) {
      if (index < 1 || index > count) continue;
      const value = dimension.hidden ? 0 : dimension[size] === void 0 ? defaultSize : dimension[size] / 9525;
      delta += value - defaultSize;
      this.#indices.push(index);
      this.#deltas.push(delta);
    }
  }
  readonly count:number;
  readonly defaultSize:number;
  #indices:number[] = [];
  #deltas:number[] = [];
  start(index:number):number {
    let lo = 0, hi = this.#indices.length;
    while (lo < hi) {
      const middle = lo + hi >>> 1;
      if (this.#indices[middle] < index) lo = middle + 1;
      else hi = middle;
    }
    return (index - 1) * this.defaultSize + (lo ? this.#deltas[lo - 1] : 0);
  }
  size(index:number):number {
    return this.start(index + 1) - this.start(index);
  }
  get total() {
    return this.start(this.count + 1);
  }
  at(offset:number):number {
    let lo = 1, hi = this.count;
    while (lo < hi) {
      const middle = Math.floor((lo + hi) / 2);
      if (this.start(middle + 1) <= offset) lo = middle + 1;
      else hi = middle;
    }
    return lo;
  }
  visible(offset:number, length:number, limit2 = 200):number[] {
    const result = [];
    let index = this.at(Math.max(0, offset));
    const end = Math.max(0, offset) + length;
    while (index <= this.count && this.start(index) < end && result.length < limit2) {
      if (this.size(index) > 0) {
        result.push(index);
        index++;
      } else {
        const next = this.at(this.start(index) + 1e-5);
        index = Math.max(index + 1, next);
      }
    }
    return result;
  }
}
export function columnLabel(column:number):string {
  let value = "";
  while (column > 0) {
    column--;
    value = String.fromCharCode(65 + column % 26) + value;
    column = Math.floor(column / 26);
  }
  return value;
}
export function cellLabel(row:number, column:number):string {
  return columnLabel(column) + row;
}
