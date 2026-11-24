// SPDX-License-Identifier: MIT
import { OfficeError } from '@office-ex/contracts';

/** Product budgets, not the theoretical limits of ZIP or XML. */
export const DEFAULT_CONTAINER_LIMITS = Object.freeze({
  maxFileBytes: 100 * 1024 * 1024,
  maxEntries: 10_000,
  maxEntryBytes: 64 * 1024 * 1024,
  maxTotalBytes: 512 * 1024 * 1024,
  maxCompressionRatio: 200,
  maxXmlBytes: 32 * 1024 * 1024,
  maxXmlDepth: 128,
  maxXmlNodes: 1_000_000,
});
export type ContainerLimits = { -readonly [K in keyof typeof DEFAULT_CONTAINER_LIMITS]: number };
export interface ContainerOptions {
  /** Callers may tighten budgets, never silently raise the product defaults. */
  limits?: Partial<ContainerLimits>;
  signal?: Pick<AbortSignal, 'aborted'>;
}
export function containerLimits(options: ContainerOptions = {}): ContainerLimits {
  const result: ContainerLimits = { ...DEFAULT_CONTAINER_LIMITS };
  if (options.limits) {
    for (const key of Object.keys(options.limits) as (keyof ContainerLimits)[]) {
      const value = options.limits[key];
      if (value === undefined) continue;
      if (!Number.isSafeInteger(value) || value < 1 || value > DEFAULT_CONTAINER_LIMITS[key]) {
        throw new OfficeError('invalid', 'Invalid container budget');
      }
      result[key] = value;
    }
  }
  return result;
}
export function checkCancelled(options: ContainerOptions): void {
  if (options.signal?.aborted) throw new OfficeError('cancelled', 'Container operation cancelled');
}
export function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}
export function joinBytes(chunks: readonly Uint8Array[], length: number): Uint8Array {
  const result = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length; }
  return result;
}
