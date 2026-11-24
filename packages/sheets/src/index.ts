import { assert, assertCellAddress } from '@office-ex/contracts';
import { patchCommand, type Cell, type DocumentCommand, type DocumentSnapshot, type PatchOperation } from '@office-ex/core';
export type { SheetsDocument, Sheet, Cell, CellValue } from '@office-ex/core';
export function cellKey(row: number, column: number): string { assertCellAddress(row, column); return `${row}:${column}`; }
function sheetIndex(snapshot: DocumentSnapshot, sheetId: string): number {
  assert(snapshot.document.kind === 'sheets', 'invalid', 'Expected sheets document');
  const index = snapshot.document.sheets.findIndex((sheet) => sheet.id === sheetId); assert(index >= 0, 'invalid', 'Sheet does not exist'); return index;
}
/** Blank removal is explicit; string "", numeric 0 and formula source remain distinct. */
export function setCells(snapshot: DocumentSnapshot, sheetId: string, edits: readonly { row: number; column: number; cell: Cell | null }[]): DocumentCommand {
  const index = sheetIndex(snapshot, sheetId); const keys = new Set<string>();
  const operations: PatchOperation[] = edits.map((edit) => {
    const key = cellKey(edit.row, edit.column); assert(!keys.has(key), 'invalid', 'Duplicate cell in edit batch'); keys.add(key);
    const path = ['sheets', index, 'cells', key];
    return edit.cell === null ? { op: 'delete', path } : { op: 'set', path, value: edit.cell };
  });
  return patchCommand(snapshot, 'SHT-02', operations);
}
export function renameSheet(snapshot: DocumentSnapshot, sheetId: string, name: string): DocumentCommand {
  const index = sheetIndex(snapshot, sheetId);
  assert(snapshot.document.kind === 'sheets', 'invalid', 'Expected sheets document');
  // Reference rewriting belongs to M3. Refuse a rename if it could invalidate a formula.
  assert(!snapshot.document.sheets.some((sheet) => Object.values(sheet.cells).some((cell) => cell.value.kind === 'formula')), 'unsupported', 'Formula-aware sheet renaming is not implemented');
  return patchCommand(snapshot, 'SHT-02', [{ op: 'set', path: ['sheets', index, 'name'], value: name }]);
}
