import { assert, replaceGraphemeRange } from '@office-ex/contracts';
import { patchCommand, type DocumentCommand, type DocumentSnapshot, type ModelPath, type WriterBlock, type DeepReadonly, type WriterDocument } from '@office-ex/core';
export type { WriterDocument, WriterSection, WriterBlock, Paragraph, TextRun, Inline, Table, TableCell } from '@office-ex/core';

function document(snapshot: DocumentSnapshot): DeepReadonly<WriterDocument> { assert(snapshot.document.kind === 'writer', 'invalid', 'Expected writer document'); return snapshot.document; }
function findTextRun(blocks: readonly DeepReadonly<WriterBlock>[], runId: string, path: ModelPath): { text: string; path: ModelPath } | undefined {
  for (let index = 0; index < blocks.length; index++) {
    const block = blocks[index]!; const blockPath = [...path, index];
    if (block.type === 'paragraph') {
      const runIndex = block.runs.findIndex((run) => run.id === runId && run.type === 'text');
      if (runIndex !== -1) { const run = block.runs[runIndex]!; if (run.type === 'text') return { text: run.text, path: [...blockPath, 'runs', runIndex, 'text'] }; }
    } else if (block.type === 'table') {
      for (let row = 0; row < block.rows.length; row++) for (let cell = 0; cell < block.rows[row]!.cells.length; cell++) {
        const result = findTextRun(block.rows[row]!.cells[cell]!.blocks, runId, [...blockPath, 'rows', row, 'cells', cell, 'blocks']);
        if (result) return result;
      }
    }
  }
  return undefined;
}
/** One composition commit / range replacement is one undoable transaction. */
export function replaceText(snapshot: DocumentSnapshot, runId: string, start: number, end: number, replacement: string): DocumentCommand {
  const model = document(snapshot);
  for (let index = 0; index < model.sections.length; index++) {
    const run = findTextRun(model.sections[index]!.blocks, runId, ['sections', index, 'blocks']);
    if (run) return patchCommand(snapshot, 'WRT-03', [{ op: 'set', path: run.path, value: replaceGraphemeRange(run.text, start, end, replacement) }]);
  }
  assert(false, 'invalid', 'Text run does not exist');
}
export function insertBlocks(snapshot: DocumentSnapshot, sectionId: string, index: number, blocks: readonly WriterBlock[]): DocumentCommand {
  const model = document(snapshot); const section = model.sections.findIndex((candidate) => candidate.id === sectionId);
  assert(section >= 0, 'invalid', 'Section does not exist');
  return patchCommand(snapshot, 'WRT-03', [{ op: 'splice', path: ['sections', section, 'blocks'], index, deleteCount: 0, items: blocks }]);
}

export * from './paragraphs.js';
