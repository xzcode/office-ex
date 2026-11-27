import {writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {fixture} from './sheets-format-fixture.ts';
import {writeSheetsXlsx} from '@office-ex/ooxml';
import {writeLegacyXls} from '@office-ex/format-service';
const out=process.argv[2];if(!out)throw Error('output directory required');
writeFileSync(join(out,'workbook.xlsx'),writeSheetsXlsx(fixture()));
writeFileSync(join(out,'workbook.xls'),writeLegacyXls(fixture()));
