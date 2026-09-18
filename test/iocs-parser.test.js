import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import * as XLSX from 'xlsx';

import { parseIocsExcel } from '../src/index.js';

const workbookPath = process.env.IOCS_TEST_WORKBOOK;

test('Utah date and similarly named columns map to the correct values', () => {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ['Tech', 'DCT', 'Day', 'Finance No.', 'Office Zip', 'Office', 'Pay Location', 'Test ID', 'Employee', 'Roster Des', 'Activity', 'Emp Start Time', 'Emp End Time', 'Read Code', 'Read Time', 'Pay Period', 'Pay Week'],
    ['CONNIE', 46284, 'Saturday', 497800, 84107, 'SAL-MURRAY BRUT', 0, '004-56-2111', 'D T TANIELU', 11, 0, 0, 0.3541666667, 2, 0.1090277778, 21, 1]
  ];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet.B2.z = 'dddd';
  XLSX.utils.book_append_sheet(workbook, sheet, 'Saturday');

  const entries = parseIocsExcel(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }));
  assert.equal(entries.length, 1);
  assert.equal(entries[0].date, '2026-09-19');
  assert.equal(entries[0].location, 'SAL-MURRAY BRUT');
  assert.equal(entries[0].bt, '00:00');
  assert.equal(entries[0].rt, '02:37');
});

test('exact Utah weekly master parses every assigned reading', { skip: !workbookPath }, () => {
  const entries = parseIocsExcel(fs.readFileSync(workbookPath));

  assert.equal(entries.length, 69);
  assert.deepEqual([...new Set(entries.map(entry => entry.state))], ['Utah']);

  const first = entries.find(entry => entry.date === '2026-09-19' && entry.dct === 'CONNIE');
  assert.ok(first);
  assert.equal(first.location, 'SAL-MURRAY BRUT');
  assert.equal(first.bt, '00:00');
  assert.equal(first.rt, '02:37');
  assert.notEqual(first.bt, first.rt);
});
