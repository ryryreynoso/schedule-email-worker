import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import * as XLSX from 'xlsx';

import { classifyWorkbook, parseIocsExcel } from '../src/index.js';

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

test('quarterly schedule filenames win over IOCS-like TEST DAY text', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['TEST DAY'],
    ['TEST SCHEDULE'],
    ['ZIP CODES', 'ZIP', 'SITE', 'TYPE', 'TEST ID', 'TECH(S)'],
    ['MONDAY 9-21-26'],
    ['FLAT STREAM', '89101', 'LAS VEGAS', 'ODIS', '123456', 'RR']
  ]);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Qtr 4');

  assert.deepEqual(classifyWorkbook(workbook, '2026 Qtr 4 Schedule NV addd.xlsx'), {
    kind: 'schedule',
    state: 'Nevada'
  });
  assert.deepEqual(classifyWorkbook(workbook, 'addd UT.xlsx'), {
    kind: 'schedule',
    state: 'Utah'
  });
});

test('IOCS filename remains authoritative', () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['TEST SCHEDULE'],
    ['MONDAY 9-21-26']
  ]), 'Monday');

  assert.deepEqual(classifyWorkbook(workbook, 'Copy of DAILY IOCS NV.xlsx'), { kind: 'iocs' });
});

test('IOCS parser keeps the active week when its header is deep or omitted', () => {
  const workbook = XLSX.utils.book_new();
  const header = ['TEST DATE', 'FINANCE', 'OFFICE', 'EIN', 'EMPLOYEE', 'ROSTER DES', 'BT', 'ET', 'RT', 'ASSIGNED DCT'];
  const deepRows = Array.from({ length: 40 }, () => ['']);
  deepRows.push(header);
  deepRows.push(['09/23/2026', '314881', 'LAS VEGAS NV NV', '04771924', 'D M LABAYAN-ARVEL', 'MAILHANDLER', '23:00', '7:30', '1:36', 'REYNOSO RYAN']);
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(deepRows), '09-19-26');

  // A sibling weekly tab can contain rows without repeating the visible
  // header. It should inherit the proven column layout from the workbook.
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['09/26/2026', '314895', 'LAS-RED ROCK NV', '04436251', 'J E MCKENZIE', 'CLERK', '3:00', '12:00', '5:06', 'AGRESOR GRACE']
  ]), '09-26-26');

  const entries = parseIocsExcel(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }));
  assert.equal(entries.length, 2);
  assert.deepEqual(entries.map(entry => entry.date), ['2026-09-23', '2026-09-26']);
  assert.deepEqual(entries.map(entry => entry.dct), ['REYNOSO RYAN', 'AGRESOR GRACE']);
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
