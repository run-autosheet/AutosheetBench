// SheetJS 0.20.3 turns "\r\n" into "\n" in every string cell it reads, so a
// workbook written and read back no longer has its original line endings.
// vendor/xlsx-0.20.3-crlf.tgz drops that step for shared and inline strings
// (vendor/xlsx-0.20.3-crlf.patch), as SheetJS 0.18.5 did. Formulas keep the
// upstream normalization, which 0.18.5 also had.
import test from 'node:test'
import assert from 'node:assert/strict'
import {createRequire} from 'node:module'
import * as XLSX from 'xlsx'

const VALUES = ['a\r\nb', 'line one\r\nline two\r\n', 'bare\rreturn', 'plain\nnewline']

function roundTrip(lib, options) {
  const workbook = lib.utils.book_new()
  lib.utils.book_append_sheet(workbook, lib.utils.aoa_to_sheet(VALUES.map(value => [value])), 'Source')
  const read = lib.read(lib.write(workbook, {type: 'buffer', bookType: 'xlsx', ...options}), {type: 'buffer'})
  return VALUES.map((_, row) => read.Sheets.Source[`A${row + 1}`].v)
}

for (const [name, lib] of [['ESM', XLSX], ['CommonJS', createRequire(import.meta.url)('xlsx')]]) {
  test(`${name} SheetJS keeps carriage returns in shared strings`, () => {
    assert.deepEqual(roundTrip(lib, {bookSST: true}), VALUES)
  })
  test(`${name} SheetJS keeps carriage returns in inline strings`, () => {
    assert.deepEqual(roundTrip(lib, {bookSST: false}), VALUES)
  })
}
