import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { auditExpectedDependencies } from '../src/audit/expected-dependencies.js'
import { buildDependencyCheckWorkbook } from '../src/audit/dependency-export.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js', import.meta.url), 'utf8'))

const MECH = { building: 'BLDG-A', discipline: 'MECHANICAL WET', systemName: '111 Chilled Water', upn: '111' }
const ELEC = { building: 'BLDG-A', discipline: 'ELECTRICAL', systemName: '603 Low Voltage', upn: '603' }
const entries = [['SYN-PANEL', ''], ['SYN-MCC', 'SYN-PANEL'], ['SYN-PUMP-1', 'SYN-MCC'], ['SYN-PUMP-2', 'SYN-PANEL']]
const check = auditExpectedDependencies({ rows: [
  { equipmentId: 'SYN-PUMP-1', ...MECH },
  { equipmentId: 'SYN-PUMP-2', ...MECH, dependencies: 'SYN-PANEL' },
  { equipmentId: 'SYN-PANEL', ...ELEC },
] }, { easyPower: { kind: 'easyPower', sheetName: 'Synthetic', entries: entries.map(([equipmentId, parent], index) => ({ equipmentId, parent, sourceRow: index + 2, sheetName: 'Synthetic' })) } })
const cells = (sheet, row) => { const range = XLSX.utils.decode_range(sheet['!ref']); const out = []; for (let c = 0; c <= range.e.c; c++) out.push(sheet[XLSX.utils.encode_cell({ r: row, c })]?.v ?? ''); return out }

test('the workbook has Summary, Dependencies and Not in registry sheets', () => {
  const workbook = buildDependencyCheckWorkbook(check, 'Synthetic registry.xlsx')
  assert.deepEqual(workbook.SheetNames, ['Summary', 'Dependencies', 'Not in registry'])
})

test('the summary carries the counts', () => {
  const sheet = buildDependencyCheckWorkbook(check, 'Synthetic registry.xlsx').Sheets.Summary
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 })
  const value = label => rows.find(row => row[0] === label)?.[1]
  assert.equal(value('Compared'), 2)
  assert.equal(value('Match'), 1)
  assert.equal(value('Missing'), 1)
  assert.equal(value('Not in registry'), 1)
  assert.ok(rows.some(row => row[0] === 'MECHANICAL WET' && row[1] === 2 && row[2] === 1 && row[3] === 1))
})

test('the dependency table lists one row per equipment with a frozen, filtered header', () => {
  const sheet = buildDependencyCheckWorkbook(check, 'Synthetic registry.xlsx').Sheets.Dependencies
  assert.deepEqual(cells(sheet, 0).slice(0, 7), ['Equipment ID', 'Status', 'Expected dependencies', 'Registry dependencies', 'Missing', 'Note', 'Source'])
  assert.deepEqual(cells(sheet, 1).slice(0, 5), ['SYN-PUMP-1', 'Missing', 'SYN-PANEL', '', 'SYN-PANEL'])
  assert.match(String(cells(sheet, 1)[5]), /SYN-MCC/)
  assert.match(String(cells(sheet, 1)[6]), /EasyPower Synthetic, row 4/)
  assert.deepEqual(cells(sheet, 2).slice(0, 2), ['SYN-PUMP-2', 'Match'])
  assert.equal(sheet['!freeze'], 'A2')
  assert.equal(sheet['!autofilter'].ref, sheet['!ref'])
  assert.equal(sheet.B2.s.fill.fgColor.rgb, 'FBE3E1')
  assert.equal(sheet.E2.s.fill.fgColor.rgb, 'FBE3E1')
  assert.notEqual(sheet.B3.s.fill.fgColor.rgb, 'FBE3E1')
})

test('tags missing from the registry list their document and the rows they affect', () => {
  const sheet = buildDependencyCheckWorkbook(check, 'Synthetic registry.xlsx').Sheets['Not in registry']
  assert.deepEqual(cells(sheet, 0), ['Source tag', 'Document', 'Sheet', 'Row', 'Also found at', 'Registry rows affected'])
  assert.deepEqual(cells(sheet, 1), ['SYN-MCC', 'EasyPower', 'Synthetic', 3, 'Synthetic, row 4', 'SYN-PUMP-1'])
  assert.equal(sheet['!freeze'], 'A2')
})

test('an empty check still exports readable sheets', () => {
  const workbook = buildDependencyCheckWorkbook(auditExpectedDependencies({ rows: [] }, {}), 'Empty.xlsx')
  const bytes = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' })
  assert.deepEqual(XLSX.read(bytes, { type: 'array' }).SheetNames, ['Summary', 'Dependencies', 'Not in registry'])
})

test('the written file really freezes the header row', async () => {
  const { workbookBytes, workbookBytesCompact } = await import('../src/core/download.js')
  const workbook = buildDependencyCheckWorkbook(check, 'Synthetic registry.xlsx')
  for (const bytes of [workbookBytes(workbook), await workbookBytesCompact(workbook)]) {
    const container = XLSX.CFB.read(new Uint8Array(bytes), { type: 'array' })
    const xml = new TextDecoder().decode(XLSX.CFB.find(container, '/xl/worksheets/sheet2.xml').content)
    assert.match(xml, /<sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/>/)
    const summary = new TextDecoder().decode(XLSX.CFB.find(container, '/xl/worksheets/sheet1.xml').content)
    assert.ok(!summary.includes('<pane'), 'sheets without a freeze stay unfrozen')
    assert.equal(XLSX.read(bytes, { type: 'array' }).SheetNames.length, 3)
  }
})
