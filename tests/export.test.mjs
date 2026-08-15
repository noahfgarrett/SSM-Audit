import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'

import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { auditExportNestLevels, auditExportOrderRows, auditExportSheetName, buildAuditWorkbook } from '../src/audit/export.js'

/* The application loads SheetJS as a plain browser script into one shared realm.
   Tests evaluate the vendored copy the same way so the modules find the global
   they expect and the objects it returns share this realm's prototypes. */
const vendor = readFileSync(new URL('../src/vendor/sheetjs.js', import.meta.url), 'utf8')
vm.runInThisContext(vendor, { filename: 'sheetjs.js' })

const headers = EXTO_REV21_COLUMNS.map(column => column.header)
const index = Object.fromEntries(EXTO_REV21_COLUMNS.map(column => [column.field, column.index]))

/* Two milestone names that only differ past the 31st character, one carrying a
   forbidden sheet character, so sanitizing and de-duplication are both exercised. */
const ALPHA = 'L2-M1-101 30%/Enabling Capacity Phase Alpha'
const BETA = 'L2-M1-101 30%[Enabling] Capacity Phase Beta'
const SYSTEM = '1820 Mechanical System'

function row(values) {
  const cells = new Array(headers.length).fill('')
  for (const [field, value] of Object.entries(values)) cells[index[field]] = value
  return cells
}

function equipment({ tag, parent, description, milestone, milestoneParent = 'L1-M1-101 Enabling', dependencies = '' }) {
  return row({
    equipmentId: tag, closestParent: parent, closestParentStatus: parent === SYSTEM ? 'NEW' : '',
    equipmentDescription: description, dependencies, milestone, milestoneParent,
    building: 'B14', upn: '1820', discipline: 'MECHANICAL DRY', systemName: SYSTEM, itemMaster: 'VF1_EQUIPMENT',
  })
}

function auditResult() {
  const snapshot = auditSnapshotFromAoa([
    headers,
    equipment({ tag: 'B14-AHU-2201', parent: SYSTEM, description: 'Air Handler Unit', milestone: ALPHA }),
    equipment({ tag: 'B14-AHU-2201-VFD', parent: 'B14-AHU-2201', description: 'Variable Frequency Drive', milestone: ALPHA }),
    equipment({ tag: 'B14-AHU-2201-VFD-IO', parent: 'B14-AHU-2201-VFD', description: 'FMS Hardwired I/O', milestone: ALPHA }),
    equipment({ tag: 'B14-AHU-2202', parent: SYSTEM, description: 'Air Handler Unit', milestone: ALPHA }),
    equipment({ tag: 'B14-AHU-2202-B', parent: 'B14-AHU-2202', description: 'Temperature Transmitter', milestone: ALPHA }),
    equipment({ tag: 'B14-AHU-2202-A', parent: 'B14-AHU-2202', description: 'Pressure Transmitter', milestone: ALPHA }),
    equipment({ tag: 'B14-PMP-3101', parent: SYSTEM, description: 'Centrifugal Pump', milestone: BETA }),
    equipment({ tag: 'B14-PMP-3102', parent: '', description: 'Centrifugal Pump', milestone: BETA }),
    equipment({ tag: 'B14-FAN-4101', parent: SYSTEM, description: 'Exhaust Fan', milestone: '' }),
  ], { file: 'synthetic-registry.xlsx', sheet: 'Registry' })
  return runSsmAudit(snapshot)
}

function workbook() {
  return buildAuditWorkbook(auditResult(), 'synthetic-registry.xlsx', { generatedAt: new Date('2026-01-05T09:30:00Z') })
}

function grid(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true })
}

function milestoneSheetNames(book) {
  return book.SheetNames.filter(name => !['Dashboard', 'Index', 'All Findings', 'Rules'].includes(name))
}

function linkTarget(cell) {
  const target = cell && cell.l && cell.l.Target
  const match = target && target.match(/^#'(.+)'!A1$/)
  return match ? match[1].replace(/''/g, "'") : ''
}

function equipmentColumn(sheet) {
  return grid(sheet).slice(2).map(line => String(line[2] || '').trim())
}

function tabHolding(book, tag) {
  const name = milestoneSheetNames(book).find(sheet => equipmentColumn(book.Sheets[sheet]).includes(tag))
  assert.ok(name, `no milestone tab holds ${tag}`)
  return book.Sheets[name]
}

test('the audit workbook opens on a dashboard and closes on the rules reference', () => {
  const book = workbook()
  assert.equal(book.SheetNames[0], 'Dashboard')
  assert.equal(book.SheetNames[1], 'Index')
  assert.deepEqual(book.SheetNames.slice(-2), ['All Findings', 'Rules'])
  assert.equal(milestoneSheetNames(book).length, 3)
})

test('every L2 milestone gets a tab and unassigned equipment lands on No milestone', () => {
  const book = workbook()
  const tabs = milestoneSheetNames(book)
  assert.equal(tabs[tabs.length - 1], 'No milestone')
  assert.deepEqual([...new Set(equipmentColumn(book.Sheets['No milestone']))], ['B14-FAN-4101'])
})

test('milestone names that collide after sanitizing stay unique, legal, and inside 31 characters', () => {
  const tabs = milestoneSheetNames(workbook()).filter(name => name !== 'No milestone')
  assert.deepEqual(new Set(tabs), new Set(['L2-M1-101 30% Enabling Capacity', 'L2-M1-101 30% Enabling Capa (2)']))
  for (const name of tabs) {
    assert.ok(name.length <= 31, name)
    assert.doesNotMatch(name, /[[\]:*?/\\]/)
  }
  assert.equal(new Set(tabs).size, tabs.length)
})

test('sheet names de-duplicate against names already claimed by the workbook', () => {
  const used = new Set(['dashboard'])
  assert.equal(auditExportSheetName('Dashboard', used), 'Dashboard (2)')
  assert.equal(auditExportSheetName('Blank/Name:*?', used), 'Blank Name')
  assert.equal(auditExportSheetName('   ', used), 'Milestone')
})

test('index rows hyperlink to milestone tabs that exist in the workbook', () => {
  const book = workbook()
  const sheet = book.Sheets.Index
  const targets = []
  for (let rowIndex = 5; rowIndex < 5 + milestoneSheetNames(book).length; rowIndex++) {
    const target = linkTarget(sheet[`A${rowIndex}`])
    assert.ok(target, `Index row ${rowIndex} has no hyperlink`)
    assert.ok(book.SheetNames.includes(target), `${target} is not a sheet in the workbook`)
    targets.push(target)
  }
  assert.deepEqual(targets, milestoneSheetNames(book))
  for (const name of ['Dashboard', 'All Findings', 'Rules']) {
    assert.ok([linkTarget(sheet.A2), linkTarget(sheet.B2), linkTarget(sheet.C2)].includes(name), name)
  }
})

test('every sheet outside Index and Dashboard offers a way back to the index', () => {
  const book = workbook()
  for (const name of book.SheetNames) {
    if (name === 'Dashboard' || name === 'Index') continue
    assert.equal(book.Sheets[name].A1.v, '← Index', name)
    assert.equal(linkTarget(book.Sheets[name].A1), 'Index', name)
  }
})

test('nest depth counts hierarchy levels from the system root', () => {
  const book = workbook()
  const sheet = tabHolding(book, 'B14-AHU-2201')
  const depths = new Map(grid(sheet).slice(2).map(line => [String(line[2] || '').trim(), line[1]]))
  assert.equal(depths.get('B14-AHU-2201'), 0)
  assert.equal(depths.get('B14-AHU-2201-VFD'), 1)
  assert.equal(depths.get('B14-AHU-2201-VFD-IO'), 2)
})

test('nest depth survives parent cycles and rows whose parent is outside the registry', () => {
  const rows = [
    { equipmentId: 'LOOP-A', closestParent: 'LOOP-B' },
    { equipmentId: 'LOOP-B', closestParent: 'LOOP-A' },
    { equipmentId: 'ORPHAN', closestParent: 'OUTSIDE-PROJECT-TAG' },
  ]
  const levelFor = auditExportNestLevels(rows)
  assert.equal(levelFor(rows[2]), 0)
  for (const row of rows) assert.ok(Number.isInteger(levelFor(row)) && levelFor(row) >= 0)
})

test('milestone rows read as a tree with parents ahead of children and siblings in tag order', () => {
  const book = workbook()
  const tags = [...new Set(equipmentColumn(tabHolding(book, 'B14-AHU-2201')))]
  assert.deepEqual(tags, ['B14-AHU-2201', 'B14-AHU-2201-VFD', 'B14-AHU-2201-VFD-IO', 'B14-AHU-2202', 'B14-AHU-2202-A', 'B14-AHU-2202-B'])
})

test('the equipment tree indents each level so the hierarchy is visible in the cell', () => {
  const book = workbook()
  const lines = grid(tabHolding(book, 'B14-AHU-2201')).slice(2)
  const child = lines.find(line => String(line[2]).trim() === 'B14-AHU-2201-VFD')
  const grandchild = lines.find(line => String(line[2]).trim() === 'B14-AHU-2201-VFD-IO')
  assert.equal(child[2], '  B14-AHU-2201-VFD')
  assert.equal(grandchild[2], '    B14-AHU-2201-VFD-IO')
})

test('a row whose parent sits on another milestone still starts a branch of its own', () => {
  const parent = { equipmentId: 'B14-AHU-9001', closestParent: SYSTEM }
  const child = { equipmentId: 'B14-AHU-9001-VFD', closestParent: 'B14-AHU-9001' }
  assert.deepEqual(auditExportOrderRows([child]).map(entry => entry.equipmentId), ['B14-AHU-9001-VFD'])
  assert.deepEqual(auditExportOrderRows([child, parent]).map(entry => entry.equipmentId), ['B14-AHU-9001', 'B14-AHU-9001-VFD'])
})

test('progress formulas count the Y column of the milestone tab they belong to', () => {
  const book = workbook()
  const tabs = milestoneSheetNames(book)
  tabs.forEach((name, offset) => {
    const indexRow = 5 + offset, dashboardRow = 9 + offset
    assert.equal(book.Sheets.Index[`D${indexRow}`].f, `COUNTIF('${name}'!A:A,"Y")`)
    assert.equal(book.Sheets.Dashboard[`D${dashboardRow}`].f, `COUNTIF('${name}'!A:A,"Y")`)
    assert.equal(book.Sheets.Index[`E${indexRow}`].f, `IF(B${indexRow}=0,0,D${indexRow}/B${indexRow})`)
    assert.equal(book.Sheets.Index[`E${indexRow}`].z, '0%')
    assert.match(book.Sheets.Index[`F${indexRow}`].f, /^REPT\("█",MIN\(10,ROUND\(E\d+\*10,0\)\)\)&REPT\("░",10-MIN\(10,ROUND\(E\d+\*10,0\)\)\)$/)
  })
  assert.match(book.Sheets.Dashboard.G6.f, /^IF\(SUM\(B9:B11\)=0,0,SUM\(D9:D11\)\/SUM\(B9:B11\)\)$/)
})

test('milestone equipment and finding counts agree with the audit result', () => {
  const result = auditResult()
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx')
  const tabs = milestoneSheetNames(book)
  let equipmentTotal = 0, findingTotal = 0
  tabs.forEach((name, offset) => {
    const indexRow = 5 + offset, lines = grid(book.Sheets[name]).slice(2)
    const equipmentCount = book.Sheets.Index[`B${indexRow}`].v, findingCount = book.Sheets.Index[`C${indexRow}`].v
    assert.ok(lines.length >= equipmentCount, name)
    assert.equal(new Set(lines.map(line => String(line[2]).trim())).size, equipmentCount, name)
    equipmentTotal += equipmentCount
    findingTotal += findingCount
  })
  assert.equal(equipmentTotal, result.summary.rows)
  assert.equal(findingTotal, result.summary.findings)
})

test('rows with several findings repeat the equipment once per finding', () => {
  const result = auditResult()
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx')
  for (const name of milestoneSheetNames(book)) {
    const lines = grid(book.Sheets[name]).slice(2)
    const counted = new Map()
    for (const line of lines) {
      const tag = String(line[2]).trim()
      counted.set(tag, (counted.get(tag) || 0) + 1)
    }
    for (const [tag, lineCount] of counted) {
      const findings = result.findings.filter(finding => finding.equipmentId === tag).length
      assert.equal(lineCount, Math.max(1, findings), `${name} / ${tag}`)
    }
  }
})

test('severity cells carry a fill so the tab reads at a glance', () => {
  const result = auditResult()
  assert.ok(result.findings.length > 0, 'the fixture should raise findings')
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx')
  const findings = book.Sheets['All Findings']
  for (let offset = 0; offset < result.findings.length; offset++) {
    const cell = findings[`A${offset + 3}`]
    assert.ok(cell.v, 'severity label')
    assert.ok(cell.s && cell.s.fill && cell.s.fill.fgColor.rgb, `no severity fill on row ${offset + 3}`)
  }
  for (const name of milestoneSheetNames(book)) {
    const sheet = book.Sheets[name], lines = grid(sheet).slice(2)
    lines.forEach((line, offset) => {
      const cell = sheet[`I${offset + 3}`]
      if (line[8]) assert.ok(cell.s && cell.s.fill, `${name} row ${offset + 3} severity fill`)
    })
  }
})

test('nest level tints the equipment cell so depth is visible without reading the number', () => {
  const book = workbook()
  const sheet = tabHolding(book, 'B14-AHU-2201')
  const lines = grid(sheet).slice(2)
  const fills = ['B14-AHU-2201', 'B14-AHU-2201-VFD', 'B14-AHU-2201-VFD-IO']
    .map(tag => sheet[`C${lines.findIndex(line => String(line[2]).trim() === tag) + 3}`].s.fill.fgColor.rgb)
  assert.deepEqual(fills, [...new Set(fills)])
})

test('milestone tabs filter and freeze on the header row', () => {
  const book = workbook()
  for (const name of milestoneSheetNames(book)) {
    const sheet = book.Sheets[name]
    assert.match(sheet['!autofilter'].ref, /^A2:N\d+$/)
    assert.equal(sheet['!freeze'], 'A3')
    assert.equal(sheet['!cols'].length, 14)
  }
  assert.match(book.Sheets['All Findings']['!autofilter'].ref, /^A2:L\d+$/)
  assert.match(book.Sheets.Rules['!autofilter'].ref, /^A2:E\d+$/)
})

test('the rules tab explains every rule that produced a finding', () => {
  const result = auditResult()
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx')
  const lines = grid(book.Sheets.Rules).slice(2)
  assert.deepEqual(lines[0].slice(0, 5).map(value => typeof value), ['string', 'string', 'string', 'string', 'number'])
  const titles = new Set(lines.map(line => line[0]))
  for (const finding of result.findings) assert.ok(titles.has(finding.rule.title), finding.rule.title)
  const total = lines.reduce((sum, line) => sum + line[4], 0)
  assert.equal(total, result.summary.findings)
})

test('the workbook survives a SheetJS write and read with its links and formulas intact', () => {
  const book = workbook()
  const buffer = XLSX.write(book, { bookType: 'xlsx', type: 'buffer', cellStyles: true })
  const reopened = XLSX.read(buffer, { type: 'buffer', cellStyles: true })
  assert.deepEqual(reopened.SheetNames, book.SheetNames)
  assert.equal(linkTarget(reopened.Sheets.Index.A5), linkTarget(book.Sheets.Index.A5))
  assert.equal(reopened.Sheets.Index.D5.f, book.Sheets.Index.D5.f)
  assert.equal(reopened.Sheets.Dashboard.G6.f, book.Sheets.Dashboard.G6.f)
})

test('an empty audit result still produces a readable workbook', () => {
  const book = buildAuditWorkbook({ rows: [], findings: [], summary: { rows: 0, findings: 0, severity: {} }, standard: 'Registry Integrity' }, '')
  assert.deepEqual(book.SheetNames, ['Dashboard', 'Index', 'All Findings', 'Rules'])
  assert.equal(book.Sheets.Dashboard.G6.v, 0)
  assert.ok(!book.Sheets.Dashboard.G6.f)
})
