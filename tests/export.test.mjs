import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'

import { EXTO_REV21_COLUMNS, extoRev21SystemsForUpn } from '../src/exto/rev21-contract.js'
import { auditMergeSnapshots, auditSnapshotFromAoa, auditSnapshotFromWorkbook } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { auditMakeCorrection, auditRecommendationContext, auditProposeCorrection, auditApplyCorrections } from '../src/audit/actions.js'
import { AUDIT_EXPORT_TICK, applyChangesToWorkbook, validateAuditCorrections, auditExportNestLevels, auditExportOrderRows, auditExportSheetName, buildAuditWorkbook, buildAuditCorrectionsWorkbook, buildAuditTrackerWorkbook, buildUpdatedRegistryBytes, exportUpdatedRegistryXlsx, exportAuditCorrectionsXlsx } from '../src/audit/export.js'
import { S, resetSession } from '../src/state.js'

// Package tests use the browser's XML DOM, or @xmldom/xmldom supplied by this
// optional test-only module path. The offline application has no new dependency.
if (process.env.SSM_EXPORT_XML_DOM) {
  const { DOMParser, XMLSerializer } = await import(process.env.SSM_EXPORT_XML_DOM)
  Object.assign(globalThis, { DOMParser, XMLSerializer })
}
const packageTest = (name, fn) => test(name, { skip: typeof DOMParser !== 'function' && 'Set SSM_EXPORT_XML_DOM to an XML DOM module to run package-preservation tests' }, fn)

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
  return book.SheetNames.filter(name => !['Dashboard', 'Index', 'All Findings', 'Rules', 'Calc'].includes(name))
}

function linkTarget(cell) {
  const target = cell && cell.l && cell.l.Target
  const match = target && target.match(/^#'(.+)'!A1$/)
  return match ? match[1].replace(/''/g, "'") : ''
}

/* Equipment cells are merged blocks: only the block's first line carries the
   tag. carriedTags fills the blanks downward so lookups keep working. */
function carriedTags(sheet) {
  let current = ''
  return grid(sheet).slice(2).map(line => {
    const tag = String(line[2] || '').trim()
    if (tag) current = tag
    return current
  })
}
function equipmentColumn(sheet) {
  return grid(sheet).slice(2).map(line => String(line[2] || '').trim()).filter(Boolean)
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
  assert.deepEqual(book.SheetNames.slice(-3), ['All Findings', 'Rules', 'Calc'])
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
    if (name === 'Dashboard' || name === 'Index' || name === 'Calc') continue
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

function dashboardMilestoneRows(book) {
  const sheet = book.Sheets.Dashboard, rows = grid(sheet)
  const header = rows.findIndex(line => line[0] === 'Milestone')
  assert.ok(header > 0, 'dashboard has a milestone table')
  return { first: header + 2, header: header + 1 }
}

test('progress formulas count the ticks in the Actioned column of the milestone tab they belong to', () => {
  const book = workbook()
  const tabs = milestoneSheetNames(book)
  const { first } = dashboardMilestoneRows(book)
  tabs.forEach((name, offset) => {
    const indexRow = 5 + offset, dashboardRow = first + offset
    assert.equal(book.Sheets.Index[`D${indexRow}`].f, `COUNTIF('${name}'!A:A,"${AUDIT_EXPORT_TICK}")`)
    assert.equal(book.Sheets.Dashboard[`D${dashboardRow}`].f, `COUNTIF('${name}'!A:A,"${AUDIT_EXPORT_TICK}")`)
    assert.equal(book.Sheets.Index[`E${indexRow}`].f, `IF(B${indexRow}=0,0,D${indexRow}/B${indexRow})`)
    assert.equal(book.Sheets.Index[`E${indexRow}`].z, '0%')
    assert.equal(book.Sheets.Index[`F${indexRow}`].f, `E${indexRow}`, 'the Index bar cell carries the percent (a data bar fills it)')
    assert.equal(book.Sheets.Index[`F${indexRow}`].z, ';;;', 'bar cells hide their number — the % column carries it')
    assert.equal(book.Sheets.Dashboard[`F${dashboardRow}`].f, `E${dashboardRow}`)
  })
  const last = first + tabs.length - 1
  assert.equal(book.Sheets.Dashboard.F6.f, `IF(SUM(B${first}:B${last})=0,0,SUM(D${first}:D${last})/SUM(B${first}:B${last}))`)
  assert.equal(book.Sheets.Dashboard.A6.f, 'F6', 'the overall bar cell carries the overall percent')
  assert.equal(book.Sheets.Dashboard.A6.z, ';;;', 'the overall bar shows the bar only')
  /* native data bars are attached as conditional formatting on Dashboard (overall + both tables) and Index */
  const dashboardBars = book.Sheets.Dashboard['!xmlExtras'].conditionalFormatting.join('')
  assert.match(dashboardBars, /sqref="A6:A6"><cfRule type="dataBar" priority="1"><dataBar minLength="0" maxLength="100" showValue="0">/)
  assert.match(dashboardBars, new RegExp(`sqref="F${first}:F${last}"><cfRule type="dataBar"`))
  assert.match(book.Sheets.Index['!xmlExtras'].conditionalFormatting.join(''), /sqref="F5:F\d+"><cfRule type="dataBar"/)
})

test('the dashboard has a per-discipline progress table that sums ticks across milestone tabs', () => {
  const book = workbook(), result = auditResult()
  const rows = grid(book.Sheets.Dashboard)
  const header = rows.findIndex(line => line[0] === 'Discipline')
  assert.ok(header > 0)
  const disciplines = new Set(result.rows.map(row => row.discipline || 'No discipline'))
  const tableRows = rows.slice(header + 1).filter(line => line[0] && line[0] !== 'Milestone' && disciplines.has(line[0]))
  assert.equal(tableRows.length, disciplines.size)
  const firstRow = header + 2
  const cell = book.Sheets.Dashboard[`D${firstRow}`]
  const tabs = milestoneSheetNames(book)
  assert.equal(cell.f, `SUM('Calc'!B2:B${1 + tabs.length})`, 'the Dashboard sums the hidden Calc column for that discipline')
  const calc = book.Sheets.Calc
  tabs.forEach((name, offset) => {
    assert.equal(calc[`A${offset + 2}`].v, name)
    assert.equal(calc[`B${offset + 2}`].f, `COUNTIFS('${name}'!G:G,"${rows[header + 1][0]}",'${name}'!A:A,"${AUDIT_EXPORT_TICK}")`)
  })
  assert.equal(book.Workbook.Sheets.find(sheet => sheet.name === 'Calc').Hidden, 1, 'Calc is hidden')
  assert.ok(book.Workbook.Sheets.filter(sheet => sheet.name !== 'Calc').every(sheet => !sheet.Hidden))
  const equipmentTotal = tableRows.reduce((sum, line) => sum + line[1], 0)
  assert.equal(equipmentTotal, result.summary.rows)
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
    assert.equal(lines.filter(line => String(line[2] || '').trim()).length, equipmentCount, name)
    equipmentTotal += equipmentCount
    findingTotal += findingCount
  })
  assert.equal(equipmentTotal, result.summary.rows)
  assert.equal(findingTotal, result.summary.findings)
})

test('an equipment with several findings is one merged block: tag once, one Actioned box, one line per finding', () => {
  const result = auditResult()
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx')
  let checkedMerge = false
  for (const name of milestoneSheetNames(book)) {
    const sheet = book.Sheets[name], lines = grid(sheet).slice(2), tags = carriedTags(sheet)
    const counted = new Map()
    tags.forEach(tag => counted.set(tag, (counted.get(tag) || 0) + 1))
    for (const [tag, lineCount] of counted) {
      const findings = result.findings.filter(finding => finding.equipmentId === tag).length
      assert.equal(lineCount, Math.max(1, findings), `${name} / ${tag}`)
      if (lineCount > 1) {
        checkedMerge = true
        const first = tags.indexOf(tag)
        /* the tag appears once; continuation lines are blank in every equipment column */
        for (let at = first + 1; at < first + lineCount; at++)
          for (let column = 0; column <= 10; column++)
            assert.equal(String(lines[at][column] ?? ''), '', `${name} ${tag} line ${at} column ${column}`)
        /* and a vertical merge covers the block for the tag column */
        assert.ok((sheet['!merges'] || []).some(range => range.s.c === 2 && range.s.r === first + 2 && range.e.r === first + 1 + lineCount), `${name} ${tag} merge`)
        /* one Actioned box: the dropdown validation range still covers the block, but only the anchor holds a value */
        assert.equal(lines[first][0], '☐')
      }
    }
  }
  assert.ok(checkedMerge, 'the fixture has at least one multi-finding equipment')
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
      const cell = sheet[`L${offset + 3}`]
      if (line[11]) assert.ok(cell.s && cell.s.fill, `${name} row ${offset + 3} severity fill`)
    })
  }
})

test('nest level tints the equipment cell with a distinct hue per level and keeps the ID black', () => {
  const book = workbook()
  const sheet = tabHolding(book, 'B14-AHU-2201')
  const lines = grid(sheet).slice(2)
  const cells = ['B14-AHU-2201', 'B14-AHU-2201-VFD', 'B14-AHU-2201-VFD-IO']
    .map(tag => sheet[`C${lines.findIndex(line => String(line[2] || '').trim() === tag) + 3}`])
  const fills = cells.map(cell => cell.s.fill.fgColor.rgb)
  assert.deepEqual(fills, [...new Set(fills)])
  /* Different hues, not shades of one: the dominant RGB channel differs between neighbours. */
  const dominant = rgb => ['R', 'G', 'B'][[0, 2, 4].map(at => parseInt(rgb.slice(at, at + 2), 16)).reduce((best, value, index, all) => value > all[best] ? index : best, 0)]
  assert.notEqual(dominant(fills[0]), dominant(fills[1]))
  for (const cell of cells) assert.equal(cell.s.font.color.rgb, '000000')
  for (const line of lines) {
    if (!String(line[2] || '').trim()) continue
    const rowIndex = lines.indexOf(line) + 3
    assert.equal(sheet[`C${rowIndex}`].s.font.color.rgb, '000000', `row ${rowIndex} equipment ID is black`)
  }
})

test('the cell a finding is about is shaded red on that finding\'s line', () => {
  const result = auditResult(), book = buildAuditWorkbook(result, 'synthetic-registry.xlsx')
  const columns = { 'Closest Parent': 'E', 'Dependencies': 'F', 'UPN': 'H', 'System Name': 'I', 'Equipment ID': 'C' }
  let checked = 0
  for (const finding of result.findings) {
    const column = columns[finding.field]; if (!column) continue
    const sheet = tabHolding(book, finding.equipmentId), tags = carriedTags(sheet), lines = grid(sheet).slice(2)
    const lineIndex = tags.findIndex((tag, at) => tag === finding.equipmentId && lines[at][12] === finding.rule.title)
    assert.ok(lineIndex >= 0, `${finding.equipmentId} / ${finding.rule.title}`)
    const anchorRow = tags.indexOf(finding.equipmentId) + 3
    assert.equal(sheet[`${column}${anchorRow}`].s.fill.fgColor.rgb, 'FBE3E1', `${finding.equipmentId} ${finding.field}`)
    checked++
  }
  assert.ok(checked > 0, 'the fixture raises findings on flaggable fields')
})

test('the Actioned column starts unticked on each equipment line and the written file carries the tick dropdown and green-row format', async () => {
  const { workbookBytes } = await import('../src/core/download.js')
  const book = workbook()
  for (const name of milestoneSheetNames(book)) {
    const lines = grid(book.Sheets[name]).slice(2)
    for (const line of lines) {
      const anchor = Boolean(String(line[2] || '').trim())
      assert.equal(line[0], anchor ? '☐' : '', name)
    }
  }
  const bytes = workbookBytes(book, { compression: true })
  const container = XLSX.CFB.read(new Uint8Array(bytes), { type: 'array' })
  const decoder = new TextDecoder()
  const styles = decoder.decode(XLSX.CFB.find(container, '/xl/styles.xml').content)
  assert.match(styles, /<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFE3F5E8"\/><\/patternFill><\/fill><\/dxf><\/dxfs>/)
  book.SheetNames.forEach((name, index) => {
    const xml = decoder.decode(XLSX.CFB.find(container, `/xl/worksheets/sheet${index + 1}.xml`).content)
    const isMilestone = milestoneSheetNames(book).includes(name)
    assert.equal(xml.includes('<dataValidation type="list"'), isMilestone, name)
    assert.equal(xml.includes('dxfId="0"'), isMilestone, `${name}: the green-row format belongs to milestone tabs only`)
    if (name === 'Dashboard' || name === 'Index') assert.ok(xml.includes('type="dataBar"'), `${name} carries data bars`)
    if (isMilestone) {
      assert.ok(xml.includes(`<formula1>"${AUDIT_EXPORT_TICK},☐"</formula1>`), name)
      assert.match(xml, /<formula>\$R3="☑"<\/formula>/, 'the green format reads the hidden Done helper column')
      const validationAt = xml.indexOf('<dataValidations'), hyperlinksAt = xml.indexOf('<hyperlinks'), sheetDataEnd = xml.indexOf('</sheetData>')
      assert.ok(validationAt > sheetDataEnd, 'extras follow sheetData')
      if (hyperlinksAt !== -1) assert.ok(validationAt < hyperlinksAt, 'extras precede hyperlinks (schema order)')
      assert.ok(xml.indexOf('<conditionalFormatting') < validationAt, 'conditional formatting precedes data validation')
    }
  })
  const reopened = XLSX.read(bytes, { type: 'array' })
  assert.deepEqual(reopened.SheetNames, book.SheetNames)
})

test('milestone tabs filter and freeze on the header row', () => {
  const book = workbook()
  for (const name of milestoneSheetNames(book)) {
    const sheet = book.Sheets[name]
    assert.match(sheet['!autofilter'].ref, /^A2:Q\d+$/)
    assert.equal(sheet['!freeze'], 'A3')
    assert.equal(sheet['!cols'].length, 18)
    assert.equal(sheet['!cols'][17].hidden, true, 'the Done helper column is hidden')
    assert.deepEqual(grid(sheet)[1].slice(2, 6), ['Equipment ID', 'Description', 'Closest Parent', 'Dependencies'], 'Dependencies follows Closest Parent')
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
  assert.equal(reopened.Sheets.Dashboard.F6.f, book.Sheets.Dashboard.F6.f)
})

test('an empty audit result still produces a readable workbook', () => {
  const book = buildAuditWorkbook({ rows: [], findings: [], summary: { rows: 0, findings: 0, severity: {} }, standard: 'Registry Integrity' }, '')
  assert.deepEqual(book.SheetNames, ['Dashboard', 'Index', 'All Findings', 'Rules', 'Calc'])
  assert.equal(book.Sheets.Dashboard.F6.v, 0)
  assert.ok(!book.Sheets.Dashboard.F6.f)
})

test('no formula in the workbook approaches Excel\'s 8,192-character limit, even with many milestone tabs', () => {
  /* 120 milestones: one row each is enough to produce 120 tabs. */
  const rows = []
  for (let i = 0; i < 120; i++) rows.push(row({ equipmentId: `B1-EQ-${String(i).padStart(4, '0')}`, closestParent: SYSTEM, closestParentStatus: 'NEW', upn: '1820', discipline: i % 2 ? 'ELECTRICAL' : 'MECHANICAL WET', systemName: SYSTEM, milestone: `L2-M1-182 Phase ${i}`, milestoneParent: 'L1-M1 30% Capacity' }))
  const snapshot = auditSnapshotFromAoa([headers, ...rows], { file: 'many.xlsx', sheet: 'Registry' })
  const book = buildAuditWorkbook(runSsmAudit(snapshot), 'many.xlsx')
  assert.equal(milestoneSheetNames(book).length, 120)
  let longest = 0
  for (const name of book.SheetNames) for (const [key, cell] of Object.entries(book.Sheets[name])) if (key[0] !== '!' && cell.f && cell.f.length > longest) longest = cell.f.length
  assert.ok(longest < 1000, `longest formula is ${longest} characters`)
})

test('the compact writer produces a smaller, valid xlsx with the extras and all sheets intact', async () => {
  const { workbookBytes, workbookBytesCompact } = await import('../src/core/download.js')
  const book = workbook()
  const plain = workbookBytes(book, { compression: true })
  const compact = await workbookBytesCompact(book)
  assert.ok(compact.length < plain.length, `compact ${compact.length} should be smaller than ${plain.length}`)
  const reopened = XLSX.read(compact, { type: 'array', cellStyles: true })
  assert.deepEqual(reopened.SheetNames, book.SheetNames)
  assert.equal(reopened.Sheets.Index.D5.f, book.Sheets.Index.D5.f)
  assert.equal(reopened.Workbook.Sheets.find(sheet => sheet.name === 'Calc').Hidden, 1)
  const container = XLSX.CFB.read(new Uint8Array(compact), { type: 'array' }), decoder = new TextDecoder()
  const firstMilestone = milestoneSheetNames(book)[0], index = book.SheetNames.indexOf(firstMilestone) + 1
  assert.ok(decoder.decode(XLSX.CFB.find(container, `/xl/worksheets/sheet${index}.xml`).content).includes('<dataValidations'), 'tick dropdown survives the compact zip')
  assert.match(decoder.decode(XLSX.CFB.find(container, '/xl/styles.xml').content), /<dxfs count="1">/)
  /* [Content_Types].xml leads the archive, as Excel expects. */
  assert.equal(new TextDecoder().decode(compact.slice(30, 30 + 19)), '[Content_Types].xml')
})

test('the export plan can pre-tick or leave out whole levels and single checks', async () => {
  const { auditExportApplyPlan, auditExportPlanMode } = await import('../src/audit/export.js')
  const result = auditResult()
  const severities = new Set(result.findings.map(finding => finding.severity))
  assert.ok(severities.size >= 2, 'fixture spans several levels')
  /* Leave out every info finding; pre-tick every warning. */
  const plan = { levels: { info: 'skip', warning: 'pretick' }, rules: {} }
  const applied = auditExportApplyPlan(result, plan)
  assert.equal(applied.result.findings.filter(finding => finding.severity === 'info').length, 0)
  assert.equal(applied.skipped, result.findings.filter(finding => finding.severity === 'info').length)
  for (const finding of applied.result.findings.filter(entry => entry.severity === 'warning')) assert.ok(applied.preticked.has(finding))
  assert.equal(applied.result.summary.findings, applied.result.findings.length)
  assert.equal(applied.result.summary.severity.info, 0)
  /* A rule override beats its level. */
  const someInfo = result.findings.find(finding => finding.severity === 'info')
  const override = auditExportApplyPlan(result, { levels: { info: 'skip' }, rules: { [someInfo.rule.id]: 'include' } })
  assert.ok(override.result.findings.some(finding => finding.rule.id === someInfo.rule.id))
  assert.equal(auditExportPlanMode({ levels: {}, rules: {} }, someInfo), 'include', 'default is include')
})

test('a skipped level vanishes from the workbook and pre-ticked equipment starts checked', async () => {
  const result = auditResult()
  const plan = { levels: { info: 'skip', warning: 'pretick', error: 'pretick', blocker: 'pretick' }, rules: {} }
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx', { plan })
  const findingsRows = grid(book.Sheets['All Findings']).slice(2)
  assert.ok(findingsRows.every(line => line[0] !== 'NOTE'), 'no NOTE lines exported')
  const expected = result.findings.filter(finding => finding.severity !== 'info').length
  assert.equal(findingsRows.length, expected)
  /* Every remaining finding is pre-ticked, so every equipment line with findings starts ☑. */
  for (const name of milestoneSheetNames(book)) {
    const sheet = book.Sheets[name], lines = grid(sheet).slice(2), tags = carriedTags(sheet)
    lines.forEach((line, at) => {
      if (line[0] === '') return
      const tag = tags[at]
      const hasFinding = result.findings.some(finding => finding.severity !== 'info' && finding.equipmentId === tag)
      assert.equal(line[0], hasFinding ? '☑' : '☐', `${name}: ${tag}`)
    })
  }
  /* Dashboard severity strip reflects the filtered counts. */
  assert.equal(book.Sheets.Dashboard.F9.v, 0, 'Notes KPI is zero when info is left out')
})

test('the level layout builds one tab per finding level with flagged equipment only and an L2 Milestone column', () => {
  const result = auditResult()
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx', { layout: 'level' })
  const tabs = milestoneSheetNames(book)
  const present = [...new Set(result.findings.map(finding => finding.severity))]
  const expected = ['blocker', 'error', 'warning', 'info'].filter(severity => present.includes(severity))
    .map(severity => ({ blocker: 'INVALID', error: 'RULE BROKEN', warning: 'CHECK THIS', info: 'NOTE' })[severity])
  assert.deepEqual(tabs, expected)
  for (const name of tabs) {
    const sheet = book.Sheets[name], header = grid(sheet)[1]
    assert.equal(header[11], 'L2 Milestone')
    assert.equal(header[12], 'Severity')
    const tags = carriedTags(sheet), lines = grid(sheet).slice(2)
    /* flagged equipment only: every block has at least one finding line */
    lines.forEach((line, at) => { if (String(line[2] || '').trim()) assert.ok(lines[at][13], `${name} ${tags[at]} has a finding`) })
  }
  /* Index labels and links follow the level tabs */
  assert.equal(grid(book.Sheets.Index)[3][0], 'Finding level')
  tabs.forEach((name, offset) => assert.equal(linkTarget(book.Sheets.Index[`A${5 + offset}`]), name))
  /* an equipment with findings on two levels appears on both tabs */
  const byLevel = new Map()
  for (const finding of result.findings) { const set = byLevel.get(finding.equipmentId) || new Set(); set.add(finding.severity); byLevel.set(finding.equipmentId, set) }
  const multi = [...byLevel.entries()].find(([, levels]) => levels.size > 1)
  assert.ok(multi, 'fixture has an equipment flagged on two levels')
  const appearances = tabs.filter(name => carriedTags(book.Sheets[name]).includes(multi[0])).length
  assert.equal(appearances, multi[1].size)
})

test('the compact writer reports monotonic per-entry progress ending at 1', async () => {
  const { workbookBytesCompact } = await import('../src/core/download.js')
  const book = workbook()
  const calls = []
  await workbookBytesCompact(book, { onProgress: (fraction, name) => { calls.push([fraction, name]) } })
  assert.ok(calls.length >= 5, 'one call per zip entry')
  for (let at = 1; at < calls.length; at++) assert.ok(calls[at][0] >= calls[at - 1][0], 'progress never goes backwards')
  assert.ok(Math.abs(calls[calls.length - 1][0] - 1) < 1e-9, 'ends at exactly 1')
  assert.ok(calls.some(([, name]) => name.includes('sheet')), 'entry names are reported')
})

test('switched-off checks are left out of the Rules sheet', () => {
  const result = auditResult()
  const firedRule = result.findings[0].rule
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx', { disabledRules: [firedRule.id] })
  const titles = grid(book.Sheets.Rules).slice(2).map(line => line[0])
  assert.ok(!titles.includes(firedRule.title), 'the switched-off check is not listed')
  assert.ok(titles.length > 40, 'the rest of the rulebook is still there')
})

test('findings actioned in the app arrive pre-ticked in the report', () => {
  const result = auditResult()
  /* action every finding on one equipment */
  const tag = result.findings[0].equipmentId
  const ids = result.findings.filter(finding => finding.equipmentId === tag).map(finding => finding.id)
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx', { actionedIds: ids })
  const sheet = tabHolding(book, tag), lines = grid(sheet).slice(2), tags = carriedTags(sheet)
  const first = tags.indexOf(tag)
  assert.equal(lines[first][0], '☑', 'its Actioned box starts ticked')
  /* an equipment with none of its findings actioned stays unticked */
  const other = tags.find((candidate, at) => candidate !== tag && String(lines[at][2] || '').trim())
  assert.equal(lines[tags.indexOf(other)][0], '☐')
})

test('a unique legacy tag is accepted and original cell formatting is retained', () => {
  const aoa = [
    headers,
    row({ equipmentId: 'B1-PMP-1', closestParent: '111  Chilled Water (R/S)', closestParentStatus: 'NEW', upn: '111', discipline: 'MECHANICAL WET', systemName: '279  CCW', equipmentDescription: 'Pump' }),
    row({ equipmentId: 'B1-PMP-2', closestParent: 'B1-PMP-1', closestParentStatus: 'NEW', upn: '111', discipline: 'MECHANICAL WET', systemName: '111  Chilled Water (R/S)', equipmentDescription: 'Pump two' }),
  ]
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(aoa), 'Full Export')
  const snapshot = auditSnapshotFromAoa(aoa, { sheet: 'Full Export' })
  const originalStyle = { font: { bold: true }, fill: { fgColor: { rgb: 'ABCDEF' } } }
  workbook.Sheets['Full Export'][XLSX.utils.encode_cell({ r: 1, c: index.systemName })].s = originalStyle
  const applied = applyChangesToWorkbook(workbook, snapshot, [
    { tag: 'B1-PMP-1', field: 'System Name', header: 'System Name', value: '111  Chilled Water (R/S)' },
  ])
  assert.equal(applied, 1)
  const sheet = workbook.Sheets['Full Export']
  const col = index.systemName
  assert.equal(sheet[XLSX.utils.encode_cell({ r: 1, c: col })].v, '111  Chilled Water (R/S)', 'the staged value landed in the right cell')
  assert.equal(sheet[XLSX.utils.encode_cell({ r: 1, c: col })].s, originalStyle, 'corrections do not replace original formatting')
  assert.equal(sheet[XLSX.utils.encode_cell({ r: 2, c: col })].s, undefined, 'untouched cells keep no highlight')
  assert.equal(sheet[XLSX.utils.encode_cell({ r: 2, c: col })].v, '111  Chilled Water (R/S)', 'the untouched row keeps its value')
  assert.equal(sheet[XLSX.utils.encode_cell({ r: 1, c: index.equipmentDescription })].v, 'Pump', 'other columns stay as they were')
})

function correctionFixture(tags = ['EQ-1', 'EQ-2']) {
  const aoa = [headers, ...tags.map((equipmentId, at) => row({ equipmentId, closestParent: 'ROOT', upn: '100', systemName: `Before ${at + 1}`, discipline: at ? 'Electrical' : 'Mechanical', milestone: at ? 'Phase B' : 'Phase A' }))]
  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(aoa), 'Registry')
  const snapshot = auditSnapshotFromAoa(aoa, { sheet: 'Registry' })
  return { book, snapshot, aoa }
}
function correction(snapshot, at = 0, value = 'After', prop = 'systemName') {
  const row = snapshot.rows[at], column = EXTO_REV21_COLUMNS.find(column => column.field === prop)
  return { tag: row.equipmentId, field: column.header, header: column.header, prop, before: row[prop], value, source: { sheet: row._source.sheet, row: row._source.row, columns: { ...row._source.columns } }, ruleId: 'synthetic-rule', findingId: `finding-${at}` }
}
function addressOf(change) { return XLSX.utils.encode_cell({ r: change.source.row - 1, c: change.source.columns[change.prop] }) }
function failsAtomically(book, snapshot, changes, code) {
  const before = structuredClone(book)
  assert.throws(() => validateAuditCorrections(book, snapshot, changes), error => error.code === code)
  assert.deepEqual(book, before, 'preview preflight must not modify any cell or worksheet metadata')
  assert.throws(() => applyChangesToWorkbook(book, snapshot, changes), error => error.code === code)
  assert.deepEqual(book, before, 'failed preflight must not modify any cell or worksheet metadata')
}

function mirroredCorrectionFixture(count = 1) {
  const values = row({ equipmentId: 'EQ-1', closestParent: 'ROOT', upn: '100', systemName: 'Before', equipmentDescription: 'Synthetic equipment' })
  const registry = [headers, ...Array.from({ length: count }, () => [...values])]
  const summary = [[...headers].reverse(), ...Array.from({ length: count }, () => [...values].reverse())]
  const rowNums = [3, ...Array.from({ length: count }, (_, at) => 6 + at * 3)], physicalSummary = []
  summary.forEach((row, at) => { physicalSummary[rowNums[at]] = row })
  const primary = auditSnapshotFromAoa(registry, { sheet: 'Registry' }), mirror = auditSnapshotFromAoa(summary, { sheet: 'Summary', rowNums })
  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(registry), 'Registry')
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(physicalSummary), 'Summary')
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Unchanged note']]), 'Notes')
  return { book, primary, mirror, snapshot: auditMergeSnapshots([primary, mirror], '') }
}

test('correction preview reports logical corrections and unique mirrored cells without mutating any input', () => {
  const { book, snapshot } = mirroredCorrectionFixture(2), change = correction(snapshot, 1)
  const changes = [change, structuredClone(change)], before = structuredClone({ book, snapshot, changes })
  assert.deepEqual(validateAuditCorrections(book, snapshot, changes), { correctionCount: 2, cellCount: 2, sheetCount: 2 })
  assert.deepEqual(validateAuditCorrections(book, snapshot, changes), { correctionCount: 2, cellCount: 2, sheetCount: 2 })
  assert.deepEqual({ book, snapshot, changes }, before)
})

test('correction preview returns zero for an empty batch and counts distinct cells on one sheet', () => {
  const { book, snapshot } = correctionFixture()
  assert.deepEqual(validateAuditCorrections(book, snapshot, []), { correctionCount: 0, cellCount: 0, sheetCount: 0 })
  assert.deepEqual(validateAuditCorrections(book, snapshot, [correction(snapshot), correction(snapshot, 1)]), { correctionCount: 2, cellCount: 2, sheetCount: 1 })
})

test('canonical corrections update proven mirrored rows using each physical column and row mapping', () => {
  const { book, snapshot, mirror } = mirroredCorrectionFixture(), change = correction(snapshot), before = JSON.stringify(snapshot)
  assert.equal(snapshot.rows[0]._sources.length, 2)
  assert.equal(applyChangesToWorkbook(book, snapshot, [change]), 1, 'return count remains logical corrections')
  assert.equal(book.Sheets.Registry[addressOf(change)].v, 'After')
  assert.equal(book.Sheets.Summary[addressOf(correction(mirror))].v, 'After')
  assert.notEqual(change.source.columns.systemName, mirror.rows[0]._source.columns.systemName)
  assert.equal(mirror.rows[0]._source.row, 7)
  assert.equal(JSON.stringify(snapshot), before, 'the canonical and leaf baselines remain immutable')
})

test('same-sheet identical duplicates remain distinct and only the selected occurrence and its mirror change', () => {
  const { book, snapshot, primary, mirror } = mirroredCorrectionFixture(2), change = correction(snapshot, 1)
  assert.equal(snapshot.rows.length, 2)
  applyChangesToWorkbook(book, snapshot, [change])
  for (const source of [primary, mirror]) {
    assert.equal(book.Sheets[source.source.sheet][addressOf(correction(source))].v, 'Before')
    assert.equal(book.Sheets[source.source.sheet][addressOf(correction(source, 1))].v, 'After')
  }
})

test('a stale mirrored row aborts all corrections even when only an untargeted field differs', () => {
  const { book, snapshot, mirror } = mirroredCorrectionFixture(2)
  book.Sheets.Summary[addressOf(correction(mirror, 1, '', 'equipmentDescription'))].v = 'Different equipment'
  failsAtomically(book, snapshot, [correction(snapshot), correction(snapshot, 1)], 'CONFLICT')
})

test('mirrored target conflicts, missing sheets, and formulas never produce a partial correction', () => {
  for (const kind of ['conflict', 'missing', 'formula']) {
    const { book, snapshot, mirror } = mirroredCorrectionFixture(2), target = book.Sheets.Summary[addressOf(correction(mirror, 1))]
    if (kind === 'conflict') target.v = 'Different original value'
    if (kind === 'missing') delete book.Sheets.Summary
    if (kind === 'formula') target.f = '"Before"'
    failsAtomically(book, snapshot, [correction(snapshot), correction(snapshot, 1)], { conflict: 'CONFLICT', missing: 'SHEET', formula: 'FORMULA' }[kind])
  }
})

test('unproven or same-sheet mirror claims are rejected instead of expanding by tag', () => {
  const { book, snapshot, primary, mirror } = mirroredCorrectionFixture(2)
  const altered = sources => ({ ...snapshot, rows: [{ ...snapshot.rows[0], _sources: sources }] })
  failsAtomically(book, altered([primary.rows[0]._source, primary.rows[1]._source]), [correction(snapshot)], 'AMBIGUOUS')
  const wrongColumns = { ...mirror.rows[0]._source, columns: { ...mirror.rows[0]._source.columns, systemName: 0 } }
  failsAtomically(book, altered([primary.rows[0]._source, wrongColumns]), [correction(snapshot)], 'COLUMN')
  const withoutProof = { rows: snapshot.rows }
  failsAtomically(book, withoutProof, [correction(snapshot)], 'SOURCE')
  const different = { ...mirror.rows[0], equipmentDescription: 'Not identical' }
  const falseProof = { ...snapshot, snapshots: [primary, { ...mirror, rows: [different, mirror.rows[1]] }] }
  failsAtomically(book, falseProof, [correction(snapshot)], 'BASELINE')
})

test('explicit writes that conflict with a canonical mirror correction fail atomically', () => {
  const { book, snapshot, mirror } = mirroredCorrectionFixture()
  failsAtomically(book, snapshot, [correction(snapshot), correction(mirror, 0, 'Conflicting after')], 'CONFLICT')
})

test('an unknown correction rejects the whole batch without silently skipping it', () => {
  const { book, snapshot } = correctionFixture()
  failsAtomically(book, snapshot, [correction(snapshot), { tag: 'MISSING', header: 'System Name', value: 'After' }], 'SOURCE')
})

test('duplicate tags require exact physical source identity', () => {
  const { book, snapshot } = correctionFixture(['DUPLICATE', 'DUPLICATE'])
  failsAtomically(book, snapshot, [{ tag: 'duplicate', header: 'System Name', value: 'After' }], 'AMBIGUOUS')
  const change = correction(snapshot, 1)
  assert.equal(applyChangesToWorkbook(book, snapshot, [change]), 1)
  assert.equal(book.Sheets.Registry[addressOf(change)].v, 'After')
  assert.equal(book.Sheets.Registry[addressOf(correction(snapshot))].v, 'Before 1')
})

test('source targeting distinguishes duplicate tags on different sheets, including deduplicated imports', () => {
  const first = correctionFixture(['DUPLICATE']), second = correctionFixture(['DUPLICATE'])
  XLSX.utils.book_append_sheet(first.book, second.book.Sheets.Registry, 'Other')
  const secondSnapshot = auditSnapshotFromAoa(second.aoa, { sheet: 'Other' })
  const merged = { rows: first.snapshot.rows, snapshots: [first.snapshot, secondSnapshot] }
  failsAtomically(first.book, merged, [{ tag: 'DUPLICATE', header: 'System Name', value: 'After' }], 'AMBIGUOUS')
  applyChangesToWorkbook(first.book, merged, [correction(secondSnapshot)])
  assert.equal(first.book.Sheets.Other[addressOf(correction(secondSnapshot))].v, 'After')
  assert.equal(first.book.Sheets.Registry[addressOf(correction(first.snapshot))].v, 'Before 1')
})

test('physical row numbers and imported column mappings override decoy headers and contract positions', () => {
  const values = row({ equipmentId: 'EQ-1', closestParent: 'ROOT', upn: '100', systemName: 'Before' }).reverse()
  const compressed = [[...headers].reverse(), values], snapshot = auditSnapshotFromAoa(compressed, { sheet: 'Registry', rowNums: [39, 72] })
  const physical = [headers]
  physical[39] = compressed[0]; physical[72] = values
  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(physical), 'Registry')
  const change = correction(snapshot)
  applyChangesToWorkbook(book, snapshot, [change])
  assert.equal(book.Sheets.Registry[addressOf(change)].v, 'After')
  assert.equal(change.source.row, 73)
  assert.notEqual(change.source.columns.systemName, index.systemName)
  assert.deepEqual(grid(book.Sheets.Registry)[0], headers)
})

test('source-targeted corrections never fall back to tags or guessed columns', () => {
  const { book, snapshot } = correctionFixture()
  const wrongSheet = correction(snapshot); wrongSheet.source.sheet = 'Missing'
  failsAtomically(book, snapshot, [wrongSheet], 'SOURCE')
  const wrongRow = correction(snapshot); wrongRow.source.row++
  failsAtomically(book, snapshot, [wrongRow], 'SOURCE')
  const wrongColumns = correction(snapshot); wrongColumns.source.columns.systemName++
  failsAtomically(book, snapshot, [wrongColumns], 'COLUMN')
  const missingColumns = correction(snapshot); delete missingColumns.source.columns
  failsAtomically(book, snapshot, [missingColumns], 'COLUMN')
  const missingBefore = correction(snapshot); delete missingBefore.before
  failsAtomically(book, snapshot, [missingBefore], 'BEFORE')
  const mismatchedField = correction(snapshot); mismatchedField.header = 'Discipline'
  failsAtomically(book, snapshot, [mismatchedField], 'FIELD')
})

test('a changed baseline or old workbook cell stops the entire correction batch', () => {
  const { book, snapshot } = correctionFixture(), first = correction(snapshot), second = correction(snapshot, 1)
  failsAtomically(book, snapshot, [first, { ...second, before: 'Not original' }], 'BASELINE')
  book.Sheets.Registry[addressOf(second)].v = 'Changed outside baseline'
  failsAtomically(book, snapshot, [first, second], 'CONFLICT')
  const draft = { ...snapshot, rows: snapshot.rows.map(row => ({ ...row, systemName: 'After' })) }
  failsAtomically(book, draft, [first], 'BASELINE')
})

test('conflicting writes to a cell fail while repeated identical corrections are accounted for', () => {
  const { book, snapshot } = correctionFixture(), change = correction(snapshot)
  failsAtomically(book, snapshot, [change, { ...change, value: 'Other after' }], 'CONFLICT')
  assert.equal(applyChangesToWorkbook(book, snapshot, [change, { ...change }]), 2)
})

test('formula, merged, unsupported, and missing-sheet targets reject before mutation', () => {
  const { book, snapshot } = correctionFixture(), first = correction(snapshot), second = correction(snapshot, 1)
  book.Sheets.Registry[addressOf(second)].f = '"Before 2"'
  failsAtomically(book, snapshot, [first, second], 'FORMULA')
  delete book.Sheets.Registry[addressOf(second)].f
  const cell = XLSX.utils.decode_cell(addressOf(second))
  book.Sheets.Registry['!merges'] = [{ s: cell, e: { r: cell.r, c: cell.c + 1 } }]
  failsAtomically(book, snapshot, [first, second], 'MERGED')
  delete book.Sheets.Registry['!merges']
  for (const value of [undefined, NaN, Infinity, {}, '\u0000', '\uD800', 'x'.repeat(32768)]) failsAtomically(book, snapshot, [first, { ...second, value }], 'VALUE')
  delete book.Sheets.Registry
  failsAtomically(book, snapshot, [first], 'SHEET')
})

test('blank cells and typed replacement values are written without mutating the baseline', () => {
  const { book, snapshot } = correctionFixture(), before = JSON.stringify(snapshot)
  const change = correction(snapshot, 0, 0, 'dependencies')
  delete book.Sheets.Registry[addressOf(change)]
  applyChangesToWorkbook(book, snapshot, [change, correction(snapshot, 1, false, 'dependencies')])
  assert.equal(book.Sheets.Registry[addressOf(change)].v, 0)
  assert.equal(book.Sheets.Registry[addressOf(change)].t, 'n')
  assert.equal(book.Sheets.Registry[addressOf(correction(snapshot, 1, '', 'dependencies'))].t, 'b')
  assert.equal(JSON.stringify(snapshot), before)
})

test('action corrections use the imported trim contract for both targeted and untargeted whitespace', () => {
  for (const field of ['UPN', 'System Name']) {
    const aoa = [headers, row({ equipmentId: 'EQ-1', closestParent: 'ROOT', upn: '602', systemName: '602  Medium Voltage ' })]
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(aoa), 'Registry')
    const snapshot = auditSnapshotFromAoa(aoa, { sheet: 'Registry' }), change = auditMakeCorrection(snapshot.rows[0], field, field === 'UPN' ? '603' : '603  Low Voltage')
    assert.deepEqual(validateAuditCorrections(book, snapshot, [change]), { correctionCount: 1, cellCount: 1, sheetCount: 1 })
    assert.equal(applyChangesToWorkbook(book, snapshot, [change]), 1)
    assert.equal(book.Sheets.Registry[addressOf(change)].v, change.value)
    if (field === 'UPN') assert.equal(book.Sheets.Registry[XLSX.utils.encode_cell({ r: 1, c: index.systemName })].v, '602  Medium Voltage ', 'untargeted original whitespace remains intact')
  }
})

test('formatted numeric before values match import without trusting cached display text or mutating cells', async () => {
  const { book } = correctionFixture(), address = XLSX.utils.encode_cell({ r: 1, c: index.upn })
  book.Sheets.Registry[address] = { t: 'n', v: 602, z: '0000', w: '0602' }
  const snapshot = await auditSnapshotFromWorkbook(book, ''), change = auditMakeCorrection(snapshot.rows[0], 'UPN', '603'), before = structuredClone(book)
  assert.equal(change.before, '0602')
  assert.deepEqual(validateAuditCorrections(book, snapshot, [change]), { correctionCount: 1, cellCount: 1, sheetCount: 1 })
  assert.deepEqual(book, before)
  book.Sheets.Registry[address].v = 604
  failsAtomically(book, snapshot, [change], 'CONFLICT')
})

test('normalized before comparison still rejects internal whitespace and meaningful value changes', () => {
  const { book, snapshot } = correctionFixture(), change = correction(snapshot)
  for (const value of ['Before  1', 'Before 2', 'before 1']) {
    book.Sheets.Registry[addressOf(change)].v = value
    failsAtomically(book, snapshot, [change], 'CONFLICT')
  }
})

test('proven mirrors may have different surrounding whitespace without blocking canonical corrections', () => {
  const { book, snapshot, mirror } = mirroredCorrectionFixture()
  book.Sheets.Summary[addressOf(correction(mirror))].v = ' Before '
  assert.deepEqual(validateAuditCorrections(book, snapshot, [correction(snapshot)]), { correctionCount: 1, cellCount: 2, sheetCount: 2 })
  applyChangesToWorkbook(book, snapshot, [correction(snapshot)])
  assert.equal(book.Sheets.Summary[addressOf(correction(mirror))].v, 'After')
})

test('read-only target sheets fail before another sheet can be changed', () => {
  const first = correctionFixture(), second = correctionFixture()
  XLSX.utils.book_append_sheet(first.book, Object.freeze(second.book.Sheets.Registry), 'Other')
  const otherSnapshot = auditSnapshotFromAoa(second.aoa, { sheet: 'Other' })
  const baseline = { rows: [...first.snapshot.rows, ...otherSnapshot.rows] }
  failsAtomically(first.book, baseline, [correction(first.snapshot), correction(otherSnapshot)], 'WORKBOOK')
})

test('correction log and review actions retain before, after, reason, owner, physical source and distinct statuses', () => {
  const { snapshot } = correctionFixture(), change = correction(snapshot, 0, '=literal text')
  const history = [
    { id: 'review-1', at: '2026-01-05T09:00:00Z', owner: 'Reviewer', reason: 'Confirmed original metadata', disposition: 'corrected-draft', findingIds: [change.findingId], changes: [change] },
    { id: 'review-2', at: '2026-01-05T09:10:00Z', owner: 'Reviewer', reason: 'Checked without editing', disposition: 'reviewed', findingIds: ['review-only'], changes: [] },
    { id: 'review-3', at: '2026-01-05T09:20:00Z', owner: 'Reviewer', reason: 'Approved exception', disposition: 'exception', findingIds: ['exception-only'], changes: [] },
  ]
  const book = buildAuditCorrectionsWorkbook([change], history), log = grid(book.Sheets['Correction Log']), actions = grid(book.Sheets.Actions)
  assert.deepEqual(book.SheetNames, ['Correction Log', 'Actions'])
  assert.deepEqual(log[1].slice(0, 10), ['EQ-1', 'System Name', 'Before 1', '=literal text', 'Confirmed original metadata', 'Registry', 2, 'J', 'J2', 'Staged in draft'])
  assert.equal(book.Sheets['Correction Log'].D2.f, undefined, 'formula-like review data stays literal text')
  assert.equal(log[1][10], 'Reviewer')
  assert.equal(actions.length, 4)
  assert.deepEqual(actions.slice(1).map(row => row[3]), ['Corrected in draft', 'Reviewed', 'Exception'])
  assert.deepEqual(actions.slice(1).map(row => row[4]), history.map(entry => entry.reason))
  assert.equal(book.Sheets.Actions['!freeze'], 'A2')
  assert.ok(book.Sheets['Correction Log']['!autofilter'])
})

test('tracker retains resolved and reviewed baseline findings and adds new draft findings without before-after duplicates', () => {
  const { snapshot } = correctionFixture(['DUPLICATE', 'DUPLICATE'])
  const finding = (id, at, ruleId) => ({ id, equipmentId: 'DUPLICATE', sheet: 'Registry', row: at + 2, rule: { id: ruleId } })
  const baseline = { rows: snapshot.rows, findings: [finding('resolved', 0, 'rule-a'), finding('reviewed', 1, 'rule-b'), finding('open', 1, 'rule-c')] }
  const current = { rows: snapshot.rows.map(row => ({ ...row, milestone: 'Changed draft milestone' })), findings: [finding('after-review', 1, 'rule-b'), finding('open', 1, 'rule-c'), finding('new', 0, 'rule-new')] }
  const sheet = buildAuditTrackerWorkbook(current, '', { baselineResult: baseline, actionedIds: new Set(['reviewed']), draftResolvedIds: new Set(['resolved']) }).Sheets.Tracker
  const lines = grid(sheet).slice(8)
  assert.deepEqual(lines.map(row => row.slice(0, 3)), [['Phase A', 2, 0], ['Phase B', 2, 0]])
  assert.equal(sheet.E6.f, `COUNTIF(F9:F10,"${AUDIT_EXPORT_TICK}")/2`, 'the overall bar still follows manual milestone ticks')
  assert.match(sheet['!xmlExtras'].dataValidations[0], /sqref="F9:F10"/)
  assert.deepEqual(grid(sheet)[7], ['L2 Milestone', 'Findings', 'Signed off', '%', 'Progress', 'Signed off?'])
  assert.equal(sheet.C9.f, `IF(F9="${AUDIT_EXPORT_TICK}",B9,0)`)
  assert.equal(sheet.D9.f, 'IF(B9=0,0,C9/B9)')
  assert.equal(sheet.E9.f, 'D9')
  const completed = buildAuditTrackerWorkbook({ rows: snapshot.rows, findings: [] }, '', { baselineResult: baseline, actionedIds: ['reviewed', 'open'], draftResolvedIds: ['resolved'] }).Sheets.Tracker
  assert.deepEqual(grid(completed).slice(8).map(row => [row[1], row[2], row[5]]), [[1, 0, '☐'], [2, 0, '☐']], 'app actioning never pre-ticks manual sign-off')
})

test('tracker honors rule and completed-equipment scope while retaining explicitly reviewed exceptions', () => {
  const { snapshot } = correctionFixture()
  const finding = (id, at, rule) => ({ id, equipmentId: snapshot.rows[at].equipmentId, sheet: 'Registry', row: at + 2, rule: { id: rule } })
  const baseline = { rows: snapshot.rows, findings: [finding('exception', 0, 'rule-a'), finding('excluded', 0, 'rule-b'), finding('disabled', 0, 'rule-c'), finding('site-complete', 1, 'rule-a')] }
  const sheet = buildAuditTrackerWorkbook(baseline, '', { baselineResult: baseline, actionedIds: ['exception'], excludedIds: ['exception', 'excluded'], disabledRules: ['rule-c'], completedEquipmentIds: ['eq-2'] }).Sheets.Tracker
  assert.deepEqual(grid(sheet).slice(8).map(row => row.slice(0, 3)), [['Phase A', 1, 0]])
})

test('tracker does not complete sibling baseline findings of the same rule when only one was reviewed', () => {
  const { snapshot } = correctionFixture()
  const baseline = { rows: snapshot.rows, findings: ['one', 'two'].map(id => ({ id, equipmentId: 'EQ-1', sheet: 'Registry', row: 2, rule: { id: 'shared-rule' } })) }
  const sheet = buildAuditTrackerWorkbook(baseline, '', { baselineResult: baseline, actionedIds: ['one'] }).Sheets.Tracker
  assert.deepEqual(grid(sheet)[8].slice(0, 3), ['Phase A', 2, 0])
})

test('tracker discipline sign-off combines milestones and has no frozen cross-discipline percentages', () => {
  const { snapshot } = correctionFixture()
  const rows = snapshot.rows.map((row, index) => ({ ...row, discipline: 'Electrical', milestone: `Phase ${index + 1}` }))
  const result = { rows, findings: rows.map(row => ({ id: row.equipmentId, equipmentId: row.equipmentId, sheet: row._source.sheet, row: row._source.row, rule: { id: 'test' } })) }
  const sheet = buildAuditTrackerWorkbook(result, '', { signOffBy: 'discipline', actionedIds: result.findings.map(f => f.id) }).Sheets.Tracker
  assert.deepEqual(grid(sheet)[7], ['Discipline', 'Findings', 'Signed off', '%', 'Progress', 'Signed off?'])
  assert.deepEqual(grid(sheet)[8], ['Electrical', rows.length, 0, 0, 0, '☐'])
  assert.equal(sheet.E6.f, `COUNTIF(F9:F9,"${AUDIT_EXPORT_TICK}")/1`)
  assert.equal(sheet.C9.f, `IF(F9="${AUDIT_EXPORT_TICK}",B9,0)`)
  assert.equal(sheet['!ref'], 'A1:F9', 'a large milestone list cannot create a spreadsheet hundreds of columns wide')
})

test('tracker safely handles empty scope and missing grouping metadata in either mode', () => {
  for (const signOffBy of ['milestone', 'discipline']) {
    const empty = buildAuditTrackerWorkbook({ rows: [], findings: [] }, '', { signOffBy }).Sheets.Tracker
    assert.equal(empty.A9.v, 'No findings in scope')
    assert.equal(empty.E6?.f, undefined)
    assert.deepEqual(empty['!xmlExtras'].dataValidations, [])
    const { snapshot } = correctionFixture()
    const row = { ...snapshot.rows[0], milestone: '', discipline: '' }
    const result = { rows: [row], findings: [{ id: 'missing', equipmentId: row.equipmentId, sheet: row._source.sheet, row: row._source.row, rule: { id: 'test' } }] }
    const sheet = buildAuditTrackerWorkbook(result, '', { signOffBy }).Sheets.Tracker
    assert.equal(sheet.A9.v, signOffBy === 'discipline' ? 'No discipline' : 'No L2 milestone')
    assert.equal(sheet.F9.v, '☐')
  }
})

test('large review batches use separate finding rows instead of overflowing a joined-ID cell', () => {
  const findingIds = Array.from({ length: 1200 }, (_, index) => `synthetic-finding-with-a-long-identifier-${index}`)
  const history = [{ id: 'large-review', at: '', owner: '', reason: 'Reviewed batch', disposition: 'reviewed', findingIds, changes: [] }]
  const sheet = buildAuditCorrectionsWorkbook([], history).Sheets.Actions, rows = grid(sheet)
  assert.equal(rows.length, findingIds.length + 1)
  assert.deepEqual(rows.slice(1).map(row => row[5]), findingIds)
  assert.ok(rows.every(row => row.every(cell => typeof cell !== 'string' || cell.length <= 32767)))
})

function captureExportDownloads(t) {
  const priorDocument = globalThis.document, priorSession = S.session, downloads = [], blobs = new Map(), elements = new Map()
  const element = () => ({ style: {}, classList: { add() {}, remove() {}, contains() { return false } }, setAttribute() {}, removeEventListener() {}, addEventListener() {} })
  globalThis.document = {
    querySelector(selector) { if (!elements.has(selector)) elements.set(selector, element()); return elements.get(selector) },
    body: { appendChild() {} },
    createElement() { return { remove() {}, click() { downloads.push(blobs.get(this.href)) } } },
  }
  t.mock.method(URL, 'createObjectURL', blob => { const url = `blob:synthetic-${blobs.size}`; blobs.set(url, blob); return url })
  t.mock.method(URL, 'revokeObjectURL', () => {})
  t.mock.method(globalThis, 'setTimeout', callback => { queueMicrotask(callback); return 0 })
  t.mock.method(globalThis, 'clearTimeout', () => {})
  t.after(() => { if (priorDocument === undefined) delete globalThis.document; else globalThis.document = priorDocument; S.session = priorSession })
  resetSession()
  return { downloads, elements }
}

test('correction export reads session reviewHistory, including review-only physical source', async t => {
  const { downloads } = captureExportDownloads(t), { snapshot } = correctionFixture(), change = correction(snapshot)
  S.session.baselineResult = { rows: snapshot.rows, findings: [{ id: 'review-only', equipmentId: 'EQ-2', sheet: 'Registry', row: 3, field: 'System Name', actual: 'Before 2', rule: { id: 'review-rule' } }] }
  S.session.changes = [change]
  S.session.reviewHistory = [{ id: 'review', at: '2026-01-05T09:00:00Z', owner: 'Reviewer', reason: 'Checked in place', disposition: 'reviewed', findingIds: ['review-only'], changes: [] }]
  assert.equal(await exportAuditCorrectionsXlsx(), true)
  assert.equal(downloads.length, 1)
  const book = XLSX.read(await downloads[0].arrayBuffer(), { type: 'array' }), actions = grid(book.Sheets.Actions)
  assert.deepEqual(actions[1].slice(6, 15), ['EQ-2', 'System Name', 'Before 2', '', 'Registry', 3, 'J', 'J3', 'review-rule'])
})

function packageEntries(bytes) {
  const container = XLSX.CFB.read(new Uint8Array(bytes), { type: 'array' }), entries = new Map()
  container.FullPaths.forEach((path, at) => {
    const name = path.slice(container.FullPaths[0].length), file = container.FileIndex[at]
    if (name && !name.endsWith('/') && !name.startsWith('\u0001') && file.type === 2) entries.set(name, new Uint8Array(file.content))
  })
  return entries
}
function packageBytes(entries) {
  const container = XLSX.CFB.utils.cfb_new()
  for (const [path, data] of entries) XLSX.CFB.utils.cfb_add(container, path, data)
  return new Uint8Array(XLSX.CFB.write(container, { fileType: 'zip', type: 'array', compression: true }))
}
const utf8 = value => new TextEncoder().encode(value)
const xmlDocument = bytes => new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml')
const xmlBytes = document => utf8(new XMLSerializer().serializeToString(document))
function syntheticPackage() {
  const { book, snapshot } = correctionFixture(), change = correction(snapshot)
  book.Sheets.Registry[addressOf(change)].s = { font: { bold: true, color: { rgb: '006633' } }, fill: { patternType: 'solid', fgColor: { rgb: 'EEEEEE' } } }
  book.Sheets.Registry.A6 = { t: 'n', v: 2, f: '1+1' }; book.Sheets.Registry['!ref'] = 'A1:AR6'
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Unchanged'], [12]]), 'Other')
  const entries = packageEntries(XLSX.write(book, { bookType: 'xlsx', type: 'array', cellStyles: true, bookSST: true }))
  const document = xmlDocument(entries.get('xl/worksheets/sheet1.xml')), root = document.documentElement
  const extras = new DOMParser().parseFromString(`<extras xmlns="${root.namespaceURI}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><mergeCells count="1"><mergeCell ref="B6:C6"/></mergeCells><conditionalFormatting sqref="J2:J3"><cfRule type="expression" priority="1"><formula>J2="After"</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" sqref="J2:J3"><formula1>"Before 1,After"</formula1></dataValidation></dataValidations><legacyDrawing r:id="rForm"/><controls><control shapeId="1025" r:id="rControl" name="Completion"/></controls></extras>`, 'application/xml')
  for (const node of Array.from(extras.documentElement.childNodes)) root.appendChild(document.importNode(node, true))
  const target = 'xl/worksheets/physical.xml'
  entries.delete('xl/worksheets/sheet1.xml'); entries.set(target, xmlBytes(document))
  const rels = xmlDocument(entries.get('xl/_rels/workbook.xml.rels'))
  for (const relation of Array.from(rels.getElementsByTagNameNS('*', 'Relationship'))) if (relation.getAttribute('Target') === 'worksheets/sheet1.xml') relation.setAttribute('Target', 'worksheets/physical.xml')
  entries.set('xl/_rels/workbook.xml.rels', xmlBytes(rels))
  const types = xmlDocument(entries.get('[Content_Types].xml'))
  for (const type of Array.from(types.getElementsByTagNameNS('*', 'Override'))) if (type.getAttribute('PartName') === '/xl/worksheets/sheet1.xml') type.setAttribute('PartName', `/${target}`)
  entries.set('[Content_Types].xml', xmlBytes(types))
  entries.set('xl/worksheets/_rels/physical.xml.rels', utf8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rForm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/form.vml"/><Relationship Id="rControl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/control.xml"/></Relationships>'))
  entries.set('xl/drawings/form.vml', utf8('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t201"><x:ClientData ObjectType="Checkbox"><x:Checked>1</x:Checked></x:ClientData></v:shape></xml>'))
  entries.set('xl/ctrlProps/control.xml', utf8('<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="Checked"/>'))
  entries.set('customXml/item.xml', utf8('<synthetic><retained>yes</retained></synthetic>'))
  entries.set('xl/media/image.bin', new Uint8Array([0, 1, 2, 200, 255]))
  return { bytes: packageBytes(entries), entries, snapshot, change, target }
}

packageTest('package export preserves forms, styles, relationships, formulas and every untargeted part payload', async () => {
  const { bytes, entries, snapshot, change, target } = syntheticPackage(), original = bytes.slice()
  change.value = ' After <&> "literal" _x0041_ \r\nSecond line '
  const output = await buildUpdatedRegistryBytes(bytes, snapshot, [change]), after = packageEntries(output)
  assert.deepEqual(bytes, original, 'the original imported bytes never change')
  assert.deepEqual([...after.keys()].sort(), [...entries.keys()].sort())
  for (const [path, data] of entries) if (path !== target && path !== 'xl/styles.xml') assert.deepEqual(after.get(path), data, `untargeted package part changed: ${path}`)
  const beforeXml = xmlDocument(entries.get(target)), afterXml = xmlDocument(after.get(target))
  const at = document => Array.from(document.getElementsByTagNameNS('*', 'c')).find(cell => cell.getAttribute('r') === addressOf(change))
  const beforeStyles=xmlDocument(entries.get('xl/styles.xml')),afterStyles=xmlDocument(after.get('xl/styles.xml'))
  const formats=document=>Array.from(document.getElementsByTagNameNS('*','cellXfs')[0].childNodes).filter(node=>node.nodeType===1)
  const oldFormats=formats(beforeStyles),newFormats=formats(afterStyles),serialize=node=>new XMLSerializer().serializeToString(node)
  oldFormats.forEach((format,index)=>assert.equal(serialize(newFormats[index]),serialize(format),'existing styles stay intact'))
  const oldFormat=oldFormats[Number(at(beforeXml).getAttribute('s')||0)],newFormat=newFormats[Number(at(afterXml).getAttribute('s'))].cloneNode(true)
  newFormat.setAttribute('fillId',oldFormat.getAttribute('fillId'));if(oldFormat.hasAttribute('applyFill'))newFormat.setAttribute('applyFill',oldFormat.getAttribute('applyFill'));else newFormat.removeAttribute('applyFill')
  assert.equal(serialize(newFormat),serialize(oldFormat),'only the fill changes on the corrected cell')
  for (const document of [beforeXml, afterXml]) { const cell = at(document); cell.parentNode.removeChild(cell) }
  assert.equal(new XMLSerializer().serializeToString(afterXml), new XMLSerializer().serializeToString(beforeXml), 'all other worksheet XML survives semantically, including form references and validation')
  const restored = XLSX.read(output, { type: 'array', cellStyles: true })
  assert.equal(restored.Sheets.Registry[addressOf(change)].v, change.value)
  assert.equal(restored.Sheets.Registry[addressOf(change)].s.fgColor.rgb,'FFF2CC')
  assert.equal(restored.Sheets.Registry.A6.f, '1+1')
})

packageTest('package conflicts are atomic and compression never starts for rejected corrections', async () => {
  const { bytes, snapshot, change } = syntheticPackage(), before = bytes.slice()
  let progress = false
  await assert.rejects(buildUpdatedRegistryBytes(bytes, snapshot, [change, { ...correction(snapshot, 1), before: 'Changed' }], { onProgress: () => { progress = true } }), error => error.code === 'BASELINE')
  assert.equal(progress, false)
  assert.deepEqual(bytes, before)
})

packageTest('export and reimport preserve logical row counts after correcting occurrence-paired summary copies', async () => {
  const { book, snapshot } = mirroredCorrectionFixture(2), changes = [correction(snapshot, 1)]
  const bytes = new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' })), before = bytes.slice(), originalParts = packageEntries(bytes)
  const output = await buildUpdatedRegistryBytes(bytes, snapshot, changes), reimported = XLSX.read(output, { type: 'array' })
  const after = await auditSnapshotFromWorkbook(reimported, '')
  assert.equal(after.rows.length, snapshot.rows.length, 'no stale summary copy becomes an extra conflicting equipment row')
  assert.deepEqual(after.rows.map(row => row.systemName), ['Before', 'After'])
  assert.deepEqual(after.rows.map(row => row._sources.length), [2, 2])
  const resultParts = packageEntries(output)
  for (const [path, data] of originalParts) if (!['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml','xl/styles.xml'].includes(path)) assert.deepEqual(resultParts.get(path), data)
  assert.deepEqual(bytes, before)
})

packageTest('package preflight and preview agree for formatted numbers and whitespace in mirrored originals', async () => {
  const { book, primary, mirror } = mirroredCorrectionFixture()
  for (const source of [primary, mirror]) book.Sheets[source.source.sheet][addressOf(correction(source, 0, '', 'upn'))] = { t: 'n', v: 602, z: '0000' }
  book.Sheets.Summary[addressOf(correction(mirror))].v = ' Before '
  const entries = packageEntries(XLSX.write(book, { type: 'array', bookType: 'xlsx', cellStyles: true }))
  // The vendored writer drops custom number formats, so supply real OOXML proof.
  const styles = xmlDocument(entries.get('xl/styles.xml')), root = styles.documentElement, ns = root.namespaceURI
  let formats = styles.getElementsByTagNameNS(ns, 'numFmts')[0]
  if (!formats) { formats = styles.createElementNS(ns, 'numFmts'); root.insertBefore(formats, root.firstChild) }
  const format = styles.createElementNS(ns, 'numFmt'); format.setAttribute('numFmtId', '164'); format.setAttribute('formatCode', '0000'); formats.appendChild(format)
  formats.setAttribute('count', String(formats.getElementsByTagNameNS(ns, 'numFmt').length))
  const xfs = styles.getElementsByTagNameNS(ns, 'cellXfs')[0], styleIndex = xfs.getElementsByTagNameNS(ns, 'xf').length, xf = xfs.firstChild.cloneNode(true)
  xf.setAttribute('numFmtId', '164'); xf.setAttribute('applyNumberFormat', '1'); xfs.appendChild(xf); xfs.setAttribute('count', String(styleIndex + 1))
  entries.set('xl/styles.xml', xmlBytes(styles))
  for (const [index, source] of [primary, mirror].entries()) {
    const path = `xl/worksheets/sheet${index + 1}.xml`, document = xmlDocument(entries.get(path)), address = addressOf(correction(source, 0, '', 'upn'))
    Array.from(document.getElementsByTagNameNS(ns, 'c')).find(cell => cell.getAttribute('r') === address).setAttribute('s', String(styleIndex))
    entries.set(path, xmlBytes(document))
  }
  const bytes = packageBytes(entries), before = bytes.slice()
  const imported = XLSX.read(bytes.slice(), { type: 'array', cellStyles: true }), snapshot = await auditSnapshotFromWorkbook(imported, '')
  assert.equal(snapshot.rows[0].upn, '0602')
  assert.equal(snapshot.rows.length, 1)
  const changes = [auditMakeCorrection(snapshot.rows[0], 'UPN', '603'), auditMakeCorrection(snapshot.rows[0], 'System Name', 'After')]
  assert.deepEqual(validateAuditCorrections(imported, snapshot, changes), { correctionCount: 2, cellCount: 4, sheetCount: 2 })
  const output = await buildUpdatedRegistryBytes(bytes, snapshot, changes), restored = await auditSnapshotFromWorkbook(XLSX.read(output, { type: 'array', cellStyles: true }), '')
  assert.equal(restored.rows.length, 1)
  assert.equal(restored.rows[0]._sources.length, 2)
  assert.equal(restored.rows[0].upn, '603')
  assert.equal(restored.rows[0].systemName, 'After')
  assert.deepEqual(bytes, before)
})

packageTest('package mirror conflicts are rejected before compression or source-byte mutation', async () => {
  const { book, snapshot, mirror } = mirroredCorrectionFixture(), target = addressOf(correction(mirror))
  book.Sheets.Summary[target].v = 'Stale summary'
  const bytes = new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' })), before = bytes.slice()
  let compressed = false
  await assert.rejects(buildUpdatedRegistryBytes(bytes, snapshot, [correction(snapshot)], { onProgress: () => { compressed = true } }), error => error.code === 'CONFLICT')
  assert.equal(compressed, false)
  assert.deepEqual(bytes, before)
})

packageTest('package export creates missing blank cells and retains inline whitespace and literal formula text', async () => {
  const { bytes, snapshot } = syntheticPackage()
  const first = correction(snapshot, 0, ' =1+1 ', 'dependencies'), second = correction(snapshot, 1, false, 'dependencies')
  const output = await buildUpdatedRegistryBytes(bytes, snapshot, [second, first]), restored = XLSX.read(output, { type: 'array' })
  assert.equal(restored.Sheets.Registry[addressOf(first)].v, first.value)
  assert.equal(restored.Sheets.Registry[addressOf(first)].f, undefined)
  assert.equal(restored.Sheets.Registry[addressOf(second)].v, false)
})

packageTest('missing cells are inserted in column order even when corrections arrive in reverse order', async () => {
  const { entries, snapshot, target } = syntheticPackage(), document = xmlDocument(entries.get(target))
  const first = correction(snapshot, 0, 'After', 'dependencies'), second = correction(snapshot, 0, 'Project', 'dependencyProject')
  for (const cell of Array.from(document.getElementsByTagNameNS('*', 'c'))) if ([addressOf(first), addressOf(second)].includes(cell.getAttribute('r'))) cell.parentNode.removeChild(cell)
  entries.set(target, xmlBytes(document))
  const output = packageEntries(await buildUpdatedRegistryBytes(packageBytes(entries), snapshot, [second, first])), result = xmlDocument(output.get(target))
  const row = Array.from(result.getElementsByTagNameNS('*', 'row')).find(row => row.getAttribute('r') === '2')
  const columns = Array.from(row.getElementsByTagNameNS('*', 'c')).map(cell => XLSX.utils.decode_cell(cell.getAttribute('r')).c)
  assert.deepEqual(columns, [...columns].sort((a, b) => a - b))
})

packageTest('package export also preserves part payloads when native deflate is unavailable', async () => {
  const { bytes, entries, snapshot, change, target } = syntheticPackage(), compression = globalThis.CompressionStream
  try {
    globalThis.CompressionStream = undefined
    const output = packageEntries(await buildUpdatedRegistryBytes(bytes, snapshot, [change]))
    for (const [path, data] of entries) if (path !== target && path !== 'xl/styles.xml') assert.deepEqual(output.get(path), data)
  } finally { globalThis.CompressionStream = compression }
})

packageTest('signed, legacy, and unsupported workbook packages fail without values-only fallback', async () => {
  const { bytes, entries, snapshot, change } = syntheticPackage()
  await assert.rejects(buildUpdatedRegistryBytes(new Uint8Array([1, 2, 3]), snapshot, [change]), error => error.code === 'PACKAGE')
  entries.set('_xmlsignatures/sig.xml', utf8('<Signature/>'))
  await assert.rejects(buildUpdatedRegistryBytes(packageBytes(entries), snapshot, [change]), error => error.code === 'PACKAGE')
  entries.delete('_xmlsignatures/sig.xml')
  const types = xmlDocument(entries.get('[Content_Types].xml'))
  Array.from(types.getElementsByTagNameNS('*', 'Override')).find(node => node.getAttribute('PartName') === '/xl/workbook.xml').setAttribute('ContentType', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
  entries.set('[Content_Types].xml', xmlBytes(types))
  await assert.rejects(buildUpdatedRegistryBytes(packageBytes(entries), snapshot, [change]), error => error.code === 'PACKAGE')
  assert.ok(bytes.length)
})

packageTest('an ambiguous worksheet cell aborts package export before mutation', async () => {
  const { entries, snapshot, change, target } = syntheticPackage(), document = xmlDocument(entries.get(target))
  const cell = Array.from(document.getElementsByTagNameNS('*', 'c')).find(cell => cell.getAttribute('r') === addressOf(change))
  cell.parentNode.appendChild(cell.cloneNode(true)); entries.set(target, xmlBytes(document))
  await assert.rejects(buildUpdatedRegistryBytes(packageBytes(entries), snapshot, [change]), error => error.code === 'XML')
})

packageTest('recommended drive corrections round-trip with every metadata cell intact and only changes yellow',async()=>{
  const system=extoRev21SystemsForUpn('101')[0];
  const records=[
    {equipmentId:'DEMO-MAH101-01-00',upn:'101',systemName:system,discipline:'MECHANICAL DRY',building:'DEMO',equipmentDescription:'MAH Makeup Air Handler'},
    {equipmentId:'DEMO-VFD101-01-00',upn:'650',systemName:extoRev21SystemsForUpn('650')[0],discipline:'FACILITIES MONITORING SYSTEM',building:'DEMO',equipmentDescription:'Variable Frequency Drive',closestParent:'DEMO-MAH101-01-00',dependencies:'DEMO-PNL-1',manufacturer:'Retain manufacturer',modelNumber:'Retain model',milestone:'Retain L2',milestoneParent:'Retain L1',itemMaster:'Retain checklist'},
  ];
  const aoa=[EXTO_REV21_COLUMNS.map(c=>c.header).concat('Unmapped custom metadata'),...records.map((record,i)=>EXTO_REV21_COLUMNS.map(c=>record[c.field]||'').concat(`Keep custom ${i}`))];
  const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
  const bytes=new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'})),original=bytes.slice(),snapshot=await auditSnapshotFromWorkbook(XLSX.read(bytes,{type:'array'}),'');
  const issue=runSsmAudit(snapshot).findings.find(f=>f.rule.id==='parent.cross-upn');
  const proposal=auditProposeCorrection(issue,auditRecommendationContext(snapshot));assert.ok(proposal);
  const draft=auditApplyCorrections(snapshot,proposal.changes),output=await buildUpdatedRegistryBytes(bytes,snapshot,proposal.changes),restored=XLSX.read(output,{type:'array',cellStyles:true});
  const changes=new Map(proposal.changes.map(c=>[addressOf(c),c.value]));
  for(let r=0;r<aoa.length;r++)for(let c=0;c<aoa[r].length;c++){
    const address=XLSX.utils.encode_cell({r,c}),cell=restored.Sheets.Registry[address];
    assert.equal(cell?.v??'',changes.has(address)?changes.get(address):aoa[r][c],`metadata ${address}`);
    if(changes.has(address))assert.equal(cell.s.fgColor.rgb,'FFF2CC');
    else assert.notEqual(cell?.s?.fgColor?.rgb,'FFF2CC',`untouched ${address} is not highlighted`);
  }
  const reimported=await auditSnapshotFromWorkbook(restored,'');
  assert.deepEqual(reimported.rows.map(r=>EXTO_REV21_COLUMNS.map(c=>r[c.field])),draft.rows.map(r=>EXTO_REV21_COLUMNS.map(c=>r[c.field])));
  assert.deepEqual(new Uint8Array(bytes),original);
});

packageTest('cleared Dependency Project and changed Building stay yellow without losing dependencies',async()=>{
  const rows=[{equipmentId:'CHILD',building:'BLDG-A',closestParent:'PARENT',dependencies:'PARENT',dependencyProject:'Old project',equipmentDescription:'Keep description'},{equipmentId:'PARENT',building:'BLDG-B'}];
  const aoa=[EXTO_REV21_COLUMNS.map(c=>c.header),...rows.map(r=>EXTO_REV21_COLUMNS.map(c=>r[c.field]||''))];
  const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
  const bytes=new Uint8Array(XLSX.write(book,{type:'array',bookType:'xlsx'})),snapshot=await auditSnapshotFromWorkbook(XLSX.read(bytes,{type:'array'}),'');
  const changes=[auditMakeCorrection(snapshot.rows[0],'Dependency Project',''),auditMakeCorrection(snapshot.rows[1],'Building','BLDG-A')];
  const restored=XLSX.read(await buildUpdatedRegistryBytes(bytes,snapshot,changes),{type:'array',cellStyles:true});
  for(const change of changes){const cell=restored.Sheets.Registry[addressOf(change)];assert.equal(cell.v,change.value);assert.equal(cell.s.fgColor.rgb,'FFF2CC');}
  const result=await auditSnapshotFromWorkbook(restored,'');assert.equal(result.rows[0].dependencies,'PARENT');assert.equal(result.rows[0].equipmentDescription,'Keep description');assert.equal(result.rows[0].closestParent,'PARENT');
});

packageTest('a workbook without a styles part gets a valid yellow style without losing metadata',async()=>{
  const {bytes,snapshot,change}=syntheticPackage(),entries=packageEntries(bytes);
  entries.delete('xl/styles.xml');
  const rels=xmlDocument(entries.get('xl/_rels/workbook.xml.rels'));
  for(const relation of Array.from(rels.getElementsByTagNameNS('*','Relationship')))if(relation.getAttribute('Type').endsWith('/styles'))relation.parentNode.removeChild(relation);
  entries.set('xl/_rels/workbook.xml.rels',xmlBytes(rels));
  const types=xmlDocument(entries.get('[Content_Types].xml'));
  for(const node of Array.from(types.getElementsByTagNameNS('*','Override')))if(node.getAttribute('PartName')==='/xl/styles.xml')node.parentNode.removeChild(node);
  entries.set('[Content_Types].xml',xmlBytes(types));
  for(const [path,bytes] of entries)if(path.startsWith('xl/worksheets/')&&path.endsWith('.xml')){
    const document=xmlDocument(bytes);for(const cell of Array.from(document.getElementsByTagNameNS('*','c')))cell.removeAttribute('s');entries.set(path,xmlBytes(document));
  }
  const output=await buildUpdatedRegistryBytes(packageBytes(entries),snapshot,[change]);
  const restored=XLSX.read(output,{type:'array',cellStyles:true});
  assert.equal(restored.Sheets.Registry[addressOf(change)].v,change.value);assert.equal(restored.Sheets.Registry[addressOf(change)].s.fgColor.rgb,'FFF2CC');
});

packageTest('updated-registry session export uses the immutable baseline and never downloads a partial correction', async t => {
  const { downloads } = captureExportDownloads(t), { bytes, snapshot, change } = syntheticPackage(), original = bytes.slice()
  S.session.sourceBytes = bytes; S.session.baselineSnapshot = snapshot; S.session.changes = [change]
  S.session.snapshot = { ...snapshot, rows: snapshot.rows.map(row => ({ ...row, systemName: 'Corrected working draft' })) }
  assert.equal(await exportUpdatedRegistryXlsx(), true)
  const output = XLSX.read(await downloads[0].arrayBuffer(), { type: 'array' })
  assert.equal(output.Sheets.Registry[addressOf(change)].v, 'After')
  S.session.changes = [change, { ...correction(snapshot, 1), before: 'Conflict' }]
  assert.equal(await exportUpdatedRegistryXlsx(), false)
  assert.equal(downloads.length, 1, 'no partial second download')
  S.session.baselineSnapshot = null
  assert.equal(await exportUpdatedRegistryXlsx(), false)
  assert.equal(downloads.length, 1, 'a working draft is never substituted for a missing baseline')
  assert.deepEqual(bytes, original)
})
