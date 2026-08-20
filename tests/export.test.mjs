import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'

import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { AUDIT_EXPORT_TICK, auditExportNestLevels, auditExportOrderRows, auditExportSheetName, buildAuditWorkbook } from '../src/audit/export.js'

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
    assert.match(book.Sheets.Index[`F${indexRow}`].f, /^REPT\("█",MIN\(10,ROUND\(E\d+\*10,0\)\)\)&REPT\("░",10-MIN\(10,ROUND\(E\d+\*10,0\)\)\)$/)
    assert.match(book.Sheets.Dashboard[`F${dashboardRow}`].f, /^REPT\("█",MIN\(25,/, 'dashboard bars are the wide 25-segment kind')
  })
  const last = first + tabs.length - 1
  assert.equal(book.Sheets.Dashboard.F6.f, `IF(SUM(B${first}:B${last})=0,0,SUM(D${first}:D${last})/SUM(B${first}:B${last}))`)
  assert.match(book.Sheets.Dashboard.A6.f, /^REPT\("█",MIN\(25,ROUND\(F6\*25,0\)\)\)/, 'the overall bar is the big one at the top')
  assert.ok(book.Sheets.Dashboard.A6.s.font.sz >= 18, 'overall bar uses a large font')
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
    assert.equal(xml.includes('<conditionalFormatting'), isMilestone, name)
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
