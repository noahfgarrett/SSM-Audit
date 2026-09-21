import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'

import { buildAuditActionsWorkbook, buildAuditTrackerWorkbook, buildAuditWorkbook } from '../src/audit/export.js'
import { auditCoverageRows, SSM_AUDIT_ENGINEERING_RULES } from '../src/audit/engineering-references.js'

const vendor = readFileSync(new URL('../src/vendor/sheetjs.js', import.meta.url), 'utf8')
vm.runInThisContext(vendor, { filename: 'sheetjs.js' })

function grid(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true })
}

function syntheticResult() {
  const registryRow = {
    equipmentId: 'REG-001',
    building: 'B01',
    discipline: 'MECHANICAL',
    upn: '1820',
    systemName: '1820 Mechanical System',
    equipmentDescription: 'Supply fan',
    equipmentClassification: 'Fan',
    closestParent: 'SYS-01',
    milestone: 'L2-M1-101 Supply Fans',
    _source: { sheet: 'Registry', row: 2 },
  }
  const missingRule = SSM_AUDIT_ENGINEERING_RULES.missing
  const findings = [
    {
      id: 'finding-registry', equipmentId: 'REG-001', severity: 'warning',
      rule: { id: 'synthetic.registry-check', title: 'Registry check', statement: 'Synthetic rule', source: 'sop', confidence: 'required' },
      field: 'Discipline', actual: 'MECHANICAL', expected: 'MECHANICAL', why: 'Synthetic registry finding.',
      recommendation: 'Review the synthetic registry finding.', sheet: 'Registry', row: 2,
    },
    {
      id: 'finding-source-electrical', equipmentId: 'SRC-101', severity: 'missing', rule: missingRule,
      field: 'Equipment ID', actual: 'Not found in registry', expected: 'SRC-101', why: 'Synthetic source-only tag.',
      recommendation: 'Confirm synthetic source scope.', sheet: 'MEL', row: 7,
      sourceContext: {
        building: 'B02', discipline: 'ELECTRICAL', upn: '1820', systemName: '1820 Electrical System',
        equipmentDescription: 'Smoke exhaust fan', equipmentClassification: 'Fan', closestParent: 'SYS-E1',
      },
    },
    {
      id: 'finding-source-unassigned', equipmentId: 'SRC-102', severity: 'missing', rule: missingRule,
      field: 'Equipment ID', actual: 'Not found in registry', expected: 'SRC-102', why: 'Synthetic source-only tag with no discipline.',
      recommendation: 'Confirm synthetic source scope.', sheet: 'MEL', row: 8,
      sourceContext: {
        building: 'B03', discipline: '', upn: '1820', systemName: '1820 Mechanical System',
        equipmentDescription: 'Relief fan', equipmentClassification: 'Fan', closestParent: '',
      },
    },
    {
      id: 'finding-source-conflict-a', equipmentId: 'SRC-103', severity: 'missing', rule: missingRule,
      field: 'Equipment ID', actual: 'Not found in registry', expected: 'SRC-103', why: 'Synthetic source context conflict A.',
      recommendation: 'Confirm synthetic source scope.', sheet: 'MEL', row: 9,
      sourceContext: {
        building: 'B04', discipline: 'CONTROLS', upn: '1820', systemName: '1820 Fire System',
        equipmentDescription: 'Fire damper actuator', equipmentClassification: 'Actuator', closestParent: 'SYS-F1',
      },
    },
    {
      id: 'finding-source-conflict-b', equipmentId: 'SRC-103', severity: 'missing', rule: missingRule,
      field: 'Equipment ID', actual: 'Not found in registry', expected: 'SRC-103', why: 'Synthetic source context conflict B.',
      recommendation: 'Confirm synthetic source scope.', sheet: 'MEL', row: 9,
      sourceContext: {
        building: 'B04', discipline: 'FIRE ALARM', upn: '1820', systemName: '1820 Fire System',
        equipmentDescription: 'Smoke damper actuator', equipmentClassification: 'Actuator', closestParent: 'SYS-F1',
      },
    },
  ]
  return {
    rows: [registryRow], findings,
    summary: { rows: 1, findings: findings.length, severity: { warning: 1, missing: 4 } },
    standard: 'Synthetic audit',
  }
}

function disciplineRows(book) {
  const rows = grid(book.Sheets.Dashboard)
  const header = rows.findIndex(line => line[0] === 'Discipline')
  const end = rows.findIndex((line, at) => at > header && ['Milestone', 'Finding level'].includes(line[0]))
  return new Map(rows.slice(header + 1, end).filter(line => line[0]).map(line => [line[0], { equipment: line[1], findings: line[2] }]))
}

function indexGroups(book) {
  return grid(book.Sheets.Index).slice(4).filter(line => line[0] && line[0] !== 'No equipment rows in this registry')
}

test('milestone reports include source-only tags and aggregate blank disciplines explicitly', () => {
  const result = syntheticResult(), originalRows = result.rows, beforeRows = structuredClone(result.rows)
  const beforeFindings = structuredClone(result.findings)
  const coverage = auditCoverageRows(result)
  assert.equal(coverage.length, 3)
  assert.ok(coverage.every(row => row._coverageOnly))
  assert.deepEqual(coverage.map(row => row.discipline), ['ELECTRICAL', '', ''])
  const conflicted = coverage.find(row => row.equipmentId === 'SRC-103')
  assert.equal(conflicted.equipmentDescription, '')

  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx')
  const noMilestone = grid(book.Sheets['No milestone']).slice(2)
  assert.deepEqual(noMilestone.map(line => line[2]).filter(Boolean), ['SRC-101', 'SRC-102', 'SRC-103'])
  assert.deepEqual(noMilestone.map(line => line[6]).filter(Boolean), ['ELECTRICAL'])

  const groups = indexGroups(book)
  assert.equal(groups.reduce((total, line) => total + line[2], 0), result.findings.length)
  assert.deepEqual(groups.find(line => line[0] === 'No milestone').slice(1, 3), [3, 4])
  assert.deepEqual(disciplineRows(book), new Map([
    ['ELECTRICAL', { equipment: 1, findings: 1 }],
    ['MECHANICAL', { equipment: 1, findings: 1 }],
    ['No discipline', { equipment: 2, findings: 3 }],
  ]))

  const calcHeaders = grid(book.Sheets.Calc)[0]
  const unassignedColumn = XLSX.utils.encode_col(calcHeaders.indexOf('No discipline'))
  const unassignedMilestoneRow = grid(book.Sheets.Calc).findIndex(line => line[0] === 'No milestone') + 1
  assert.equal(book.Sheets.Calc[`${unassignedColumn}${unassignedMilestoneRow}`].f,
    `COUNTIFS('No milestone'!G:G,"",'No milestone'!A:A,"☑")`)

  assert.strictEqual(result.rows, originalRows)
  assert.deepEqual(result.rows, beforeRows)
  assert.deepEqual(result.findings, beforeFindings)
  assert.equal(result.rows.some(row => row._coverageOnly), false)
  assert.deepEqual(result.rows.map(row => row.equipmentId), ['REG-001'])
})

test('level reports retain each row milestone in All Findings and count every finding', () => {
  const result = syntheticResult()
  const book = buildAuditWorkbook(result, 'synthetic-registry.xlsx', { layout: 'level' })
  const levelTabs = book.SheetNames.filter(name => !['Dashboard', 'Index', 'All Findings', 'Rules', 'Calc'].includes(name))
  assert.deepEqual(levelTabs, ['MISSING TAGS', 'CHECK THIS'])

  const allFindings = grid(book.Sheets['All Findings']).slice(2)
  assert.deepEqual(allFindings.map(line => [line[2], line[1]]), [
    ['REG-001', 'L2-M1-101 Supply Fans'],
    ['SRC-101', 'No milestone'],
    ['SRC-102', 'No milestone'],
    ['SRC-103', 'No milestone'],
    ['SRC-103', 'No milestone'],
  ])

  const missingRows = grid(book.Sheets['MISSING TAGS']).slice(2)
  assert.deepEqual(missingRows.map(line => line[2]).filter(Boolean), ['SRC-101', 'SRC-102', 'SRC-103'])
  assert.ok(missingRows.every(line => !line[11] || line[11] === 'No milestone'))
  assert.equal(indexGroups(book).reduce((total, line) => total + line[2], 0), result.findings.length)
  assert.deepEqual(disciplineRows(book), new Map([
    ['ELECTRICAL', { equipment: 1, findings: 1 }],
    ['MECHANICAL', { equipment: 1, findings: 1 }],
    ['No discipline', { equipment: 2, findings: 3 }],
  ]))
})

test('Actions and Tracker use coverage context without adding rows to the registry result', () => {
  const result = syntheticResult(), beforeRows = structuredClone(result.rows)
  const actions = buildAuditActionsWorkbook(result, 'synthetic-registry.xlsx')
  const missingSheet = actions.Sheets[SSM_AUDIT_ENGINEERING_RULES.missing.title.slice(0, 31).trim()]
  const missingRows = grid(missingSheet).slice(5)
  const electrical = missingRows.find(line => line[3] === 'SRC-101')
  const unassigned = missingRows.find(line => line[3] === 'SRC-102')
  assert.equal(electrical[14], 'ELECTRICAL')
  assert.equal(electrical[4], 'Smoke exhaust fan')
  assert.equal(unassigned[14], '')
  assert.equal(unassigned[4], 'Relief fan')

  const tracker = buildAuditTrackerWorkbook(result, 'synthetic-registry.xlsx', { signOffBy: 'discipline' })
  const trackerRows = grid(tracker.Sheets.Tracker).slice(8).filter(line => line[0])
  assert.deepEqual(new Map(trackerRows.map(line => [line[0], line[1]])), new Map([
    ['ELECTRICAL', 1], ['MECHANICAL', 1], ['No discipline', 3],
  ]))
  assert.deepEqual(result.rows, beforeRows)
  assert.equal(result.rows.some(row => row._coverageOnly), false)
})
