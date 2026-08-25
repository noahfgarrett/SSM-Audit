import test from 'node:test'
import assert from 'node:assert/strict'

import { auditStatusCompleted, auditStatusSheetName } from '../src/audit/status-report.js'
import { applyCompletedEquipment } from '../src/ui/audit.js'
import { auditNormId, auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'

const HEADERS = ['Equipment Name', 'Area', 'RR OA/BT', 'DIST OA/BT', 'EQ OA/BT', 'SYS OA/BT']

test('the Equipment Status Report tab is found by name, forgiving of spacing and case', () => {
  assert.equal(auditStatusSheetName(['Full Export', 'Equipment Status Report']), 'Equipment Status Report')
  assert.equal(auditStatusSheetName(['Full Export', 'equipment  status report ']), 'equipment  status report ')
  assert.equal(auditStatusSheetName(['Full Export', 'Summary']), '')
})

test('equipment counts as completed only when a step says Completed and none says Not Started', () => {
  const status = auditStatusCompleted([
    ['Some banner row'],
    HEADERS,
    ['B1-AHU-1', 'FAB', 'Completed', '', '', ''],
    ['B1-AHU-2', 'FAB', 'Not Started', '', '', ''],
    ['B1-AHU-3', 'FAB', 'In Progress', '', '', ''],
    ['B1-AHU-4', 'FAB', 'Completed', 'Not Started', '', ''],
    ['B1-AHU-5', 'FAB', 'Completed', 'In Progress', '', ''],
    ['B1-AHU-6', 'FAB', '', '', '', 'completed'],
    ['', 'FAB', 'Completed', '', '', ''],
  ])
  assert.equal(status.totalRows, 6)
  assert.equal(status.completedRows, 3)
  assert.deepEqual([...status.completed].sort(), ['B1-AHU-1', 'B1-AHU-5', 'B1-AHU-6'].map(auditNormId).sort())
  assert.ok(!status.completed.has(auditNormId('B1-AHU-4')), 'Completed beside Not Started is treated as not started')
})

test('a sheet without the expected headers yields nothing instead of guessing', () => {
  assert.equal(auditStatusCompleted([['Equipment', 'Status'], ['B1-AHU-1', 'Completed']]), null)
  assert.equal(auditStatusCompleted(null), null)
})

test('completed equipment leaves the metrics entirely while registry-wide findings stay', () => {
  const headers = EXTO_REV21_COLUMNS.map(column => column.header)
  const index = Object.fromEntries(EXTO_REV21_COLUMNS.map(column => [column.field, column.index]))
  const row = values => {
    const cells = new Array(headers.length).fill('')
    for (const [field, value] of Object.entries(values)) cells[index[field]] = value
    return cells
  }
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'DONE-PMP', closestParent: 'MISSING-PARENT', closestParentStatus: 'NEW', upn: '111', discipline: 'MECHANICAL WET', systemName: '111  Chilled Water (R/S)' }),
    row({ equipmentId: 'OPEN-PMP', closestParent: 'ALSO-MISSING', closestParentStatus: 'NEW', upn: '111', discipline: 'MECHANICAL WET', systemName: '111  Chilled Water (R/S)' }),
  ], { file: 'synthetic.xlsx', sheet: 'Full Export' })
  const raw = runSsmAudit(snapshot)
  const doneBefore = raw.findings.filter(f => f.equipmentId === 'DONE-PMP').length
  assert.ok(doneBefore > 0, 'the completed equipment fires findings before filtering')
  const trimmed = applyCompletedEquipment(raw, new Set([auditNormId('DONE-PMP')]))
  assert.equal(trimmed.findings.filter(f => f.equipmentId === 'DONE-PMP').length, 0)
  assert.ok(trimmed.findings.filter(f => f.equipmentId === 'OPEN-PMP').length > 0, 'other equipment keeps its findings')
  assert.equal(trimmed.summary.findings, trimmed.findings.length)
  const levels = ['blocker', 'error', 'warning', 'info']
  assert.equal(levels.reduce((sum, level) => sum + trimmed.summary.severity[level], 0), trimmed.findings.length)
  assert.equal(applyCompletedEquipment(raw, new Set()), raw, 'no completed equipment returns the result untouched')
})

test('each completed equipment records which OA/BT step said Completed', () => {
  const status = auditStatusCompleted([
    HEADERS,
    ['B1-AHU-1', 'FAB', 'Completed', '', '', ''],
    ['B1-AHU-5', 'FAB', 'Completed', 'In Progress', '', ''],
    ['B1-AHU-6', 'FAB', '', '', '', 'completed'],
    ['B1-AHU-6', 'FAB', '', 'Completed', '', ''],
  ])
  const byName = Object.fromEntries(status.equipment.map(entry => [entry.name, entry.steps]))
  assert.deepEqual(byName['B1-AHU-1'], ['RR OA/BT'])
  assert.deepEqual(byName['B1-AHU-5'], ['RR OA/BT'], 'In Progress columns are not listed as completed steps')
  assert.deepEqual(byName['B1-AHU-6'].sort(), ['DIST OA/BT', 'SYS OA/BT'], 'duplicate rows merge their steps')
  assert.equal(status.completedRows, 3, 'completed count is distinct equipment, not rows')
})
