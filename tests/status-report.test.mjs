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

test('every listed tag is completed regardless of step statuses or conflicting cells', () => {
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
  assert.equal(status.completedRows, 6)
  assert.deepEqual([...status.completed].sort(), Array.from({length:6},(_,i)=>`B1-AHU-${i+1}`))
  assert.ok(status.completed.has(auditNormId('B1-AHU-4')), 'step statuses never override membership in the completed-only report')
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

test('duplicate tags merge by normalized identity without step-based classification', () => {
  const status = auditStatusCompleted([
    HEADERS,
    ['B1-AHU-1', 'FAB', 'Completed', '', '', ''],
    ['B1-AHU-5', 'FAB', 'Completed', 'In Progress', '', ''],
    ['B1-AHU-6', 'FAB', '', '', '', 'completed'],
    [' b1-ahu-6 ', 'FAB', '', 'Not Started', '', ''],
  ])
  assert.deepEqual(status.equipment,[{name:'B1-AHU-1'},{name:'B1-AHU-5'},{name:'B1-AHU-6'}])
  assert.equal(status.completedRows, 3, 'completed count is distinct equipment, not rows')
  assert.equal(status.totalRows,4)
})

for(const header of ['Equipment ID','Equipment Tag','Equipment Name','Equipment Tag/ID','Equipment Tag / ID'])test(`completed-only report accepts ${header} without status columns`,()=>{
  const status=auditStatusCompleted([['Report banner'],[header],['TAG-1'],[''],[null],['TAG-2'],['TAG-2-A']]);
  assert.deepEqual([...status.completed],['TAG-1','TAG-2','TAG-2-A']);
  assert.equal(status.totalRows,3);
});

test('explicit equipment identity takes precedence over descriptive names',()=>{
  const status=auditStatusCompleted([['Equipment Name','Equipment ID','Other Completion Step'],['Pump','DEMO-PUMP-1','Approved'],['Second pump','DEMO-PUMP-2',''],['No ID','','Completed']]);
  assert.deepEqual([...status.completed],['DEMO-PUMP-1','DEMO-PUMP-2']);
  assert.equal(status.completed.has('PUMP'),false);
});
