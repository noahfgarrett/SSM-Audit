import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'

import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { auditCyclePaths, runSsmAudit, SSM_AUDIT_RULES, SSM_AUDIT_SOURCES } from '../src/audit/engine.js'

const headers = EXTO_REV21_COLUMNS.map(column => column.header)
const index = Object.fromEntries(EXTO_REV21_COLUMNS.map(column => [column.field, column.index]))

function row(values) {
  const cells = new Array(headers.length).fill('')
  for (const [field, value] of Object.entries(values)) cells[index[field]] = value
  return cells
}

function rules(result) {
  return new Set(result.findings.map(finding => finding.rule.id))
}

test('SSM Audit detects the official header row beyond the first line', () => {
  const snapshot = auditSnapshotFromAoa([
    ['Project export generated locally'],
    headers,
    row({ equipmentId: 'B1-GIS-01', closestParent: '602  Medium Voltage', upn: '602', discipline: 'ELECTRICAL' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.ok(snapshot)
  assert.equal(snapshot.headerRow, 2)
  assert.equal(snapshot.rows[0]._source.row, 3)
  assert.equal(snapshot.rows[0]._source.file, 'synthetic.xlsx')
})

test('SSM Audit preserves physical Excel row numbers after blank rows are compacted', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-GIS-01', closestParent: '602  Medium Voltage', upn: '602', discipline: 'ELECTRICAL' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry', rowNums: [1, 3] })
  assert.equal(snapshot.headerRow, 2)
  assert.equal(snapshot.rows[0]._source.row, 4)
})

test('cycle detection handles deeply nested hierarchies without recursion', () => {
  const count = 25000, nodes = Array.from({ length: count }, (_, index) => `N${index}`), edges = new Map()
  for (let index = 0; index < count - 1; index++) edges.set(nodes[index], new Set([nodes[index + 1]]))
  edges.set(nodes[count - 1], new Set([nodes[0]]))
  const cycles = auditCyclePaths(nodes, edges)
  assert.equal(cycles.length, 1)
  assert.equal(cycles[0].length, count + 1)
})

test('SSM Audit produces explainable structural and dependency findings', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-GIS-01', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }),
    row({ equipmentId: 'B1-XFM-01', closestParent: 'B1-GIS-01', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', dependencies: 'B1-GIS-01' }),
    row({ equipmentId: 'B1-LVS-01', closestParent: 'B1-LVS-01', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', dependencies: 'N/A' }),
    row({ equipmentId: 'B1-RIO650-01', closestParent: 'B1-GIS-01', upn: '650', discipline: 'I&C', systemName: '650  Facility Management System' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const found = rules(result)
  assert.ok(found.has('dependency.repeats-parent'))
  assert.ok(found.has('dependency.literal-na'))
  assert.ok(found.has('parent.self'))
  assert.ok(found.has('parent.cross-upn'))
  assert.ok(found.has('parent.cross-discipline'))
  assert.ok(found.has('metadata.ic-discipline'))
  const ic = result.findings.find(finding => finding.rule.id === 'metadata.ic-discipline')
  assert.match(ic.actual, /650/)
  assert.match(ic.recommendation, /UPN 650/)
  assert.ok(ic.why && ic.expected && ic.rule.standardRef && ic.fingerprint)
})

test('SSM Audit reports missing upload columns instead of silently guessing', () => {
  const shortHeaders = ['UPN', 'Discipline', 'Equipment ID', 'Closest Parent']
  const snapshot = auditSnapshotFromAoa([shortHeaders, ['602', 'ELECTRICAL', 'B1-GIS-01', '602  Medium Voltage']], { sheet: 'Partial' })
  const result = runSsmAudit(snapshot)
  const finding = result.findings.find(item => item.rule.id === 'upload.missing-header' && item.severity === 'blocker')
  assert.ok(finding)
  assert.equal(finding.sheet, 'Partial')
  assert.equal(finding.row, 1)
})

test('repeated New-header references collapse to one actionable finding', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-A', closestParent: 'SYSTEM HEADER', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }),
    row({ equipmentId: 'B1-B', closestParent: 'SYSTEM HEADER', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }),
  ], { sheet: 'Registry' })
  const findings = runSsmAudit(snapshot).findings.filter(finding => finding.rule.id === 'parent.generated-header-review')
  assert.equal(findings.length, 1)
  assert.match(findings[0].why, /2 equipment rows/)
})

test('every audit rule has a neutral source, title, and concise statement', () => {
  const sourceIds = new Set(SSM_AUDIT_SOURCES.map(source => source.id))
  assert.deepEqual([...sourceIds], ['upload', 'sop', 'logic'])
  for (const rule of Object.values(SSM_AUDIT_RULES)) {
    assert.ok(sourceIds.has(rule.source), rule.id)
    assert.ok(rule.title && rule.statement, rule.id)
    assert.doesNotMatch(`${rule.title} ${rule.statement} ${rule.standardRef}`, /Rev21|golden SSM|Exto[- ]Cx/i)
  }
})

test('commissioning logic distinguishes healthy paths from missing links', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'BLD-PNL-01', closestParent: '602 Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602 Medium Voltage', equipmentDescription: 'Electrical Panel' }),
    row({ equipmentId: 'BLD-VFD-01', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', dependencies: 'BLD-PNL-01', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'BLD-PUMP-OK', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', dependencies: 'BLD-VFD-01', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Centrifugal Pump' }),
    row({ equipmentId: 'BLD-PUMP-MISSING', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Centrifugal Pump' }),
    row({ equipmentId: 'BLD-RIO-OK', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', dependencies: 'BLD-PNL-01', upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 Facility Management System', equipmentDescription: 'Remote I/O Panel' }),
    row({ equipmentId: 'BLD-RIO-MISSING', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 Facility Management System', equipmentDescription: 'Remote I/O Panel' }),
    row({ equipmentId: 'BLD-CS-OK', closestParent: 'BLD-PUMP-OK', dependencies: 'BLD-RIO-OK', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Current Switch' }),
    row({ equipmentId: 'BLD-CS-MISSING', closestParent: 'BLD-PUMP-OK', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Current Switch' }),
    row({ equipmentId: '1820 Mechanical System - Pumps', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Centrifugal Pump System', itemMaster: 'VF_Blank' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const byRule = ruleId => result.findings.filter(finding => finding.rule.id === ruleId).map(finding => finding.equipmentId)
  assert.deepEqual(byRule('logic.driven-electrical-path-missing'), ['BLD-PUMP-MISSING'])
  assert.deepEqual(byRule('logic.control-electrical-path-missing'), ['BLD-RIO-MISSING'])
  assert.deepEqual(byRule('logic.control-link-missing'), ['BLD-CS-MISSING'])
  assert.equal(result.summary.source.logic, 3)
})

test('the audit path is isolated from profile learning and hierarchy imports', () => {
  const model = readFileSync(new URL('../src/audit/model.js', import.meta.url), 'utf8')
  const engine = readFileSync(new URL('../src/audit/engine.js', import.meta.url), 'utf8')
  const ui = readFileSync(new URL('../src/ui/audit.js', import.meta.url), 'utf8')
  assert.doesNotMatch(model + engine, /from ['"]\.\.\/(?:compiler|hierarchy|profile|rules)\//)
  assert.doesNotMatch(model + engine, /from ['"]\.\.\/state\.js|persistProfiles|learnItemMaster|learnNesting|buildHierarchy/)
  assert.doesNotMatch(ui, /persistProfiles|learnItemMaster|learnNesting|buildHierarchy|S\.files|S\.extoSel|Site Profile/)
  assert.match(ui, /S\.session/)
})
