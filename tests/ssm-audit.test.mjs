import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'

import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { auditCyclePaths, auditItemMasterCanonicalCandidates, auditMilestoneBranchCandidates, auditMilestoneCohortCandidates, auditMilestoneLevelIssues, runSsmAudit, SSM_AUDIT_RULES, SSM_AUDIT_SOURCES } from '../src/audit/engine.js'

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

function equipmentIdsForRule(result, ruleKey) {
  const rule = SSM_AUDIT_RULES[ruleKey]
  assert.ok(rule, `Missing audit rule metadata: ${ruleKey}`)
  return result.findings
    .filter(finding => finding.rule.id === rule.id)
    .map(finding => finding.equipmentId)
    .sort()
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
  assert.ok(!found.has('dependency.repeats-parent'))
  assert.ok(!found.has('dependency.literal-na'))
  assert.ok(found.has('parent.self'))
  assert.ok(found.has('parent.cross-upn'))
  assert.ok(found.has('parent.cross-discipline'))
  assert.ok(found.has('metadata.ic-discipline'))
  const ic = result.findings.find(finding => finding.rule.id === 'metadata.ic-discipline')
  assert.match(ic.actual, /650/)
  assert.match(ic.recommendation, /UPN 650/)
  assert.ok(ic.why && ic.expected && ic.rule.standardRef && ic.fingerprint)
})

test('post-upload audit omits import completeness and vocabulary rules', () => {
  for (const ruleKey of ['missingHeader', 'missingGating', 'invalidDropdown', 'literalNa', 'parentAsDependency']) {
    assert.equal(SSM_AUDIT_RULES[ruleKey], undefined, `${ruleKey} should not be part of a post-upload audit`)
  }

  const shortHeaders = ['UPN', 'Discipline', 'Equipment ID', 'Closest Parent']
  const snapshot = auditSnapshotFromAoa([shortHeaders, ['602', 'ELECTRICAL', 'B1-GIS-01', '602  Medium Voltage']], { sheet: 'Partial' })
  const result = runSsmAudit(snapshot)
  assert.ok(!rules(result).has('upload.missing-header'))
})

test('post-upload audit accepts blanks, project vocabulary, N/A, and a repeated parent dependency', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-PARENT', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Project-specific equipment', equipmentClassification: 'LOCAL' }),
    row({ equipmentId: 'B1-CHILD', closestParent: 'B1-PARENT', dependencies: 'B1-PARENT; N/A', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Another local description', equipmentClassification: 'SITE' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const found = rules(runSsmAudit(snapshot))
  for (const ruleId of ['upload.missing-gating-field', 'metadata.invalid-dropdown-value', 'dependency.literal-na', 'dependency.repeats-parent']) {
    assert.ok(!found.has(ruleId), `${ruleId} should not be emitted by a post-upload audit`)
  }
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
  assert.deepEqual([...sourceIds], ['registry', 'sop', 'logic'])
  for (const rule of Object.values(SSM_AUDIT_RULES)) {
    assert.ok(sourceIds.has(rule.source), rule.id)
    assert.ok(rule.title && rule.statement, rule.id)
    assert.doesNotMatch(`${rule.title} ${rule.statement} ${rule.standardRef}`, /Rev21|golden SSM|Exto[- ]Cx/i)
  }
})

test('Item Master standard remains visible but disabled and emits no prefix-based findings', () => {
  assert.ok(SSM_AUDIT_RULES.itemMasterStandard)
  assert.equal(SSM_AUDIT_RULES.itemMasterStandard.enabled, false)

  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-CA', closestParent: '602 Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602 Medium Voltage', itemMaster: 'CA_LEGACY_PANEL' }),
    row({ equipmentId: 'B1-SP', closestParent: '602 Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602 Medium Voltage', itemMaster: 'SP_LEGACY_PANEL' }),
    row({ equipmentId: 'B1-EA', closestParent: '602 Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602 Medium Voltage', itemMaster: 'EA_LEGACY_PANEL' }),
    row({ equipmentId: 'B1-CUSTOM', closestParent: '602 Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602 Medium Voltage', itemMaster: 'PROJECT99_LEGACY_PANEL' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.equal(result.findings.filter(finding => finding.rule.id === SSM_AUDIT_RULES.itemMasterStandard.id).length, 0)
  assert.equal(result.findings.filter(finding => finding.rule.id.startsWith('item-master.')).length, 0)
})

test('milestone rules remain visible but disabled and emit no findings', () => {
  for (const ruleKey of ['milestonePair', 'milestoneIntent', 'milestoneInconsistent', 'milestoneCohort', 'milestoneLevel', 'milestoneBranch']) {
    assert.equal(SSM_AUDIT_RULES[ruleKey].enabled, false)
  }
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PAIR-CONFLICT', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', milestoneParent: 'L1 30 Percent Enabling', milestone: 'L2 100 Percent Capacity' }),
    row({ equipmentId: 'MISSING-L1', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', milestone: 'L2 Capacity' }),
    row({ equipmentId: 'MISSING-L2', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', milestoneParent: 'L1 Capacity' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.equal(result.findings.filter(finding => finding.category === 'milestones').length, 0)
})

test('milestone cohort candidates require ten strongly agreeing local peers', () => {
  const values = Array.from({ length: 10 }, (_, index) => row({ equipmentId: `COHORT-${index}`, closestParent: 'SYSTEM', building: 'BLD1', upn: '402', discipline: 'MECHANICAL DRY', equipmentDescription: 'Purge Panel', milestoneParent: 'L1-EXPECTED', milestone: 'L2-EXPECTED' }))
  values.push(row({ equipmentId: 'COHORT-OUTLIER', closestParent: 'SYSTEM', building: 'BLD1', upn: '402', discipline: 'MECHANICAL DRY', equipmentDescription: 'Purge Panel', milestoneParent: 'L1-OTHER', milestone: 'L2-OTHER' }))
  const snapshot = auditSnapshotFromAoa([headers, ...values], { sheet: 'Registry' })
  const candidates = auditMilestoneCohortCandidates(snapshot.rows)
  assert.deepEqual(candidates.map(candidate => candidate.row.equipmentId), ['COHORT-OUTLIER'])
  assert.equal(candidates[0].agreementCount, 10)
  assert.equal(candidates[0].expectedMilestone, 'L2-EXPECTED')
})

test('milestone level candidates recognize only explicit identifiers in the wrong field', () => {
  const issues = auditMilestoneLevelIssues({ milestoneParent: 'SP-L2-M1-1418 100% Blowdown', milestone: 'SP-L1-M1-133 100% Ready' })
  assert.deepEqual(issues.map(issue => issue.field), ['L1 Milestone Parent', 'L2 Milestone'])
  assert.deepEqual(auditMilestoneLevelIssues({ milestoneParent: 'Local readiness gate', milestone: 'Capacity complete' }), [])
})

test('milestone branch candidates require the parent and five siblings to agree', () => {
  const common = { building: 'BLD1', upn: '1820', discipline: 'MECHANICAL DRY', milestoneParent: 'L1-EXPECTED', milestone: 'L2-EXPECTED' }
  const values = [row({ equipmentId: 'BRANCH-PARENT', closestParent: 'SYSTEM', ...common })]
  for (let index = 0; index < 5; index++) values.push(row({ equipmentId: `BRANCH-PEER-${index}`, closestParent: 'BRANCH-PARENT', ...common }))
  values.push(row({ equipmentId: 'BRANCH-OUTLIER', closestParent: 'BRANCH-PARENT', ...common, milestoneParent: 'L1-OTHER', milestone: 'L2-OTHER' }))
  const snapshot = auditSnapshotFromAoa([headers, ...values], { sheet: 'Registry' })
  const candidates = auditMilestoneBranchCandidates(snapshot.rows)
  assert.deepEqual(candidates.map(candidate => candidate.row.equipmentId), ['BRANCH-OUTLIER'])
  assert.equal(candidates[0].agreementCount, 5)
})

test('Item Master candidates ignore arbitrary prefixes and preserve ambiguity', () => {
  const catalog = ['VF1_PUMP_CENTRIFUGAL', 'VF2_PUMP_CENTRIFUGAL', 'VF1_AIR_HANDLER']
  assert.deepEqual(auditItemMasterCanonicalCandidates('EA_SITE_PUMP_CENTRIFUGAL', catalog), ['VF1_PUMP_CENTRIFUGAL', 'VF2_PUMP_CENTRIFUGAL'])
  assert.deepEqual(auditItemMasterCanonicalCandidates('PROJECT99_AIR_HANDLER', catalog), ['VF1_AIR_HANDLER'])
  assert.deepEqual(auditItemMasterCanonicalCandidates('EA_EXHAUST_FAN', catalog), [])
})

test('HEADER description defines an organizational header while VF_Blank alone does not', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PHYSICAL-VF-BLANK', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Centrifugal Pump', itemMaster: 'VF_Blank' }),
    row({ equipmentId: 'REAL-HEADER', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'HEADER', itemMaster: 'SP_GROUPING' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const headerFindings = result.findings.filter(finding => finding.rule.id.startsWith('header.'))
  assert.ok(headerFindings.some(finding => finding.rule.id === SSM_AUDIT_RULES.unusedHeader.id && finding.equipmentId === 'REAL-HEADER'))
  assert.ok(!headerFindings.some(finding => finding.equipmentId === 'PHYSICAL-VF-BLANK'))
  assert.deepEqual(equipmentIdsForRule(result, 'drivenElectricalPath'), ['PHYSICAL-VF-BLANK'])
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
    row({ equipmentId: '1820 Mechanical System - Pumps', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'HEADER', itemMaster: 'VF_Blank' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const byRule = ruleId => result.findings.filter(finding => finding.rule.id === ruleId).map(finding => finding.equipmentId)
  assert.deepEqual(byRule('logic.driven-electrical-path-missing'), ['BLD-PUMP-MISSING'])
  assert.deepEqual(byRule('logic.control-electrical-path-missing'), ['BLD-RIO-MISSING'])
  assert.deepEqual(byRule('logic.control-link-missing'), ['BLD-CS-MISSING'])
})

test('RIO panels require a controller path, not only an electrical dependency', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PLC-01', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 Facility Management System', equipmentDescription: 'Programmable Logic Controller' }),
    row({ equipmentId: 'PNL-01', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', upn: '650', discipline: 'ELECTRICAL', systemName: '650 Facility Management System', equipmentDescription: 'Electrical Panel' }),
    row({ equipmentId: 'RIO-HEALTHY', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', dependencies: 'PLC-01; PNL-01', upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 Facility Management System', equipmentDescription: 'Remote I/O Panel' }),
    row({ equipmentId: 'RIO-NO-CONTROLLER', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', dependencies: 'PNL-01', upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 Facility Management System', equipmentDescription: 'Remote I/O Panel' }),
    row({ equipmentId: 'PLC-RIO-COMBINED', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', dependencies: 'PNL-01', upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 Facility Management System', equipmentDescription: 'FMS PLC/RIO Cabinet' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot), 'rioControlPath'), ['RIO-NO-CONTROLLER'])
})

test('System Name consistency supports three- and four-digit UPNs', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'UPN-OK', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System' }),
    row({ equipmentId: 'UPN-WRONG', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1821 Different System' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot), 'systemUpn'), ['UPN-WRONG'])
})

test('drives and motor starters are parented by the driven equipment they serve', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PUMP-01', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Centrifugal Pump' }),
    row({ equipmentId: 'VFD-HEALTHY', closestParent: 'PUMP-01', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'STARTER-HEALTHY', closestParent: 'PUMP-01', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Motor Starter' }),
    row({ equipmentId: 'COMPRESSOR-01', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Chiller Compressor W/VFD' }),
    row({ equipmentId: 'VFD-COMPRESSOR', closestParent: 'COMPRESSOR-01', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'VFD-WRONG-PARENT', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'STARTER-WRONG-PARENT', closestParent: '1820 Mechanical System', closestParentStatus: 'NEW', upn: '1820', discipline: 'MECHANICAL DRY', systemName: '1820 Mechanical System', equipmentDescription: 'Motor Starter' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot), 'driveParent'), ['STARTER-WRONG-PARENT', 'VFD-WRONG-PARENT'])
})

test('heat-trace panels and junction boxes maintain the transformer-to-panel chain', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'XFM-01', closestParent: '610 Heat Trace', closestParentStatus: 'NEW', upn: '610', discipline: 'ELECTRICAL', systemName: '610 Heat Trace', equipmentDescription: 'Transformer' }),
    row({ equipmentId: 'HTP-HEALTHY', closestParent: 'XFM-01', upn: '610', discipline: 'ELECTRICAL', systemName: '610 Heat Trace', equipmentDescription: 'Heat Trace Panel' }),
    row({ equipmentId: 'HTJB-HEALTHY', closestParent: 'HTP-HEALTHY', upn: '610', discipline: 'ELECTRICAL', systemName: '610 Heat Trace', equipmentDescription: 'Heat Trace Power Connection' }),
    row({ equipmentId: 'HTP-NO-XFM', closestParent: '610 Heat Trace', closestParentStatus: 'NEW', upn: '610', discipline: 'ELECTRICAL', systemName: '610 Heat Trace', equipmentDescription: 'Heat Trace Panel' }),
    row({ equipmentId: 'HTJB-NO-PANEL', closestParent: '610 Heat Trace', closestParentStatus: 'NEW', upn: '610', discipline: 'ELECTRICAL', systemName: '610 Heat Trace', equipmentDescription: 'Heat Trace Power Connection' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot), 'heatTraceChain'), ['HTJB-NO-PANEL', 'HTP-NO-XFM'])
})

test('FDU equipment identifies a commissioning dependency', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'FAP-01', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'Fire Alarm Panel' }),
    row({ equipmentId: 'FDU-HEALTHY', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', dependencies: 'FAP-01', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'Fiber Optic Distribution Unit' }),
    row({ equipmentId: 'FDU-NO-DEPENDENCY', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'Fiber Optic Distribution Unit' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot), 'fduDependency'), ['FDU-NO-DEPENDENCY'])
})

test('VESDA systems identify the associated fire alarm panel as a dependency', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'FAP-01', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'Fire Alarm Panel' }),
    row({ equipmentId: 'FDU-01', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', dependencies: 'FAP-01', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'Fiber Optic Distribution Unit' }),
    row({ equipmentId: 'VESDA-HEALTHY', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', dependencies: 'FAP-01', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'VESDA System' }),
    row({ equipmentId: 'VESDA-WRONG-DEPENDENCY', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', dependencies: 'FDU-01', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'Very Early Smoke Detection Apparatus (VESDA)' }),
    row({ equipmentId: 'VESDA-NO-DEPENDENCY', closestParent: '630 Life Safety System', closestParentStatus: 'NEW', upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630 Life Safety System', equipmentDescription: 'VESDA System' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot), 'vesdaFireAlarm'), ['VESDA-NO-DEPENDENCY', 'VESDA-WRONG-DEPENDENCY'])
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
