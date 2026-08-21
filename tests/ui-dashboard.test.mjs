import test from 'node:test'
import assert from 'node:assert/strict'

import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit, SSM_AUDIT_RULES } from '../src/audit/engine.js'
import { S, resetSession } from '../src/state.js'
import { activeRules, applyRulePreferences, dashCheckOverview, dashDependencyStats, dashHierarchyHealth, dashMilestoneReadiness, isRuleActive, scopedResult } from '../src/ui/audit.js'

const headers = EXTO_REV21_COLUMNS.map(column => column.header)
const index = Object.fromEntries(EXTO_REV21_COLUMNS.map(column => [column.field, column.index]))

function row(values) {
  const cells = new Array(headers.length).fill('')
  for (const [field, value] of Object.entries(values)) cells[index[field]] = value
  return cells
}

const MV = '602  Medium Voltage', EXH = '130 Process Exhaust and Environmental Testing'

/* Two systems in two buildings: a four-deep electrical chain in FAB1 and a
   two-row exhaust system in CUB1, plus a parent loop that must not hang the
   chain walk. */
const FIXTURE = [
  headers,
  row({ building: 'FAB1', upn: '602', discipline: 'ELECTRICAL', systemName: MV, equipmentId: 'FAB1-602-MCC-01', equipmentDescription: 'MOTOR CONTROL CENTER', closestParent: MV, itemMaster: 'VF_EL_LV_SWGR', equipmentClassification: 'MCC', milestone: 'P1', milestoneParent: 'L1' }),
  row({ building: 'FAB1', upn: '602', discipline: 'ELECTRICAL', systemName: MV, equipmentId: 'FAB1-602-PNL-01', equipmentDescription: 'DISTRIBUTION PANEL', closestParent: 'FAB1-602-MCC-01', itemMaster: 'VF_EL_PANEL', equipmentClassification: 'PNL', milestone: 'P1', milestoneParent: 'L1', dependencies: 'FAB1-602-MCC-01' }),
  row({ building: 'FAB1', upn: '602', discipline: 'ELECTRICAL', systemName: MV, equipmentId: 'FAB1-602-PNL-02', equipmentDescription: 'DISTRIBUTION PANEL', closestParent: 'FAB1-602-PNL-01', itemMaster: 'VF_EL_PANEL', equipmentClassification: 'PNL', milestone: 'P1', milestoneParent: 'L1' }),
  row({ building: 'FAB1', upn: '602', discipline: 'ELECTRICAL', systemName: MV, equipmentId: 'FAB1-602-BUS-01', equipmentDescription: 'BUSWAY RUN', closestParent: 'FAB1-602-PNL-02', itemMaster: 'VF_EL_LV_CKT', equipmentClassification: 'BUS', milestone: 'P1', milestoneParent: 'L1', dependencies: 'CUB1-130-FAN-01' }),
  row({ building: 'CUB1', upn: '130', discipline: 'MECHANICAL EXHAUST', systemName: EXH, equipmentId: 'CUB1-130-FAN-01', equipmentDescription: 'EXHAUST FAN', closestParent: EXH, itemMaster: 'VF_MD_EXHAUST FAN', equipmentClassification: 'FAN', milestone: 'P2', milestoneParent: 'L1' }),
  row({ building: 'CUB1', upn: '130', discipline: 'MECHANICAL EXHAUST', systemName: EXH, equipmentId: 'CUB1-130-FAN-02', equipmentDescription: 'EXHAUST FAN', closestParent: 'CUB1-130-FAN-01', itemMaster: 'SP_FAN', equipmentClassification: 'FAN', dependencies: 'OFFSITE-FARM-01', dependencyProject: 'PHASE 1' }),
  row({ building: 'CUB1', upn: '130', discipline: 'MECHANICAL EXHAUST', systemName: EXH, equipmentId: 'CUB1-130-LOOP-A', equipmentDescription: 'ISOLATION DAMPER', closestParent: 'CUB1-130-LOOP-B', itemMaster: 'SP_DMP', equipmentClassification: 'FAN', milestone: 'P2', milestoneParent: 'L1' }),
  row({ building: 'CUB1', upn: '130', discipline: 'MECHANICAL EXHAUST', systemName: EXH, equipmentId: 'CUB1-130-LOOP-B', equipmentDescription: 'ISOLATION DAMPER', closestParent: 'CUB1-130-LOOP-A', itemMaster: 'SP_DMP', equipmentClassification: 'FAN', milestone: 'P2', milestoneParent: 'L1' }),
]

function loadFixture() {
  resetSession()
  const snapshot = auditSnapshotFromAoa(FIXTURE, { file: 'synthetic.xlsx', sheet: 'Registry' })
  S.session.snapshot = snapshot
  S.session.result = runSsmAudit(snapshot)
  return S.session.result
}

test('a dimension filter narrows the scoped rows; a level filter leaves them alone', () => {
  const result = loadFixture()
  assert.equal(scopedResult().rows.length, result.rows.length)

  S.session.dimFilters.building = ['FAB1']
  S.session.scopedCacheKey = ''
  S.session.filteredCacheKey = ''
  const byBuilding = scopedResult()
  assert.equal(byBuilding.rows.length, 4)
  assert.ok(byBuilding.rows.every(row => row.building === 'FAB1'))
  assert.ok(byBuilding.findings.every(finding => finding.equipmentId.startsWith('FAB1')))

  /* Hiding a level removes findings but must not remove rows: a row with
     nothing left to show is still part of the registry being looked at. */
  S.session.dimFilters.building = []
  S.session.hiddenSeverities = ['blocker', 'error', 'warning', 'info']
  S.session.scopedCacheKey = ''
  S.session.filteredCacheKey = ''
  const noFindings = scopedResult()
  assert.equal(noFindings.findings.length, 0)
  assert.equal(noFindings.rows.length, result.rows.length)
})

test('milestone readiness counts every row once and reports clean rows per phase', () => {
  const result = loadFixture()
  const buckets = dashMilestoneReadiness(scopedResult())
  assert.equal(buckets.reduce((total, bucket) => total + bucket.rows, 0), result.rows.length)

  const unassigned = buckets.find(bucket => bucket.label === 'No L2 milestone')
  assert.ok(unassigned, 'rows without an L2 milestone get their own bucket')
  assert.equal(unassigned.rows, 1)

  for (const bucket of buckets) {
    assert.ok(bucket.clean <= bucket.rows)
    assert.equal(bucket.findings, bucket.blocker + bucket.error + bucket.warning + bucket.info)
  }
})

test('hierarchy health measures depth from the top of a system and survives a parent loop', () => {
  loadFixture()
  const stats = dashHierarchyHealth(scopedResult())
  /* FAB1-602-BUS-01 sits three levels under FAB1-602-MCC-01, which is a root. */
  assert.equal(stats.maxDepth, 3)
  assert.equal(stats.deep, 0)
  assert.equal(stats.roots, 2)
  assert.ok(stats.cleanBranchPercent >= 0 && stats.cleanBranchPercent <= 100)
})

test('dependency references are resolved against the rows in scope', () => {
  loadFixture()
  const stats = dashDependencyStats(scopedResult())
  assert.equal(stats.withDependencies, 3)
  assert.equal(stats.references, 3)
  assert.equal(stats.sameUpn, 1, 'FAB1-602-PNL-01 depends on its own UPN')
  assert.equal(stats.crossUpn, 1, 'FAB1-602-BUS-01 depends on UPN 130')
  assert.equal(stats.external, 1, 'the unresolvable tag names a Dependency Project')

  /* Scoping to one building leaves the cross-UPN target outside the scope, so
     that reference stops resolving locally. */
  S.session.dimFilters.building = ['FAB1']
  S.session.scopedCacheKey = ''
  S.session.filteredCacheKey = ''
  const scopedStats = dashDependencyStats(scopedResult())
  assert.equal(scopedStats.crossUpn, 0)
  assert.equal(scopedStats.sameUpn, 1)
})

test('the checks overview reports every enabled check, firing or not', () => {
  loadFixture()
  const overview = dashCheckOverview(scopedResult())
  const enabled = Object.values(SSM_AUDIT_RULES).filter(rule => rule.enabled).length
  const listed = overview.groups.reduce((total, group) => total + group.firing.length + group.passing.length, 0)
  assert.equal(listed, enabled)
  assert.equal(overview.total, enabled)
  assert.ok(overview.firing > 0 && overview.firing < enabled)
  assert.ok(overview.groups.every(group => group.firing.every(entry => entry.count > 0)))
  assert.ok(overview.groups.every(group => group.passing.every(entry => entry.count === 0 && entry.severity === '')))
})

test('switching a check off removes its findings, recounts the summary, and drops it from the checks overview', () => {
  const raw = S.session.result || (() => { throw new Error('fixture session expected') })()
  const fired = raw.findings[0].rule
  const before = raw.findings.filter(finding => finding.rule.id === fired.id).length
  assert.ok(before > 0)
  const filtered = applyRulePreferences(raw, [fired.id])
  assert.equal(filtered.findings.filter(finding => finding.rule.id === fired.id).length, 0)
  assert.equal(filtered.summary.findings, raw.summary.findings - before)
  const levels = ['blocker', 'error', 'warning', 'info']
  assert.equal(levels.reduce((sum, level) => sum + filtered.summary.severity[level], 0), filtered.findings.length, 'severity counts add up')
  assert.equal(applyRulePreferences(raw, []), raw, 'nothing switched off returns the engine result untouched')
  S.rules.disabled = [fired.id]
  try {
    assert.equal(isRuleActive(fired), false)
    assert.equal(activeRules().some(rule => rule.id === fired.id), false)
    const overview = dashCheckOverview(scopedResult())
    assert.ok(!overview.groups.some(group => [...group.firing, ...group.passing].some(entry => entry.rule.id === fired.id)), 'the overview no longer lists it')
  } finally { S.rules.disabled = [] }
})
