import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'

import { EXTO_REV21_COLUMNS, extoRev21IsUpn, extoRev21SystemsForUpn } from '../src/exto/rev21-contract.js'
import { auditMergeSnapshots, auditSnapshotFromAoa } from '../src/audit/model.js'
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
  assert.ok(!found.has('parent.cross-discipline'))
  assert.ok(found.has('metadata.ic-discipline'))
  const ic = result.findings.find(finding => finding.rule.id === 'metadata.ic-discipline')
  assert.match(ic.actual, /650/)
  assert.match(ic.recommendation, /UPN 650/)
  assert.ok(ic.why && ic.expected && ic.rule.standardRef && ic.fingerprint)
})

test('multi-sheet imports remove only exact duplicate rows and preserve conflicts', () => {
  const first=auditSnapshotFromAoa([headers,row({equipmentId:'B1-EQ-01',closestParent:'101 System',upn:'101',discipline:'MECHANICAL WET',systemName:'101 System'})],{file:'synthetic.xlsx',sheet:'Registry'})
  const repeated=auditSnapshotFromAoa([headers,row({equipmentId:'B1-EQ-01',closestParent:'101 System',upn:'101',discipline:'MECHANICAL WET',systemName:'101 System'})],{file:'synthetic.xlsx',sheet:'Registry Copy'})
  const conflict=auditSnapshotFromAoa([headers,row({equipmentId:'B1-EQ-01',closestParent:'B1-OTHER',upn:'101',discipline:'MECHANICAL WET',systemName:'101 System'})],{file:'synthetic.xlsx',sheet:'Changed Phase'})
  const merged=auditMergeSnapshots([first,repeated,conflict],'synthetic.xlsx')
  assert.equal(merged.rows.length,2)
  assert.equal(merged.source.ignoredDuplicateRows,1)
  assert.equal(runSsmAudit(merged).findings.filter(finding=>finding.rule.id==='identity.duplicate-equipment-id').length,2)
})

test('cross-discipline controls children are allowed while other crossings remain reviewable', () => {
  const snapshot=auditSnapshotFromAoa([
    headers,
    row({equipmentId:'B1-MAH-01',closestParent:'101 Example System',closestParentStatus:'NEW',upn:'101',discipline:'MECHANICAL DRY',systemName:'101 Example System',equipmentDescription:'Makeup Air Handler'}),
    row({equipmentId:'B1-TET-01',closestParent:'B1-MAH-01',upn:'101',discipline:'FACILITIES MONITORING SYSTEM',systemName:'101 Example System',equipmentDescription:'Temperature Transmitter'}),
    row({equipmentId:'B1-PUMP-01',closestParent:'B1-MAH-01',upn:'101',discipline:'MECHANICAL WET',systemName:'101 Example System',equipmentDescription:'Centrifugal Pump'}),
  ],{file:'synthetic.xlsx',sheet:'Registry'})
  const crossings=runSsmAudit(snapshot).findings.filter(finding=>finding.rule.id==='parent.cross-discipline')
  assert.deepEqual(crossings.map(finding=>finding.equipmentId),['B1-PUMP-01'])
  assert.equal(crossings[0].severity,'warning')
})

test('System Name must be an approved Rev21 name for the row UPN', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    // approved: 602 -> "602  Medium Voltage"
    row({ equipmentId: 'B1-OK', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }),
    // wrong: a real Rev21 name, but for a different UPN
    row({ equipmentId: 'B1-OTHER-UPN', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '104  General Air Handler System (GAH)' }),
    // wrong: not a Rev21 name at all
    row({ equipmentId: 'B1-MADE-UP', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602 Chilled Water System' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(result, 'systemUpn'), ['B1-MADE-UP', 'B1-OTHER-UPN'])
  const other = result.findings.find(f => f.equipmentId === 'B1-OTHER-UPN')
  assert.match(other.why, /different UPN/)
  const madeUp = result.findings.find(f => f.equipmentId === 'B1-MADE-UP')
  assert.match(madeUp.why, /not an approved/)
})

test('letter-code UPNs use their own approved System Names, not a numeric prefix', () => {
  const approved = extoRev21SystemsForUpn('RR')
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'RR-OK', closestParent: 'X', closestParentStatus: 'NEW', upn: 'RR', discipline: 'ROOM/AREA/BAY-READY', systemName: approved[0] || 'RR Placeholder' }),
    row({ equipmentId: 'RR-WRONG', closestParent: 'X', closestParentStatus: 'NEW', upn: 'RR', discipline: 'ROOM/AREA/BAY-READY', systemName: '602  Medium Voltage' }),
    row({ equipmentId: 'BAD-UPN', closestParent: 'X', closestParentStatus: 'NEW', upn: '9999', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.equal(extoRev21IsUpn('RR'), true, 'RR is an approved letter-code UPN')
  assert.equal(extoRev21IsUpn('9999'), false, 'no four-digit UPNs exist in Rev21')
  assert.deepEqual(equipmentIdsForRule(result, 'upnNotApproved'), ['BAD-UPN'])
  const systemFindings = equipmentIdsForRule(result, 'systemUpn')
  assert.ok(systemFindings.includes('RR-WRONG'), 'a numeric-UPN system on a letter UPN is wrong')
  if (approved.length) assert.ok(!systemFindings.includes('RR-OK'), 'an approved name for the letter code passes')
  assert.ok(!systemFindings.includes('BAD-UPN'), 'an unapproved UPN is reported once, as the UPN problem, not as a system mismatch too')
})

test('UPN SEC belongs to Security, in either spelling, and MISC rows always get a second-look note', () => {
  assert.deepEqual(extoRev21SystemsForUpn('SEC'), ['Security Systems', 'Security'])
  assert.deepEqual(extoRev21SystemsForUpn('MISC'), [], 'MISC is deliberately unlinked — it is a catch-all')
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'SEC-FULL', closestParent: 'X', closestParentStatus: 'NEW', upn: 'SEC', discipline: 'SECURITY', systemName: 'Security Systems' }),
    row({ equipmentId: 'SEC-BARE', closestParent: 'X', closestParentStatus: 'NEW', upn: 'SEC', discipline: 'SECURITY', systemName: 'Security' }),
    row({ equipmentId: 'SEC-BAD', closestParent: 'X', closestParentStatus: 'NEW', upn: 'SEC', discipline: 'SECURITY', systemName: '602  Medium Voltage' }),
    row({ equipmentId: 'MISC-1', closestParent: 'X', closestParentStatus: 'NEW', upn: 'MISC', discipline: 'MECHANICAL MISC', systemName: 'MISC' }),
    row({ equipmentId: 'MISC-2', closestParent: 'X', closestParentStatus: 'NEW', upn: 'MISC', discipline: 'MECHANICAL MISC', systemName: 'Anything At All' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const systemFindings = equipmentIdsForRule(result, 'systemUpn')
  assert.ok(!systemFindings.includes('SEC-FULL') && !systemFindings.includes('SEC-BARE'), 'both Security spellings pass on SEC')
  assert.ok(systemFindings.includes('SEC-BAD'), 'a numeric system on SEC still flags')
  assert.ok(!systemFindings.includes('MISC-1') && !systemFindings.includes('MISC-2'), 'MISC rows never get the unwinnable system-name blocker')
  assert.deepEqual(equipmentIdsForRule(result, 'miscUpnReview'), ['MISC-1', 'MISC-2'], 'every MISC row gets the second-look note')
  const note = result.findings.find(f => f.rule.id === 'metadata.misc-upn-review')
  assert.equal(note.severity, 'info')
  assert.match(note.why, /catch-all/)
})

test('UPN RR belongs to the CSA system, in either spelling', () => {
  assert.deepEqual(extoRev21SystemsForUpn('RR'), ['Civil Structural Architectural Systems (CSA)', 'CSA'])
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'RR-LONG', closestParent: 'X', closestParentStatus: 'NEW', upn: 'RR', discipline: 'ROOM/AREA/BAY-READY', systemName: 'Civil Structural Architectural Systems (CSA)' }),
    row({ equipmentId: 'RR-SHORT', closestParent: 'X', closestParentStatus: 'NEW', upn: 'RR', discipline: 'ROOM/AREA/BAY-READY', systemName: 'CSA' }),
    row({ equipmentId: 'RR-BAD', closestParent: 'X', closestParentStatus: 'NEW', upn: 'RR', discipline: 'ROOM/AREA/BAY-READY', systemName: '602  Medium Voltage' }),
    /* CSA on a numeric UPN is still that UPN's problem, not a free pass */
    row({ equipmentId: 'NUM-CSA', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: 'CSA' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const findings = equipmentIdsForRule(runSsmAudit(snapshot), 'systemUpn')
  assert.ok(!findings.includes('RR-LONG'), 'the full CSA name passes on RR')
  assert.ok(!findings.includes('RR-SHORT'), 'the bare CSA code passes on RR')
  assert.ok(findings.includes('RR-BAD'), 'a numeric system on RR still flags')
  assert.ok(findings.includes('NUM-CSA'), 'CSA on a numeric UPN still flags')
})

test('SSM Audit recognizes common Instrumentation and Controls discipline wording', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-RIO650-01', closestParent: '650 Facility Management System', closestParentStatus: 'NEW', upn: '650', discipline: 'Instrumentation and Controls', systemName: '650 Facility Management System', equipmentDescription: 'Remote I/O Panel' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.ok(rules(runSsmAudit(snapshot)).has('metadata.ic-discipline'))
})

test('SSM Audit accepts explicitly external parent and dependency paths', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-RIO650-01', closestParent: 'OTHER-PROJECT-PANEL', closestParentStatus: 'EXISTING', upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 Facility Management System', equipmentDescription: 'Remote I/O Panel', dependencies: 'OTHER-PROJECT-PLC', dependencyProject: 'Other project' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const found = rules(runSsmAudit(snapshot))
  for (const ruleId of ['parent.unresolved','dependency.unresolved','dependency.control-link-missing','dependency.control-electrical-path','dependency.rio-controller-path']) assert.ok(!found.has(ruleId),ruleId)
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
  assert.match(findings[0].why, /2 rows nest under a parent/)
})

test('every audit rule has a neutral source, title, and plain-language statement', () => {
  const sourceIds = new Set(SSM_AUDIT_SOURCES.map(source => source.id))
  assert.deepEqual([...sourceIds], ['registry', 'sop', 'logic'])
  for (const rule of Object.values(SSM_AUDIT_RULES)) {
    assert.ok(sourceIds.has(rule.source), rule.id)
    assert.ok(rule.title && rule.statement, rule.id)
    assert.ok(rule.statement.length <= 320, `${rule.id} statement should stay short: ${rule.statement.length} chars`)
    assert.doesNotMatch(`${rule.title} ${rule.statement} ${rule.standardRef}`, /golden SSM|Exto[- ]Cx|regex|normali[sz]ed|fingerprint/i, rule.id)
  }
})

test('every finding explains itself in plain language, without implementation talk', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PLAIN-A', closestParent: '', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }),
    row({ equipmentId: 'PLAIN-B', closestParent: 'PLAIN-B', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', dependencies: 'PLAIN-B; PLAIN-B; GHOST' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  for (const finding of runSsmAudit(snapshot).findings) {
    assert.ok(finding.why && finding.why.length <= 260, `${finding.rule.id}: why should be one or two short sentences`)
    assert.ok(finding.recommendation, `${finding.rule.id}: every finding says what to do`)
    assert.doesNotMatch(finding.why, /canonical|normali[sz]e|regex|token|fingerprint|snapshot/i, finding.rule.id)
  }
})

test('Item Master standard flags legacy CA_ names and proposes the VF equivalent', () => {
  assert.equal(SSM_AUDIT_RULES.itemMasterStandard.enabled, true)
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B1-CA', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', itemMaster: 'CA_NB_ELEC_PANEL' }),
    row({ equipmentId: 'B1-VF', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', itemMaster: 'VF_ELEC_PANEL' }),
    row({ equipmentId: 'B1-UNKNOWN', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', itemMaster: 'PROJECT99_PANEL' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const withVocabulary = runSsmAudit(snapshot, { itemMasterVocabulary: ['VF_ELEC_PANEL', 'VF_MECH_PUMP'] })
  assert.deepEqual(equipmentIdsForRule(withVocabulary, 'itemMasterStandard'), ['B1-CA', 'B1-UNKNOWN'])
  const ca = withVocabulary.findings.find(f => f.equipmentId === 'B1-CA')
  assert.equal(ca.expected, 'VF_ELEC_PANEL', 'the VF equivalent is proposed by suffix match')
  assert.match(ca.recommendation, /Replace with VF_ELEC_PANEL/)
  assert.match(ca.why, /site-prefixed Item Master \(CA_\)/, 'any site prefix is called out by name, not only CA_')
  assert.equal(ca.severity, 'warning', 'in a mostly-VF registry a stray legacy name is a warning')
  // an explicit empty vocabulary disables the unknown-name check; only site-prefixed legacy names can fire
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot, { itemMasterVocabulary: [] }), 'itemMasterStandard'), ['B1-CA'])
  // by default the shipped VF list is used, so names absent from it — even VF-looking ones — are flagged
  const shipped = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(shipped, 'itemMasterStandard'), ['B1-CA', 'B1-UNKNOWN', 'B1-VF'], 'the invented VF_ELEC_PANEL is not a real Rev14 name, so it is flagged too')
})

test('milestone rules are enabled as review-grade checks', () => {
  for (const ruleKey of ['milestonePair', 'milestoneIntent', 'milestoneInconsistent', 'milestoneCohort', 'milestoneLevel', 'milestoneBranch', 'milestoneUpn']) {
    assert.equal(SSM_AUDIT_RULES[ruleKey].enabled, true, ruleKey)
  }
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PAIR-CONFLICT', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', milestoneParent: 'L1 30 Percent Enabling', milestone: 'L2 602 100 Percent Capacity' }),
    row({ equipmentId: 'MISSING-L1', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', milestone: 'L2 602 Capacity' }),
    row({ equipmentId: 'MISSING-L2', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', milestoneParent: 'L1 Capacity' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(result, 'milestonePair'), ['MISSING-L1', 'MISSING-L2'])
  assert.deepEqual(equipmentIdsForRule(result, 'milestoneIntent'), ['PAIR-CONFLICT'])
  assert.ok(result.findings.filter(f => f.category === 'milestones').every(f => f.severity !== 'blocker' && f.severity !== 'error'), 'milestone checks are review-grade, never blocking')
})

test('the L2 milestone should name the row UPN; a different UPN is flagged for review', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'L2-OWN', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', milestoneParent: 'L1 Energization', milestone: 'L2-M1 602 Medium Voltage Energization' }),
    row({ equipmentId: 'L2-OTHER', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', milestoneParent: 'L1 Energization', milestone: 'L2-M2 604 Normal Power Energization' }),
    row({ equipmentId: 'L2-NONE', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', milestoneParent: 'L1 Energization', milestone: 'L2 Building Ready' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const byId = Object.fromEntries(result.findings.filter(f => f.rule.id === SSM_AUDIT_RULES.milestoneUpn.id).map(f => [f.equipmentId, f]))
  assert.ok(!byId['L2-OWN'], 'an L2 naming the row UPN passes')
  assert.equal(byId['L2-OTHER'].severity, 'warning')
  assert.match(byId['L2-OTHER'].why, /names UPN 604.*on UPN 602/)
  assert.equal(byId['L2-NONE'].severity, 'info', 'an L2 with no UPN in its name cannot be checked, so it is only noted')
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

test('an organizational header is a parented row with a Blank Item Master, not a description keyword', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    // a real header: parented by others, Blank item master
    row({ equipmentId: 'RIO-RING-HEADER', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'ODD CLUSTER RIO RING', itemMaster: 'VF_Blank', dependencies: 'SOME-PANEL' }),
    row({ equipmentId: 'RIO-01', closestParent: 'RIO-RING-HEADER', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Remote IO Panel', itemMaster: 'VF_IC_RIO' }),
    // a Blank-item-master row nobody nests under: unused header
    row({ equipmentId: 'LONELY-HEADER', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Spare grouping', itemMaster: 'VF_Blank' }),
    // real equipment with a Blank item master and NO children is just an unused header, never a physical-equipment finding by description
    // a parented grouping row that forgot the Blank item master
    row({ equipmentId: 'GROUP-NO-BLANK', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'PIPING ROLLUP', itemMaster: 'VF_ELEC_PANEL' }),
    row({ equipmentId: 'PT-1', closestParent: 'GROUP-NO-BLANK', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Pressure Transmitter' }),
    row({ equipmentId: 'PT-2', closestParent: 'GROUP-NO-BLANK', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Pressure Transmitter' }),
    row({ equipmentId: 'PT-3', closestParent: 'GROUP-NO-BLANK', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Pressure Transmitter' }),
    // an AHU with children keeps its real item master and is never called a header
    row({ equipmentId: 'AHU-1', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Air Handler', itemMaster: 'VF_MECH_AHU' }),
    row({ equipmentId: 'TT-1', closestParent: 'AHU-1', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Temperature Transmitter' }),
    row({ equipmentId: 'TT-2', closestParent: 'AHU-1', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Temperature Transmitter' }),
    row({ equipmentId: 'TT-3', closestParent: 'AHU-1', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Temperature Transmitter' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual([...result.headerIds].sort(), ['RIO-RING-HEADER'], 'only parented Blank-item-master rows are headers')
  assert.deepEqual(equipmentIdsForRule(result, 'headerDependency'), ['RIO-RING-HEADER'])
  assert.deepEqual(equipmentIdsForRule(result, 'unusedHeader'), ['LONELY-HEADER'])
  assert.deepEqual(equipmentIdsForRule(result, 'headerItemMaster'), ['GROUP-NO-BLANK'], 'a grouping row without a Blank item master is flagged; the AHU is not')
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

test('rows on one UPN must agree on System Name and Discipline', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'UPN-A', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }),
    row({ equipmentId: 'UPN-B', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'MECHANICAL DRY', systemName: '602  Medium Voltage' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.equal(result.findings.filter(f => f.rule.id === SSM_AUDIT_RULES.upnInconsistent.id).length, 1)
})

test('drives and motor starters are parented by the driven equipment they serve', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PUMP-01', closestParent: '104  General Air Handler System (GAH)', closestParentStatus: 'NEW', upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)', equipmentDescription: 'Centrifugal Pump' }),
    row({ equipmentId: 'VFD-HEALTHY', closestParent: 'PUMP-01', upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)', equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'STARTER-HEALTHY', closestParent: 'PUMP-01', upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)', equipmentDescription: 'Motor Starter' }),
    row({ equipmentId: 'COMPRESSOR-01', closestParent: '104  General Air Handler System (GAH)', closestParentStatus: 'NEW', upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)', equipmentDescription: 'Chiller Compressor W/VFD' }),
    row({ equipmentId: 'VFD-COMPRESSOR', closestParent: 'COMPRESSOR-01', upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)', equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'VFD-WRONG-PARENT', closestParent: '104  General Air Handler System (GAH)', closestParentStatus: 'NEW', upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)', equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'STARTER-WRONG-PARENT', closestParent: '104  General Air Handler System (GAH)', closestParentStatus: 'NEW', upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)', equipmentDescription: 'Motor Starter' }),
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
  // Real coupling only: calls into profile learning or the hierarchy build. Prose
  // that merely PROMISES the registry is never stored in a site profile is fine.
  assert.doesNotMatch(ui, /persistProfiles\(|learnItemMaster|learnNesting|buildHierarchy\(|S\.files\b|S\.extoSel/)
  assert.match(ui, /S\.(?:audit[Ss]ession|session)\b/, 'the audit UI keeps its own session slot')
})

/* ---- SOP nesting conventions (the tag's UPN is the guiderail) ---- */

const MECH = { upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)' }
const MECH_ROOT = '104  General Air Handler System (GAH)'

test('an instrument nests under equipment in the UPN its tag carries', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'B14-AHU-104-01', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Air Handler' }),
    row({ equipmentId: 'B14-TT-104-01', closestParent: 'B14-AHU-104-01', ...MECH, equipmentDescription: 'Temperature Transmitter' }),
    row({ equipmentId: 'B14-TT-104-02', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Temperature Transmitter' }),
    row({ equipmentId: 'B14-PMP-101-01', closestParent: '101  Cleanroom Makeup Air System', closestParentStatus: 'NEW', upn: '101', discipline: 'MECHANICAL DRY', systemName: '101  Cleanroom Makeup Air System', equipmentDescription: 'Pump' }),
    // row UPN 104 (matches its tag) but nested under UPN 101 equipment: a real cross-UPN nesting
    row({ equipmentId: 'B14-PT-104-03', closestParent: 'B14-PMP-101-01', ...MECH, equipmentDescription: 'Pressure Transmitter' }),
    // row UPN 101 agrees with its parent; only the TAG says 104: a tag question, not a nesting error
    row({ equipmentId: 'B14-PT-104-04', closestParent: 'B14-PMP-101-01', upn: '101', discipline: 'MECHANICAL DRY', systemName: '101  Cleanroom Makeup Air System', equipmentDescription: 'Pressure Transmitter' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const findings = result.findings.filter(f => f.rule.id === SSM_AUDIT_RULES.instrumentUpn.id)
  const byId = Object.fromEntries(findings.map(f => [f.equipmentId, f]))
  assert.ok(!byId['B14-TT-104-01'], 'nested under equipment in its own UPN: fine')
  assert.equal(byId['B14-TT-104-02'].severity, 'warning', 'sitting on the System Name instead of equipment')
  assert.equal(byId['B14-PT-104-03'].severity, 'error', 'nested under equipment in a different UPN than the tag and the row')
  assert.match(byId['B14-PT-104-03'].why, /carries UPN 104.*on UPN 101/)
  assert.equal(byId['B14-PT-104-04'].severity, 'info', 'row and parent agree; only the tag disagrees — a tag question')
  assert.match(byId['B14-PT-104-04'].why, /tag may be wrong/)
})

test('a VFD lists its electrical panel and its PLC as dependencies', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'AHU-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Air Handler' }),
    row({ equipmentId: 'PNL-1', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Distribution Panel' }),
    row({ equipmentId: 'PLC-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'PLC' }),
    row({ equipmentId: 'VFD-COMPLETE', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'Variable Frequency Drive', dependencies: 'PNL-1; PLC-1' }),
    row({ equipmentId: 'VFD-NO-PLC', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'Variable Frequency Drive', dependencies: 'PNL-1' }),
    row({ equipmentId: 'VFD-NOTHING', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'Variable Frequency Drive' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const byId = Object.fromEntries(result.findings.filter(f => f.rule.id === SSM_AUDIT_RULES.vfdDependencies.id).map(f => [f.equipmentId, f]))
  assert.ok(!byId['VFD-COMPLETE'])
  assert.match(byId['VFD-NO-PLC'].why, /No PLC is listed/)
  assert.match(byId['VFD-NOTHING'].why, /Neither is listed/)
})

test('FMS hardwired I/O nests under its VFD with the PLC as a dependency', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'AHU-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Air Handler' }),
    row({ equipmentId: 'VFD-1', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'Variable Frequency Drive' }),
    row({ equipmentId: 'PLC-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'PLC' }),
    row({ equipmentId: 'IO-GOOD', closestParent: 'VFD-1', ...MECH, equipmentDescription: 'FMS Hardwired I/O', dependencies: 'PLC-1' }),
    row({ equipmentId: 'IO-WRONG-PARENT', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'FMS Hardwired I/O', dependencies: 'PLC-1' }),
    row({ equipmentId: 'IO-NO-PLC', closestParent: 'VFD-1', ...MECH, equipmentDescription: 'FMS Hardwired I/O' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const byId = Object.fromEntries(result.findings.filter(f => f.rule.id === SSM_AUDIT_RULES.fmsIoUnderVfd.id).map(f => [f.equipmentId, f]))
  assert.ok(!byId['IO-GOOD'])
  assert.match(byId['IO-WRONG-PARENT'].why, /not nested under a VFD/)
  assert.match(byId['IO-NO-PLC'].why, /does not list its PLC/)
})

test('an LCP sits with its equipment or skid in its own UPN (MAH > SKID > LCP)', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'MAH-101', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Makeup Air Handler' }),
    row({ equipmentId: 'SKD-101', closestParent: 'MAH-101', ...MECH, equipmentDescription: 'Skid' }),
    row({ equipmentId: 'LCP-101', closestParent: 'SKD-101', ...MECH, equipmentDescription: 'Local Control Panel' }),
    row({ equipmentId: 'LCP-ROOT', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Local Control Panel' }),
    row({ equipmentId: 'PMP-OTHER', closestParent: '101  Cleanroom Makeup Air System', closestParentStatus: 'NEW', upn: '101', discipline: 'MECHANICAL DRY', systemName: '101  Cleanroom Makeup Air System', equipmentDescription: 'Pump' }),
    row({ equipmentId: 'LCP-CROSS', closestParent: 'PMP-OTHER', ...MECH, equipmentDescription: 'Local Control Panel' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(result, 'lcpPlacement'), ['LCP-CROSS'], 'only the LCP under another UPN is flagged; nested and root LCPs pass')
})

test('control valves and room sensors nest under equipment, not the system or room', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'AHU-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Air Handler' }),
    row({ equipmentId: 'TCV-GOOD', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'Temperature Control Valve' }),
    row({ equipmentId: 'TCV-LOOSE', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Temperature Control Valve' }),
    row({ equipmentId: 'RT-GOOD', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'Room Temperature Sensor' }),
    row({ equipmentId: 'RT-LOOSE', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Room Temperature Sensor' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(result, 'controlValveParent'), ['TCV-LOOSE'])
  assert.deepEqual(equipmentIdsForRule(result, 'roomSensorParent'), ['RT-LOOSE'])
})

test('a loose instrument may sit under a piping roll-up header only in its own UPN', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PIPE-104', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Distribution Piping', itemMaster: 'VF_Blank' }),
    row({ equipmentId: 'PIPE-101', closestParent: '101  Cleanroom Makeup Air System', closestParentStatus: 'NEW', upn: '101', discipline: 'MECHANICAL DRY', systemName: '101  Cleanroom Makeup Air System', equipmentDescription: 'Distribution Piping', itemMaster: 'VF_Blank' }),
    row({ equipmentId: 'B14-PT-104-09', closestParent: 'PIPE-104', ...MECH, equipmentDescription: 'Pressure Transmitter' }),
    row({ equipmentId: 'B14-PT-104-10', closestParent: 'PIPE-101', upn: '101', discipline: 'MECHANICAL DRY', systemName: '101  Cleanroom Makeup Air System', equipmentDescription: 'Pressure Transmitter' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(result, 'untiedInstrumentRollup'), ['B14-PT-104-10'])
  assert.ok(!equipmentIdsForRule(result, 'instrumentUpn').includes('B14-PT-104-09'), 'a roll-up header in the right UPN is an accepted parent for an instrument')
})

test('same-UPN dependencies are normal in top-down disciplines and on instruments; elsewhere they are noted', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PNL-A', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Panel' }),
    row({ equipmentId: 'PNL-B', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Panel', dependencies: 'PNL-A' }),
    row({ equipmentId: 'PMP-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Pump' }),
    row({ equipmentId: 'PMP-2', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Pump', dependencies: 'PMP-1' }),
    row({ equipmentId: 'PT-1', closestParent: 'PMP-1', ...MECH, equipmentDescription: 'Pressure Transmitter', dependencies: 'PMP-2' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(result, 'sameUpnBottomUp'), ['PMP-2'], 'electrical (top-down) and the instrument pass; the mechanical pump-to-pump dependency is noted')
  const finding = result.findings.find(f => f.equipmentId === 'PMP-2' && f.rule.id === SSM_AUDIT_RULES.sameUpnBottomUp.id)
  assert.equal(finding.severity, 'info')
})


test('a registry that is wholesale on a site item-master scheme is a migration, not thousands of warnings', () => {
  const rows = [headers]
  for (let i = 0; i < 60; i++) rows.push(row({ equipmentId: `B1-PNL-${String(i).padStart(3, '0')}`, closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', itemMaster: 'SP_NB_ELEC_PANEL' }))
  const result = runSsmAudit(auditSnapshotFromAoa(rows, { file: 'synthetic.xlsx', sheet: 'Registry' }))
  const findings = result.findings.filter(f => f.rule.id === SSM_AUDIT_RULES.itemMasterStandard.id)
  assert.equal(findings.length, 60, 'every row still gets its migration note')
  assert.ok(findings.every(f => f.severity === 'info'), 'but as info, so structural findings stay on top')
  assert.match(findings[0].why, /site-prefixed Item Master \(SP_\)/)
})


/* ---- checks added after auditing the auditor against a real registry ---- */

test('electrical path reaches a child through its parent feed', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PNL-1', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Panel' }),
    row({ equipmentId: 'AHU-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Air Handler', dependencies: 'PNL-1' }),
    row({ equipmentId: 'FFU-1', closestParent: 'AHU-1', ...MECH, equipmentDescription: 'Fan Filter Unit' }),
    row({ equipmentId: 'PMP-LOOSE', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Pump' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  assert.deepEqual(equipmentIdsForRule(runSsmAudit(snapshot), 'drivenElectricalPath'), ['PMP-LOOSE'], 'the FFU is powered through its parent; only the loose pump has no path')
})

test('a parent duplicated as a dependency is a note; a dependency on a header is a warning', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'GIS-1', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'GIS' }),
    row({ equipmentId: 'XFM-1', closestParent: 'GIS-1', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Transformer', dependencies: 'GIS-1' }),
    row({ equipmentId: 'HDR-1', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Ring header', itemMaster: 'VF_Blank' }),
    row({ equipmentId: 'RIO-1', closestParent: 'HDR-1', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Remote IO Panel' }),
    row({ equipmentId: 'PNL-2', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', equipmentDescription: 'Panel', dependencies: 'HDR-1' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const dup = result.findings.find(f => f.rule.id === SSM_AUDIT_RULES.parentAlsoDependency.id)
  assert.equal(dup.equipmentId, 'XFM-1'); assert.equal(dup.severity, 'info')
  assert.match(dup.expected, /Common in electrical/)
  assert.deepEqual(equipmentIdsForRule(result, 'dependencyOnHeader'), ['PNL-2'])
})

test('a parent in another building is a rule break; a stale Dependency Project is a note', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'PNL-B14', closestParent: '602  Medium Voltage', closestParentStatus: 'NEW', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', building: 'B14', equipmentDescription: 'Panel' }),
    row({ equipmentId: 'PNL-B31', closestParent: 'PNL-B14', upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage', building: 'B31', equipmentDescription: 'Panel', dependencies: 'PNL-B14', dependencyProject: 'OTHER-PROJECT' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  assert.deepEqual(equipmentIdsForRule(result, 'crossBuildingParent'), ['PNL-B31'])
  assert.match(result.findings.find(f => f.rule.id === SSM_AUDIT_RULES.crossBuildingParent.id).why, /in B31 but its parent is in B14/)
  assert.deepEqual(equipmentIdsForRule(result, 'staleDependencyProject'), ['PNL-B31'])
})

test('a system with no root row, a description typed as a tag, and a site classification are each surfaced', () => {
  const snapshot = auditSnapshotFromAoa([
    headers,
    row({ equipmentId: 'AHU-1', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Air Handler', equipmentClassification: 'AH' }),
    // every row on 101 nests under 104 equipment: 101 has no top
    row({ equipmentId: 'FCU-1', closestParent: 'AHU-1', upn: '101', discipline: 'MECHANICAL DRY', systemName: '101  Cleanroom Makeup Air System', equipmentDescription: 'Fan Coil', equipmentClassification: 'ZZ-SITE' }),
    row({ equipmentId: 'Distribution piping east wing', closestParent: MECH_ROOT, closestParentStatus: 'NEW', ...MECH, equipmentDescription: 'Piping', itemMaster: 'VF_Blank' }),
  ], { file: 'synthetic.xlsx', sheet: 'Registry' })
  const result = runSsmAudit(snapshot)
  const noRoot = result.findings.find(f => f.rule.id === SSM_AUDIT_RULES.systemNoRoot.id)
  assert.ok(noRoot && /101/.test(noRoot.actual), 'the 101 system has no root row')
  assert.deepEqual(equipmentIdsForRule(result, 'tagLooksLikeText'), ['Distribution piping east wing'])
  assert.deepEqual(equipmentIdsForRule(result, 'siteClassification'), ['FCU-1'])
  assert.equal(result.findings.find(f => f.rule.id === SSM_AUDIT_RULES.siteClassification.id).severity, 'info')
})
