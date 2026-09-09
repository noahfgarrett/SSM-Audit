import test from 'node:test'
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { basename, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import * as model from '../src/audit/model.js'
import * as engine from '../src/audit/engine.js'

const mech = { upn: '104', discipline: 'MECHANICAL DRY', systemName: '104  General Air Handler System (GAH)' }
const electrical = { upn: '602', discipline: 'ELECTRICAL', systemName: '602  Medium Voltage' }
const controls = { upn: '650', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650  Facility Management System' }
const fire = { upn: '630', discipline: 'LIFE SAFETY SYSTEM', systemName: '630  Life Safety System' }
const headers = EXTO_REV21_COLUMNS.map(column => column.header)
function equipment(equipmentId, values = {}, system = mech) {
  return { ...system, closestParent: system.systemName, closestParentStatus: 'NEW', equipmentId, ...values }
}
function snapshot(rows, sheet = 'Registry') {
  return model.auditSnapshotFromAoa([headers, ...rows.map(row => EXTO_REV21_COLUMNS.map(column => row[column.field] || ''))], { file: 'synthetic.xlsx', sheet })
}
function findings(result, key) {
  return result.findings.filter(finding => finding.rule.id === engine.SSM_AUDIT_RULES[key].id)
}
function tags(result, key) {
  return findings(result, key).map(finding => finding.equipmentId).sort()
}
const panel = equipment('PNL-1', { equipmentDescription: 'Electrical Panel' }, electrical)

test('source identity is exactly the sheet and physical row, independent of tag and file', () => {
  assert.equal(model.auditSourceKey({}), '["",0]')
  assert.equal(model.auditSourceKey({ _source: { sheet: 'A|B', row: 12, file: 'one.xlsx' } }), '["A|B",12]')
  assert.equal(model.auditSourceKey({ equipmentId: 'FIRST', _source: { sheet: 'Registry', row: 2 } }), model.auditSourceKey({ equipmentId: 'RENAMED', _source: { sheet: 'Registry', row: 2 } }))
  assert.notEqual(model.auditSourceKey({ _source: { sheet: 'A', row: 12 } }), model.auditSourceKey({ _source: { sheet: 'A1', row: 2 } }))
})

test('same-sheet identical rows and their physical order survive a single-sheet merge', () => {
  const repeated = equipment('PMP-1', { equipmentDescription: 'Pump' })
  const source = snapshot([repeated, equipment('OTHER'), repeated])
  const merged = model.auditMergeSnapshots([source], 'synthetic.xlsx')
  assert.deepEqual(merged.rows.map(model.auditSourceKey), ['["Registry",2]', '["Registry",3]', '["Registry",4]'])
  assert.equal(merged.source.ignoredDuplicateRows, 0)
  const duplicates = findings(engine.runSsmAudit(merged), 'duplicateId')
  assert.deepEqual(duplicates.map(finding => finding.row), [2, 4])
  assert.equal(new Set(duplicates.map(finding => finding.id)).size, 2)
})

test('summary copies retain the largest per-sheet duplicate group and every source location', () => {
  const repeated = equipment('PMP-1', { equipmentDescription: 'Pump' })
  const summary = snapshot([repeated], 'Summary')
  const registry = snapshot([repeated, repeated, repeated], 'Registry')
  const copy = snapshot([repeated, repeated], 'Summary Copy')
  for (const sources of [[summary, registry, copy], [registry, copy, summary], [copy, summary, registry]]) {
    const merged = model.auditMergeSnapshots(sources, 'synthetic.xlsx')
    assert.equal(merged.rows.length, 3)
    assert.equal(merged.source.ignoredDuplicateRows, 3)
    assert.deepEqual(merged.rows.map(model.auditSourceKey), ['["Registry",2]', '["Registry",3]', '["Registry",4]'])
    assert.equal(findings(engine.runSsmAudit(merged), 'duplicateId').length, 3)
    assert.deepEqual(merged.rows.flatMap(row => row._sources.map(_source => model.auditSourceKey({ _source }))).sort(), sources.flatMap(source => source.rows.map(model.auditSourceKey)).sort())
    assert.ok(merged.rows.every(row => Object.isFrozen(row) && Object.isFrozen(row._sources) && row._sources.every(Object.isFrozen)))
  }
  assert.ok(registry.rows.every(row => !('_sources' in row)), 'the original snapshots remain unchanged')
})

test('case differences, field boundaries, and unknown sheet identities are not silently deduplicated', () => {
  const first = equipment('PMP-1', { equipmentDescription: 'Pump', building: 'A\u001fB' })
  const second = { ...first, equipmentDescription: 'PUMP' }
  const merged = model.auditMergeSnapshots([snapshot([first], 'Registry'), snapshot([second], 'Summary')], 'synthetic.xlsx')
  assert.equal(merged.rows.length, 2)
  assert.equal(findings(engine.runSsmAudit(merged), 'duplicateId').length, 2)
  const boundaryA = { ...first, site: 'A\u001fB', building: 'C' }
  const boundaryB = { ...first, site: 'A', building: 'B\u001fC' }
  assert.equal(model.auditMergeSnapshots([snapshot([boundaryA], 'Registry'), snapshot([boundaryB], 'Summary')], 'synthetic.xlsx').rows.length, 2, 'field separators inside values cannot create signature collisions')
  const unknown = model.auditMergeSnapshots([snapshot([first], ''), snapshot([first], '')], 'synthetic.xlsx')
  assert.equal(unknown.rows.length, 2, 'without distinct sheets there is no evidence of a summary copy')
})

test('one fed pump cannot power its sibling through a mechanical or electrical Blank header', () => {
  for (const headerSystem of [mech, electrical]) {
    const rows = [panel,
      equipment('Pump group', { itemMaster: 'VF_Blank', equipmentDescription: 'Pump header' }, headerSystem),
      equipment('PMP-1', { closestParent: 'Pump group', equipmentDescription: 'Pump', dependencies: 'PNL-1' }),
      equipment('PMP-2', { closestParent: 'Pump group', equipmentDescription: 'Pump' }),
    ]
    for (const ordered of [rows, [...rows].reverse()]) {
      const result = engine.runSsmAudit(snapshot(ordered))
      assert.deepEqual(tags(result, 'drivenElectricalPath'), ['PMP-2'])
      assert.deepEqual(tags(result, 'tagLooksLikeText'), [])
    }
  }
})

test('headers neither relay a parent feed nor contribute their own power or control dependencies', () => {
  const result = engine.runSsmAudit(snapshot([
    panel, equipment('PLC-1', { equipmentDescription: 'PLC' }, controls),
    equipment('AHU-1', { equipmentDescription: 'Air Handler', dependencies: 'PNL-1; PLC-1' }),
    equipment('Support group', { closestParent: 'AHU-1', itemMaster: 'VF_Blank', equipmentDescription: 'Control panel', dependencies: 'PNL-1; PLC-1' }),
    equipment('PMP-1', { closestParent: 'Support group', equipmentDescription: 'Pump' }),
    equipment('RIO-1', { closestParent: 'Support group', equipmentDescription: 'Remote I/O Panel' }, controls),
    equipment('PMP-2', { equipmentDescription: 'Pump', dependencies: 'Support group' }),
  ]))
  assert.deepEqual(tags(result, 'drivenElectricalPath'), ['PMP-1', 'PMP-2'])
  assert.deepEqual(tags(result, 'controlElectricalPath'), ['PLC-1', 'RIO-1'])
  assert.deepEqual(tags(result, 'rioControlPath'), ['RIO-1'])
  assert.deepEqual(tags(result, 'dependencyOnHeader'), ['PMP-2'])
})

test('real parent feeds and nested drive supplies work without reversing general commissioning edges', () => {
  const result = engine.runSsmAudit(snapshot([
    panel,
    equipment('AHU-FED', { equipmentDescription: 'Air Handler', dependencies: 'PNL-1' }),
    equipment('FAN-CHILD', { closestParent: 'AHU-FED', equipmentDescription: 'Fan Filter Unit' }),
    equipment('AHU-NOT-FED', { equipmentDescription: 'Air Handler' }),
    equipment('PMP-CHILD', { closestParent: 'AHU-NOT-FED', equipmentDescription: 'Pump', dependencies: 'PNL-1' }),
    equipment('PMP-DRIVEN', { equipmentDescription: 'Pump' }),
    equipment('VFD-1', { closestParent: 'PMP-DRIVEN', equipmentDescription: 'Variable Frequency Drive', dependencies: 'PNL-1' }),
    equipment('PMP-SEQUENCED', { equipmentDescription: 'Pump', dependencies: 'PMP-CHILD' }),
  ]))
  assert.deepEqual(tags(result, 'drivenElectricalPath'), ['AHU-NOT-FED', 'PMP-SEQUENCED'])
  assert.deepEqual(tags(result, 'precedenceCycle'), [], 'supply relationships do not alter commissioning cycle detection')
})

test('a powered controller dependency is control evidence, not proof of motor power', () => {
  const result = engine.runSsmAudit(snapshot([
    panel,
    equipment('PLC-1', { equipmentDescription: 'PLC', dependencies: 'PNL-1' }, controls),
    equipment('PMP-1', { equipmentDescription: 'Pump', dependencies: 'PLC-1' }),
  ]))
  assert.deepEqual(tags(result, 'drivenElectricalPath'), ['PMP-1'])
})

test('external supply uncertainty follows real equipment parents but cannot cross headers', () => {
  const result = engine.runSsmAudit(snapshot([
    equipment('AHU-EXT', { equipmentDescription: 'Air Handler', dependencies: 'EXT-PNL', dependencyProject: 'Synthetic Project' }),
    equipment('FAN-EXT', { closestParent: 'AHU-EXT', equipmentDescription: 'Fan Filter Unit' }),
    equipment('External group', { closestParent: 'AHU-EXT', itemMaster: 'VF_Blank', dependencies: 'EXT-PNL', dependencyProject: 'Synthetic Project' }),
    equipment('PMP-ABSENT', { closestParent: 'External group', equipmentDescription: 'Pump' }),
  ]))
  assert.deepEqual(tags(result, 'externalPathReview'), ['AHU-EXT', 'FAN-EXT'])
  assert.deepEqual(tags(result, 'drivenElectricalPath'), ['PMP-ABSENT'])
  assert.equal(result.summary.unverified, 2)
})

test('external power and control checks produce at most one informational unknown per equipment row', () => {
  const result = engine.runSsmAudit(snapshot([
    equipment('RIO-EXT', { equipmentDescription: 'Remote I/O Panel', dependencies: 'EXT-PLC; EXT-PNL', dependencyProject: 'Synthetic Project' }, controls),
    equipment('PMP-EXT', { equipmentDescription: 'Pump', closestParent: 'EXT-SKID', closestParentStatus: 'EXISTING' }),
    equipment('VESDA-EXT', { equipmentDescription: 'VESDA', closestParent: 'EXT-FACP', closestParentStatus: 'EXISTING' }, fire),
    equipment('VFD-EXT', { closestParent: 'PMP-EXT', equipmentDescription: 'Variable Frequency Drive', dependencies: 'EXT-PNL; EXT-PLC', dependencyProject: 'Synthetic Project' }),
  ]))
  assert.deepEqual(tags(result, 'externalPathReview'), ['PMP-EXT', 'RIO-EXT', 'VESDA-EXT', 'VFD-EXT'])
  assert.ok(findings(result, 'externalPathReview').every(finding => finding.severity === 'info'))
  for (const key of ['unresolvedParent', 'unresolvedDependency', 'drivenElectricalPath', 'controlElectricalPath', 'rioControlPath', 'vesdaFireAlarm', 'vfdDependencies']) assert.deepEqual(tags(result, key), [], key)
  const rio = findings(result, 'externalPathReview').find(finding => finding.equipmentId === 'RIO-EXT')
  assert.match(rio.why, /power supply/)
  assert.match(rio.why, /controller link/)
  assert.equal(result.summary.unverified, 4)
  assert.equal(result.summary.status, 'review')
})

test('external-heavy registries remain reviewable without hard missing-path errors', () => {
  const rows = Array.from({ length: 1200 }, (_, index) => equipment(`RIO-${index}`, { equipmentDescription: 'Remote I/O Panel', dependencies: 'EXT-CONTROL', dependencyProject: 'Synthetic Project', milestoneParent: 'L1 Readiness', milestone: 'L2 UPN 650 Readiness' }, controls))
  const result = engine.runSsmAudit(snapshot(rows))
  assert.equal(result.summary.unverified, rows.length)
  assert.equal(result.findings.length, rows.length)
  assert.deepEqual(result.summary.severity, { blocker: 0, error: 0, warning: 0, info: rows.length })
  assert.equal(result.summary.status, 'review')
})

test('verified local paths are not obscured by unrelated external references or stale project fields', () => {
  const result = engine.runSsmAudit(snapshot([
    panel,
    equipment('PLC-1', { equipmentDescription: 'PLC', dependencies: 'PNL-1' }, controls),
    equipment('RIO-KNOWN', { equipmentDescription: 'Remote I/O Panel', dependencies: 'PNL-1; PLC-1; EXT-AUX', dependencyProject: 'Synthetic Project' }, controls),
    equipment('PMP-ABSENT', { equipmentDescription: 'Pump', dependencyProject: 'Synthetic Project' }),
    equipment('PMP-UNRESOLVED', { equipmentDescription: 'Pump', dependencies: 'EXT-PNL' }),
    equipment('PMP-ROOT', { equipmentDescription: 'Pump', closestParentStatus: 'EXISTING' }),
  ]))
  assert.deepEqual(tags(result, 'externalPathReview'), [])
  assert.deepEqual(tags(result, 'drivenElectricalPath'), ['PMP-ABSENT', 'PMP-ROOT', 'PMP-UNRESOLVED'])
  assert.deepEqual(tags(result, 'unresolvedDependency'), ['PMP-UNRESOLVED'])
})

test('ambiguous, unmatched, and approved legacy names never get mandatory migration findings', () => {
  const rows = [equipment('IM-CLEAR', { itemMaster: 'IM_SITE_AIR_HANDLER' }), equipment('CA-AMBIGUOUS', { itemMaster: 'CA_SITE_PUMP' }), equipment('SP-UNMATCHED', { itemMaster: 'SP_SITE_WIDGET' }), equipment('IM-APPROVED', { itemMaster: 'IM_APPROVED' })]
  const result = engine.runSsmAudit(snapshot(rows), { itemMasterVocabulary: ['VF1_AIR_HANDLER', 'VF1_PUMP', 'VF2_PUMP', 'IM_APPROVED'] })
  assert.deepEqual(tags(result, 'itemMasterStandard'), [])
  assert.deepEqual(tags(result, 'itemMasterMigration'), ['IM-CLEAR'])
  const advisory = findings(result, 'itemMasterMigration')[0]
  assert.equal(advisory.severity, 'info')
  assert.equal(advisory.expected, 'VF1_AIR_HANDLER')
  assert.match(advisory.rule.title, /Optional/)
  assert.doesNotMatch(advisory.recommendation, /replace|must|should be used/i)
})

test('unknown milestone names have their own rule and preserve Blank-header and wrong-field guards', () => {
  const values = { milestoneParent: 'L1 Enabling', milestone: 'L2 Building Ready' }
  const result = engine.runSsmAudit(snapshot([
    equipment('EQ-UNKNOWN', values),
    equipment('EQ-WRONG', { ...values, milestone: 'L2 UPN 602 Enabling' }),
    equipment('Milestone group', { ...values, itemMaster: 'VF_Blank' }),
    equipment('CHILD', { closestParent: 'Milestone group' }),
    equipment('EQ-WRONG-LEVEL', { ...values, milestone: 'SITE-L1-M1-119 Building ready' }),
  ]))
  assert.deepEqual(tags(result, 'milestoneUpnUnknown'), ['EQ-UNKNOWN'])
  assert.deepEqual(tags(result, 'milestoneUpn'), ['EQ-WRONG'])
  assert.deepEqual(tags(result, 'milestoneLevel'), ['EQ-WRONG-LEVEL'])
  assert.deepEqual(tags(result, 'tagLooksLikeText'), [])
})

const root = fileURLToPath(new URL('..', import.meta.url))
const counterpart = resolve(root, '..', basename(root) === 'SSM-Audit' ? 'SSManagement' : 'SSM-Audit')
test('both repositories share identical engine/model source, exports, and synthetic behavior', { skip: !existsSync(counterpart) }, async () => {
  for (const name of ['model', 'engine']) {
    const path = `src/audit/${name}.js`
    assert.equal(readFileSync(resolve(root, path), 'utf8'), readFileSync(resolve(counterpart, path), 'utf8'), path)
  }
  const otherModel = await import(pathToFileURL(resolve(counterpart, 'src/audit/model.js')).href)
  const otherEngine = await import(pathToFileURL(resolve(counterpart, 'src/audit/engine.js')).href)
  assert.deepEqual(Object.keys(model).sort(), Object.keys(otherModel).sort())
  assert.deepEqual(Object.keys(engine).sort(), Object.keys(otherEngine).sort())
  assert.equal(engine.auditCommissioningRole({ equipmentDescription: 'Variable Frequency Drive' }), 'drive')
  assert.equal(engine.auditCommissioningRole({ equipmentClassification: 'RIO' }), 'rio')
  assert.equal(otherEngine.auditCommissioningRole({ equipmentClassification: 'RIO' }), 'rio')
  const input = snapshot([panel, equipment('Pump group', { itemMaster: 'VF_Blank' }), equipment('PMP-1', { closestParent: 'Pump group', equipmentDescription: 'Pump', dependencies: 'PNL-1' }), equipment('PMP-2', { closestParent: 'Pump group', equipmentDescription: 'Pump', milestone: 'Building Ready' })])
  assert.deepEqual(engine.runSsmAudit(input), otherEngine.runSsmAudit(input))
  assert.deepEqual(model.auditMergeSnapshots([input, input], 'synthetic.xlsx'), otherModel.auditMergeSnapshots([input, input], 'synthetic.xlsx'))
})
