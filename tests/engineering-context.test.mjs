import test from 'node:test'
import assert from 'node:assert/strict'
import {
  auditCoverageRows,
  auditEngineeringFindings,
  auditReadEngineeringAoa,
  SSM_AUDIT_ENGINEERING_RULES,
} from '../src/audit/engineering-references.js'

const parse = (rows, kind = 'mel', sheet = 'Synthetic') => auditReadEngineeringAoa(rows, kind, sheet)
const missing = (equipmentId, sourceContext, sheet = '', row = 0) => ({
  equipmentId,
  rule: SSM_AUDIT_ENGINEERING_RULES.missing,
  category: 'missing-tags',
  sourceContext,
  sheet,
  row,
})

test('source-only missing findings keep only normalized consensus context', () => {
  const reference = parse([
    ['Equipment Tag', 'UPN', 'Building', 'Discipline', 'System Name', 'Equipment Description', 'Equipment Classification', 'System Parent Equipment Tag(s)'],
    ['SYN-PUMP-1', '111', 'BLDG-A', 'MECHANICAL', '111 Chilled Water', 'Pump', 'CENTRIFUGAL', 'SYS-ROOT'],
    [' syn-pump-1 ', ' 111 ', ' bldg-a ', ' mechanical ', '111 chilled water', '', 'centrifugal', ' sys-root '],
  ])
  const finding = auditEngineeringFindings({ rows: [] }, { mel: reference })[0]

  assert.deepEqual(finding.sourceContext, {
    building: 'BLDG-A',
    discipline: 'MECHANICAL',
    upn: '111',
    systemName: '111 Chilled Water',
    equipmentDescription: 'Pump',
    equipmentClassification: 'CENTRIFUGAL',
    closestParent: 'SYS-ROOT',
  })
})

test('conflicting metadata stays blank while one nonblank normalized value survives', () => {
  const reference = parse([
    ['Equipment Tag', 'UPN', 'Building', 'Discipline', 'System Name', 'Equipment Description', 'Equipment Classification', 'System Parent Equipment Tag(s)'],
    ['SYN-PUMP-2', '111', 'BLDG-A', 'Mechanical', 'System A', 'Pump', 'Fan', 'SYS-ROOT'],
    ['SYN-PUMP-2', '112', 'BLDG-B', ' MECHANICAL ', ' system a ', '', 'PUMP', ' sys-root '],
  ])
  const finding = auditEngineeringFindings({ rows: [] }, { mel: reference })[0]

  assert.deepEqual(finding.sourceContext, {
    building: '',
    discipline: 'Mechanical',
    upn: '',
    systemName: 'System A',
    equipmentDescription: 'Pump',
    equipmentClassification: '',
    closestParent: 'SYS-ROOT',
  })
})

test('coverage merges unique context values from different engineering sources', () => {
  const mel = parse([
    ['Equipment Tag', 'UPN', 'Building', 'System Parent Equipment Tag(s)'],
    ['SYN-EXT-2', '111', 'BLDG-A', 'PARENT-A'],
  ], 'mel', 'MEL')
  const pmd = { entries: [{ equipmentId: 'syn-ext-2', parent: 'PARENT-B' }], sheetName: 'PMD' }
  const cable = { entries: [{ equipmentId: 'SYN-EXT-2', parent: 'parent-b' }], sheetName: 'Cable' }
  const findings = auditEngineeringFindings({ rows: [] }, { mel, pmd, cable })
  const coverage = auditCoverageRows({ rows: [], findings })

  assert.equal(coverage.length, 1)
  assert.equal(coverage[0].building, 'BLDG-A')
  assert.equal(coverage[0].upn, '111')
  assert.equal(coverage[0].closestParent, '')
})

test('coverage rows merge unique source context by normalized tag and export location', () => {
  const registryRow = { equipmentId: 'REGISTRY-ONLY', building: 'DO-NOT-COPY', _source: { sheet: 'Registry', row: 8 } }
  const result = {
    rows: [registryRow],
    findings: [
      missing('SYN-EXT-1', { building: 'BLDG-A', upn: '111', closestParent: 'SYN-PARENT' }, 'MEL', 4),
      missing(' syn-ext-1 ', { building: ' bldg-a ', discipline: 'ELECTRICAL', closestParent: 'syn-parent' }, 'MEL', 4),
      missing('SYN-EXT-1', { building: 'BLDG-B' }, 'MEL', 4),
      missing('SYN-EXT-1', { building: 'BLDG-B' }, 'MEL', 5),
      missing('REGISTRY-ONLY', { building: 'SHOULD-NOT-APPEAR' }, 'MEL', 4),
      { ...missing('OTHER-RULE', { building: 'SHOULD-NOT-APPEAR' }), rule: { id: 'reference.other' } },
    ],
  }

  const coverage = auditCoverageRows(result)
  assert.equal(coverage.length, 2)
  assert.deepEqual(coverage[0], {
    equipmentId: 'SYN-EXT-1',
    building: '',
    discipline: 'ELECTRICAL',
    upn: '111',
    systemName: '',
    equipmentDescription: '',
    equipmentClassification: '',
    closestParent: 'SYN-PARENT',
    _source: { sheet: 'MEL', row: 4 },
    _coverageOnly: true,
  })
  assert.equal(coverage[1].building, 'BLDG-B')
  assert.equal(coverage[1]._source.row, 5)
  assert.ok(!JSON.stringify(coverage).includes('REGISTRY-ONLY'))
  assert.deepEqual(result.rows, [registryRow])
})

test('coverage rows accept older findings without sourceContext', () => {
  const coverage = auditCoverageRows({
    rows: [],
    findings: [{ ...missing('SYN-LEGACY'), sourceContext: undefined }],
  })

  assert.deepEqual(coverage, [{
    equipmentId: 'SYN-LEGACY',
    building: '',
    discipline: '',
    upn: '',
    systemName: '',
    equipmentDescription: '',
    equipmentClassification: '',
    closestParent: '',
    _source: { sheet: '', row: 0 },
    _coverageOnly: true,
  }])
})

test('N/A and NA placeholders are skipped, but tag suffixes and prefixes remain', () => {
  const mel = parse([
    ['Equipment Tag', 'UPN', 'System Parent Equipment Tag(s)'],
    ['N/A', '111', 'N/A'],
    ['NA', '111', 'NA'],
    ['NA-PUMP', '111', 'PARENT-NA'],
    ['PUMP-NA', '111', 'N/A-ROOT'],
    ['N/A-PUMP', '111', 'NA-ROOT'],
  ])
  assert.deepEqual(mel.entries.map(entry => [entry.equipmentId, entry.closestParent]), [
    ['NA-PUMP', 'PARENT-NA'],
    ['PUMP-NA', 'N/A-ROOT'],
    ['N/A-PUMP', 'NA-ROOT'],
  ])

  const cable = parse([
    ['Load Name (To)', 'Panel (From)'],
    ['N/A; LOAD-NA; NA', 'NA; PANEL-NA; N/A'],
  ], 'cable')
  assert.deepEqual(cable.entries.map(entry => [entry.equipmentId, entry.parent]), [
    ['LOAD-NA', 'PANEL-NA'],
    ['PANEL-NA', ''],
  ])

  const pmd = parse([
    ['Panel', 'Instrument Tag'],
    ['N/A; PANEL-NA; NA', 'N/A; SENSOR-NA; NA'],
  ], 'pmd')
  assert.deepEqual(pmd.entries.map(entry => [entry.equipmentId, entry.parent]), [
    ['SENSOR-NA', 'PANEL-NA'],
    ['PANEL-NA', ''],
  ])

  const easyPower = parse([
    ['Starting Source', 'Downstream 1', 'Final Source', 'ID Name'],
    ['NA', 'N/A', 'POWER-NA', 'N/A-LOAD'],
  ], 'easyPower')
  assert.deepEqual(easyPower.entries.map(entry => [entry.equipmentId, entry.parent]), [
    ['POWER-NA', ''],
    ['N/A-LOAD', 'POWER-NA'],
  ])
})

test('dependency locality compares Dependency Project with Project, not Site', () => {
  const snapshot = {
    rows: [
      { equipmentId: 'SYN-CHILD', closestParent: 'SYN-ROOT', dependencies: 'SYN-PANEL', dependencyProject: 'DEMO', project: 'DEMO', site: 'OTHER' },
      { equipmentId: 'SYN-PANEL' },
    ],
  }
  const references = { pmd: { entries: [{ equipmentId: 'SYN-CHILD', parent: 'SYN-PANEL' }] } }

  assert.deepEqual(auditEngineeringFindings(snapshot, references), [])
})
