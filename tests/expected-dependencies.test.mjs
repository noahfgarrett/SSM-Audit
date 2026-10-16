import test from 'node:test'
import assert from 'node:assert/strict'
import { auditExpectedDependencies } from '../src/audit/expected-dependencies.js'

const MECH = { building: 'BLDG-A', discipline: 'MECHANICAL WET', systemName: '111 Chilled Water', upn: '111' }
const ELEC = { building: 'BLDG-A', discipline: 'ELECTRICAL', systemName: '603 Low Voltage', upn: '603' }
const row = (equipmentId, context, extra = {}) => ({ equipmentId, closestParent: context.systemName, ...context, ...extra })
const edges = (kind, pairs) => ({ [kind]: { kind, sheetName: 'Sheet1', entries: pairs.map(([equipmentId, parent], index) => ({ equipmentId, parent, sourceRow: index + 2, sheetName: 'Sheet1' })) } })
const check = (rows, references) => auditExpectedDependencies({ rows }, references)
const only = (result, tag) => result.rows.find(entry => entry.equipmentId === tag)

test('a feeder in another group is an expected dependency', () => {
  const result = check([row('SYN-PUMP', MECH), row('SYN-PANEL', ELEC)], edges('cable', [['SYN-PUMP', 'SYN-PANEL']]))
  const pump = only(result, 'SYN-PUMP')
  assert.equal(pump.status, 'missing')
  assert.deepEqual(pump.missing, ['SYN-PANEL'])
  assert.deepEqual(pump.expected.map(entry => entry.tag), ['SYN-PANEL'])
  assert.match(pump.expected[0].reason, /discipline/i)
  assert.deepEqual(pump.expected[0].sources, [{ kind: 'cable', label: 'Cable Schedule', sheet: 'Sheet1', row: 2 }])
})

test('an expected dependency already listed is a match', () => {
  const result = check([row('SYN-PUMP', MECH, { dependencies: 'syn-panel' }), row('SYN-PANEL', ELEC)], edges('cable', [['SYN-PUMP', 'SYN-PANEL']]))
  assert.equal(only(result, 'SYN-PUMP').status, 'match')
  assert.deepEqual(result.counts, { compared: 1, match: 1, missing: 0, notInRegistry: 0 })
})

test('a feeder in the same group is a parent, not a dependency', () => {
  const result = check([row('SYN-PUMP', MECH), row('SYN-SKID', MECH)], edges('cable', [['SYN-PUMP', 'SYN-SKID']]))
  assert.equal(result.rows.length, 0)
})

test('electrical feeds nest inside one building but become dependencies across buildings', () => {
  const sub = { ...ELEC, systemName: '602 Medium Voltage', upn: '602' }
  const same = check([row('SYN-PANEL', ELEC), row('SYN-SWGR', sub)], edges('easyPower', [['SYN-SWGR', ''], ['SYN-PANEL', 'SYN-SWGR']]))
  assert.equal(same.rows.length, 0)
  const other = check([row('SYN-PANEL', ELEC), row('SYN-SWGR', { ...sub, building: 'BLDG-B' })], edges('easyPower', [['SYN-SWGR', ''], ['SYN-PANEL', 'SYN-SWGR']]))
  assert.deepEqual(only(other, 'SYN-PANEL').missing, ['SYN-SWGR'])
  assert.match(only(other, 'SYN-PANEL').expected[0].reason, /building/i)
})

test('an unknown group never creates an expectation', () => {
  const result = check([row('SYN-PUMP', { ...MECH, systemName: '' }), row('SYN-PANEL', ELEC)], edges('cable', [['SYN-PUMP', 'SYN-PANEL']]))
  assert.equal(result.rows.length, 0)
})

test('a source tag missing from the registry is skipped, flagged, and listed', () => {
  const result = check([row('SYN-PUMP', MECH), row('SYN-PANEL', ELEC)], edges('easyPower', [['SYN-PANEL', ''], ['SYN-MCC', 'SYN-PANEL'], ['SYN-PUMP', 'SYN-MCC']]))
  const pump = only(result, 'SYN-PUMP')
  assert.deepEqual(pump.missing, ['SYN-PANEL'])
  assert.deepEqual(pump.expected[0].skipped, ['SYN-MCC'])
  assert.ok(pump.notes.some(note => /SYN-MCC/.test(note) && /not in the registry/i.test(note)))
  assert.deepEqual(result.notInRegistry.map(entry => [entry.tag, entry.affects]), [['SYN-MCC', ['SYN-PUMP']]])
  assert.equal(result.counts.notInRegistry, 1)
})

test('a feeder sitting in Closest Parent is noted', () => {
  const result = check([row('SYN-PUMP', MECH, { closestParent: 'SYN-PANEL' }), row('SYN-PANEL', ELEC)], edges('pmd', [['SYN-PUMP', 'SYN-PANEL']]))
  const pump = only(result, 'SYN-PUMP')
  assert.equal(pump.status, 'missing')
  assert.ok(pump.notes.some(note => /listed as Closest Parent/i.test(note)))
})

test('a dependency in another Dependency Project does not satisfy the expectation', () => {
  const local = check([row('SYN-PUMP', MECH, { dependencies: 'SYN-PANEL', dependencyProject: 'DEMO', project: 'DEMO' }), row('SYN-PANEL', ELEC)], edges('cable', [['SYN-PUMP', 'SYN-PANEL']]))
  assert.equal(only(local, 'SYN-PUMP').status, 'match')
  const remote = check([row('SYN-PUMP', MECH, { dependencies: 'SYN-PANEL', dependencyProject: 'REMOTE', project: 'DEMO' }), row('SYN-PANEL', ELEC)], edges('cable', [['SYN-PUMP', 'SYN-PANEL']]))
  assert.equal(only(remote, 'SYN-PUMP').status, 'missing')
})

test('several feeders from several documents are all expected, each once', () => {
  const references = { ...edges('cable', [['SYN-PUMP', 'SYN-PANEL'], ['SYN-PUMP', 'SYN-PANEL']]), ...edges('pmd', [['SYN-PUMP', 'SYN-RIO']]) }
  const rio = { building: 'BLDG-A', discipline: 'FACILITIES MONITORING SYSTEM', systemName: '650 FMS', upn: '650' }
  const result = check([row('SYN-PUMP', MECH, { dependencies: 'SYN-RIO' }), row('SYN-PANEL', ELEC), row('SYN-RIO', rio)], references)
  const pump = only(result, 'SYN-PUMP')
  assert.deepEqual(pump.expected.map(entry => entry.tag), ['SYN-PANEL', 'SYN-RIO'])
  assert.deepEqual(pump.missing, ['SYN-PANEL'])
  assert.equal(pump.expected[0].sources.length, 2)
})

test('a chain that returns to its start is flagged as a loop', () => {
  const result = check([row('SYN-PUMP', MECH), row('SYN-PANEL', ELEC)], edges('cable', [['SYN-PUMP', 'SYN-PANEL'], ['SYN-PUMP', 'SYN-X'], ['SYN-X', 'SYN-PUMP']]))
  const pump = only(result, 'SYN-PUMP')
  assert.ok(pump.notes.some(note => /loop/i.test(note)))
})

test('a diamond reaching one feeder twice is not a loop', () => {
  const result = check([row('SYN-PUMP', MECH), row('SYN-PANEL', ELEC)], edges('easyPower', [['SYN-PUMP', 'SYN-A'], ['SYN-PUMP', 'SYN-B'], ['SYN-A', 'SYN-PANEL'], ['SYN-B', 'SYN-PANEL']]))
  const pump = only(result, 'SYN-PUMP')
  assert.deepEqual(pump.missing, ['SYN-PANEL'])
  assert.ok(!pump.notes.some(note => /loop/i.test(note)))
})

test('no source documents means nothing to compare', () => {
  const result = check([row('SYN-PUMP', MECH)], {})
  assert.deepEqual(result, { rows: [], notInRegistry: [], counts: { compared: 0, match: 0, missing: 0, notInRegistry: 0 }, kinds: [] })
})

/* EasyPower lists one row per load with its whole upstream chain, so the
   links near the top of the tree repeat on every row. */
function chainExport(loads) {
  const entries = [], rows = []
  const add = (equipmentId, parent, sourceRow) => entries.push({ equipmentId, parent, sourceRow, sheetName: 'Sheet1' })
  for (let i = 0; i < loads; i++) {
    const chain = ['SYN-GIS', `SYN-SWBD-${i % 4}`, `SYN-PNL-${i % 40}`, `SYN-SUB-${i % 400}`, `SYN-LOAD-${i}`]
    chain.forEach((tag, index) => add(tag, index ? chain[index - 1] : '', i + 2))
  }
  for (let p = 0; p < 40; p++) rows.push(row(`SYN-PNL-${p}`, ELEC))
  for (let i = 0; i < loads; i++) rows.push(row(`SYN-LOAD-${i}`, MECH))
  return { rows, references: { easyPower: { kind: 'easyPower', sheetName: 'Sheet1', entries } } }
}

test('a large EasyPower export with repeated upstream links stays fast', () => {
  const { rows, references } = chainExport(30000)
  const started = performance.now()
  const result = check(rows, references)
  assert.ok(performance.now() - started < 2000, `took ${Math.round(performance.now() - started)} ms`)
  assert.equal(result.counts.compared, 30000)
})

test('source locations are trimmed for display but counted exactly', () => {
  const { rows, references } = chainExport(3000)
  const result = check(rows, references)
  const gis = result.notInRegistry.find(entry => entry.tag === 'SYN-GIS')
  assert.equal(gis.sourceCount, 3000)
  assert.ok(gis.sources.length <= 20)
  assert.equal(new Set(gis.sources.map(source => source.row)).size, gis.sources.length, 'no duplicate locations')
  const load = only(result, 'SYN-LOAD-7')
  assert.equal(load.expected[0].sourceCount, 1)
  assert.deepEqual(load.expected[0].skipped, ['SYN-SUB-7'])
})
