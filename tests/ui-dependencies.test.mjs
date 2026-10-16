import test from 'node:test'
import assert from 'node:assert/strict'
import { auditExpectedDependencies } from '../src/audit/expected-dependencies.js'
import { dependencyFiltered, dependencyItems, dependencyRowHtml } from '../src/ui/dependencies.js'

const MECH = { building: 'BLDG-A', discipline: 'MECHANICAL WET', systemName: '111 Chilled Water', upn: '111' }
const ELEC = { building: 'BLDG-A', discipline: 'ELECTRICAL', systemName: '603 Low Voltage', upn: '603' }
const entries = [['SYN-PANEL', ''], ['SYN-MCC<b>', 'SYN-PANEL'], ['SYN-PUMP-1', 'SYN-MCC<b>'], ['SYN-PUMP-2', 'SYN-PANEL']]
const check = auditExpectedDependencies({ rows: [
  { equipmentId: 'SYN-PUMP-1', ...MECH },
  { equipmentId: 'SYN-PUMP-2', ...MECH, building: 'BLDG-B', dependencies: 'SYN-PANEL' },
  { equipmentId: 'SYN-PANEL', ...ELEC },
] }, { cable: { kind: 'cable', sheetName: 'Synthetic', entries: entries.map(([equipmentId, parent], index) => ({ equipmentId, parent, sourceRow: index + 2, sheetName: 'Synthetic' })) } })
const items = dependencyItems(check)

test('the tab lists compared rows and tags missing from the registry together', () => {
  assert.deepEqual(items.map(item => [item.tag, item.kind]), [['SYN-PUMP-1', 'missing'], ['SYN-PUMP-2', 'match'], ['SYN-MCC<b>', 'absent']])
})

test('status, discipline, building and search filters narrow the list', () => {
  const names = session => dependencyFiltered(items, session).map(item => item.tag)
  assert.deepEqual(names({ depStatus: 'missing' }), ['SYN-PUMP-1'])
  assert.deepEqual(names({ depStatus: 'absent' }), ['SYN-MCC<b>'])
  assert.deepEqual(names({ depBuilding: 'BLDG-B' }), ['SYN-PUMP-2'])
  assert.deepEqual(names({ depDiscipline: 'MECHANICAL WET' }), ['SYN-PUMP-1', 'SYN-PUMP-2'])
  assert.deepEqual(names({ depSearch: 'pump-1' }), ['SYN-PUMP-1', 'SYN-MCC<b>'], 'a missing upstream tag shows beside the rows it affects')
  assert.deepEqual(names({ depSearch: 'syn-panel', depStatus: 'missing' }), ['SYN-PUMP-1'], 'search also matches expected dependencies')
})

test('rows escape document text and show the chain when expanded', () => {
  const missing = items.find(item => item.kind === 'missing'), absent = items.find(item => item.kind === 'absent')
  const closed = dependencyRowHtml(missing, false), open = dependencyRowHtml(missing, true)
  assert.match(closed, /aria-expanded="false"/)
  assert.match(closed, />Missing</)
  assert.ok(!closed.includes('dep-detail'))
  assert.match(open, /dep-detail/)
  assert.match(open, /class="skipped">SYN-MCC&lt;b&gt;</)
  assert.ok(!open.includes('SYN-MCC<b>'))
  assert.ok(!dependencyRowHtml(absent, true).includes('<b>SYN'))
  assert.match(dependencyRowHtml(absent, false), /Affects 1 registry row</)
})
