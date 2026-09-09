import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { esc } from '../src/core/text.js'

test('Tracker export selector keeps its choice and passes it to the download', async () => {
  const source = readFileSync(new URL('../src/ui/audit.js', import.meta.url), 'utf8')
  const start = source.indexOf('function renderExportOptions(){')
  const end = source.indexOf('\nexport function openExportOptions()', start)
  const session = { exportKind: 'tracker', result: { findings: [] } }, nodes = new Map(), downloads = []
  const inputs = [{ value: 'milestone' }, { value: 'discipline' }]
  const node = selector => {
    if (!nodes.has(selector)) nodes.set(selector, { innerHTML: '' })
    return nodes.get(selector)
  }
  const context = vm.createContext({
    S: { session }, esc, ic: () => '', $: node,
    $$: selector => selector === 'input[name="tracker-signoff"]' ? inputs : [],
    exportPlan: () => ({ levels: {}, rules: {} }), exportPlanSummary: () => ({}),
    SSM_AUDIT_SEVERITIES: [], closeExportOptions() {},
    exportTrackerXlsx: async mode => downloads.push(mode),
  })
  vm.runInContext(source.slice(start, end), context)
  vm.runInContext('renderExportOptions()', context)
  assert.match(node('#exportModalBody').innerHTML, /value="milestone" checked/)
  assert.match(node('#exportModalBody').innerHTML, /All checkmarks start empty/)
  inputs[1].onchange()
  assert.equal(session.trackerSignOffBy, 'discipline')
  assert.match(node('#exportModalBody').innerHTML, /value="discipline" checked/)
  await node('#exportGo').onclick()
  inputs[0].onchange()
  await node('#exportGo').onclick()
  assert.deepEqual(downloads, ['discipline', 'milestone'])
})
