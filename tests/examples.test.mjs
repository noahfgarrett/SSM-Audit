import test from 'node:test'
import assert from 'node:assert/strict'

import { SSM_AUDIT_RULES, runSsmAudit } from '../src/audit/engine.js'
import { SSM_AUDIT_EXAMPLES, auditExampleColumns, auditExampleSnapshot } from '../src/audit/examples.js'

const enabled = Object.values(SSM_AUDIT_RULES).filter(rule => rule.enabled)

test('every enabled check has a worked example', () => {
  const missing = enabled.filter(rule => !SSM_AUDIT_EXAMPLES[rule.id]).map(rule => rule.id)
  assert.deepEqual(missing, [])
  const stray = Object.keys(SSM_AUDIT_EXAMPLES).filter(id => !enabled.some(rule => rule.id === id))
  assert.deepEqual(stray, [], 'examples for checks that are not in use')
})

test('the real engine flags the focused row of every example for that example\'s check', () => {
  const failures = []
  for (const rule of enabled) {
    const example = SSM_AUDIT_EXAMPLES[rule.id]
    const result = runSsmAudit(auditExampleSnapshot(example), example.options || {})
    const hits = result.findings.filter(finding => finding.rule.id === rule.id)
    if (!hits.length) { failures.push(`${rule.id}: no finding (got ${[...new Set(result.findings.map(f => f.rule.id))].join(', ') || 'nothing'})`); continue }
    const focusedTags = new Set(example.focus.map(cell => example.rows[cell.row].equipmentId))
    if (!hits.some(finding => focusedTags.has(finding.equipmentId))) failures.push(`${rule.id}: fired on ${hits.map(f => f.equipmentId).join(', ')} but the example highlights ${[...focusedTags].join(', ')}`)
  }
  assert.deepEqual(failures, [])
})

test('every example explains itself and highlights at least one cell', () => {
  for (const [id, example] of Object.entries(SSM_AUDIT_EXAMPLES)) {
    assert.ok(example.caption.length > 20, id)
    assert.ok(example.fix.length > 5, id)
    assert.ok(example.focus.length > 0, id)
    for (const cell of example.focus) assert.ok(example.rows[cell.row], `${id} focus row ${cell.row}`)
    for (const mark of example.marks) {
      const joined = mark.parts.map(part => part[0]).join('')
      assert.equal(joined, String(example.rows[mark.row][mark.field]), `${id} marks must spell the cell exactly`)
      assert.ok(mark.parts.some(part => part[1]), `${id} marks must flag something`)
    }
    const columns = auditExampleColumns(example)
    for (const cell of example.focus) assert.ok(columns.includes(cell.field), `${id} shows the focused column ${cell.field}`)
  }
})

test('examples never carry site or client identifiers', () => {
  const text = JSON.stringify(SSM_AUDIT_EXAMPLES)
  for (const row of Object.values(SSM_AUDIT_EXAMPLES).flatMap(example => example.rows)) {
    assert.match(row.equipmentId, /^(B[0-9]-|Chilled water)/, `mock tags start with a generic building prefix: ${row.equipmentId}`)
    assert.match(row.building, /^B[0-9]$/)
  }
  assert.doesNotMatch(text, /SP_|CA_NB|_NB_/, 'no site prefixes')
})
