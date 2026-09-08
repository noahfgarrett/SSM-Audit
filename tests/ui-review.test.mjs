import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { clean, esc } from '../src/core/text.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { auditApplyCorrections, auditCorrectionImpact, auditMakeCorrection, auditReadReviewDocument, auditRegistryRevision, auditReviewDocument } from '../src/audit/actions.js'
import { auditReadReferenceAoa, auditReadReferenceWorkbook, auditReferenceFindings, auditReferenceSheets, SSM_AUDIT_REFERENCE_RULES } from '../src/audit/references.js'
import { validateAuditCorrections } from '../src/audit/export.js'
import { resetSession } from '../src/state.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js', import.meta.url), 'utf8'), { filename: 'sheetjs.js' })
const ui = readFileSync(new URL('../src/ui/audit.js', import.meta.url), 'utf8')
const start = ui.indexOf('export function sessionAudit('), end = ui.indexOf('\nfunction renderActionPreview(', start)
assert.ok(start >= 0 && end > start, 'private review helper boundaries must remain identifiable')
const reviewSource = ui.slice(start, end).replace('export function sessionAudit(', 'function sessionAudit(')
const headers = EXTO_REV21_COLUMNS.map(column => column.header)
const system = '602  Medium Voltage'

function registry(overrides = {}) {
  const values = { equipmentId: 'EQ-1', building: 'DEMO', upn: '602', discipline: 'ELECTRICAL', systemName: '650  Facility Management System', closestParent: system, equipmentDescription: 'Electrical panel', ...overrides }
  const aoa = [headers, EXTO_REV21_COLUMNS.map(column => values[column.field] || '')]
  const baseline = auditSnapshotFromAoa(aoa, { sheet: 'Registry' }), book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(aoa), 'Registry')
  const result = runSsmAudit(baseline)
  return {
    sourceBytes: XLSX.write(book, { type: 'array', bookType: 'xlsx' }), baselineSnapshot: baseline, baselineResult: result,
    snapshot: baseline, rawResult: result, references: {}, changes: [], changesRev: 0, actionedRev: 0,
    actioned: new Set(), reviewedIds: new Set(), draftResolved: new Set(), excluded: new Set(), reviewHistory: [], reviewUndo: [], filterViews: [],
  }
}

// Run production helpers unchanged. Only rendering and async checkpoints are
// substituted; workbook parsing, correction validation, digests and audits are real.
function reviewHarness(session) {
  const nodes = new Map(), lists = new Map(), hooks = {}, messages = []
  const calls = { checkpoints: 0, preflight: 0, progress: 0, refresh: 0, render: 0, readReview: 0 }
  const node = selector => {
    if (!nodes.has(selector)) {
      const classes = new Set()
      nodes.set(selector, { innerHTML: '', disabled: false, setAttribute() {}, focus() {}, classList: { add: name => classes.add(name), remove: name => classes.delete(name), contains: name => classes.has(name) } })
    }
    return nodes.get(selector)
  }
  const context = vm.createContext({
    S: { session, comparison: { targetSnapshot: session.snapshot, result: null } }, XLSX, clean, esc,
    auditApplyCorrections, auditCorrectionImpact, auditReadReferenceAoa, auditReadReferenceWorkbook, auditReferenceFindings, auditReferenceSheets, SSM_AUDIT_REFERENCE_RULES, runSsmAudit,
    auditReadReviewDocument: async (...args) => { calls.readReview++; return auditReadReviewDocument(...args) },
    validateAuditCorrections: (...args) => { calls.preflight++; return validateAuditCorrections(...args) },
    auditReviewDocument, modifyRecommendationContext: null, currentNavigate() {},
    $: node, $$: selector => lists.get(selector) || [], ic: () => '',
    document: { activeElement: null, contains: () => false },
    animateOpen: element => element.classList.add('show'), animateClose: element => element.classList.remove('show'),
    activateFocusTrap: () => () => {}, readArrayBuffer: async file => file.bytes,
    refreshSessionResult: () => { calls.refresh++ }, rerenderModifications: () => { calls.render++ }, toast: message => messages.push(message),
    runWithProgress: async (_title, _description, run) => {
      calls.progress++
      return run(async () => { calls.checkpoints++; await hooks.checkpoint?.(calls.checkpoints) }, () => {})
    },
  })
  vm.runInContext(reviewSource, context, { filename: 'audit-review-helpers.js' })
  const api = vm.runInContext('({loadReviewFile,reviewPrepare,reviewRememberUndo,reviewInstallDraft,reviewUndoLast,openReferencesDialog,sessionAudit})', context)
  return { context, api, calls, hooks, messages, node, lists }
}

function pauseOnce() {
  let entered, release, paused = false
  const reached = new Promise(resolve => { entered = resolve }), waiting = new Promise(resolve => { release = resolve })
  return { reached, release, pause: async () => { if (paused) return; paused = true; entered(); await waiting } }
}

async function savedReview(session) {
  const change = auditMakeCorrection(session.baselineSnapshot.rows[0], 'System Name', system)
  return auditReviewDocument({ ...session, changes: [change] })
}

for (const stage of ['file', 'digest', 'before-preflight', 'after-preflight']) {
  for (const invalidation of ['new session', 'draft revision']) {
    test(`review restore aborts at ${stage} when the ${invalidation} changes`, { timeout: 5000 }, async t => {
      const session = registry(), replacement = registry({ equipmentDescription: 'Electrical distribution panel' })
      assert.notEqual(await auditRegistryRevision(session.baselineSnapshot), await auditRegistryRevision(replacement.baselineSnapshot))
      // Cache population is legitimate on the original session, not a review edit.
      if (invalidation === 'draft revision') session.reviewSourceWorkbook = XLSX.read(new Uint8Array(session.sourceBytes), { type: 'array', cellStyles: true })
      const data = await savedReview(session), harness = reviewHarness(session), gate = pauseOnce()
      const file = { size: 100, text: async () => { if (stage === 'file') await gate.pause(); return JSON.stringify(data) } }
      if (stage === 'digest') {
        const digest = globalThis.crypto.subtle.digest.bind(globalThis.crypto.subtle)
        t.mock.method(globalThis.crypto.subtle, 'digest', async (...args) => { const result = await digest(...args); await gate.pause(); return result })
      }
      harness.hooks.checkpoint = count => {
        if (stage === 'before-preflight' && count === 1 || stage === 'after-preflight' && count === 2) return gate.pause()
      }
      const job = harness.api.loadReviewFile(file)
      try {
        await Promise.race([gate.reached, job.then(() => { throw new Error('Review finished before the requested async boundary') })])
        if (invalidation === 'new session') harness.context.S.session = replacement
        else session.changesRev++
        const current = harness.context.S.session, before = structuredClone(current), comparison = structuredClone(harness.context.S.comparison)
        gate.release(); await job
        assert.equal(harness.context.S.session, current)
        assert.deepEqual(structuredClone(current), before, 'an aborted restore must not alter decisions, draft, history, undo, references or source bytes')
        assert.deepEqual(structuredClone(harness.context.S.comparison), comparison)
        assert.equal(harness.calls.refresh, 0)
        assert.equal(harness.calls.render, 0)
        assert.ok(harness.messages.some(message => /changed/i.test(message)))
        if (stage === 'file') assert.equal(harness.calls.readReview, 0)
        if (stage === 'digest') assert.equal(harness.calls.progress, 0)
        if (stage === 'after-preflight') assert.equal(harness.calls.preflight, 1, 'the pause must follow actual raw-cell validation')
      } finally { gate.release(); await job }
    })
  }
}

test('a matching review restores corrections, reviewed decisions and filter views, and undo restores the prior views', async () => {
  const session = registry(), finding = session.baselineResult.findings.find(finding => finding.rule.id === 'metadata.system-upn-mismatch')
  assert.ok(finding)
  const view = name => ({ name, filters: { hiddenSeverities: [], hiddenSources: [], hiddenCategories: [], hiddenRules: [], dimFilters: { discipline: [], milestone: [], upn: ['602'], building: [] } } })
  const priorViews = [view('Prior view')], restoredViews = [view('Restored view')]
  session.filterViews = priorViews
  const data = await savedReview({ ...session, reviewedIds: new Set([finding.id]), actioned: new Set([finding.id]), filterViews: restoredViews })
  const harness = reviewHarness(session), baseline = session.baselineSnapshot, bytes = session.sourceBytes.slice(0)
  await harness.api.loadReviewFile({ size: 100, text: async () => JSON.stringify(data) })
  assert.equal(session.snapshot.rows[0].systemName, system)
  assert.equal(session.reviewedIds.has(finding.id), true)
  assert.equal(session.actioned.has(finding.id), true)
  assert.equal(session.draftResolved.has(finding.id), true)
  assert.equal(session.reviewUndo.length, 1)
  assert.equal(harness.calls.preflight, 1)
  assert.equal(harness.calls.refresh, 1)
  assert.equal(session.baselineSnapshot, baseline)
  assert.deepEqual(session.sourceBytes, bytes)
  assert.deepEqual(structuredClone(session.filterViews), restoredViews)
  await harness.api.reviewUndoLast()
  assert.deepEqual(structuredClone(session.filterViews), priorViews)
  assert.equal(session.reviewedIds.size, 0)
  assert.equal(session.snapshot.rows[0].systemName, baseline.rows[0].systemName)
  assert.equal(session.reviewUndo.length, 0)
})

for (const reviewed of [true, false]) {
  test(`draft apply, revert and undo ${reviewed ? 'retain explicit reviewedIds' : 'do not turn automatic resolution into an explicit review'}`, async () => {
    const dependencies = 'REMOTE-1; REMOTE-1'
    const session = registry({ systemName: system, dependencies, dependencyProject: 'Synthetic project' }), harness = reviewHarness(session)
    const finding = session.baselineResult.findings.find(finding => finding.rule.id === 'dependency.duplicate')
    assert.ok(finding)
    const baseline = session.baselineSnapshot, baselineResult = session.baselineResult, bytes = session.sourceBytes.slice(0)
    if (reviewed) {
      session.reviewedIds.add(finding.id); session.actioned.add(finding.id)
      session.reviewHistory.push({ id: 'review-1', at: '2026-01-01T00:00:00Z', owner: 'Reviewer', reason: 'Checked source relationship', disposition: 'reviewed', findingIds: [finding.id], changes: [] })
    }
    const history = structuredClone(session.reviewHistory), change = auditMakeCorrection(baseline.rows[0], 'Dependencies', 'REMOTE-1', finding)
    const install = async changes => {
      const prepared = await harness.api.reviewPrepare(changes)
      assert.equal(prepared.impact.unsafe.length, 0)
      harness.api.reviewRememberUndo(); harness.api.reviewInstallDraft(prepared)
    }
    const assertDecisions = resolved => {
      assert.equal(session.reviewedIds.has(finding.id), reviewed)
      assert.equal(session.draftResolved.has(finding.id), resolved)
      assert.equal(session.actioned.has(finding.id), reviewed || resolved)
      assert.deepEqual(structuredClone(session.reviewHistory), history)
      assert.equal(session.baselineSnapshot, baseline)
      assert.equal(session.baselineResult, baselineResult)
      assert.deepEqual(session.sourceBytes, bytes)
    }
    await install([change]); assertDecisions(true)
    assert.equal(session.snapshot.rows[0].dependencies, 'REMOTE-1')
    await install([]); assertDecisions(false)
    assert.equal(session.snapshot.rows[0].dependencies, dependencies)
    await harness.api.reviewUndoLast(); assertDecisions(true)
    await harness.api.reviewUndoLast(); assertDecisions(false)
    assert.equal(session.snapshot.rows[0].dependencies, dependencies)
    assert.equal(session.reviewUndo.length, 0)
  })
}

test('filter views stay session-only without reading, writing or deleting localStorage', () => {
  const from = ui.indexOf('function loadFilterViews('), to = ui.indexOf('\nfunction captureFilterView(', from)
  assert.ok(from >= 0 && to > from)
  const context = vm.createContext({ S: { session: null } })
  let storageAccesses = 0
  Object.defineProperty(context, 'localStorage', { get() { storageAccesses++; throw new Error('Session filter views must not access persistent storage') } })
  vm.runInContext(`${resetSession.toString()}\n${ui.slice(from, to)}`, context)
  const api = vm.runInContext('({loadFilterViews,saveFilterViews,resetSession})', context)
  api.resetSession()
  const session = context.S.session
  const view = { name: 'Synthetic view', filters: { dimFilters: { upn: ['602'] } } }
  api.saveFilterViews([view])
  assert.deepEqual(structuredClone(api.loadFilterViews()), [view])
  assert.equal(session.reviewDirty, true)
  api.resetSession()
  assert.deepEqual(structuredClone(api.loadFilterViews()), [])
  assert.deepEqual(session.filterViews, [view], 'changing sessions does not delete the previous session or migrate stored data')
  assert.equal(storageAccesses, 0)
})

test('undo interrupted by a session switch cannot copy old decisions or its undo entry into the new session', async () => {
  const session = registry(), replacement = registry({ equipmentDescription: 'Electrical distribution panel' }), harness = reviewHarness(session), gate = pauseOnce()
  harness.api.reviewRememberUndo()
  harness.hooks.checkpoint = () => gate.pause()
  const job = harness.api.reviewUndoLast()
  try {
    await gate.reached
    harness.context.S.session = replacement
    const before = structuredClone(replacement)
    gate.release(); await job
    assert.deepEqual(structuredClone(replacement), before)
    assert.equal(harness.calls.refresh, 0)
    assert.equal(harness.calls.render, 0)
  } finally { gate.release(); await job }
})

for (const kind of ['itemMasters', 'milestones']) {
  test(`an empty selected ${kind} reference cannot be applied and a usable selection can recover`, async () => {
    const session = registry(), harness = reviewHarness(session)
    const columns = kind === 'itemMasters' ? ['Item Master Unique Identifier'] : ['L2 ID', 'L1 ID']
    const values = kind === 'itemMasters' ? ['VF_SYNTHETIC_EQUIPMENT'] : ['L2-DEMO-1', 'L1-DEMO-1']
    const current = auditReadReferenceAoa([columns, values], kind, 'Current')
    assert.equal(current.entries.length, 1)
    session.references = { [kind]: current }
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([columns, values]), 'Current')
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([columns]), 'Empty')
    const file = { bytes: XLSX.write(book, { type: 'array', bookType: 'xlsx' }) }
    const input = { dataset: { referenceFile: kind }, files: [file] }, select = { dataset: { referenceSheet: kind }, value: 'Empty' }
    harness.lists.set('[data-reference-file]', [input]); harness.lists.set('[data-reference-sheet]', [select])
    harness.api.openReferencesDialog()
    await input.onchange()
    select.onchange()
    const before = structuredClone(session)
    await harness.node('#referencesApply').onclick()
    assert.deepEqual(structuredClone(session), before)
    assert.equal(session.references[kind], current)
    assert.equal(harness.calls.progress, 0, 'empty references must fail before any re-audit')
    assert.equal(harness.calls.refresh, 0)
    assert.match(harness.messages.at(-1), /current entries/i)
    assert.equal(harness.node('#referencesApply').disabled, false)
    select.value = 'Current'; select.onchange()
    await harness.node('#referencesApply').onclick()
    assert.equal(session.references[kind].entries.length, 1)
    assert.equal(session.references[kind].sheetName, 'Current')
    assert.equal(harness.calls.refresh, 1)
    assert.equal(harness.calls.render, 1)
  })
}
