import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { GUIDE_SECTIONS, referenceHelpHtml } from '../src/ui/guide-content.js'
import { ICONS, ic } from '../src/ui/icons.js'

test('Guide starts with the audit workflow and covers each current review surface', () => {
  assert.equal(GUIDE_SECTIONS[0].id, 'start')
  const ids = new Set(GUIDE_SECTIONS.map(section => section.id))
  assert.equal(ids.size, GUIDE_SECTIONS.length)
  for (const id of ['actions', 'references', 'save', 'checks', 'reading', 'hierarchy', 'compare', 'report', 'privacy']) assert(ids.has(id))
  for (const section of GUIDE_SECTIONS) {
    assert(ICONS[section.icon], `Guide icon ${section.icon} must render`)
    assert(section.title && section.label && section.body)
    assert.doesNotMatch(section.body, /20,000 rows take about a second|src\/|SSM_EXPORT_|\.js\b/)
  }
})

test('reference help and Guide share the same plain-language scope and limitations', () => {
  const help = referenceHelpHtml()
  assert.equal(GUIDE_SECTIONS.find(section => section.id === 'references').body, help)
  for (const text of ['Milestone register', 'Item Master catalog', 'does not choose', 'Apply references', 'Apply changes', 'Cancel', 'same registry and reference data', 'Compare Projects']) assert(help.includes(text))
  assert.match(help, /does not apply corrections/)
  assert.match(help, /different prefix alone is not an error/)
  assert.match(help, /Confirm the equipment and checklist requirements/)
})

test('Guide distinguishes drafts, reviewed findings, saved sessions and offline exports', () => {
  const body = GUIDE_SECTIONS.map(section => section.body).join('\n')
  for (const text of ['Apply changes', 'Cancel', 'Cleared in draft', 'Save review', 'Load review', 'Updated Registry', 'Correction Log', 'Tracker']) assert(body.includes(text))
  assert.doesNotMatch(body,/Preview changes|Apply to draft|Mark reviewed|Accept an exception/)
  assert.match(body, /Workbook ticks do not change Exto or sync back into the app/)
  assert.match(body, /different registry revision is rejected/)
  assert.match(body, /fallback checks may run/)
  assert.match(body, /do not include workbook data/)
})

test('References and Actions controls have real icons and scoped centered geometry', () => {
  for (const name of ['circle-help', 'folder-open', 'history', 'save', 'undo-2', 'x']) {
    assert(ICONS[name], `${name} must not reserve empty icon space`)
    assert.match(ic(name), /<(?:path|circle|rect)\b/)
  }
  const css = readFileSync(new URL('../src/styles/app.css', import.meta.url), 'utf8')
  assert.match(css, /\.action-modal \.btn,\.modify-shell \.btn\{justify-content:center;text-align:center\}/)
  assert.match(css, /\.action-modal \.btn\.icon-btn,\.modify-shell \.btn\.icon-btn\{[^}]*padding:0;gap:0;[^}]*justify-content:center/)
  assert.match(css, /\.reference-row\{display:grid;grid-template-columns:minmax\(0,1fr\) auto 36px/)
  assert.match(css, /\.reference-selection\{grid-column:1\/-1;min-width:0/)
  assert.match(css, /\.reference-help\[hidden\]\{display:none\}/)
})
