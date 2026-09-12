import test from 'node:test'
import assert from 'node:assert/strict'

import { buildMilestoneReference, calibrateMilestoneLayers, milestoneBuildingType, milestoneL2Number, milestoneScheduleFromRegister, milestoneTagCore, recommendMilestones } from '../src/audit/milestone-recommend.js'

const row = (equipmentId, values = {}) => ({ equipmentId, upn: '111', discipline: 'MECHANICAL WET', building: 'B1 (F9.1 FAB)', systemName: '111  Chilled Water', closestParent: '111  Chilled Water', milestone: '', milestoneParent: '', ...values })
const L2 = 'L2-M1-1220 - UPN 111 CHW Enabling System Green Tag', L1 = 'L1-M1-110 - BC1 Available'

test('milestone numbers parse across site prefixes and sub-numbers', () => {
  assert.equal(milestoneL2Number('L2-M1-1185 - FAB NPS LVSS Green Tag'), '1185')
  assert.equal(milestoneL2Number('SP-L2-M1-1101_1 UPN 243 Distribution'), '1101.1')
  assert.equal(milestoneL2Number('L2-M1-3400.1-FAB Remaining Equipment'), '3400.1')
  assert.equal(milestoneL2Number('CUB1 Op Ready Misc'), '')
  assert.equal(milestoneTagCore('LC23-PMP114-1-00'), 'PMP114-1-00')
  assert.equal(milestoneTagCore('OC33-PMP114-1-00'), 'PMP114-1-00')
  assert.equal(milestoneBuildingType('LC21 (F38.1 FAB)'), 'FAB')
  assert.equal(milestoneBuildingType('OC33 (BC2)'), 'CUB')
})

test('the closest-parent chain hands its milestone pair down', () => {
  const rows = [row('PMP-1', { milestone: L2, milestoneParent: L1 }), row('VFD-1', { closestParent: 'PMP-1' }), row('XV-1', { closestParent: 'VFD-1' })]
  const { byId, summary } = recommendMilestones(rows)
  assert.equal(summary.assigned, 2)
  const grandchild = byId.get('XV-1')
  assert.equal(grandchild.layer, 'parent-chain')
  assert.equal(grandchild.l2Label, L2)
  assert.equal(grandchild.l1Label, L1)
  assert.equal(grandchild.confidence, .95)
})

test('a branch that already agrees fills its unlabelled members, strongly when it is consistent', () => {
  const rows = [row('HDR', { milestone: '' })]
  for (let i = 0; i < 4; i++) rows.push(row(`PMP-${i}`, { closestParent: 'HDR', milestone: L2, milestoneParent: L1 }))
  rows.push(row('PMP-9', { closestParent: 'HDR' }))
  const { byId } = recommendMilestones(rows)
  const late = byId.get('PMP-9')
  assert.equal(late.layer, 'branch-strong')
  assert.equal(late.number, '1220')
  assert.equal(late.confidence, .99)
  // the header itself is a root with no parent -- branch evidence still applies
  assert.equal(byId.get('HDR').layer, 'branch-strong')
})

test('a reference site counterpart supplies the milestone number, translated into the local schedule', () => {
  const reference = buildMilestoneReference([row('OC33-PMP114-1-00', { upn: '114', discipline: 'ELECTRICAL', milestone: 'L2-M1-1118 - AZ wording', milestoneParent: 'L1-M1-125 - AZ L1' })], 'AZ')
  const schedule = milestoneScheduleFromRegister([{ id: 'SP-L2-M1-1118', title: 'UPN 114 local wording', parentId: 'SP-L1-M1-125', parentTitle: 'Local L1' }])
  const rows = [row('LC23-PMP114-1-00', { upn: '114', discipline: 'ELECTRICAL' })]
  const { byId } = recommendMilestones(rows, { schedule, references: [reference] })
  const hit = byId.get('LC23-PMP114-1-00')
  assert.equal(hit.layer, 'counterpart-electrical')
  assert.equal(hit.number, '1118')
  assert.equal(hit.l2Label, 'SP-L2-M1-1118 UPN 114 local wording')
  assert.equal(hit.l1Label, 'SP-L1-M1-125 Local L1')
  assert.equal(hit.confidence, .95)
  assert.equal(hit.local, true)
})

test('a UPN the schedule names in exactly one milestone is recommended, otherwise candidates are listed', () => {
  const schedule = milestoneScheduleFromRegister([
    { id: 'SP-L2-M1-1104', title: 'UPN 231 IW Enabling', parentId: 'SP-L1-M1-123', parentTitle: 'Utility' },
    { id: 'SP-L2-M1-1121', title: 'UPN 101 Enabling', parentId: 'SP-L1-M1-129', parentTitle: 'Fab' },
    { id: 'SP-L2-M1-3051', title: 'UPN 101 100% Blowdown', parentId: 'SP-L1-M1-133', parentTitle: 'Fab' },
  ])
  const rows = [row('PMP-231', { upn: '231' }), row('AHU-101', { upn: '101' }), row('XX-999', { upn: '999' })]
  const { byId, summary } = recommendMilestones(rows, { schedule })
  assert.equal(byId.get('PMP-231').layer, 'schedule-upn')
  assert.equal(byId.get('PMP-231').number, '1104')
  assert.equal(byId.get('AHU-101').layer, 'candidates')
  assert.deepEqual(byId.get('AHU-101').candidates.map(candidate => candidate.number).sort(), ['1121', '3051'])
  assert.equal(byId.get('XX-999').layer, 'none')
  assert.deepEqual(summary, { rows: 3, assigned: 1, candidates: 1, none: 1 })
})

test('reference registries lend the usual number for a UPN, building type, and discipline', () => {
  const reference = buildMilestoneReference([
    ...Array.from({ length: 9 }, (_, i) => row(`OC31-AHU-${i}`, { upn: '101', building: 'OC31 (F52.1 FAB)', discipline: 'MECHANICAL DRY', milestone: 'L2-M1-1121 - UPN 101 MAH', milestoneParent: 'L1-M1-129' })),
    row('OC31-AHU-X', { upn: '101', building: 'OC31 (F52.1 FAB)', discipline: 'MECHANICAL DRY', milestone: 'L2-M1-3051 - UPN 101 100%', milestoneParent: 'L1-M1-133' }),
  ], 'AZ')
  const rows = [row('LC21-AHU-77', { upn: '101', building: 'LC21 (F38.1 FAB)', discipline: 'MECHANICAL DRY' })]
  const { byId } = recommendMilestones(rows, { references: [reference] })
  const hit = byId.get('LC21-AHU-77')
  assert.equal(hit.layer, 'pattern-good')
  assert.equal(hit.number, '1121')
  assert.equal(hit.local, false, 'no local milestone carries the number yet, and the label says so')
  assert.match(hit.l2Label, /no local milestone carries number 1121/)
  assert.equal(hit.candidates.length, 2)
})

test('calibration replays the layers against labelled rows and reports measured accuracy', () => {
  const rows = [row('PMP-1', { milestone: L2, milestoneParent: L1 })]
  for (let i = 0; i < 6; i++) rows.push(row(`VFD-${i}`, { closestParent: 'PMP-1', milestone: L2, milestoneParent: L1 }))
  rows.push(row('VFD-ODD', { closestParent: 'PMP-1', milestone: 'L2-M1-9999 - odd one', milestoneParent: L1 }))
  const layers = calibrateMilestoneLayers(rows)
  assert.ok(layers['parent-chain'].tested >= 6)
  assert.equal(layers['parent-chain'].right, 6)
  assert.equal(layers['parent-chain'].accuracy, Math.round(6 / layers['parent-chain'].tested * 100) / 100)
})
