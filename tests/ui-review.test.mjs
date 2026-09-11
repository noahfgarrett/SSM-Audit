import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { clean, esc } from '../src/core/text.js'
import { EXTO_REV21_COLUMNS, extoRev21SystemsForUpn } from '../src/exto/rev21-contract.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { runSsmAudit } from '../src/audit/engine.js'
import { AUDIT_ACTION_FIELDS, auditFindingRow, auditCorrectionKey, auditProposeCorrection, auditRecommendationContext, auditCustomCorrection, auditMergeCorrections, auditApplyCorrections, auditCorrectionImpact, auditMakeCorrection, auditReadReviewDocument, auditRegistryRevision, auditReviewDocument } from '../src/audit/actions.js'
import { auditReadReferenceAoa, auditReadReferenceWorkbook, auditReferenceFindings, auditReferenceSheets, SSM_AUDIT_REFERENCE_RULES } from '../src/audit/references.js'
import { validateAuditCorrections } from '../src/audit/export.js'
import { resetSession } from '../src/state.js'
import { referenceHelpHtml } from '../src/ui/guide-content.js'
import { auditReadMilestoneMigration, auditReadMigrationSettings, auditMigrationReferences, auditMilestoneMigrationRows, auditMigrationImpact } from '../src/audit/milestone-migration.js'
import { auditActionEntry, auditActionPolicy } from '../src/audit/actions.js'
import { auditSessionResult, auditPrepareInWorker } from '../src/audit/review.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js', import.meta.url), 'utf8'), { filename: 'sheetjs.js' })
const ui = readFileSync(new URL('../src/ui/audit.js', import.meta.url), 'utf8')
const start = ui.indexOf('export function sessionAudit('), end = ui.indexOf('\nfunction renderActionPreview(', start)
assert.ok(start >= 0 && end > start, 'private review helper boundaries must remain identifiable')
const reviewSource = ui.slice(start, end).replace('export function sessionAudit(', 'function sessionAudit(')
const headers = EXTO_REV21_COLUMNS.map(column => column.header)
const system = '602  Medium Voltage'

function registry(overrides = {}, additional = []) {
  const values = { equipmentId: 'EQ-1', building: 'DEMO', upn: '602', discipline: 'ELECTRICAL', systemName: '650  Facility Management System', closestParent: system, equipmentDescription: 'Electrical panel', ...overrides }
  const aoa = [headers, ...[values,...additional].map(row=>EXTO_REV21_COLUMNS.map(column => row[column.field] || ''))]
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
  const reviewCache={};
  const nodes = new Map(), lists = new Map(), hooks = {}, messages = []
  const calls = { checkpoints: 0, preflight: 0, progress: 0, refresh: 0, render: 0, readReview: 0 }
  const node = selector => {
    if (!nodes.has(selector)) {
      const classes = new Set()
      nodes.set(selector, { innerHTML: '', disabled: false, querySelector:()=>({value:'suggested'}), setAttribute() {}, focus() {}, classList: { add: name => classes.add(name), remove: name => classes.delete(name), contains: name => classes.has(name) } })
    }
    return nodes.get(selector)
  }
  const context = vm.createContext({
    S: { session, comparison: { targetSnapshot: session.snapshot, result: null } }, XLSX, clean, esc,auditSessionResult,auditActionEntry,auditActionPolicy,
    importAuditWorkbook:async(file,{referenceKind})=>{
      const workbook=XLSX.read(file.bytes,{type:'array',dense:true});
      return {references:auditReferenceSheets(workbook,referenceKind).map(name=>auditReadReferenceWorkbook(workbook,referenceKind,name))};
    },
    prepareAuditReview:async(s,changes,migration,migrationChanged,report,options={})=>{
      if(changes.length)calls.preflight++;
      return auditPrepareInWorker(reviewCache,{changes,previousChanges:s.changes||[],references:options.references||s.references||{},referencesChanged:options.references!==undefined,migration,migrationChanged,...(!reviewCache.baseline?{baseline:s.baselineSnapshot,file:new Blob([s.sourceBytes||new Uint8Array()])}:{})},report);
    },
    crypto:globalThis.crypto,AUDIT_ACTION_FIELDS,auditFindingRow,auditMakeCorrection,auditCorrectionKey,auditRecommendationContext,auditCustomCorrection,auditMergeCorrections,isExcludedId:()=>false,
    auditReadMilestoneMigration,auditReadMigrationSettings,auditMigrationReferences,auditMilestoneMigrationRows,auditMigrationImpact,
    auditSparrowMilestoneMigration:()=>auditReadMilestoneMigration({format:'ssm-audit-milestone-map',version:1,project:'Demo',mappings:[{from:'DEMO-L1-M1-01',to:'DEMO-L1-M1-02',label:'DEMO-L1-M1-02 New scope'}]}),
    modifySuggestedFix:finding=>auditProposeCorrection(finding,auditRecommendationContext(session.snapshot,session.references)),
    auditApplyCorrections, auditCorrectionImpact, auditReadReferenceAoa, auditReadReferenceWorkbook, auditReferenceFindings, auditReferenceSheets, SSM_AUDIT_REFERENCE_RULES, runSsmAudit,
    auditReadReviewDocument: async (...args) => { calls.readReview++; return auditReadReviewDocument(...args) },
    validateAuditCorrections: (...args) => { calls.preflight++; return validateAuditCorrections(...args) },
    auditReviewDocument, modifyRecommendationContext: null, currentNavigate() {},
    $: node, $$: selector => lists.get(selector) || [], ic: () => '', referenceHelpHtml,
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
  vm.runInContext(ui.slice(end,ui.indexOf('\nfunction syncModifyPatternBox(',end)),context)
  vm.runInContext(ui.slice(ui.indexOf('function milestoneMigrationControls('),ui.indexOf('export function renderModifications(')),context)
  const api = vm.runInContext('({loadReviewFile,reviewPrepare,reviewRememberUndo,reviewInstallDraft,reviewUndoLast,openReferencesDialog,openActionDialog,sessionAudit,setMilestoneMigration,previewMilestoneMigration,milestoneMigrationControls,reviewMilestoneMappings,wireMilestoneMigration})', context)
  return { context, api, calls, hooks, messages, node, lists }
}

test('project milestone replacement previews and applies only L1, then restores with undo',async()=>{
  const session=registry({systemName:system,milestoneParent:'DEMO-L1-M1-01 Old scope',milestone:'DEMO-L2-M1-20 Equipment scope'}),h=reviewHarness(session);
  const profile=auditReadMilestoneMigration({format:'ssm-audit-milestone-map',version:1,project:'Demo',mappings:[{from:'DEMO-L1-M1-01',to:'DEMO-L1-M1-02',label:'DEMO-L1-M1-02 New scope'}]});
  await h.api.setMilestoneMigration({enabled:true,profile});
  assert.equal(session.snapshot.rows[0].milestoneParent,'DEMO-L1-M1-01 Old scope');
  assert.equal(session.milestoneMigration.enabled,true);
  h.api.previewMilestoneMigration();
  assert.match(h.node('#actionPreviewRows').innerHTML,/Milestone Parent/);
  assert.match(h.node('#actionPreviewRows').innerHTML,/DEMO-L1-M1-02 New scope/);
  await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[0].milestoneParent,'DEMO-L1-M1-01 Old scope');
  await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[0].milestoneParent,'DEMO-L1-M1-02 New scope');
  assert.equal(session.snapshot.rows[0].milestone,'DEMO-L2-M1-20 Equipment scope');
  assert.equal(session.changes.length,1);
  await h.api.reviewUndoLast();
  assert.equal(session.snapshot.rows[0].milestoneParent,'DEMO-L1-M1-01 Old scope');
  assert.equal(session.milestoneMigration.enabled,true);
});

test('milestone mapping summary works while off and reviews every affected row before applying',async()=>{
  const session=registry({systemName:system,milestoneParent:'DEMO-L1-M1-01 Old scope',milestone:'DEMO-L2-M1-20 Equipment scope'}),h=reviewHarness(session);
  session.milestoneMigration={enabled:false,profile:auditReadMilestoneMigration({format:'ssm-audit-milestone-map',version:1,project:'Demo',mappings:[{from:'DEMO-L1-M1-01',to:'DEMO-L1-M1-02',label:'DEMO-L1-M1-02 New scope'}]})};
  session.modifySearch='not this equipment';
  const controls=h.api.milestoneMigrationControls();assert.match(controls,/1 rows with replacements/);assert.doesNotMatch(controls,/disabled/);
  h.api.wireMilestoneMigration();h.node('#migrationPreview').onclick();
  assert.match(h.node('#actionModalBody').innerHTML,/DEMO-L1-M1-01/);assert.match(h.node('#actionModalBody').innerHTML,/DEMO-L1-M1-02/);assert.match(h.node('#actionModalBody').innerHTML,/1 equipment rows/);
  assert.equal(session.milestoneMigration.enabled,false);assert.equal(session.changes.length,0);
  await h.node('#migrationReviewApply').onclick();
  assert.equal(session.milestoneMigration.enabled,true);assert.equal(session.changes.length,0);
  assert.match(h.node('#actionPreviewRows').innerHTML,/DEMO-L1-M1-02 New scope/);
  await h.node('#actionApply').onclick();await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[0].milestoneParent,'DEMO-L1-M1-02 New scope');assert.equal(session.snapshot.rows[0].milestone,'DEMO-L2-M1-20 Equipment scope');
});

test('built-in milestone replacements activate without reference or mapping files and start with a preview',async()=>{
  const session=registry({systemName:system,milestoneParent:'DEMO-L1-M1-01 Old scope'}),h=reviewHarness(session);
  const controls=h.api.milestoneMigrationControls();assert.doesNotMatch(controls,/disabled|checked|migrationFile|Load mapping/);assert.match(controls,/1 rows with replacements/);
  h.api.wireMilestoneMigration();await h.node('#newMilestones').onchange({target:{checked:true}});
  assert.equal(session.milestoneMigration.enabled,true);assert.equal(session.changes.length,0);
  assert.match(h.node('#actionModalBody').innerHTML,/DEMO-L1-M1-01/);assert.match(h.node('#actionModalBody').innerHTML,/DEMO-L1-M1-02/);
  await h.node('#migrationReviewApply').onclick();await h.node('#actionApply').onclick();await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[0].milestoneParent,'DEMO-L1-M1-02 New scope');
  assert.equal(session.snapshot.rows[0].milestone,'');assert.ok(session.rawResult.findings.some(finding=>finding.rule.id==='milestone.incomplete-pair'));
  h.api.wireMilestoneMigration();await h.node('#newMilestones').onchange({target:{checked:false}});
  assert.equal(session.milestoneMigration.enabled,false);assert.equal(session.snapshot.rows[0].milestoneParent,'DEMO-L1-M1-02 New scope');
});

test('built-in milestone preview with no matching equipment cannot apply replacements',()=>{
  const session=registry({milestoneParent:'OTHER-L1-M1-01 Scope'}),h=reviewHarness(session);
  h.api.reviewMilestoneMappings();assert.match(h.node('#actionModalBody').innerHTML,/0 equipment rows/);
  assert.match(h.node('#actionModalBody').innerHTML,/id="migrationReviewApply" disabled/);
  assert.equal(session.changes.length,0);assert.equal(session.milestoneMigration,undefined);
});

test('metadata target switch preserves edits and applies to the chosen parent row',async()=>{
  const session=registry({systemName:system,building:'DEMO-A',closestParent:'PARENT'},[{equipmentId:'PARENT',building:'DEMO-B',upn:'602',systemName:system,discipline:'ELECTRICAL',closestParent:system}]);
  const h=reviewHarness(session),child={value:'child'},parent={value:'parent'};h.lists.set('input[name="actionTarget"]',[child,parent]);
  const issue=session.rawResult.findings.find(f=>f.rule.id==='parent.cross-building');assert.ok(issue);
  h.api.openActionDialog('Building mismatch',[issue]);
  assert.match(h.node('#actionModalBody').innerHTML,/Edit metadata on/);
  assert.match(h.node('#actionPreviewRows').innerHTML,/Building/);assert.doesNotMatch(h.node('#actionPreviewRows').innerHTML,/aria-label="Closest Parent/);
  h.node('#actionPreviewRows').oninput({target:{closest:()=>({dataset:{actionEntry:'0',actionCell:'0'},value:'DEMO-C'})}});
  parent.onchange();assert.match(h.node('#actionPreviewRows').innerHTML,/<b>PARENT<\/b>/);
  child.onchange();assert.match(h.node('#actionPreviewRows').innerHTML,/value="DEMO-C"/);
  parent.onchange();assert.match(h.node('#actionPreviewRows').innerHTML,/value="DEMO-B"/);
  h.node('#actionPreviewRows').oninput({target:{closest:()=>({dataset:{actionEntry:'0',actionCell:'0'},value:'DEMO-A'})}});
  await h.node('#actionApply').onclick();await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[0].building,'DEMO-A');assert.equal(session.snapshot.rows[1].building,'DEMO-A');assert.equal(session.changes[0].tag,'PARENT');
  assert.equal(session.snapshot.rows[0].closestParent,'PARENT');assert.match(h.node('#actionModalBody').innerHTML,/Changes applied/);
});

test('group editor defaults to bulk and preserves individual overrides and child-parent drafts',async()=>{
  const parentRow={equipmentId:'PARENT',building:'DEMO-B',upn:'602',systemName:system,discipline:'ELECTRICAL',closestParent:system};
  const childRow={equipmentId:'EQ-2',building:'DEMO-A',upn:'602',systemName:system,discipline:'ELECTRICAL',closestParent:'PARENT'};
  const session=registry({systemName:system,building:'DEMO-A',closestParent:'PARENT'},[childRow,parentRow]);
  const h=reviewHarness(session),bulk={dataset:{actionMode:'bulk'}},individual={dataset:{actionMode:'individual'}},child={value:'child'},parent={value:'parent'};
  h.lists.set('[data-action-mode]',[bulk,individual]);h.lists.set('input[name="actionTarget"]',[child,parent]);
  const issues=session.rawResult.findings.filter(f=>f.rule.id==='parent.cross-building');assert.equal(issues.length,2);
  await h.api.openActionDialog('Building mismatch',issues);
  assert.match(h.node('#actionModalBody').innerHTML,/data-action-mode="bulk" aria-selected="true"/);
  assert.match(h.node('#actionModalBody').innerHTML,/data-action-bulk="building"[^>]*value="DEMO-A"/);
  h.node('#actionBulkFields').oninput({target:{closest:()=>({dataset:{actionBulk:'building'},value:'DEMO-C'})}});
  individual.onclick();assert.equal((h.node('#actionPreviewRows').innerHTML.match(/value="DEMO-C"/g)||[]).length,2);
  h.node('#actionPreviewRows').oninput({target:{closest:()=>({dataset:{actionEntry:'1',actionCell:'0'},value:'DEMO-D'})}});
  bulk.onclick();assert.match(h.node('#actionModalBody').innerHTML,/Mixed suggestions/);
  individual.onclick();assert.match(h.node('#actionPreviewRows').innerHTML,/value="DEMO-D"/);
  bulk.onclick();parent.onchange();assert.match(h.node('#actionModalBody').innerHTML,/1 affected parent rows/);
  child.onchange();individual.onclick();assert.match(h.node('#actionPreviewRows').innerHTML,/value="DEMO-C"/);assert.match(h.node('#actionPreviewRows').innerHTML,/value="DEMO-D"/);
  parent.onchange();
  h.node('#actionPreviewRows').oninput({target:{closest:()=>({dataset:{actionEntry:'0',actionCell:'0'},value:'DEMO-A'})}});
  await h.node('#actionApply').onclick();assert.equal(session.changes.length,0);
  await h.node('#actionApply').onclick();assert.deepEqual(session.snapshot.rows.map(r=>r.building),['DEMO-A','DEMO-A','DEMO-A']);
  assert.equal(session.changes.length,1);assert.equal(session.changes[0].tag,'PARENT');
});

test('bulk edit applies the shared value to every selected child only after confirmation',async()=>{
  const session=registry({systemName:system,dependencyProject:'DEMO',project:'DEMO'},[{equipmentId:'EQ-2',systemName:system,upn:'602',project:'DEMO',dependencyProject:'DEMO',discipline:'ELECTRICAL',building:'DEMO'}]);
  const h=reviewHarness(session),issues=session.rawResult.findings.filter(f=>f.rule.id==='dependency.project-not-needed');assert.equal(issues.length,2);
  await h.api.openActionDialog('Clear dependency project',issues);
  h.node('#actionBulkFields').oninput({target:{closest:()=>({dataset:{actionBulk:'dependencyProject'},value:''})}});
  await h.node('#actionApply').onclick();assert.equal(session.changes.length,0);
  await h.node('#actionApply').onclick();assert.deepEqual(session.snapshot.rows.map(r=>r.dependencyProject),['','']);assert.equal(session.changes.length,2);
});

test('catalog migration proposes a VF replacement and applies only its Item Master cell',async()=>{
  const session=registry({systemName:system,itemMaster:'CAMPUS_DEMO_EL_PANEL'}),h=reviewHarness(session);
  const catalog=auditReadReferenceAoa([['Item Master Name'],['VF_EL_PANEL']],'itemMasters','VF Current');
  session.references={itemMasters:catalog};session.rawResult=auditSessionResult(session.snapshot,session.references);session.baselineResult=session.rawResult;
  const issue=session.rawResult.findings.find(f=>f.rule.id==='item-master.migration-advisory');assert.ok(issue);
  h.api.openActionDialog('Optional VF migration',[issue]);
  assert.match(h.node('#actionPreviewRows').innerHTML,/value="VF_EL_PANEL"/);
  await h.node('#actionApply').onclick();assert.equal(session.snapshot.rows[0].itemMaster,'CAMPUS_DEMO_EL_PANEL');
  await h.node('#actionApply').onclick();assert.equal(session.snapshot.rows[0].itemMaster,'VF_EL_PANEL');
  assert.deepEqual(session.changes.map(c=>c.field),['Item Master Unique Identifier']);
  assert.equal(session.baselineSnapshot.rows[0].itemMaster,'CAMPUS_DEMO_EL_PANEL');
  assert.ok(!session.rawResult.findings.some(f=>f.rule.id==='item-master.migration-advisory'));
});
test('bulk action preparation shows progress and yields before displaying editable blanks',async()=>{
  const session=registry({systemName:system,dependencyProject:'DEMO',project:'DEMO'}),h=reviewHarness(session);
  const issue=session.rawResult.findings.find(f=>f.rule.id==='dependency.project-not-needed');assert.ok(issue);
  const findings=Array.from({length:2872},(_,i)=>({...issue,id:`bulk-${i}`}));
  await h.api.openActionDialog('Clear project',findings);
  assert.ok(h.calls.progress>0);assert.ok(h.calls.checkpoints>20);
  assert.match(h.node('#actionPreviewRows').innerHTML,/value=""/);
  assert.equal(session.reviewBusy,false);assert.equal(session.changes.length,0);
});

test('large L1 reviews yield with controls locked before merging and validating every change',async()=>{
  const values={systemName:system,milestoneParent:'DEMO-L1-M1-01 Old scope',milestone:'DEMO-L2-M1-10'};
  const session=registry(values,Array.from({length:499},(_,i)=>({...values,equipmentId:`EQ-${i+2}`,building:'DEMO',upn:'602',discipline:'ELECTRICAL',closestParent:system,equipmentDescription:'Electrical panel'}))),h=reviewHarness(session);
  session.milestoneMigration={enabled:true,profile:h.context.auditSparrowMilestoneMigration()};
  await h.api.previewMilestoneMigration();
  const before=h.calls.checkpoints;let sawLocked=false;
  h.hooks.checkpoint=()=>{sawLocked||=session.reviewBusy&&h.node('#actionApply').disabled;};
  await h.node('#actionApply').onclick();
  assert.ok(sawLocked);assert.ok(h.calls.checkpoints-before>=5);assert.equal(h.calls.preflight,1);
  assert.match(h.node('#actionModalBody').innerHTML,/500 cells will change/);
  assert.equal(session.changes.length,0,'review alone must not apply the batch');
  await h.node('#actionApply').onclick();
  assert.equal(session.changes.length,500);
  assert.ok(session.changes.every(change=>change.prop==='milestoneParent'));
  assert.ok(session.snapshot.rows.every(row=>row.milestone==='DEMO-L2-M1-10'));
});

test('unchanged reviews reuse audit results but still validate cells and invalidate changed reference context',async()=>{
  const session=registry(),cache={},changes=[auditMakeCorrection(session.snapshot.rows[0],'System Name',system)];
  const data={baseline:session.baselineSnapshot,file:new Blob([session.sourceBytes]),changes,previousChanges:[],references:{},migration:{enabled:false,profile:null}};
  const first=await auditPrepareInWorker(cache,data);
  const {baseline,file,...repeat}=data;
  const second=await auditPrepareInWorker(cache,repeat);
  assert.equal(second.result,first.result);assert.deepEqual(second.impact,first.impact);
  const third=await auditPrepareInWorker(cache,{...repeat,references:{itemMasters:{entries:[{name:'VF_DEMO'}]}}});
  assert.notEqual(third.result,first.result);
  const column=session.baselineSnapshot.rows[0]._source.columns.systemName;
  cache.workbook.Sheets.Registry[XLSX.utils.encode_cell({r:1,c:column})].v='Conflicting original';
  await assert.rejects(auditPrepareInWorker(cache,repeat),/original cell value has changed/i);
});

test('Actions show the issue and editable drive corrections, preview all cells, then confirm success',async()=>{
  const system101=extoRev21SystemsForUpn('101')[0],rows=[
    {equipmentId:'DEMO-MAH101-01-00',building:'DEMO',upn:'101',systemName:system101,discipline:'MECHANICAL DRY',equipmentDescription:'MAH Makeup Air Handler'},
    {equipmentId:'DEMO-VFD101-01-00',building:'DEMO',upn:'650',systemName:extoRev21SystemsForUpn('650')[0],discipline:'FACILITIES MONITORING SYSTEM',equipmentDescription:'Variable Frequency Drive',closestParent:'DEMO-MAH101-01-00'},
  ];
  const aoa=[headers,...rows.map(row=>EXTO_REV21_COLUMNS.map(c=>row[c.field]||''))],snapshot=auditSnapshotFromAoa(aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
  const session={...registry(),baselineSnapshot:snapshot,snapshot,baselineResult:runSsmAudit(snapshot),rawResult:runSsmAudit(snapshot),sourceBytes:XLSX.write(book,{type:'array',bookType:'xlsx'})};
  const h=reviewHarness(session),issue=session.rawResult.findings.find(f=>f.rule.id==='parent.cross-upn');
  h.api.openActionDialog('Parent UPN mismatch',[issue]);
  assert.match(h.node('#actionPreviewRows').innerHTML,/UPN/);assert.match(h.node('#actionPreviewRows').innerHTML,/value="101"/);assert.match(h.node('#actionPreviewRows').innerHTML,/System Name/);
  assert.match(h.node('#actionPreviewRows').innerHTML,/This row is on UPN 650/);
  assert.equal(session.snapshot.rows[1].upn,'650','opening the action changes no data');
  assert.doesNotMatch(h.node('#actionModalBody').innerHTML,/action-mode|actionValue|actionOwner|actionReason|actionSelectAll|Preview changes/);
  assert.doesNotMatch(h.node('#actionPreviewRows').innerHTML,/checkbox/);
  assert.match(h.node('#actionModalBody').innerHTML,/>Review changes</);
  await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[1].upn,'650','review does not apply edits');
  assert.match(h.node('#actionModalBody').innerHTML,/3 cells will change/);
  assert.doesNotMatch(h.node('#actionPreviewRows').innerHTML,/<input/);
  await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[1].upn,'101');assert.equal(session.snapshot.rows[1].systemName,system101);
  assert.equal(session.baselineSnapshot.rows[1].upn,'650');assert.equal(session.changes.length,3);
  assert.ok(session.draftResolved.has(issue.id));assert.equal(session.reviewHistory.at(-1).disposition,'corrected-draft');assert.equal(h.calls.refresh,1);
  assert.match(h.node('#actionModalBody').innerHTML,/Changes applied/);
  assert.match(h.node('#actionModalBody').innerHTML,/3 cells updated/);
  assert.equal(h.node('#actionModal').classList.contains('show'),true,'confirmation stays visible until dismissed');
  await h.node('#actionApply').onclick();assert.match(h.node('#actionModalBody').innerHTML,/Corrections and review history/);
  assert.match(h.node('#actionModalBody').innerHTML,/3 cells changed/);
});

test('an unsupported action allows verified input without inventing a correction',async()=>{
  const session=registry({closestParent:'UNKNOWN-PARENT'}),h=reviewHarness(session),issue=session.rawResult.findings.find(f=>f.rule.id==='parent.unresolved');
  assert.ok(issue);h.api.openActionDialog('Missing parent',[issue]);
  assert.match(h.node('#actionPreviewRows').innerHTML,/No reliable suggestion/);
  assert.match(h.node('#actionPreviewRows').innerHTML,/value="UNKNOWN-PARENT"/);
  await h.node('#actionApply').onclick();
  assert.match(h.node('#actionImpact').textContent,/No values have changed/);h.node('#actionCancel').onclick();
  assert.equal(session.changes.length,0);assert.equal(h.calls.preflight,0);
});
test('a user-entered fix is previewed, can be edited again, and is only applied after confirmation',async()=>{
  const session=registry({systemName:system,closestParent:'UNKNOWN-PARENT'},[{equipmentId:'DEMO-PANEL',building:'DEMO',upn:'602',discipline:'ELECTRICAL',systemName:system,closestParent:system,closestParentStatus:'NEW'}]),h=reviewHarness(session),issue=session.rawResult.findings.find(f=>f.rule.id==='parent.unresolved');
  h.api.openActionDialog('Parent missing',[issue]);
  const edit=value=>h.node('#actionPreviewRows').oninput({target:{closest:()=>({dataset:{actionEntry:'0',actionCell:'0'},value})}});
  edit('EQ-1');await h.node('#actionApply').onclick();
  assert.match(h.node('#actionImpact').textContent,/No changes applied/);
  edit('DEMO-PANEL');await h.node('#actionApply').onclick();
  assert.equal(session.changes.length,0);
  assert.match(h.node('#actionPreviewRows').innerHTML,/UNKNOWN-PARENT/);
  assert.match(h.node('#actionPreviewRows').innerHTML,/DEMO-PANEL/,h.node('#actionImpact').textContent);
  h.node('#actionBack').onclick();
  assert.match(h.node('#actionPreviewRows').innerHTML,/value="DEMO-PANEL"/);
  await h.node('#actionApply').onclick();await h.node('#actionApply').onclick();
  assert.equal(session.snapshot.rows[0].closestParent,'DEMO-PANEL');assert.equal(session.changes.length,1);
  assert.ok(session.draftResolved.has(issue.id));assert.match(h.node('#actionModalBody').innerHTML,/1 cell updated/);
  h.node('#actionCancel').onclick();await h.api.reviewUndoLast();
  assert.equal(session.snapshot.rows[0].closestParent,'UNKNOWN-PARENT');assert.equal(session.changes.length,0);
});
test('edits survive pagination and every page is included in the reviewed batch',async()=>{
  const records=Array.from({length:81},(_,i)=>({equipmentId:`DEMO-${i}`,building:'DEMO',upn:'602',discipline:'ELECTRICAL',systemName:extoRev21SystemsForUpn('650')[0],closestParent:system}));
  const aoa=[headers,...records.map(row=>EXTO_REV21_COLUMNS.map(c=>row[c.field]||''))],snapshot=auditSnapshotFromAoa(aoa,{sheet:'Registry'}),book=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(aoa),'Registry');
  const session={...registry(),baselineSnapshot:snapshot,snapshot,baselineResult:runSsmAudit(snapshot),rawResult:runSsmAudit(snapshot),sourceBytes:XLSX.write(book,{type:'array',bookType:'xlsx'})},h=reviewHarness(session);
  const issues=session.rawResult.findings.filter(f=>f.rule.id==='metadata.system-upn-mismatch');assert.equal(issues.length,81);
  h.api.openActionDialog('System mismatch',issues);
  assert.equal((h.node('#actionPreviewRows').innerHTML.match(/<tr>/g)||[]).length,80);
  h.node('#actionNext').onclick();assert.match(h.node('#actionPreviewRows').innerHTML,/DEMO-80/);
  h.node('#actionPreviewRows').oninput({target:{closest:()=>({dataset:{actionEntry:'80',actionCell:'0'},value:system+' '})}});
  h.node('#actionPrevious').onclick();h.node('#actionNext').onclick();
  assert.match(h.node('#actionPreviewRows').innerHTML,/value="602  Medium Voltage "/);
  await h.node('#actionApply').onclick();assert.match(h.node('#actionModalBody').innerHTML,/81 cells will change/,h.node('#actionImpact').textContent);
  await h.node('#actionApply').onclick();assert.equal(session.changes.length,81);assert.equal(session.snapshot.rows[80].systemName,system);
});
test('changing the draft after preview invalidates the final confirmation',async()=>{
  const session=registry(),h=reviewHarness(session),issue=session.rawResult.findings.find(f=>f.rule.id==='metadata.system-upn-mismatch');
  h.api.openActionDialog('System mismatch',[issue]);await h.node('#actionApply').onclick();
  session.changesRev++;await h.node('#actionApply').onclick();
  assert.equal(session.changes.length,0);assert.match(h.node('#actionImpact').textContent,/registry changed/);
});
for(const cancelDuringValidation of [false,true])test(`Cancel leaves suggestions unapplied ${cancelDuringValidation?'during validation':'before validation'}`,async()=>{
  const session=registry(),h=reviewHarness(session),issue=session.rawResult.findings.find(f=>f.rule.id==='metadata.system-upn-mismatch');
  h.api.openActionDialog('System mismatch',[issue,issue]);
  assert.equal((h.node('#actionPreviewRows').innerHTML.match(/<tr>/g)||[]).length,1);
  if(cancelDuringValidation){
    const gate=pauseOnce();h.hooks.checkpoint=gate.pause;
    const job=h.node('#actionApply').onclick();await gate.reached;
    h.node('#actionCancel').onclick();gate.release();await job;
  }else h.node('#actionCancel').onclick();
  assert.equal(session.changes.length,0);assert.equal(session.reviewHistory.length,0);
  assert.equal(session.snapshot,session.baselineSnapshot);assert.equal(h.calls.refresh,0);
});

test('one-click apply blocks a suggestion that introduces a new hierarchy error',async()=>{
  const session=registry(),h=reviewHarness(session),issue=session.rawResult.findings.find(f=>f.rule.id==='metadata.system-upn-mismatch');
  h.context.auditActionEntry=()=>({row:session.snapshot.rows[0],finding:issue,reason:'Synthetic unsafe suggestion',changes:[auditMakeCorrection(session.snapshot.rows[0],'Closest Parent','EQ-1',issue)]});
  h.api.openActionDialog('Test correction',[issue]);await h.node('#actionApply').onclick();
  assert.match(h.node('#actionImpact').textContent,/No changes applied/);
  assert.equal(session.changes.length,0);assert.equal(session.reviewHistory.length,0);assert.equal(h.calls.refresh,0);
  assert.equal(session.reviewBusy,false);
});
test('an action opened against an older draft cannot apply stale suggestions',async()=>{
  const session=registry(),h=reviewHarness(session),issue=session.rawResult.findings.find(f=>f.rule.id==='metadata.system-upn-mismatch');
  h.api.openActionDialog('System mismatch',[issue]);session.changesRev++;
  await h.node('#actionApply').onclick();
  assert.match(h.node('#actionImpact').textContent,/registry changed/);
  assert.equal(session.changes.length,0);assert.equal(h.calls.preflight,0);
});

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
    assert.equal(harness.calls.progress, 1, 'only the import ran; empty references must fail before any re-audit')
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

test('reference help is a reversible disclosure and preserves sheet selection through repaint', async () => {
  const session = registry(), harness = reviewHarness(session)
  const book = XLSX.utils.book_new()
  for (const [sheet, id] of [['Current', 'L2-DEMO-1'], ['Alternate', 'L2-DEMO-2']]) {
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['L2 ID', 'L1 ID'], [id, 'L1-DEMO-1']]), sheet)
  }
  const input = { dataset: { referenceFile: 'milestones' }, files: [{ bytes: XLSX.write(book, { type: 'array', bookType: 'xlsx' }) }] }
  const select = { dataset: { referenceSheet: 'milestones' }, value: 'Alternate' }
  harness.lists.set('[data-reference-file]', [input]); harness.lists.set('[data-reference-sheet]', [select])
  harness.api.openReferencesDialog()
  const before = structuredClone(session), attributes = {}
  harness.node('#referencesHelp').setAttribute = (key, value) => { attributes[key] = value }
  assert.match(harness.node('#actionModalBody').innerHTML, /aria-controls="referencesHelpBody" aria-expanded="false"/)
  harness.node('#referencesHelp').onclick()
  assert.equal(harness.node('#referencesHelpBody').hidden, false)
  assert.equal(attributes['aria-expanded'], 'true')
  await input.onchange(); select.onchange()
  assert.match(harness.node('#actionModalBody').innerHTML, /aria-expanded="true"/)
  assert.match(harness.node('#actionModalBody').innerHTML, /value="Alternate" selected/)
  harness.node('#referencesHelp').onclick()
  assert.equal(harness.node('#referencesHelpBody').hidden, true)
  assert.equal(attributes['aria-expanded'], 'false')
  assert.deepEqual(structuredClone(session), before, 'help and pending selections do not alter the active review')
  await harness.node('#referencesApply').onclick()
  assert.equal(session.references.milestones.sheetName, 'Alternate')
  assert.equal(harness.calls.refresh, 1)
})
