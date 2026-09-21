import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import vm from 'node:vm'
import { S, resetSession } from '../src/state.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { auditSessionResult } from '../src/audit/review.js'
import { auditReadEngineeringAoa } from '../src/audit/engineering-references.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { findingContextFor, dimRowMatches, dimOptions, modifyDimMatch, dashRankBuckets, dashMilestoneReadiness, scopedResult, coverageReviewOnly } from '../src/ui/audit.js'

function fixture(){
  resetSession();
  const row={equipmentId:'TEST-PUMP111-01',building:'TEST',discipline:'MECHANICAL WET',upn:'111',systemName:'111 Chilled Water',milestone:'TEST-L2'};
  const snapshot=auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),EXTO_REV21_COLUMNS.map(c=>row[c.field]||'')],{sheet:'Registry'});
  const mel=auditReadEngineeringAoa([['Equipment Tag','Building','Discipline','UPN'],['TEST-PUMP111-02','ANNEX','MECHANICAL WET','111'],['TEST-PUMP111-03','','','']],'mel','MEL');
  const result=auditSessionResult(snapshot,{mel});
  S.session.snapshot=snapshot;S.session.rawResult=result;S.session.result=result;
  return {result,known:result.findings.find(f=>f.equipmentId==='TEST-PUMP111-02'),unknown:result.findings.find(f=>f.equipmentId==='TEST-PUMP111-03')};
}

test('source-only findings retain metadata for dimension filters without adding registry rows',()=>{
  const {known,unknown}=fixture();
  assert.equal(findingContextFor(known).building,'ANNEX');
  assert.equal(dimRowMatches(known,[['building',new Set(['ANNEX'])]]),true);
  assert.equal(dimRowMatches(known,[['building',new Set(['TEST'])]]),false);
  assert.equal(dimRowMatches(unknown,[['building',new Set([''])]]),true);
  assert.ok(dimOptions('building').some(option=>option.label==='Unassigned'&&option.key===''));
  assert.ok(dimOptions('building').some(option=>option.key==='ANNEX'));
  S.session.dimFilters.building=['ANNEX'];
  const scoped=scopedResult();
  assert.equal(scoped.rows.length,0);
  assert.deepEqual(scoped.findings.map(f=>f.id),[known.id]);
  assert.equal(S.session.snapshot.rows.length,1);
});

test('unknown discipline and milestone selections in Actions match the actual dropdown values',()=>{
  const {known,unknown}=fixture();
  S.session.modifyDiscipline='MECHANICAL WET';
  assert.equal(modifyDimMatch(known),true);assert.equal(modifyDimMatch(unknown),false);
  S.session.modifyDiscipline='none';
  assert.equal(modifyDimMatch(known),false);assert.equal(modifyDimMatch(unknown),true);
  S.session.modifyDiscipline='all';S.session.modifyMilestone='none';
  assert.equal(modifyDimMatch(known),true);
});

test('dashboard dimensions and milestone groups count coverage findings but not fabricated equipment',()=>{
  const {result,known,unknown}=fixture();
  const findings=[known,unknown],scoped={rows:result.rows,findings};
  for(const kind of ['building','discipline','upn','milestone']){
    const buckets=dashRankBuckets(kind,findings);
    assert.equal(buckets.reduce((sum,b)=>sum+b.count,0),2);
    assert.ok(buckets.every(b=>b.label));
  }
  const milestones=dashMilestoneReadiness(scoped);
  assert.equal(milestones.reduce((sum,b)=>sum+b.findings,0),2);
  assert.equal(milestones.reduce((sum,b)=>sum+b.rows,0),1);
  assert.equal(milestones.find(b=>b.label==='No L2 milestone').missing,2);
  assert.equal(coverageReviewOnly(known),true);
});

test('coverage context is rebuilt when references change',()=>{
  const {known}=fixture();assert.equal(findingContextFor(known).building,'ANNEX');
  S.session.rawResult={...S.session.rawResult,findings:S.session.rawResult.findings.map(f=>f.id===known.id?{...f,sourceContext:{...f.sourceContext,building:'OTHER'}}:f)};
  assert.equal(findingContextFor(known).building,'OTHER');
});

test('coverage review records only review status, never a correction or export-ready row',()=>{
  const source=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
  const code=source.slice(source.indexOf('function openCoverageReview('),source.indexOf('\nfunction openActionDialog('));
  const elements=new Map(),session={reviewHistory:[],changes:[],draftResolved:new Set(),reviewedIds:new Set()};
  let undo=0,closed=0,refreshed=0;
  const $=selector=>{if(!elements.has(selector))elements.set(selector,{setAttribute(){}});return elements.get(selector);};
  const context=vm.createContext({S:{session,screen:'modify'},$,esc:String,document:{activeElement:null},animateOpen(){},activateFocusTrap:()=>()=>{},closeActionDialog:()=>closed++,reviewRememberUndo:()=>undo++,setActionedMany:findings=>findings.forEach(f=>session.reviewedIds.add(f.id)),invalidateFindingCaches(){},rerenderModifications:()=>refreshed++,toast(){},crypto:{randomUUID:()=> 'review-test'},currentNavigate:()=>{}});
  vm.runInContext(`let actionScope,actionOpener,actionTrapCleanup;${code}`,context);
  context.openCoverageReview('Tag coverage',[{id:'missing-1',equipmentId:'TEST-01',why:'Not in registry'}],()=>{});
  assert.match($('#actionModalBody').innerHTML,/Registry values unchanged/);
  $('#coverageReviewed').onclick();
  assert.equal(undo,1);assert.equal(closed,1);assert.equal(refreshed,1);
  assert.equal(session.reviewedIds.has('missing-1'),true);
  assert.equal(session.reviewHistory[0].disposition,'reviewed');
  assert.equal(session.reviewHistory[0].changes.length,0);
  assert.equal(session.changes.length,0);assert.equal(session.draftResolved.size,0);
});
