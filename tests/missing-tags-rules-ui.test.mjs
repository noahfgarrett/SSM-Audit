import test from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import { readFileSync } from 'node:fs'
import { S, resetSession } from '../src/state.js'
import { SSM_AUDIT_RULES } from '../src/audit/engine.js'
import { SSM_AUDIT_ENGINEERING_RULES, auditReadEngineeringAoa, AUDIT_FINDING_LEVELS } from '../src/audit/engineering-references.js'
import { auditSessionResult } from '../src/audit/review.js'
import { auditSnapshotFromAoa } from '../src/audit/model.js'
import { EXTO_REV21_COLUMNS } from '../src/exto/rev21-contract.js'
import { ruleGroupsHtml, applyFindingExclusions, applyRulePreferences } from '../src/ui/audit.js'
import { auditExportLevelGroups, buildAuditWorkbook, buildAuditActionsWorkbook } from '../src/audit/export.js'

vm.runInThisContext(readFileSync(new URL('../src/vendor/sheetjs.js',import.meta.url),'utf8'));
function result(){
  const row={equipmentId:'TEST-PUMP111-01',upn:'111',closestParent:'111 Chilled Water',systemName:'111 Chilled Water',building:'TEST',discipline:'MECHANICAL WET'};
  const snapshot=auditSnapshotFromAoa([EXTO_REV21_COLUMNS.map(c=>c.header),EXTO_REV21_COLUMNS.map(c=>row[c.field]||'')],{sheet:'Registry'});
  const mel=auditReadEngineeringAoa([['Equipment Tag','UPN'],['TEST-PUMP111-02','111'],['TEST-PUMP111-03','111']],'mel','MEL');
  return auditSessionResult(snapshot,{mel});
}
function resetRules(){S.rules={search:'',source:'all',category:'all',disabled:[]};resetSession();}

test('finding levels and severity sorting put Missing Tags directly after Rule broken',()=>{
  assert.deepEqual(AUDIT_FINDING_LEVELS,['blocker','error','missing','warning','info']);
  const ui=readFileSync(new URL('../src/ui/audit.js',import.meta.url),'utf8');
  const rank=ui.slice(ui.indexOf('function severityRank('),ui.indexOf('\n',ui.indexOf('function severityRank(')));
  const filter=ui.slice(ui.indexOf('function filteredFindings('),ui.indexOf('\nfunction findingGroup('));
  const findings=['info','missing','blocker','warning','error'].map((severity,i)=>({id:String(i),severity,rule:{id:'test',source:'reference'},searchKey:''}));
  const session={result:{findings},sort:'severity-desc'},ctx=vm.createContext({S:{session},AUDIT_FINDING_LEVELS,clean:value=>String(value||'').trim(),dimActiveSets:()=>[],DIM_KEYS:[],dimFilterMap:()=>({})});
  vm.runInContext(`${rank}\n${filter}`,ctx);
  assert.deepEqual(Array.from(ctx.filteredFindings(),finding=>finding.severity),AUDIT_FINDING_LEVELS);
  session.sort='severity-asc';
  assert.deepEqual(Array.from(ctx.filteredFindings(),finding=>finding.severity),[...AUDIT_FINDING_LEVELS].reverse());
});
test('severity palette has readable contrast and a distinct color for each level',()=>{
  const css=readFileSync(new URL('../src/styles/app.css',import.meta.url),'utf8'),inks=new Set();
  const luminance=hex=>hex.match(/../g).map(value=>parseInt(value,16)/255).map(value=>value<=.04045?value/12.92:((value+.055)/1.055)**2.4).reduce((sum,value,index)=>sum+value*[.2126,.7152,.0722][index],0);
  for(const level of AUDIT_FINDING_LEVELS){
    const colors=css.match(new RegExp(`\\.${level}\\{--severity-ink:#([a-f0-9]{6});--severity-bg:#([a-f0-9]{6});--severity-line:#([a-f0-9]{6})`));
    assert.ok(colors,level);inks.add(colors[1]);
    assert.ok((luminance(colors[2])+.05)/(luminance(colors[1])+.05)>=4.5,`${level} text contrast`);
  }
  assert.equal(inks.size,5);
});

test('missing tags occupy their own level and topic, not error, warning or note',()=>{
  const audit=result(),coverage=audit.findings.filter(f=>f.category==='missing-tags');
  assert.equal(coverage.length,3);assert.ok(coverage.every(f=>f.severity==='missing'));
  assert.equal(audit.summary.severity.missing,3);assert.equal(audit.summary.category['missing-tags'],3);
  assert.equal(Object.values(audit.summary.severity).reduce((a,b)=>a+b,0),audit.findings.length);
});
test('dismissal and rule switches recount coverage without changing registry rows',()=>{
  const audit=result(),coverage=audit.findings.filter(f=>f.severity==='missing');
  const dismissed=applyFindingExclusions(audit,new Set([coverage[0].id]));
  assert.equal(dismissed.summary.severity.missing,2);
  const disabled=applyRulePreferences(audit,[SSM_AUDIT_ENGINEERING_RULES.missing.id]);
  assert.equal(disabled.summary.severity.missing,1);assert.equal(disabled.rows.length,1);
});
test('Missing Tags level export retains tags that have no registry row',()=>{
  const audit=result(),group=auditExportLevelGroups(audit).find(g=>g.severity==='missing');
  assert.equal(group.lines.length,3);assert.equal(group.findingCount,3);
  const book=buildAuditWorkbook(audit,'Synthetic',{layout:'level'}),sheet=book.Sheets['MISSING TAGS'];
  assert.ok(sheet);const text=JSON.stringify(XLSX.utils.sheet_to_json(sheet,{header:1}));
  for(const tag of ['TEST-PUMP111-01','TEST-PUMP111-02','TEST-PUMP111-03'])assert.ok(text.includes(tag));
  assert.equal(book.Sheets.Dashboard.E8.v,'Missing Tags');assert.equal(book.Sheets.Dashboard.E9.v,3);
  assert.equal(book.Sheets.Dashboard.D8.v,'Rule broken');assert.equal(book.Sheets.Dashboard.F8.v,'Check this');
  const bytes=XLSX.write(book,{type:'array',bookType:'xlsx'}),roundtrip=XLSX.read(bytes,{type:'array'});
  assert.ok(roundtrip.Sheets['MISSING TAGS']);assert.equal(audit.rows.length,1,'coverage report rows must not be inserted into the registry');
});
test('Missing Tags can be excluded independently from the level export',()=>{
  const book=buildAuditWorkbook(result(),'Synthetic',{layout:'level',plan:{levels:{missing:'skip'}}});
  assert.ok(!book.Sheets['MISSING TAGS']);assert.ok(!JSON.stringify(XLSX.utils.sheet_to_json(book.Sheets['All Findings'],{header:1})).includes('MISSING TAGS'));
});
test('Actions workbook uses the coverage label without calling a missing tag invalid',()=>{
  const audit=result(),findings=audit.findings.filter(f=>f.severity==='missing');
  const book=buildAuditActionsWorkbook({...audit,findings},'Synthetic');
  const text=JSON.stringify(book.SheetNames.map(name=>XLSX.utils.sheet_to_json(book.Sheets[name],{header:1})));
  assert.ok(text.includes('MISSING TAGS'));assert.equal(audit.rows.length,1);
});
test('Rules use source and topic disclosures, retaining toggles, examples and finding links',()=>{
  resetRules();const rule=SSM_AUDIT_RULES.blankParent;
  const html=ruleGroupsHtml([rule],new Map([[rule.id,2]]));
  assert.match(html,/<details class="rule-source-section/);assert.match(html,/<details class="rule-category/);
  assert.match(html,/2 found/);assert.ok(html.includes(`data-rule-toggle="${rule.id}"`));
  assert.ok(html.includes(`data-rule-example="${rule.id}"`));assert.ok(html.includes(`data-rule-findings="${rule.id}"`));
  assert.ok(!/<details[^>]*\sopen[\s>]/.test(html));
});
test('Rules remember open groups and show search matches inside opened sections',()=>{
  resetRules();const rule=SSM_AUDIT_RULES.blankParent;
  S.rules.expandedGroups={[`source:${rule.source}`]:true,[`${rule.source}:${rule.category}`]:true};
  assert.equal((ruleGroupsHtml([rule],null).match(/\sopen>/g)||[]).length,2);
  S.rules.expandedGroups={};S.rules.search='Parent';
  assert.equal((ruleGroupsHtml([rule],null,'PARENT').match(/\sopen>/g)||[]).length,2);
  S.rules.autoExpand=false;
  assert.equal((ruleGroupsHtml([rule],null).match(/\sopen>/g)||[]).length,0);
  resetRules();
});
test('Missing Tags has a dedicated Rules topic and disabled switches retain their state',()=>{
  resetRules();const rule=SSM_AUDIT_ENGINEERING_RULES.missing;S.rules.disabled=[rule.id];
  const html=ruleGroupsHtml([rule],new Map());
  assert.ok(html.includes('Missing Tags'));assert.ok(html.includes('aria-checked="false"'));assert.ok(html.includes('Not checked'));
  resetRules();
});
