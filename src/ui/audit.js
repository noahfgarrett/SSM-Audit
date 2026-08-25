import { $, $$, clean, esc, natCmp } from '../core/text.js'
import { S, resetSession } from '../state.js'
import { readArrayBuffer } from '../io/workbook.js'
import { auditNormId, auditSnapshotFromWorkbook, auditSplitReferences, auditFingerprint } from '../audit/model.js'
import { auditIsBlankItemMaster, auditPolarity, runSsmAudit, SSM_AUDIT_CATEGORIES, SSM_AUDIT_RULES, SSM_AUDIT_SEVERITIES, SSM_AUDIT_SOURCES } from '../audit/engine.js'
import { extoRev21Canonical } from '../exto/rev21-contract.js'
import { compareSsmRegistries, comparisonSystemTypes } from '../audit/compare.js'
import { buildSsmHierarchy } from '../audit/hierarchy.js'
import { auditExportPlanMode, exportSsmAuditXlsx, exportSsmComparisonXlsx } from '../audit/export.js'
import { ic } from './icons.js'
import { activateFocusTrap, copyTagHtml, runWithProgress, toast, wireCopyTags, animateOpen, animateClose } from './feedback.js'
import { AUDIT_EXAMPLE_FIELD_LABELS, SSM_AUDIT_EXAMPLES, auditExampleColumns, auditExampleSnapshot } from '../audit/examples.js'

const AUDIT_ROW_HEIGHT=64,AUDIT_OVERSCAN=18,AUDIT_MAX_ROWS=160;
const COMPARE_ROW_HEIGHT=96,COMPARE_OVERSCAN=14,COMPARE_MAX_ROWS=120;
const COMPARE_TREE_ROW_HEIGHT=52,COMPARE_TREE_OVERSCAN=10,COMPARE_TREE_MAX_ROWS=120;
const HIERARCHY_ROW_HEIGHT=52,HIERARCHY_OVERSCAN=18,HIERARCHY_MAX_ROWS=180;
const COMPARE_MAX_OBSERVATIONS=500;
const SEARCH_DEBOUNCE_MS=150;
const CATEGORY_LABELS={structure:'Structure',dependencies:'Dependencies',metadata:'Metadata',milestones:'Milestones','item-masters':'Item Masters',headers:'Headers / Rollups'};
const RULE_CATEGORY_LABELS={structure:'Hierarchy',dependencies:'Dependencies',metadata:'Registry consistency',milestones:'Milestones','item-masters':'Item Masters',headers:'Headers / Rollups'};
const RULE_CONFIDENCE_LABELS={required:'Required',strong:'Strong pattern','description-rated':'Description based'};
const RULE_SOURCE_DESCRIPTIONS={registry:'Identity, references, and metadata consistency within the registry.',sop:'Required parent-child, dependency, sequencing, and header checks.',logic:'Confidence-rated control, electrical, and process-enabling relationships.'};
const SEVERITY_LABELS={blocker:'Invalid',error:'Rule broken',warning:'Check this',info:'Note'};
const SEVERITY_PLURALS={blocker:'Invalid',error:'Rule broken',warning:'Check this',info:'Notes'};
const SOURCE_LABELS=Object.fromEntries(SSM_AUDIT_SOURCES.map(source=>[source.id,source.label]));
let drawerTrapCleanup=null,searchDebounceTimer=0;

/* ---------------------------------------------------------------- shell nav */

const NAV_SECTIONS=[
  {id:'dashboard',icon:'layout-dashboard',label:'Dashboard',hint:'Everything this registry turned up, at a glance'},
  {id:'audit',icon:'check-check',label:'Audit Findings',hint:'Every issue found in this registry'},
  {id:'modify',icon:'sliders-horizontal',label:'Modifications',hint:'Set aside findings you disagree with'},
  {id:'hierarchy',icon:'list-tree',label:'SSM Hierarchy',hint:'Browse the registry as a tree'},
  {id:'compare',icon:'square-stack',label:'Compare Projects',hint:'Line this registry up beside a finished one'},
  {id:'rules',icon:'book-open',label:'Rules',hint:'What the audit checks, in plain language'},
];
function navActiveId(){
  if(S.screen==='dashboard')return 'dashboard';
  if(S.screen==='audit')return 'audit';
  if(S.screen==='modify')return 'modify';
  if(S.screen==='hierarchy')return 'hierarchy';
  if(S.screen==='compare')return 'compare';
  if(S.screen==='rules')return 'rules';
  return S.homeMode==='compare'?'compare':S.homeMode==='rules'?'rules':'audit';
}
/* Badges never crowd the label out of the sidebar: four digits and up show
   compactly (7.1k, 23k) with the exact number in the tooltip. */
function navCount(value){
  const count=Number(value)||0;
  const text=count>=10000?`${Math.round(count/1000)}k`:count>=1000?`${(count/1000).toFixed(1).replace(/\.0$/,'')}k`:String(count);
  return {text,title:count.toLocaleString()};
}
function navCountHtml(value,extraClass,titleSuffix){
  const {text,title}=navCount(value);
  return `<b class="nav-count${extraClass?` ${extraClass}`:''}" title="${esc(title)}${titleSuffix?` ${titleSuffix}`:''}">${esc(text)}</b>`;
}
function navBadges(id){
  const result=S.session&&S.session.result,comparison=S.comparison&&S.comparison.result;
  if(id==='audit'&&result){
    const blockers=result.summary.severity.blocker;
    return `${navCountHtml(result.summary.findings,'','findings')}${blockers?navCountHtml(blockers,'blocker',SEVERITY_LABELS.blocker.toLowerCase()):''}`;
  }
  if(id==='modify'&&S.session&&S.session.rawResult){const aside=excludedInBase();return aside?navCountHtml(aside,'','set aside'):'';}
  if(id==='hierarchy'&&S.session&&S.session.snapshot)return navCountHtml(S.session.snapshot.rows.length,'','rows');
  if(id==='compare'&&comparison)return navCountHtml(comparison.summary.differentSystems+comparison.summary.targetOnlySystems+comparison.summary.referenceOnlySystems,'','systems with differences');
  return '';
}
function navDisabled(id){
  if(id==='hierarchy')return !(S.session&&S.session.result&&S.session.snapshot);
  if(id==='dashboard')return !(S.session&&S.session.result);
  if(id==='modify')return !(S.session&&S.session.rawResult);
  return false;
}
function navItemHtml(section,active,collapsed){
  const disabled=navDisabled(section.id),title=collapsed||disabled?`${section.label}${disabled?' — run an audit first':''}`:section.hint;
  return `<button class="sidenav-item ${active===section.id?'active':''}" type="button" data-nav="${section.id}" ${disabled?'disabled':''} title="${esc(title)}" aria-current="${active===section.id?'page':'false'}">${ic(section.icon)}<span>${esc(section.label)}</span>${navBadges(section.id)}</button>`;
}
export function renderSideNav(navigate){
  const nav=$('#sideNav');if(!nav)return;
  const active=navActiveId(),collapsed=!!S.ui.navCollapsed,canExport=S.screen==='compare'?!!(S.comparison&&S.comparison.result):!!(S.session&&S.session.result);
  nav.classList.toggle('collapsed',collapsed);document.body.classList.toggle('nav-collapsed',collapsed);
  nav.innerHTML=`<div class="sidenav-items">${NAV_SECTIONS.map(section=>navItemHtml(section,active,collapsed)).join('')}</div>
  <div class="sidenav-foot">
    <button class="sidenav-item" type="button" id="navGuide" title="Guide">${ic('book-open')}<span>Guide</span></button>
    <button class="sidenav-item" type="button" id="navExport" ${canExport?'':'disabled'} title="${canExport?'Download an Excel report':'Run an audit first'}">${ic('file-down')}<span>Export</span></button>
    <button class="sidenav-collapse" type="button" id="navCollapse" title="${collapsed?'Expand menu':'Collapse menu'}" aria-label="${collapsed?'Expand menu':'Collapse menu'}">${ic(collapsed?'panel-left-open':'panel-left-close')}<span>Collapse</span></button>
  </div>`;
  $$('[data-nav]',nav).forEach(button=>button.onclick=()=>navigateSection(button.dataset.nav,navigate));
  $('#navGuide').onclick=()=>document.dispatchEvent(new CustomEvent('ssm-audit:guide'));
  $('#navExport').onclick=()=>{if(S.screen==='compare')exportSsmComparisonXlsx();else openExportOptions();};
  $('#navCollapse').onclick=()=>{S.ui.navCollapsed=!S.ui.navCollapsed;renderSideNav(navigate);$('#navCollapse')?.focus();};
}
function navigateSection(id,navigate){
  if(id==='dashboard'){S.homeMode='audit';navigate(S.session&&S.session.result?'dashboard':'upload');return;}
  if(id==='rules'){S.homeMode='rules';navigate('rules');return;}
  if(id==='compare'){S.homeMode='compare';navigate(S.comparison&&S.comparison.result?'compare':'upload');return;}
  if(id==='modify'){if(S.session&&S.session.rawResult)navigate('modify');return;}
  if(id==='hierarchy'){if(S.session&&S.session.result)navigate('hierarchy');return;}
  S.homeMode='audit';navigate(S.session&&S.session.result?'audit':'upload');
}

/* ------------------------------------------------------------ small helpers */

function highlightHtml(value,query){
  const text=String(value==null?'':value);if(!query)return esc(text);
  const upper=text.toUpperCase();let out='',from=0,at=upper.indexOf(query);
  while(at>=0){out+=esc(text.slice(from,at))+'<mark>'+esc(text.slice(at,at+query.length))+'</mark>';from=at+query.length;at=upper.indexOf(query,from);}
  return out+esc(text.slice(from));
}
function debounceSearch(run){clearTimeout(searchDebounceTimer);searchDebounceTimer=setTimeout(run,SEARCH_DEBOUNCE_MS);}
function severityRank(severity){return {blocker:4,error:3,warning:2,info:1}[severity]||0;}

/* ------------------------------------------------------------- rules screen */

function ruleCatalog(){return Object.values(SSM_AUDIT_RULES);}
/* ---- rule preferences ----
   The engine runs every check; the user decides which ones count. Switched-off
   checks are dropped from the result the rest of the app sees (findings,
   dashboard, hierarchy badges, export), and the choice is remembered on this
   device. Nothing about the registry is stored -- only rule ids. */
const RULE_PREFERENCES_KEY='ssm-audit.disabled-rules';
export function loadRulePreferences(){
  try{const raw=localStorage.getItem(RULE_PREFERENCES_KEY);const list=raw?JSON.parse(raw):[];S.rules.disabled=Array.isArray(list)?list.filter(id=>typeof id==='string'):[];}catch(_){S.rules.disabled=[];}
}
function saveRulePreferences(){try{localStorage.setItem(RULE_PREFERENCES_KEY,JSON.stringify(S.rules.disabled));}catch(_){/* private mode: the choice lasts for this session */}}
export function isRuleActive(rule){return !!rule&&rule.enabled&&!(S.rules.disabled||[]).includes(rule.id);}
export function activeRules(){return ruleCatalog().filter(isRuleActive);}
/* The result the app works from: the engine's output minus switched-off checks,
   with the summary recounted. `rawResult` keeps the engine output so toggling
   never needs the workbook again. */
export function applyRulePreferences(raw,disabled){
  if(!raw)return raw;const off=new Set(disabled||[]);
  if(!off.size)return raw;
  const findings=raw.findings.filter(finding=>!off.has(finding.rule.id));
  const severity={blocker:0,error:0,warning:0,info:0};for(const finding of findings)severity[finding.severity]=(severity[finding.severity]||0)+1;
  const status=severity.blocker?'blocked':raw.summary.status==='blocked'?'ready':raw.summary.status;
  return Object.assign({},raw,{findings,summary:Object.assign({},raw.summary,{findings:findings.length,severity,status})});
}
/* Individual findings the user set aside on the Modifications screen are dropped
   the same way switched-off checks are, with the summary recounted properly. */
export function applyFindingExclusions(result,excluded){
  if(!result||!excluded||!excluded.size)return result;
  const findings=result.findings.filter(finding=>!excluded.has(finding.id));
  if(findings.length===result.findings.length)return result;
  const severity={blocker:0,error:0,warning:0,info:0};for(const finding of findings)severity[finding.severity]=(severity[finding.severity]||0)+1;
  const status=severity.blocker?'blocked':severity.error||severity.warning?'review':'ready';
  return Object.assign({},result,{findings,summary:Object.assign({},result.summary,{findings:findings.length,severity,status})});
}
function refreshSessionResult(){
  if(!S.session||!S.session.rawResult)return;
  S.session.result=applyFindingExclusions(applyRulePreferences(S.session.rawResult,S.rules.disabled),S.session.excluded);
  S.session.hierarchy=null;S.session.hierarchyCacheKey='';S.session.hierarchyCacheRows=null;S.session.headerIdSet=null;S.session.rowIndex=null;S.session.exportPlan=null;
  invalidateFindingCaches();
}
function setRuleDisabled(ruleId,off){
  const set=new Set(S.rules.disabled||[]);if(off)set.add(ruleId);else set.delete(ruleId);
  S.rules.disabled=[...set];saveRulePreferences();refreshSessionResult();
}
/* ---- actioned in the app ----
   Marks live in localStorage keyed by a fingerprint of the registry's contents,
   so reloading the same workbook (even renamed) brings the progress back.
   Finding ids are content hashes, so they line up run after run. */
function registryStoreKey(prefix){
  const snapshot=S.session&&S.session.snapshot;if(!snapshot)return '';
  const rows=snapshot.rows||[];
  return 'ssm-audit.'+prefix+'.'+auditFingerprint(rows.length+'|'+rows.slice(0,64).map(row=>auditNormId(row.equipmentId)).join(','));
}
function registryActionKey(){return registryStoreKey('actioned');}
function loadActioned(){
  S.session.actioned=new Set();S.session.actionedRev=0;
  try{const raw=localStorage.getItem(registryActionKey());if(raw)S.session.actioned=new Set(JSON.parse(raw).filter(id=>typeof id==='string'));}catch(_){}
}
function saveActioned(){try{const key=registryActionKey();if(key)localStorage.setItem(key,JSON.stringify([...S.session.actioned]));}catch(_){/* private mode: kept for this session */}}
export function isActioned(finding){return !!(finding&&S.session&&S.session.actioned&&S.session.actioned.has(finding.id));}
function setActioned(finding,on){
  if(!S.session.actioned)S.session.actioned=new Set();
  if(on)S.session.actioned.add(finding.id);else S.session.actioned.delete(finding.id);
  S.session.actionedRev=(S.session.actionedRev||0)+1;saveActioned();
}
function actionedInResult(){
  const result=S.session&&S.session.result;if(!result||!S.session.actioned||!S.session.actioned.size)return 0;
  let count=0;for(const finding of result.findings)if(S.session.actioned.has(finding.id))count++;return count;
}
/* ---- modifications: findings set aside ----
   The user's per-finding overrides. A set-aside finding disappears from every
   metric -- findings list, dashboard, hierarchy badges, export -- until it is
   restored on the Modifications screen. Stored like actioned marks: keyed by the
   registry's content fingerprint, so the same workbook brings the decisions back. */
function loadExcluded(){
  S.session.excluded=new Set();S.session.excludedRev=0;
  try{const raw=localStorage.getItem(registryStoreKey('excluded'));if(raw)S.session.excluded=new Set(JSON.parse(raw).filter(id=>typeof id==='string'));}catch(_){}
}
function saveExcluded(){try{const key=registryStoreKey('excluded');if(key)localStorage.setItem(key,JSON.stringify([...S.session.excluded]));}catch(_){/* private mode: kept for this session */}}
export function isExcludedId(id){return !!(S.session&&S.session.excluded&&S.session.excluded.has(id));}
function setExcluded(id,on){
  if(!S.session.excluded)S.session.excluded=new Set();
  if(on)S.session.excluded.add(id);else S.session.excluded.delete(id);
  S.session.excludedRev=(S.session.excludedRev||0)+1;saveExcluded();refreshSessionResult();
}
/* The Modifications screen works from the result BEFORE exclusions (but after
   switched-off checks), so set-aside findings stay visible to restore. */
function modifyBaseResult(){
  const raw=S.session&&S.session.rawResult;if(!raw)return null;
  return applyRulePreferences(raw,S.rules.disabled);
}
function excludedInBase(){
  const excluded=S.session&&S.session.excluded;if(!excluded||!excluded.size)return 0;
  const base=modifyBaseResult();if(!base)return 0;
  let count=0;for(const finding of base.findings)if(excluded.has(finding.id))count++;return count;
}
function ruleFindingCounts(){
  const result=S.session&&S.session.result;if(!result)return null;
  const counts=new Map();for(const finding of result.findings)counts.set(finding.rule.id,(counts.get(finding.rule.id)||0)+1);return counts;
}
function ruleCatalogRows(){
  const query=clean(S.rules.search).toUpperCase(),source=S.rules.source,category=S.rules.category;
  return ruleCatalog().filter(rule=>(source==='all'||rule.source===source)&&(category==='all'||rule.category===category)&&(!query||[rule.title,rule.statement,SOURCE_LABELS[rule.source],RULE_CATEGORY_LABELS[rule.category]].join(' ').toUpperCase().includes(query)));
}
function ruleRowHtml(rule,counts,query){
  const confidence=RULE_CONFIDENCE_LABELS[rule.confidence]||'Guidance',found=counts?counts.get(rule.id)||0:null,active=isRuleActive(rule);
  const findings=found===null?'<span class="rule-slot-empty" aria-hidden="true"></span>':found?`<button class="rule-finding-count" type="button" data-rule-findings="${esc(rule.id)}" title="Show only these findings">${found.toLocaleString()} found${ic('arrow-right')}</button>`:`<span class="rule-finding-count clear">${ic('check')}None found</span>`;
  return `<article class="rule-reference-row ${active?'':'is-off'}">
    <span class="rule-reference-icon">${ic(rule.category==='dependencies'?'git-branch':rule.category==='metadata'?'database':rule.category==='headers'?'folder-tree':rule.category==='milestones'?'clipboard-list':rule.category==='item-masters'?'tag':'list-tree')}</span>
    <div><h4>${highlightHtml(rule.title,query)}</h4><p>${highlightHtml(rule.statement,query)}</p></div>
    <div class="rule-reference-tags">${SSM_AUDIT_EXAMPLES[rule.id]?`<button class="rule-example-btn" type="button" data-rule-example="${esc(rule.id)}" title="See a worked example of what this check flags">${ic('eye')}Example</button>`:'<span class="rule-slot-empty" aria-hidden="true"></span>'}<span class="confidence-${esc(rule.confidence)}">${esc(confidence)}</span>${rule.enabled?`<button class="rule-switch ${active?'on':''}" type="button" role="switch" aria-checked="${active?'true':'false'}" data-rule-toggle="${esc(rule.id)}" title="${active?'Checked — click to switch this check off':'Switched off — click to check it again'}"><i></i><span>${active?'On':'Off'}</span></button>`:`<span class="rule-state off">Off</span>`}${active?findings:'<span class="rule-finding-count muted">Not checked</span>'}</div>
  </article>`;
}

/* ---- export options modal ----
   One step between Export and the file: per level (and per fired check) choose
   Include, Pre-ticked, or Leave out. The plan lives on the session, so the
   next report starts from the same choices. */
const EXPORT_MODE_LABELS={include:'Include',pretick:'Pre-ticked',skip:'Leave out'};
let exportTrapCleanup=null,exportOpener=null;
function exportPlan(){S.session.exportPlan=S.session.exportPlan||{levels:{},rules:{}};return S.session.exportPlan;}
function exportPlanSummary(result,plan){
  let included=0,preticked=0,skipped=0;
  for(const finding of result.findings){const mode=auditExportPlanMode(plan,finding);if(mode==='skip')skipped++;else{included++;if(mode==='pretick'||isActioned(finding))preticked++;}}
  return {included,preticked,skipped};
}
function exportModeGroupHtml(name,current,compact){
  return `<div class="export-modes${compact?' sm':''}" role="radiogroup">${['include','pretick','skip'].map(mode=>`<label class="export-mode ${mode} ${current===mode?'on':''}"><input type="radio" name="${esc(name)}" value="${mode}" ${current===mode?'checked':''}>${EXPORT_MODE_LABELS[mode]}</label>`).join('')}</div>`;
}
export function closeExportOptions(){
  const modal=$('#exportModal');if(!modal||!modal.classList.contains('show'))return;
  exportTrapCleanup?.();exportTrapCleanup=null;modal.setAttribute('aria-hidden','true');animateClose(modal);
  const opener=exportOpener;exportOpener=null;if(opener&&document.contains(opener)&&typeof opener.focus==='function')opener.focus();
}
function renderExportOptions(){
  const result=S.session.result,plan=exportPlan();
  const bySeverity=new Map();
  for(const finding of result.findings){const list=bySeverity.get(finding.severity)||new Map();list.set(finding.rule.id,{rule:finding.rule,count:(list.get(finding.rule.id)?.count||0)+1});bySeverity.set(finding.severity,list);}
  const summary=exportPlanSummary(result,plan);
  const levelSections=SSM_AUDIT_SEVERITIES.filter(severity=>bySeverity.has(severity)).map(severity=>{
    const checks=[...bySeverity.get(severity).values()].sort((left,right)=>right.count-left.count);
    const total=checks.reduce((sum,entry)=>sum+entry.count,0);
    const levelMode=plan.levels[severity]==='pretick'||plan.levels[severity]==='skip'?plan.levels[severity]:'include';
    const overrides=checks.filter(entry=>plan.rules[entry.rule.id]&&plan.rules[entry.rule.id]!==levelMode).length;
    return `<section class="export-level ${esc(severity)}">
      <header><span class="audit-severity ${esc(severity)}">${esc(SEVERITY_LABELS[severity])}</span><b>${total.toLocaleString()} finding${total===1?'':'s'} &middot; ${checks.length} check${checks.length===1?'':'s'}</b>${exportModeGroupHtml('level-'+severity,levelMode,false)}</header>
      <details ${overrides?'open':''}><summary>Fine-tune the ${checks.length} check${checks.length===1?'':'s'}${overrides?` &middot; ${overrides} overridden`:''}</summary>
        <div class="export-checks">${checks.map(entry=>{
          const ruleMode=plan.rules[entry.rule.id]&&plan.rules[entry.rule.id]!==levelMode?plan.rules[entry.rule.id]:levelMode;
          return `<div class="export-check"><div><b>${esc(entry.rule.title)}</b><span>${entry.count.toLocaleString()} finding${entry.count===1?'':'s'}</span></div>${exportModeGroupHtml('rule-'+entry.rule.id,ruleMode,true)}</div>`;
        }).join('')}</div>
      </details>
    </section>`;
  }).join('');
  const layout=plan.layout==='level'?'level':'milestone';
  $('#exportModalBody').innerHTML=`<span class="eyebrow">Excel report</span><h3 id="exportTitle">Choose what goes in the report</h3>
    <div class="export-layout"><b>Report layout</b>
      <label class="export-layout-choice ${layout==='milestone'?'on':''}"><input type="radio" name="export-layout" value="milestone" ${layout==='milestone'?'checked':''}><span><b>One tab per L2 milestone</b><small>The full equipment tree of each phase, findings beside it — work phase by phase.</small></span></label>
      <label class="export-layout-choice ${layout==='level'?'on':''}"><input type="radio" name="export-layout" value="level" ${layout==='level'?'checked':''}><span><b>One tab per finding level</b><small>Invalid, Rule broken, Check this, Note — flagged equipment only, ordered by milestone.</small></span></label>
    </div>
    <p class="export-intro">${layout==='level'?'Level tabs list only flagged equipment; the milestone layout carries the full tree.':'Every equipment row is always exported.'} Findings follow the choice for their level — or for the individual check. <b>Pre-ticked</b> findings arrive with the Actioned box already ${esc('☑')} (the block is green and counts as done); <b>Leave out</b> findings are not written at all.</p>
    ${levelSections||'<p class="export-intro">This registry has no findings — the report will contain the equipment tree only.</p>'}
    <footer class="export-foot">
      <span id="exportSummary">${summary.included.toLocaleString()} exported &middot; ${summary.preticked.toLocaleString()} pre-ticked &middot; ${summary.skipped.toLocaleString()} left out</span>
      <div><button class="btn ghost" type="button" id="exportReset">Reset</button><button class="btn primary" type="button" id="exportGo">${ic('file-down')}Export report</button></div>
    </footer>`;
  const refreshSummary=()=>{const next=exportPlanSummary(result,exportPlan());$('#exportSummary').textContent=`${next.included.toLocaleString()} exported · ${next.preticked.toLocaleString()} pre-ticked · ${next.skipped.toLocaleString()} left out`;};
  $$('#exportModalBody input[type=radio]').forEach(input=>input.onchange=()=>{
    const name=input.name,mode=input.value,current=exportPlan();
    if(name==='export-layout'){current.layout=mode;renderExportOptions();}
    else if(name.startsWith('level-')){
      const severity=name.slice(6);current.levels[severity]=mode;
      /* A level change clears its checks' overrides — the level now speaks for them. */
      for(const map of [bySeverity.get(severity)||new Map()])for(const id of map.keys())delete current.rules[id];
      renderExportOptions();
    }else{
      const ruleId=name.slice(5);current.rules[ruleId]=mode;
      $$(`#exportModalBody label`).forEach(()=>{});input.closest('.export-modes').querySelectorAll('.export-mode').forEach(label=>label.classList.toggle('on',label.querySelector('input').checked));
      refreshSummary();
    }
  });
  $('#exportReset').onclick=()=>{S.session.exportPlan={levels:{},rules:{}};renderExportOptions();};
  $('#exportGo').onclick=async()=>{const plan=JSON.parse(JSON.stringify(exportPlan()));closeExportOptions();await exportSsmAuditXlsx(plan);};
}
export function openExportOptions(){
  if(!(S.session&&S.session.result)){toast('Run an SSM Audit first');return;}
  renderExportOptions();
  const modal=$('#exportModal');exportOpener=document.activeElement;animateOpen(modal);modal.setAttribute('aria-hidden','false');
  exportTrapCleanup?.();exportTrapCleanup=activateFocusTrap(modal,closeExportOptions);
  $('#exportModalClose').onclick=closeExportOptions;modal.onclick=event=>{if(event.target===modal)closeExportOptions();};
  $('#exportGo').focus();
}

/* ---- worked example modal ----
   The mock rows are run through the real engine when the modal opens, so the
   "What the audit says" line is the live finding text, not a copy of it. */
let exampleTrapCleanup=null,exampleOpener=null;let currentNavigate=null;
function exampleCellHtml(example,rowIndex,field){
  const value=String(example.rows[rowIndex][field]||'');
  const mark=example.marks.find(entry=>entry.row===rowIndex&&entry.field===field);
  const focused=example.focus.some(cell=>cell.row===rowIndex&&cell.field===field);
  let inner;
  if(mark)inner=mark.parts.map(([text,bad])=>bad?`<mark class="ex-bad">${esc(text)}</mark>`:esc(text)).join('');
  else inner=value?esc(value):'<span class="ex-empty">(blank)</span>';
  return `<td class="${focused?'ex-flag':''}">${inner}</td>`;
}
export function closeRuleExample(){
  const modal=$('#exampleModal');if(!modal||!modal.classList.contains('show'))return;
  exampleTrapCleanup?.();exampleTrapCleanup=null;modal.setAttribute('aria-hidden','true');animateClose(modal);
  const opener=exampleOpener;exampleOpener=null;if(opener&&document.contains(opener)&&typeof opener.focus==='function')opener.focus();
}
export function openRuleExample(ruleId){
  const rule=Object.values(SSM_AUDIT_RULES).find(entry=>entry.id===ruleId),example=SSM_AUDIT_EXAMPLES[ruleId];if(!rule||!example)return;
  const result=runSsmAudit(auditExampleSnapshot(example),example.options||{});
  const focusedTags=new Set(example.focus.map(cell=>example.rows[cell.row].equipmentId));
  const finding=result.findings.find(entry=>entry.rule.id===ruleId&&focusedTags.has(entry.equipmentId))||result.findings.find(entry=>entry.rule.id===ruleId)||null;
  const columns=auditExampleColumns(example);
  const table=`<div class="ex-table-wrap"><table class="ex-table"><thead><tr>${columns.map(field=>`<th>${esc(AUDIT_EXAMPLE_FIELD_LABELS[field]||field)}</th>`).join('')}</tr></thead><tbody>${example.rows.map((row,rowIndex)=>`<tr class="${example.focus.some(cell=>cell.row===rowIndex)?'ex-row-flag':''}">${columns.map(field=>exampleCellHtml(example,rowIndex,field)).join('')}</tr>`).join('')}</tbody></table></div>`;
  const severity=finding?finding.severity:'';
  $('#exampleBody').innerHTML=`<span class="eyebrow">Worked example · mock data</span><h3 id="exampleTitle">${esc(rule.title)}</h3><p class="ex-statement">${esc(rule.statement)}</p>
    <div class="ex-legend"><span><i class="ex-swatch flag"></i>Cell the check is about</span><span><i class="ex-swatch bad"></i>The exact characters</span></div>
    ${table}
    <div class="ex-notes">
      <div class="ex-note why"><b>Why it is flagged</b><p>${esc(example.caption)}</p></div>
      <div class="ex-note says"><b>What the audit says</b>${finding?`<p><span class="audit-severity ${esc(severity)}">${esc(SEVERITY_LABELS[severity]||severity)}</span> ${esc(finding.why)}</p>`:'<p>Run on this data, the check did not fire — the example needs attention.</p>'}</div>
      <div class="ex-note fix"><b>What right looks like</b><p>${esc(example.fix)}</p></div>
    </div>`;
  const modal=$('#exampleModal');exampleOpener=document.activeElement;animateOpen(modal);modal.setAttribute('aria-hidden','false');
  exampleTrapCleanup?.();exampleTrapCleanup=activateFocusTrap(modal,closeRuleExample);
  $('#exampleClose').onclick=closeRuleExample;modal.onclick=event=>{if(event.target===modal)closeRuleExample();};
  const firstFlag=$('#exampleBody .ex-table td.ex-flag'),wrap=$('#exampleBody .ex-table-wrap');if(firstFlag&&wrap)wrap.scrollLeft=Math.max(0,firstFlag.offsetLeft-wrap.clientWidth/2+firstFlag.offsetWidth/2);
  $('#exampleClose').focus();
}
function renderRuleCatalog(navigate){
  const rows=ruleCatalogRows(),container=$('#ruleCatalog'),count=$('#ruleResultCount'),counts=ruleFindingCounts(),query=clean(S.rules.search).toUpperCase();if(!container)return;
  if(count)count.textContent=`${rows.length} of ${ruleCatalog().length} checks`;
  if(!rows.length){container.innerHTML=`<div class="rule-reference-empty">${ic('search')}<b>No checks match that search</b><span>Try a shorter word, or set both menus back to All.</span><button class="btn" type="button" id="ruleClear">Clear the search</button></div>`;
    $('#ruleClear').onclick=()=>{S.rules.search='';S.rules.source='all';S.rules.category='all';renderRules(navigate);};return;}
  container.innerHTML=SSM_AUDIT_SOURCES.map(source=>{
    const matches=rows.filter(rule=>rule.source===source.id);if(!matches.length)return '';
    const categories=[...new Set(matches.map(rule=>rule.category))].sort((a,b)=>natCmp(RULE_CATEGORY_LABELS[a]||a,RULE_CATEGORY_LABELS[b]||b));
    return `<section class="rule-source-section"><header><span>${ic(source.id==='sop'?'book-open':source.id==='logic'?'network':'shield-check')}</span><div><h3>${esc(source.label)}</h3><p>${esc(RULE_SOURCE_DESCRIPTIONS[source.id]||source.description)}</p></div><b>${matches.length}</b></header>
      ${categories.map(category=>`<div class="rule-category"><h5>${esc(RULE_CATEGORY_LABELS[category]||category)}<span>${matches.filter(rule=>rule.category===category).length}</span></h5>${matches.filter(rule=>rule.category===category).map(rule=>ruleRowHtml(rule,counts,query)).join('')}</div>`).join('')}</section>`;
  }).join('');
  $$('[data-rule-findings]',container).forEach(button=>button.onclick=()=>showOnlyRule(button.dataset.ruleFindings,navigate));
  $$('[data-rule-example]',container).forEach(button=>button.onclick=()=>openRuleExample(button.dataset.ruleExample));
  $$('[data-rule-toggle]',container).forEach(button=>button.onclick=()=>{setRuleDisabled(button.dataset.ruleToggle,button.getAttribute('aria-checked')==='true');renderRules(navigate);});
  const offCount=(S.rules.disabled||[]).filter(id=>ruleCatalog().some(rule=>rule.id===id&&rule.enabled)).length,banner=$('#ruleOffBanner');
  if(banner){banner.hidden=!offCount;banner.querySelector('b').textContent=`${offCount} check${offCount===1?' is':'s are'} switched off`;}
  const allOn=$('#ruleAllOn');if(allOn)allOn.onclick=()=>{S.rules.disabled=[];saveRulePreferences();refreshSessionResult();renderRules(navigate);};
}
/* `keepScope` leaves the four registry dimensions alone. The dashboard passes it
   so clicking a number inside a scoped dashboard stays inside that scope; the
   Rules screen does not, because "42 found" there means 42 across the registry. */
function showOnlyRule(ruleId,navigate,keepScope){
  const result=S.session&&S.session.result;if(!result)return;
  const others=[...new Set(result.findings.map(finding=>finding.rule.id))].filter(id=>id!==ruleId);
  S.session.hiddenRules=others;S.session.hiddenSources=[];S.session.hiddenSeverities=[];S.session.hiddenCategories=[];S.session.search='';S.session.scrollTop=0;S.session.cursor=-1;
  if(!keepScope)clearDimFilters();
  invalidateFindingCaches();S.homeMode='audit';navigate('audit');
}

/* ------------------------------------------------------ modifications screen */
/* Every fired check, grouped by topic, with each match individually ticked.
   Unticking sets the finding aside: it leaves every metric until restored here.
   Long lists render in chunks so a 20k-findings registry stays responsive. */
const MODIFY_CHUNK=300;
function modifyQueryNorm(){return auditNormId(clean(S.session&&S.session.modifySearch));}
/* Marks every place the search hits, matching the way the search itself works:
   case-insensitive with whitespace runs collapsed, so "UPN 650" lights up in
   "UPN  650" too. The normalized text keeps a map back to the original indices. */
function modifyHighlight(text){
  const value=String(text==null?'':text),query=modifyQueryNorm();
  if(!query)return esc(value);
  let norm='';const map=[];
  for(let i=0;i<value.length;i++){
    const ch=value[i];
    if(/\s/.test(ch)){if(norm&&norm[norm.length-1]!==' '){norm+=' ';map.push(i);}continue;}
    norm+=ch.toUpperCase();map.push(i);
  }
  let out='',from=0,at=norm.indexOf(query);
  while(at>=0){
    const start=map[at],end=map[at+query.length-1]+1;
    if(start>=from){out+=esc(value.slice(from,start))+'<mark>'+esc(value.slice(start,end))+'</mark>';from=end;}
    at=norm.indexOf(query,at+query.length);
  }
  return out+esc(value.slice(from));
}
/* Grouping walks every finding, so the result is memoized on what it depends
   on -- the search, the exclusion set, and the switched-off checks. Count
   updates and expand-all then reuse one pass instead of one per card. */
function modifyGroups(){
  const key=modifyQueryNorm()+''+(S.session&&S.session.excludedRev||0)+''+(S.rules.disabled||[]).join(',');
  if(S.session&&S.session.modifyGroupsVal&&S.session.modifyGroupsKey===key)return S.session.modifyGroupsVal;
  const groups=modifyGroupsBuild();
  if(S.session){S.session.modifyGroupsKey=key;S.session.modifyGroupsVal=groups;}
  return groups;
}
function modifyGroupsBuild(){
  const base=modifyBaseResult();if(!base)return [];
  const query=modifyQueryNorm();
  const byRule=new Map();
  for(const finding of base.findings){
    let entry=byRule.get(finding.rule.id);
    if(!entry){entry={rule:finding.rule,findings:[],matches:[],kept:0};byRule.set(finding.rule.id,entry);}
    entry.findings.push(finding);
    if(!isExcludedId(finding.id))entry.kept++;
    if(!query||finding.searchKey.includes(query))entry.matches.push(finding);
  }
  return Object.keys(RULE_CATEGORY_LABELS).map(category=>{
    const rules=[...byRule.values()].filter(entry=>entry.rule.category===category&&(!query||entry.matches.length)).sort((left,right)=>right.findings.length-left.findings.length);
    const total=rules.reduce((sum,entry)=>sum+entry.findings.length,0),kept=rules.reduce((sum,entry)=>sum+entry.kept,0);
    return {category,label:RULE_CATEGORY_LABELS[category]||category,rules,total,kept};
  }).filter(group=>group.rules.length);
}
function modifyRowHtml(finding,compact){
  const excluded=isExcludedId(finding.id);
  const tag=finding.equipmentId?`<button type="button" class="modify-tag" data-mod-open="${esc(finding.id)}" title="Open this finding's details">${modifyHighlight(finding.equipmentId)}</button>`:'<b class="modify-tag"><i>Registry-wide</i></b>';
  if(compact)return `<label class="modify-row compact ${excluded?'is-excluded':''}"><input type="checkbox" data-mod-finding="${esc(finding.id)}" ${excluded?'':'checked'}>${tag}<span class="modify-where">${esc(finding.sheet||'Registry')} &middot; row ${finding.row||'—'}</span></label>`;
  return `<label class="modify-row ${excluded?'is-excluded':''}"><input type="checkbox" data-mod-finding="${esc(finding.id)}" ${excluded?'':'checked'}><span class="audit-severity ${esc(finding.severity)}">${esc(SEVERITY_LABELS[finding.severity]||finding.severity)}</span>${tag}<span class="modify-why" title="${esc(finding.why)}">${modifyHighlight(finding.why)}</span><span class="modify-where">Row ${finding.row||'—'}</span></label>`;
}
/* Findings with the same explanation are one pattern -- "row on UPN 603, parent
   on UPN RR" -- so a whole family of matches is kept or set aside with a single
   box. Explanations unique to one row stay as plain rows. The map lets the
   delegated handlers find a pattern's findings without re-deriving the groups. */
let modifyPatternMap=new Map();
function modifyPatternKey(ruleId,why){return auditFingerprint(ruleId+'|'+(why||''));}
function modifyPatterns(entry){
  const byWhy=new Map();
  for(const finding of entry.matches){const why=finding.why||'';const list=byWhy.get(why)||[];list.push(finding);byWhy.set(why,list);}
  const groups=[],singles=[];
  for(const [why,findings] of byWhy){if(findings.length>1)groups.push({why,findings,key:modifyPatternKey(entry.rule.id,why)});else singles.push(findings[0]);}
  groups.sort((left,right)=>right.findings.length-left.findings.length);
  return {groups,singles};
}
function modifyPatternCountText(findings){
  const kept=findings.filter(finding=>!isExcludedId(finding.id)).length,pctMode=S.session&&S.session.modifyCountMode==='percent';
  if(kept!==findings.length)return `${modifyAmount(kept)} of ${modifyAmount(findings.length)} kept`;
  return pctMode?`${modifyAmount(findings.length)} of tags`:`${findings.length.toLocaleString()} kept`;
}
/* The equipment list inside a pattern is only built the first time it opens. */
function modifyPatternHtml(group){
  const kept=group.findings.filter(finding=>!isExcludedId(finding.id)).length;
  const severity=group.findings[0].severity;
  return `<div class="modify-pattern ${kept?'':'is-excluded'}" data-mod-pattern="${group.key}">
    <div class="modify-pattern-head"><input type="checkbox" data-mod-group="${group.key}" ${kept===group.findings.length?'checked':''} aria-label="Keep every finding of this pattern"><span class="audit-severity ${esc(severity)}">${esc(SEVERITY_LABELS[severity]||severity)}</span><span class="modify-pattern-why">${modifyHighlight(group.why)}</span><b class="modify-pattern-count" data-mod-group-count="${group.key}">${modifyPatternCountText(group.findings)}</b>${modifyPctBtn()}<button class="btn ghost sm modify-pattern-expand" type="button" data-mod-expand="${group.key}" aria-expanded="false">${ic('chevron-down')}Equipment</button></div>
    <div class="modify-pattern-rows" hidden data-mod-empty="1"></div>
  </div>`;
}
function modifyFillPatternRows(pattern){
  const body=pattern.querySelector('.modify-pattern-rows');if(!body||!body.dataset.modEmpty)return;
  delete body.dataset.modEmpty;
  const key=pattern.dataset.modPattern,findings=modifyPatternMap.get(key)||[];
  const rest=findings.length-MODIFY_CHUNK;
  body.innerHTML=findings.slice(0,MODIFY_CHUNK).map(finding=>modifyRowHtml(finding,true)).join('')+(rest>0?`<button class="btn ghost sm modify-more" type="button" data-mod-more-group="${key}" data-mod-offset="${MODIFY_CHUNK}">${ic('chevrons-down')}Show ${Math.min(rest,MODIFY_CHUNK).toLocaleString()} more of ${rest.toLocaleString()}</button>`:'');
}
function modifyListHtml(entry){
  const {groups,singles}=modifyPatterns(entry);
  for(const group of groups)modifyPatternMap.set(group.key,group.findings);
  const groupHtml=groups.map(modifyPatternHtml).join('');
  const first=singles.slice(0,MODIFY_CHUNK).map(finding=>modifyRowHtml(finding,false)).join('');
  const rest=singles.length-MODIFY_CHUNK;
  return `<div class="modify-list" data-mod-list="${esc(entry.rule.id)}">${groupHtml}${first}${rest>0?`<button class="btn ghost sm modify-more" type="button" data-mod-more="${esc(entry.rule.id)}" data-mod-offset="${MODIFY_CHUNK}">${ic('chevrons-down')}Show ${Math.min(rest,MODIFY_CHUNK).toLocaleString()} more of ${rest.toLocaleString()}</button>`:''}</div>`;
}
/* A pattern's box mirrors its rows: all kept, none kept, or (indeterminate) mixed. */
function syncModifyPatternBox(key){
  const findings=modifyPatternMap.get(key);if(!findings)return;
  const box=$(`input[data-mod-group="${key}"]`);if(!box)return;
  const kept=findings.filter(finding=>!isExcludedId(finding.id)).length;
  box.checked=kept===findings.length;box.indeterminate=kept>0&&kept<findings.length;
  const count=$(`[data-mod-group-count="${key}"]`);if(count)count.textContent=modifyPatternCountText(findings);
  const wrap=$(`[data-mod-pattern="${key}"]`);if(wrap)wrap.classList.toggle('is-excluded',!kept);
}
function syncAllModifyPatternBoxes(){for(const key of modifyPatternMap.keys())syncModifyPatternBox(key);}
function setExcludedMany(ids,on){
  if(!S.session.excluded)S.session.excluded=new Set();
  for(const id of ids){if(on)S.session.excluded.add(id);else S.session.excluded.delete(id);}
  S.session.excludedRev=(S.session.excludedRev||0)+1;saveExcluded();refreshSessionResult();
}
function modifyRuleSeverity(entry){
  let top='info';for(const finding of entry.findings)if(severityRank(finding.severity)>severityRank(top))top=finding.severity;
  return top;
}
/* Counts can read as a share of every tag in the registry instead -- the tiny
   %/# button beside any count flips the whole tab at once. */
function modifyTotalTags(){const snapshot=S.session&&S.session.snapshot;return snapshot&&snapshot.rows?snapshot.rows.length:0;}
function modifyShare(count){
  const total=modifyTotalTags();if(!total)return '0%';
  const pct=count/total*100;
  if(count&&pct<0.1)return '<0.1%';
  return (pct>=10?Math.round(pct):Number(pct.toFixed(1)))+'%';
}
function modifyAmount(count){return S.session&&S.session.modifyCountMode==='percent'?modifyShare(count):count.toLocaleString();}
function modifyPctBtn(){
  const pctMode=S.session&&S.session.modifyCountMode==='percent';
  return `<button class="modify-pct-toggle" type="button" data-mod-pct title="${pctMode?'Show counts':'Show as % of all tags'}">${pctMode?'#':'%'}</button>`;
}
function modifyRuleCountText(entry){
  const aside=entry.findings.length-entry.kept,pctMode=S.session&&S.session.modifyCountMode==='percent';
  if(aside)return `${modifyAmount(entry.kept)} of ${modifyAmount(entry.findings.length)} kept`;
  return pctMode?`${modifyAmount(entry.findings.length)} of tags`:`${modifyAmount(entry.findings.length)} kept`;
}
function modifyCategoryCountText(group){
  const aside=group.total-group.kept,pctMode=S.session&&S.session.modifyCountMode==='percent';
  if(aside)return `${modifyAmount(group.kept)} of ${modifyAmount(group.total)} kept`;
  return pctMode?`${modifyAmount(group.total)} of tags`:`${modifyAmount(group.total)} findings`;
}
/* A closed rule card holds a placeholder instead of its list -- on a registry
   with thousands of findings, building every hidden row up front is what made
   the screen (and every full re-render) stall for seconds. The list is built
   the first time the card opens. */
function modifyRuleHtml(entry,forceOpen){
  const open=forceOpen||(S.session.modifyOpenRules||[]).includes(entry.rule.id);
  const severity=modifyRuleSeverity(entry);
  return `<details class="modify-rule" data-mod-rule="${esc(entry.rule.id)}" ${open?'open':''}>
    <summary><span class="audit-severity ${esc(severity)}">${esc(SEVERITY_LABELS[severity])}</span><b>${modifyHighlight(entry.rule.title)}</b><span class="modify-rule-count" data-mod-count="${esc(entry.rule.id)}">${modifyRuleCountText(entry)}</span>${modifyPctBtn()}<span class="modify-rule-buttons"><button class="btn ghost sm" type="button" data-mod-keep="${esc(entry.rule.id)}" title="${modifyQueryNorm()?'Keep every match shown for this check':'Keep every finding of this check'}">${ic('check')}Keep all</button><button class="btn ghost sm" type="button" data-mod-aside="${esc(entry.rule.id)}" title="${modifyQueryNorm()?'Set every match shown for this check aside':'Set every finding of this check aside'}">${ic('circle-x')}Set all aside</button></span></summary>
    ${open?modifyListHtml(entry):`<div class="modify-lazy" data-mod-lazy="${esc(entry.rule.id)}"></div>`}
  </details>`;
}
function modifyFillLazyList(details){
  const lazy=details.querySelector('[data-mod-lazy]');if(!lazy)return;
  const groupsNow=modifyGroups();let entry=null;
  for(const group of groupsNow){entry=group.rules.find(item=>item.rule.id===lazy.dataset.modLazy)||entry;}
  if(!entry){lazy.remove();return;}
  const holder=document.createElement('div');holder.innerHTML=modifyListHtml(entry);
  lazy.replaceWith(holder.firstElementChild);
  $$('[data-mod-pattern]',details).forEach(pattern=>syncModifyPatternBox(pattern.dataset.modPattern));
}
function updateModifyCounts(navigate){
  const groups=modifyGroups();
  for(const group of groups){
    const label=$(`[data-mod-cat-count="${group.category}"]`);if(label)label.textContent=modifyCategoryCountText(group);
    for(const entry of group.rules){const count=$(`[data-mod-count="${entry.rule.id}"]`);if(count)count.textContent=modifyRuleCountText(entry);}
  }
  const result=S.session.result,aside=excludedInBase();
  const included=$('#modifyIncluded');if(included&&result)included.textContent=`${result.summary.findings.toLocaleString()} counted`;
  const asideChip=$('#modifyAside');if(asideChip)asideChip.textContent=`${aside.toLocaleString()} set aside`;
  const restore=$('#modifyRestore');if(restore)restore.disabled=!aside;
  renderSideNav(navigate);
}
export function renderModifications(navigate){
  if(!(S.session&&S.session.rawResult)){navigate('upload');return;}
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='modify';S.homeMode='audit';
  currentNavigate=navigate;modifyPatternMap=new Map();
  const groups=modifyGroups(),query=clean(S.session.modifySearch),aside=excludedInBase(),result=S.session.result;
  const matchTotal=query?groups.reduce((sum,group)=>sum+group.rules.reduce((inner,entry)=>inner+entry.matches.length,0),0):0;
  const scrollTop=S.session.modifyScrollTop||0;
  $('#view').innerHTML=`<section class="modify-shell">
    <div class="screen-heading"><div><span class="eyebrow">Your judgement, applied</span><h2>Modifications</h2><p>${esc(S.session.name)} — untick any finding you disagree with and it is set aside: dropped from the findings list, the Dashboard, and the Excel report. Decisions are remembered for this registry, even after a reload.</p></div></div>
    <div class="modify-toolbar"><div class="searchbox">${ic('search')}<input id="modifySearch" aria-label="Search findings" placeholder="Search tags and findings" value="${esc(S.session.modifySearch||'')}"></div><span class="modify-chip" id="modifyIncluded">${result?result.summary.findings.toLocaleString():0} counted</span><span class="modify-chip aside" id="modifyAside">${aside.toLocaleString()} set aside</span>${query?`<span class="modify-chip match">${matchTotal.toLocaleString()} match${matchTotal===1?'':'es'}</span>`:''}<span class="spacer"></span>${query&&matchTotal?`<button class="btn ghost" type="button" id="modifyKeepMatches" title="Keep every finding the search surfaced">${ic('check')}Keep matches</button><button class="btn ghost" type="button" id="modifyAsideMatches" title="Set every finding the search surfaced aside">${ic('circle-x')}Set matches aside</button>`:''}<button class="btn ghost" type="button" id="modifyExpandAll" title="Open every group and check">${ic('chevrons-down')}Expand all</button><button class="btn ghost" type="button" id="modifyCollapseAll" title="Close every group and check">${ic('chevrons-up')}Collapse all</button><button class="btn ghost" type="button" id="modifyRestore" ${aside?'':'disabled'}>${ic('rotate-ccw')}Restore all</button></div>
    <div class="modify-body" id="modifyBody">${groups.length?groups.map(group=>{
      const closed=!query&&(S.session.modifyClosedCats||[]).includes(group.category);
      return `<section class="modify-category ${closed?'is-closed':''}" data-mod-cat="${esc(group.category)}"><header class="modify-cat-head" data-mod-cat-toggle="${esc(group.category)}"><span class="modify-cat-chevron" aria-hidden="true">${ic('chevron-down')}</span><h3>${esc(group.label)}</h3><b data-mod-cat-count="${esc(group.category)}">${modifyCategoryCountText(group)}</b>${modifyPctBtn()}</header><div class="modify-cat-body" ${closed?'hidden':''}>${group.rules.map(entry=>modifyRuleHtml(entry,!!query)).join('')}</div></section>`;
    }).join(''):`<div class="rule-reference-empty">${ic(query?'search':'check-check')}<b>${query?'No findings match that search':'Nothing to modify'}</b><span>${query?'Try a shorter word or a different tag.':'This registry has no findings from the checks that are switched on.'}</span></div>`}</div>
  </section>`;
  const body=$('#modifyBody');
  body.onchange=event=>{
    const groupBox=event.target.closest('[data-mod-group]');
    if(groupBox){
      const key=groupBox.dataset.modGroup,findings=modifyPatternMap.get(key);if(!findings)return;
      setExcludedMany(findings.map(finding=>finding.id),!groupBox.checked);
      const wrap=$(`[data-mod-pattern="${key}"]`);
      if(wrap)$$('[data-mod-finding]',wrap).forEach(row=>{row.checked=groupBox.checked;row.closest('.modify-row').classList.toggle('is-excluded',!groupBox.checked);});
      syncModifyPatternBox(key);updateModifyCounts(navigate);
      return;
    }
    const input=event.target.closest('[data-mod-finding]');if(!input)return;
    setExcluded(input.dataset.modFinding,!input.checked);
    input.closest('.modify-row').classList.toggle('is-excluded',!input.checked);
    const pattern=input.closest('[data-mod-pattern]');if(pattern)syncModifyPatternBox(pattern.dataset.modPattern);
    updateModifyCounts(navigate);
  };
  body.onclick=event=>{
    const pct=event.target.closest('[data-mod-pct]');
    if(pct){
      event.preventDefault();
      S.session.modifyCountMode=S.session.modifyCountMode==='percent'?'count':'percent';
      updateModifyCounts(navigate);syncAllModifyPatternBoxes();
      const pctMode=S.session.modifyCountMode==='percent';
      $$('.modify-pct-toggle',body).forEach(button=>{button.textContent=pctMode?'#':'%';button.title=pctMode?'Show counts':'Show as % of all tags';});
      return;
    }
    const openBtn=event.target.closest('[data-mod-open]');
    if(openBtn){event.preventDefault();openFinding(openBtn.dataset.modOpen,openBtn);return;}
    const catHead=event.target.closest('[data-mod-cat-toggle]');
    if(catHead){
      const section=catHead.closest('.modify-category'),catBody=section.querySelector('.modify-cat-body');
      const closing=!catBody.hidden;
      catBody.hidden=closing;section.classList.toggle('is-closed',closing);
      const set=new Set(S.session.modifyClosedCats||[]);
      if(closing)set.add(section.dataset.modCat);else set.delete(section.dataset.modCat);
      S.session.modifyClosedCats=[...set];
      return;
    }
    const head=event.target.closest('.modify-pattern-head');
    if(head&&!event.target.closest('input')){
      const pattern=head.closest('.modify-pattern'),rows=pattern.querySelector('.modify-pattern-rows'),expand=pattern.querySelector('[data-mod-expand]');
      if(rows.hidden)modifyFillPatternRows(pattern);
      rows.hidden=!rows.hidden;if(expand){expand.setAttribute('aria-expanded',rows.hidden?'false':'true');expand.classList.toggle('open',!rows.hidden);}
      return;
    }
    const moreGroup=event.target.closest('[data-mod-more-group]');
    if(moreGroup){
      const findings=modifyPatternMap.get(moreGroup.dataset.modMoreGroup)||[];
      const offset=Number(moreGroup.dataset.modOffset)||0,slice=findings.slice(offset,offset+MODIFY_CHUNK);
      const fragment=document.createElement('div');fragment.innerHTML=slice.map(finding=>modifyRowHtml(finding,true)).join('');
      const rest=findings.length-offset-slice.length;
      moreGroup.before(...fragment.children);
      if(rest>0){moreGroup.dataset.modOffset=String(offset+slice.length);moreGroup.innerHTML=`${ic('chevrons-down')}Show ${Math.min(rest,MODIFY_CHUNK).toLocaleString()} more of ${rest.toLocaleString()}`;}
      else moreGroup.remove();
      return;
    }
    const more=event.target.closest('[data-mod-more]');
    if(more){
      const groupsNow=modifyGroups();let entry=null;
      for(const group of groupsNow){entry=group.rules.find(item=>item.rule.id===more.dataset.modMore)||entry;}
      if(!entry)return;
      const singles=modifyPatterns(entry).singles;
      const offset=Number(more.dataset.modOffset)||0,slice=singles.slice(offset,offset+MODIFY_CHUNK);
      const fragment=document.createElement('div');fragment.innerHTML=slice.map(finding=>modifyRowHtml(finding,false)).join('');
      const rest=singles.length-offset-slice.length;
      more.before(...fragment.children);
      if(rest>0){more.dataset.modOffset=String(offset+slice.length);more.innerHTML=`${ic('chevrons-down')}Show ${Math.min(rest,MODIFY_CHUNK).toLocaleString()} more of ${rest.toLocaleString()}`;}
      else more.remove();
      return;
    }
    const keep=event.target.closest('[data-mod-keep]'),asideAll=event.target.closest('[data-mod-aside]');
    if(keep||asideAll){
      event.preventDefault();
      const ruleId=(keep||asideAll).dataset.modKeep||(keep||asideAll).dataset.modAside;
      /* Scoped to the matches on screen -- with a search active, "Set all aside"
         on a check touches only the findings the search surfaced. */
      const groupsNow=modifyGroups();let entry=null;
      for(const group of groupsNow){entry=group.rules.find(item=>item.rule.id===ruleId)||entry;}
      if(!entry)return;
      const on=!!asideAll;
      setExcludedMany(entry.matches.map(finding=>finding.id),on);
      /* The card is updated in place -- a full screen rebuild here is what made
         these buttons stall on a large registry. Rows not rendered yet pick the
         state up from the store when they are. */
      const card=body.querySelector(`.modify-rule[data-mod-rule="${CSS.escape(ruleId)}"]`);
      if(card){
        $$('[data-mod-finding]',card).forEach(row=>{row.checked=!on;row.closest('.modify-row').classList.toggle('is-excluded',on);});
        $$('[data-mod-pattern]',card).forEach(pattern=>syncModifyPatternBox(pattern.dataset.modPattern));
      }
      updateModifyCounts(navigate);
    }
  };
  syncAllModifyPatternBoxes();
  $$('.modify-rule',body).forEach(details=>details.addEventListener('toggle',()=>{
    if(details.open)modifyFillLazyList(details);
    const set=new Set(S.session.modifyOpenRules||[]);
    if(details.open)set.add(details.dataset.modRule);else set.delete(details.dataset.modRule);
    S.session.modifyOpenRules=[...set];
  }));
  $('#modifySearch').oninput=event=>{S.session.modifySearch=event.target.value;debounceSearch(()=>rerenderModifications(navigate,true));};
  $('#modifyRestore').onclick=()=>{
    if(!excludedInBase())return;
    S.session.excluded=new Set();S.session.excludedRev=(S.session.excludedRev||0)+1;saveExcluded();refreshSessionResult();
    rerenderModifications(navigate);toast('Every finding counts again');
  };
  /* Expand/collapse work directly on the DOM -- no screen rebuild. Expand all
     fills every still-lazy card from the one memoized grouping pass. */
  $('#modifyExpandAll').onclick=()=>{
    S.session.modifyClosedCats=[];
    $$('.modify-category',body).forEach(section=>{section.classList.remove('is-closed');const catBody=section.querySelector('.modify-cat-body');if(catBody)catBody.hidden=false;});
    const ids=[];
    $$('.modify-rule',body).forEach(details=>{if(!details.open){modifyFillLazyList(details);details.open=true;}ids.push(details.dataset.modRule);});
    S.session.modifyOpenRules=ids;
  };
  $('#modifyCollapseAll').onclick=()=>{
    $$('.modify-rule',body).forEach(details=>{details.open=false;});
    S.session.modifyOpenRules=[];
    const cats=[];
    $$('.modify-category',body).forEach(section=>{section.classList.add('is-closed');const catBody=section.querySelector('.modify-cat-body');if(catBody)catBody.hidden=true;cats.push(section.dataset.modCat);});
    S.session.modifyClosedCats=cats;
  };
  const collectMatchIds=()=>{const ids=[];for(const group of modifyGroups())for(const entry of group.rules)for(const finding of entry.matches)ids.push(finding.id);return ids;};
  const keepMatches=$('#modifyKeepMatches');
  if(keepMatches)keepMatches.onclick=()=>{setExcludedMany(collectMatchIds(),false);rerenderModifications(navigate);};
  const asideMatches=$('#modifyAsideMatches');
  if(asideMatches)asideMatches.onclick=()=>{
    const ids=collectMatchIds();setExcludedMany(ids,true);rerenderModifications(navigate);
    toast(`${ids.length.toLocaleString()} finding${ids.length===1?'':'s'} set aside`);
  };
  $('#view').scrollTop=scrollTop;
  $('#view').onscroll=()=>{if(S.screen==='modify')S.session.modifyScrollTop=$('#view').scrollTop;};
  renderSideNav(navigate);
}
/* Re-render keeping the reading position; a search change starts back at the top. */
function rerenderModifications(navigate,resetScroll){
  if(resetScroll)S.session.modifyScrollTop=0;else S.session.modifyScrollTop=$('#view')?$('#view').scrollTop:0;
  const focused=document.activeElement&&document.activeElement.id==='modifySearch';
  renderModifications(navigate);
  if(focused){const input=$('#modifySearch');if(input){input.focus();input.setSelectionRange(input.value.length,input.value.length);}}
}

export function renderRules(navigate){
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='rules';S.homeMode='rules';
  const rules=ruleCatalog(),sources=new Set(rules.map(rule=>rule.source)),topics=new Set(rules.map(rule=>rule.category)),hasResult=!!(S.session&&S.session.result);
  $('#view').innerHTML=`<section class="rules-shell">
    <div class="screen-heading"><div><span class="eyebrow">Plain-language reference</span><h2>What SSM Audit checks</h2><p>Every check the audit runs, grouped by where it comes from. Switch any check off to leave it out of the findings, the Dashboard, and the Excel report. ${hasResult?'Counts show how many times each check fired on the registry you loaded.':'Load a registry to see how many times each check fires.'}</p></div></div>
    <div class="rules-overview"><div><span>Checks in use</span><b>${rules.length}</b><p>Applied to every equipment row.</p></div><div><span>Sources</span><b>${sources.size}</b><p>Registry integrity, the SSM SOP, and commissioning logic.</p></div><div><span>Topics</span><b>${topics.size}</b><p>Hierarchy, dependencies, consistency, milestones, and headers.</p></div></div>
    <div class="rules-toolbar"><div class="searchbox">${ic('search')}<input id="ruleSearch" aria-label="Search audit rules" placeholder="Search checks and explanations" value="${esc(S.rules.search)}"></div><select id="ruleSource" aria-label="Filter by rule source"><option value="all">All sources</option>${SSM_AUDIT_SOURCES.map(source=>`<option value="${esc(source.id)}" ${S.rules.source===source.id?'selected':''}>${esc(source.label)}</option>`).join('')}</select><select id="ruleCategory" aria-label="Filter by topic"><option value="all">All topics</option>${Object.entries(RULE_CATEGORY_LABELS).filter(([key])=>rules.some(rule=>rule.category===key)).map(([key,label])=>`<option value="${esc(key)}" ${S.rules.category===key?'selected':''}>${esc(label)}</option>`).join('')}</select><span id="ruleResultCount"></span></div>
    <div class="rules-note rule-off-banner" id="ruleOffBanner" hidden>${ic('info')}<span><b></b> — switched-off checks are not counted in findings, the Dashboard, or the Excel report. The choice is remembered on this device.</span><button class="btn-link" type="button" id="ruleAllOn">Turn all on</button></div>
    <div class="rule-catalog" id="ruleCatalog"></div>
  </section>`;
  renderSideNav(navigate);
  $('#ruleSearch').oninput=event=>{S.rules.search=event.target.value;debounceSearch(()=>renderRuleCatalog(navigate));};
  $('#ruleSource').onchange=event=>{S.rules.source=event.target.value;renderRuleCatalog(navigate);};
  $('#ruleCategory').onchange=event=>{S.rules.category=event.target.value;renderRuleCatalog(navigate);};
  renderRuleCatalog(navigate);
}

/* ------------------------------------------------------------ upload screen */

function importStatus(){
  const result=S.session.result,summary=result&&result.summary;
  if(result)return `<div class="audit-import ready"><span class="file-icon">${ic(summary.status==='ready'?'check-check':'triangle-alert')}</span><div class="file-meta"><b>${esc(S.session.name)}</b><span>${summary.rows.toLocaleString()} rows checked &middot; ${summary.findings.toLocaleString()} findings</span></div><button class="btn primary" id="openAuditResult">Open findings</button><button class="xbtn icon-btn" id="removeAuditTarget" type="button" aria-label="Remove this registry">${ic('x')}</button></div>`;
  if(S.session.error)return `<div class="audit-import error"><span class="file-icon">${ic('triangle-alert')}</span><div class="file-meta"><b>That workbook could not be read</b><span>${esc(S.session.error)}</span></div><button class="btn" id="chooseAuditAgain">Choose another</button></div>`;
  return '';
}

export function renderUpload(navigate){
  if(S.homeMode==='rules'){renderRules(navigate);return;}
  if(S.homeMode==='compare'){renderComparisonUpload(navigate);return;}
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='upload';
  $('#view').innerHTML=`<section class="upload-shell">
    <div class="screen-heading"><div><span class="eyebrow">Registry integrity + SSM rules</span><h2>Audit a Cx Registry</h2><p>Pick a completed registry workbook. Everything runs here on your machine, and you get a list of what to fix.</p></div></div>
    <div class="dropzone" id="dropzone">
      <span class="drop-icon">${ic('file-spreadsheet')}</span>
      <h3>Drop a registry workbook here</h3>
      <p>Or browse this device. Excel only &mdash; .xlsx or .xls.</p>
      <button class="btn primary" id="auditBrowse">${ic('upload')}Choose workbook</button>
      <span class="file-types">${ic('lock')}Nothing leaves this browser</span>
    </div>
    <div id="importStatus">${importStatus()}</div>
    <input id="auditFile" type="file" accept=".xlsx,.xls" hidden>
    <div class="checks-panel">
      <div class="checks-head"><b>What gets checked</b><button class="btn-link" type="button" id="openRulesFromUpload">See every check${ic('arrow-right')}</button></div>
      <ul class="checks-list">
        <li>${ic('list-tree')}<div><b>Hierarchy</b><span>Parents that exist, no loops, no crossing into another UPN or discipline.</span></div></li>
        <li>${ic('git-branch')}<div><b>Dependencies</b><span>Every dependency resolves, nothing repeats, nothing depends on itself.</span></div></li>
        <li>${ic('database')}<div><b>Registry consistency</b><span>UPN, System Name, discipline, and Item Masters agree across rows.</span></div></li>
        <li>${ic('network')}<div><b>Commissioning logic</b><span>Control paths, power paths, drives, headers, and milestone roll-ups.</span></div></li>
      </ul>
    </div>
  </section>`;
  renderSideNav(navigate);wireUpload(navigate);
}

function chooseFile(){const input=$('#auditFile');if(input)input.click();}
function wireUpload(navigate){
  const input=$('#auditFile'),drop=$('#dropzone'),browse=$('#auditBrowse'),again=$('#chooseAuditAgain');
  if(browse)browse.onclick=event=>{event.stopPropagation();chooseFile();};if(again)again.onclick=chooseFile;
  if(input)input.onchange=()=>{const file=input.files[0];input.value='';if(file)addAuditTarget(file,navigate);};
  if(drop){
    drop.onclick=event=>{if(!event.target.closest('button'))chooseFile();};
    drop.ondragover=event=>{event.preventDefault();drop.classList.add('dragging');};
    drop.ondragleave=()=>drop.classList.remove('dragging');
    drop.ondrop=event=>{event.preventDefault();drop.classList.remove('dragging');const file=event.dataTransfer.files[0];if(file)addAuditTarget(file,navigate);};
  }
  const open=$('#openAuditResult');if(open)open.onclick=()=>navigate('audit');
  const remove=$('#removeAuditTarget');if(remove)remove.onclick=()=>{resetSession();clearComparisonTarget();renderUpload(navigate);};
  const rules=$('#openRulesFromUpload');if(rules)rules.onclick=()=>{S.homeMode='rules';navigate('rules');};
}

export async function addAuditTarget(file,navigate){
  if(!/\.(xlsx|xls)$/i.test(file.name)){toast('Choose an Excel workbook (.xlsx or .xls)');return;}
  resetSession();clearComparisonTarget();S.session.name=file.name;
  try{
    await runWithProgress('Running SSM Audit',file.name,async(checkpoint,report)=>{
      const bytes=new Uint8Array(await readArrayBuffer(file));await checkpoint();
      const workbook=XLSX.read(bytes,{type:'array',dense:true});report(.1,'Registry opened');await checkpoint();
      const snapshot=await auditSnapshotFromWorkbook(workbook,file.name,checkpoint,(fraction,label)=>report(.1+fraction*.6,label));
      report(.74,`${snapshot.rows.length.toLocaleString()} rows parsed`);await checkpoint();
      report(.8,'Running every check');await checkpoint();
      const rawResult=runSsmAudit(snapshot),result=applyRulePreferences(rawResult,S.rules.disabled);report(1,`${result.findings.length.toLocaleString()} findings`);
      S.session={...S.session,snapshot,rawResult,result,error:'',auditedAt:Date.now()};loadActioned();loadExcluded();refreshSessionResult();
      S.comparison.targetName=file.name;S.comparison.targetSnapshot=snapshot;S.comparison.targetError='';S.comparison.result=null;
    });
    navigate('dashboard');
  }catch(error){
    console.error('SSM Audit failed',error);S.session.error=error&&error.message||'Could not read this registry';S.session.snapshot=null;S.session.result=null;renderUpload(navigate);
  }
}

function clearComparisonTarget(){S.comparison.targetName='';S.comparison.targetSnapshot=null;S.comparison.targetError='';S.comparison.result=null;S.comparison.selectedUpn='';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};}
function clearComparisonReference(){S.comparison.referenceName='';S.comparison.referenceSnapshot=null;S.comparison.referenceError='';S.comparison.result=null;S.comparison.selectedUpn='';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};}

function comparisonSlot(side){
  const target=side==='target',name=target?S.comparison.targetName:S.comparison.referenceName,snapshot=target?S.comparison.targetSnapshot:S.comparison.referenceSnapshot,error=target?S.comparison.targetError:S.comparison.referenceError;
  const title=target?'Registry to audit':'Finished project to compare against',detail=target?'Gets the full audit plus the comparison.':'Used for this session only.';
  if(snapshot)return `<section class="compare-upload-slot ready" data-compare-side="${side}"><div class="compare-slot-label"><span>${target?'01':'02'}</span><div><b>${title}</b><small>${detail}</small></div></div><div class="compare-ready-icon">${ic(target?'check-check':'square-stack')}</div><h3>${esc(name)}</h3><p>${snapshot.rows.length.toLocaleString()} rows ready</p><div class="compare-slot-actions"><button class="btn sm" type="button" data-compare-replace="${side}">${ic('upload')}Replace</button><button class="xbtn icon-btn" type="button" data-compare-remove="${side}" aria-label="Remove ${esc(title)}">${ic('x')}</button></div></section>`;
  return `<section class="compare-upload-slot ${error?'error':''}" data-compare-side="${side}"><div class="compare-slot-label"><span>${target?'01':'02'}</span><div><b>${title}</b><small>${detail}</small></div></div><div class="compare-ready-icon">${ic(error?'triangle-alert':'file-spreadsheet')}</div><h3>${error?'Could not read that workbook':'Choose workbook'}</h3><p>${esc(error||'Drop an .xlsx or .xls file here, or browse this device.')}</p><button class="btn ${target?'primary':''} sm" type="button" data-compare-browse="${side}">${ic('upload')}Browse</button></section>`;
}

function renderComparisonUpload(navigate){
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='upload';
  if(!S.comparison.targetSnapshot&&S.session.snapshot){S.comparison.targetName=S.session.name;S.comparison.targetSnapshot=S.session.snapshot;}
  const ready=!!(S.comparison.targetSnapshot&&S.comparison.referenceSnapshot);
  $('#view').innerHTML=`<section class="upload-shell compare-upload-shell">
    <div class="screen-heading"><div><span class="eyebrow">Building-neutral comparison</span><h2>Compare two project hierarchies</h2><p>Line your registry up against a finished project, UPN by UPN, to see how equipment is nested. Building values are left out so different sites do not look like differences.</p></div></div>
    <div class="compare-upload-grid">${comparisonSlot('target')}${comparisonSlot('reference')}</div>
    <div class="compare-runbar"><div class="compare-method">${ic('lock')}<span><b>Local and session-only.</b> The finished project is only a reference. It never changes the audit rules or future results.</span></div><button class="btn primary" id="runComparison" type="button" ${ready?'':'disabled'}>${ic('square-stack')}Compare registries</button></div>
    <input id="compareTargetFile" type="file" accept=".xlsx,.xls" hidden><input id="compareReferenceFile" type="file" accept=".xlsx,.xls" hidden>
    <div class="checks-panel"><div class="checks-head"><b>What the comparison looks at</b></div><ul class="checks-list">
      <li>${ic('hash')}<div><b>Matching by UPN</b><span>Systems line up by UPN, never by Building.</span></div></li>
      <li>${ic('list-tree')}<div><b>How equipment nests</b><span>Equipment is paired by tag, description, classification, and parent role.</span></div></li>
      <li>${ic('folder-tree')}<div><b>Headers and I&amp;C</b><span>Shows where the two projects group and place things differently.</span></div></li>
    </ul></div>
  </section>`;
  renderSideNav(navigate);wireComparisonUpload(navigate);
}

function wireComparisonUpload(navigate){
  const choose=side=>$('#compare'+(side==='target'?'Target':'Reference')+'File')?.click();
  for(const side of ['target','reference']){
    const input=$('#compare'+(side==='target'?'Target':'Reference')+'File'),slot=$(`[data-compare-side="${side}"]`),browse=$(`[data-compare-browse="${side}"]`),replace=$(`[data-compare-replace="${side}"]`),remove=$(`[data-compare-remove="${side}"]`);
    if(input)input.onchange=()=>{const file=input.files[0];input.value='';if(file)addComparisonFile(file,side,navigate);};
    if(browse)browse.onclick=event=>{event.stopPropagation();choose(side);};if(replace)replace.onclick=()=>choose(side);
    if(remove)remove.onclick=()=>{if(side==='target'){resetSession();clearComparisonTarget();}else clearComparisonReference();renderUpload(navigate);};
    if(slot&&!slot.classList.contains('ready')){
      slot.ondragover=event=>{event.preventDefault();slot.classList.add('dragging');};slot.ondragleave=()=>slot.classList.remove('dragging');slot.ondrop=event=>{event.preventDefault();slot.classList.remove('dragging');const file=event.dataTransfer.files[0];if(file)addComparisonFile(file,side,navigate);};
    }
  }
  $('#runComparison').onclick=()=>runComparison(navigate);
}

async function addComparisonFile(file,side,navigate){
  if(!/\.(xlsx|xls)$/i.test(file.name)){toast('Choose an Excel workbook (.xlsx or .xls)');return;}
  const errorKey=side==='target'?'targetError':'referenceError';S.comparison[errorKey]='';
  try{
    let snapshot,auditResult;
    await runWithProgress(side==='target'?'Loading registry to audit':'Loading finished project',file.name,async(checkpoint,report)=>{
      const bytes=new Uint8Array(await readArrayBuffer(file));await checkpoint();const workbook=XLSX.read(bytes,{type:'array',dense:true});report(.1,'Registry opened');await checkpoint();
      snapshot=await auditSnapshotFromWorkbook(workbook,file.name,checkpoint,(fraction,label)=>report(.1+fraction*.68,label));
      report(.82,`${snapshot.rows.length.toLocaleString()} rows parsed`);await checkpoint();if(side==='target'){auditResult=runSsmAudit(snapshot);report(1,`${auditResult.findings.length.toLocaleString()} audit findings`);}else report(1,'Reference ready');
    });
    if(side==='target'){resetSession();S.session={...S.session,name:file.name,snapshot,rawResult:auditResult,result:applyRulePreferences(auditResult,S.rules.disabled),error:'',auditedAt:Date.now()};loadActioned();loadExcluded();refreshSessionResult();S.comparison.targetName=file.name;S.comparison.targetSnapshot=snapshot;S.comparison.targetError='';}
    else{S.comparison.referenceName=file.name;S.comparison.referenceSnapshot=snapshot;S.comparison.referenceError='';}
    S.comparison.result=null;S.comparison.selectedUpn='';S.comparison.detailTab='hierarchy';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};renderUpload(navigate);
  }catch(error){console.error('Registry comparison import failed',error);if(side==='target'){resetSession();clearComparisonTarget();}else clearComparisonReference();S.comparison[errorKey]=error&&error.message||'Could not read this registry';renderUpload(navigate);}
}

async function runComparison(navigate){
  if(!S.comparison.targetSnapshot||!S.comparison.referenceSnapshot){toast('Choose both registries first');return;}
  await runWithProgress('Comparing project hierarchies','Building values are excluded',async(checkpoint,report)=>{report(.15,'Indexing target systems');await checkpoint();report(.38,'Aligning equipment by UPN and meaning');await checkpoint();const result=compareSsmRegistries(S.comparison.targetSnapshot,S.comparison.referenceSnapshot);S.comparison.result=result;const selected=result.systems.find(system=>system.status!=='aligned')||result.systems[0];if(selected&&selected.status==='aligned'&&S.comparison.systemFilter==='different')S.comparison.systemFilter='all';S.comparison.selectedUpn=selected&&selected.upn||'';S.comparison.detailTab='hierarchy';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};report(1,`${result.summary.systems.toLocaleString()} systems compared`);});
  navigate('compare');
}

/* ---------------------------------------------------------- findings screen */

function invalidateFindingCaches(){S.session.filteredCacheKey='';S.session.filteredCacheRows=null;S.session.displayCacheKey='';S.session.displayCacheRows=null;S.session.scopedCacheKey='';S.session.scopedCache=null;}
function sessionRowIndex(){
  if(S.session.rowIndex)return S.session.rowIndex;
  const bySource=new Map(),byId=new Map(),rows=S.session.snapshot&&S.session.snapshot.rows||[];
  for(const row of rows){
    const source=row._source||{};bySource.set(`${auditNormId(source.sheet)}|${source.row||0}`,row);
    const id=auditNormId(row.equipmentId);if(id&&!byId.has(id))byId.set(id,row);
  }
  S.session.rowIndex={bySource,byId};return S.session.rowIndex;
}
function registryRowFor(finding){
  const index=sessionRowIndex();
  return index.bySource.get(`${auditNormId(finding.sheet)}|${finding.row||0}`)||index.byId.get(auditNormId(finding.equipmentId))||null;
}
/* ------------------------------------------------- registry dimension filters */

/* Four registry columns a finding does not carry itself, so every match runs
   through the finding's registry row. A dimension holds the values that stay
   visible; an empty list means that dimension is not filtering at all, so
   "everything ticked" and "nothing chosen yet" are the same state. */
const DIM_KEYS=['discipline','milestone','upn','building'];
const DIM_FIELDS={discipline:'discipline',milestone:'milestone',upn:'upn',building:'building'};
const DIM_LABELS={discipline:'Discipline',milestone:'L2 milestone',upn:'UPN / system',building:'Building'};

function dimFilterMap(){return S.session.dimFilters||(S.session.dimFilters=Object.fromEntries(DIM_KEYS.map(key=>[key,[]])));}
/* "Keep nothing" needs a representable state: an impossible value that matches
   no row. Empty means "not filtering", so unchecking the last box stores this. */
const DIM_NONE='\u0001none';
function dimSelected(dimension){return new Set(dimFilterMap()[dimension]||[]);}
function clearDimFilters(){S.session.dimFilters=Object.fromEntries(DIM_KEYS.map(key=>[key,[]]));}
/* Options are the same buckets the dashboard ranks by — the dashboard just ranks
   the findings currently in scope — so a filter count and a dashboard count are
   the same number counted the same way. Options always come from the whole result
   so a value never disappears from the panel because it was filtered out. */
function dimOptions(dimension){
  const cache=S.session.dimOptionCache||(S.session.dimOptionCache={});
  if(cache[dimension])return cache[dimension];
  const options=dashRankBuckets(dimension,S.session.result.findings).map(bucket=>({key:auditNormId(bucket.value),label:bucket.label,count:bucket.count}))
    .sort((a,b)=>(a.key?0:1)-(b.key?0:1)||natCmp(a.label,b.label));
  cache[dimension]=options;return options;
}
function dimOptionLabel(dimension,key){const match=dimOptions(dimension).find(option=>option.key===key);return match?match.label:key||DASH_NO_MILESTONE;}
/* Ticking every value is the same as filtering on none of them. */
function setDimSelection(dimension,keys){
  const options=dimOptions(dimension);
  dimFilterMap()[dimension]=options.every(option=>keys.has(option.key))?[]:options.filter(option=>keys.has(option.key)).map(option=>option.key);
}
function dimActiveSets(){return DIM_KEYS.map(dimension=>[DIM_FIELDS[dimension],dimSelected(dimension)]).filter(entry=>entry[1].size);}
function dimRowMatches(finding,active){
  const row=registryRowFor(finding);if(!row)return false;
  for(const [field,values] of active)if(!values.has(auditNormId(row[field])))return false;
  return true;
}

function filteredFindings(){
  const query=clean(S.session.search).toUpperCase(),hiddenSources=new Set(S.session.hiddenSources||[]),hiddenSeverities=new Set(S.session.hiddenSeverities||[]),hiddenCategories=new Set(S.session.hiddenCategories||[]),hiddenRules=new Set(S.session.hiddenRules||[]),dimActive=dimActiveSets();
  const key=[...hiddenSources,...hiddenSeverities,...hiddenCategories,...hiddenRules,query,S.session.sort,S.session.hideActioned?'HA':'',S.session.actionedRev||0,DIM_KEYS.map(dimension=>(dimFilterMap()[dimension]||[]).join('~')).join('|')].join('');
  if(S.session.filteredCacheKey===key&&S.session.filteredCacheRows)return S.session.filteredCacheRows;
  let rows=S.session.result.findings.filter(finding=>!(S.session.hideActioned&&isActioned(finding))&&!hiddenSources.has(finding.rule.source)&&!hiddenSeverities.has(finding.severity)&&!hiddenCategories.has(finding.category)&&!hiddenRules.has(finding.rule.id)&&(!query||finding.searchKey.includes(query))&&(!dimActive.length||dimRowMatches(finding,dimActive)));
  const natural=(a,b)=>String(a||'').localeCompare(String(b||''),undefined,{numeric:true,sensitivity:'base'}),sort=S.session.sort;
  if(sort==='severity-asc')rows=[...rows].reverse();
  else if(sort==='equipment-asc'||sort==='equipment-desc')rows=[...rows].sort((a,b)=>(sort.endsWith('desc')?-1:1)*(natural(a.equipmentId,b.equipmentId)||a.row-b.row));
  else if(sort==='rule-asc'||sort==='rule-desc')rows=[...rows].sort((a,b)=>(sort.endsWith('desc')?-1:1)*(natural(a.rule.title,b.rule.title)||natural(a.rule.id,b.rule.id)||a.row-b.row));
  else if(sort==='row-asc'||sort==='row-desc')rows=[...rows].sort((a,b)=>(sort.endsWith('desc')?-1:1)*((a.row-b.row)||natural(a.sheet,b.sheet)));
  S.session.filteredCacheKey=key;S.session.filteredCacheRows=rows;return rows;
}
function findingGroup(finding){
  if(S.session.groupBy==='rule')return {key:`rule|${finding.rule.id}`,label:finding.rule.title,note:SOURCE_LABELS[finding.rule.source]||finding.rule.source};
  const row=registryRowFor(finding),milestone=clean(row&&row.milestone);
  return {key:`milestone|${auditNormId(milestone)}`,label:milestone||'No L2 milestone assigned',note:milestone?'L2 milestone':'These rows have no L2 milestone'};
}
/* One flat list of fixed-height rows — group headers and findings alike — so the
   virtual scroller keeps a single row-height calculation whatever the grouping. */
function displayRows(){
  const findings=filteredFindings(),groupBy=S.session.groupBy||'none',collapsed=new Set(S.session.collapsedGroups||[]);
  const key=[groupBy,S.session.filteredCacheKey,[...collapsed].join(',')].join('');
  if(S.session.displayCacheKey===key&&S.session.displayCacheRows)return S.session.displayCacheRows;
  let rows;
  if(groupBy==='none')rows=findings.map(finding=>({type:'finding',finding}));
  else{
    const groups=new Map();
    for(const finding of findings){const group=findingGroup(finding);let bucket=groups.get(group.key);if(!bucket){bucket={...group,items:[],worst:''};groups.set(group.key,bucket);}bucket.items.push(finding);if(severityRank(finding.severity)>severityRank(bucket.worst))bucket.worst=finding.severity;}
    rows=[];
    for(const bucket of [...groups.values()].sort((a,b)=>severityRank(b.worst)-severityRank(a.worst)||b.items.length-a.items.length||natCmp(a.label,b.label))){
      const isCollapsed=collapsed.has(bucket.key);rows.push({type:'group',group:bucket,collapsed:isCollapsed});
      if(!isCollapsed)for(const finding of bucket.items)rows.push({type:'finding',finding});
    }
  }
  S.session.displayCacheKey=key;S.session.displayCacheRows=rows;return rows;
}
function severityStrip(summary){
  const hidden=new Set(S.session.hiddenSeverities||[]),active=SSM_AUDIT_SEVERITIES.filter(level=>!hidden.has(level));
  const chip=level=>`<button class="sev-chip ${level} ${hidden.has(level)?'':'on'}" type="button" data-audit-severity="${level}" aria-pressed="${hidden.has(level)?'false':'true'}"><span class="sev-dot"></span>${SEVERITY_PLURALS[level]}<b>${summary.severity[level].toLocaleString()}</b></button>`;
  return `<div class="sev-strip"><button class="sev-chip all ${active.length===SSM_AUDIT_SEVERITIES.length?'on':''}" type="button" data-audit-severity="all">All findings<b>${summary.findings.toLocaleString()}</b></button>${SSM_AUDIT_SEVERITIES.map(chip).join('')}<span class="sev-meta">${summary.rows.toLocaleString()} rows checked &middot; ${summary.checks.toLocaleString()} checks run</span></div>`;
}
/* ------------------------------------------------------------- filter panel */

/* One filter model behind everything. Each section is a checklist over the same
   result and reduces to one of two state shapes:
     kind 'hidden' — S.session[stateKey] lists the values that are hidden
     kind 'dim'    — S.session.dimFilters[id] lists the values that are kept
   An empty list means the section is not filtering, so "everything ticked" and
   "nothing chosen yet" stay the same state in both shapes. */
const FILTER_SECTIONS=[
  {id:'severity',kind:'hidden',stateKey:'hiddenSeverities',label:'Level',hint:'How serious the finding is'},
  {id:'discipline',kind:'dim',label:DIM_LABELS.discipline,hint:'The discipline on the registry row'},
  {id:'milestone',kind:'dim',label:DIM_LABELS.milestone,hint:'The commissioning phase the row sits in'},
  {id:'upn',kind:'dim',label:DIM_LABELS.upn,hint:'The system the row belongs to'},
  {id:'building',kind:'dim',label:DIM_LABELS.building,hint:'Where the equipment is'},
  {id:'source',kind:'hidden',stateKey:'hiddenSources',label:'Source',hint:'Where the check comes from'},
  {id:'category',kind:'hidden',stateKey:'hiddenCategories',label:'Topic',hint:'What the check is about'},
  {id:'rule',kind:'hidden',stateKey:'hiddenRules',label:'Check',hint:'The individual check that fired'},
];
let filterPanelOpener=null,filterPanelTrap=null,filterRerender=null;

function filterSection(id){return FILTER_SECTIONS.find(section=>section.id===id)||null;}
/* Counts come from the whole result, never from the current scope: a number that
   moved every time you ticked a box would be impossible to aim with. */
function filterSectionOptions(section){
  const result=S.session&&S.session.result;if(!result)return [];
  if(section.kind==='dim')return dimOptions(section.id);
  const cache=S.session.filterOptionCache||(S.session.filterOptionCache={});
  if(cache[section.id])return cache[section.id];
  let options=[];
  if(section.id==='severity')options=SSM_AUDIT_SEVERITIES.map(level=>({key:level,label:SEVERITY_LABELS[level],count:result.summary.severity[level]||0}));
  else if(section.id==='source')options=SSM_AUDIT_SOURCES.map(source=>({key:source.id,label:source.label,count:result.summary.source[source.id]||0}));
  else if(section.id==='category')options=SSM_AUDIT_CATEGORIES.filter(category=>(result.summary.category[category]||0)>0).map(category=>({key:category,label:CATEGORY_LABELS[category]||category,count:result.summary.category[category]||0}));
  else if(section.id==='rule'){
    const counts=new Map();for(const finding of result.findings)counts.set(finding.rule.id,(counts.get(finding.rule.id)||0)+1);
    options=[...new Map(result.findings.map(finding=>[finding.rule.id,finding.rule])).values()]
      .map(rule=>({key:rule.id,label:rule.title,count:counts.get(rule.id)||0}))
      .sort((a,b)=>b.count-a.count||natCmp(a.label,b.label));
  }
  cache[section.id]=options;return options;
}
function filterSectionIsOn(section,key){
  if(section.kind==='dim'){const selected=dimSelected(section.id);return !selected.size||selected.has(key);}
  return !(S.session[section.stateKey]||[]).includes(key);
}
function filterSectionActive(section){
  if(section.kind==='dim')return dimSelected(section.id).size>0;
  return (S.session[section.stateKey]||[]).length>0;
}
function filterSectionKept(section){
  const options=filterSectionOptions(section);
  if(section.kind==='dim'){const selected=dimSelected(section.id);return selected.size?options.filter(option=>selected.has(option.key)):options;}
  const hidden=new Set(S.session[section.stateKey]||[]);
  return options.filter(option=>!hidden.has(option.key));
}
function filterSectionReset(section){if(section.kind==='dim')dimFilterMap()[section.id]=[];else S.session[section.stateKey]=[];}
function filterSectionNone(section){
  if(section.kind==='dim'){dimFilterMap()[section.id]=[DIM_NONE];return;}
  S.session[section.stateKey]=filterSectionOptions(section).map(option=>option.key);
}
function noneAllFilters(){for(const section of FILTER_SECTIONS)if(filterSectionOptions(section).length)filterSectionNone(section);}
/* Ticking is computed from state rather than from the checkboxes on screen, so a
   search that hides half the list cannot silently unselect what it hid. */
function filterSectionSet(section,key,on){
  if(section.kind==='dim'){
    const options=filterSectionOptions(section),selected=dimSelected(section.id),keys=selected.size?new Set(selected):new Set(options.map(option=>option.key));
    if(on)keys.add(key);else keys.delete(key);
    keys.delete(DIM_NONE);
    if(!keys.size){dimFilterMap()[section.id]=[DIM_NONE];return;}
    setDimSelection(section.id,keys);return;
  }
  const hidden=new Set(S.session[section.stateKey]||[]);
  if(on)hidden.delete(key);else hidden.add(key);
  S.session[section.stateKey]=[...hidden];
}
/* ---- saved filter views ----
   A view is a named copy of every filter list, kept on the device. Values that
   do not exist in the current registry simply match nothing. */
const FILTER_VIEWS_KEY='ssm-audit.filter-views';
function loadFilterViews(){try{const raw=localStorage.getItem(FILTER_VIEWS_KEY);const list=raw?JSON.parse(raw):[];return Array.isArray(list)?list.filter(view=>view&&typeof view.name==='string'&&view.filters):[];}catch(_){return [];}}
function saveFilterViews(views){try{localStorage.setItem(FILTER_VIEWS_KEY,JSON.stringify(views));}catch(_){/* private mode */}}
function captureFilterView(name){
  return {name,filters:{hiddenSeverities:[...(S.session.hiddenSeverities||[])],hiddenSources:[...(S.session.hiddenSources||[])],hiddenCategories:[...(S.session.hiddenCategories||[])],hiddenRules:[...(S.session.hiddenRules||[])],dimFilters:JSON.parse(JSON.stringify(dimFilterMap()))}};
}
function applyFilterView(view){
  const filters=view.filters||{};
  S.session.hiddenSeverities=[...(filters.hiddenSeverities||[])];S.session.hiddenSources=[...(filters.hiddenSources||[])];
  S.session.hiddenCategories=[...(filters.hiddenCategories||[])];S.session.hiddenRules=[...(filters.hiddenRules||[])];
  S.session.dimFilters=Object.fromEntries(DIM_KEYS.map(key=>[key,[...((filters.dimFilters||{})[key]||[])]]));
}
function filterViewsRowHtml(){
  const views=loadFilterViews();
  const chips=views.map(view=>`<span class="filter-view-chip"><button type="button" data-filter-view-load="${esc(view.name)}" title="Load this view">${esc(view.name)}</button><button type="button" class="filter-view-x" data-filter-view-delete="${esc(view.name)}" aria-label="Delete the view ${esc(view.name)}">${ic('x')}</button></span>`).join('');
  return `<div class="filter-views" id="filterViewsRow"><span class="filter-views-label">Views</span>${chips||'<span class="filter-views-empty">None saved yet</span>'}<button class="btn ghost sm" type="button" id="filterViewSave">${ic('plus')}Save view</button></div>`;
}
function wireFilterViewsRow(){
  $$('[data-filter-view-load]').forEach(button=>button.onclick=()=>{
    const view=loadFilterViews().find(entry=>entry.name===button.dataset.filterViewLoad);if(!view)return;
    applyFilterView(view);applyFilterChange();toast(`View \u201c${view.name}\u201d loaded`);
  });
  $$('[data-filter-view-delete]').forEach(button=>button.onclick=()=>{
    saveFilterViews(loadFilterViews().filter(entry=>entry.name!==button.dataset.filterViewDelete));
    renderFilterViewsRow();
  });
  const save=$('#filterViewSave');
  if(save)save.onclick=()=>{
    const row=$('#filterViewsRow');if(!row)return;
    row.innerHTML=`<span class="filter-views-label">Views</span><input id="filterViewName" class="filter-view-name" maxlength="40" placeholder="Name this view" aria-label="Name for the saved view"><button class="btn primary sm" type="button" id="filterViewConfirm">Save</button><button class="btn ghost sm" type="button" id="filterViewCancel">Cancel</button>`;
    const input=$('#filterViewName');input.focus();
    const commit=()=>{const name=clean(input.value);if(!name){renderFilterViewsRow();return;}
      const views=loadFilterViews().filter(entry=>entry.name!==name);views.push(captureFilterView(name));saveFilterViews(views);renderFilterViewsRow();toast(`View \u201c${name}\u201d saved`);};
    $('#filterViewConfirm').onclick=commit;$('#filterViewCancel').onclick=()=>renderFilterViewsRow();
    input.onkeydown=event=>{if(event.key==='Enter'){event.preventDefault();commit();}if(event.key==='Escape'){event.stopPropagation();renderFilterViewsRow();}};
  };
}
function renderFilterViewsRow(){
  const row=$('#filterViewsRow');if(!row)return;
  const replacement=document.createElement('div');replacement.innerHTML=filterViewsRowHtml();
  row.replaceWith(replacement.firstElementChild);wireFilterViewsRow();
}
function filterActiveCount(){return FILTER_SECTIONS.filter(filterSectionActive).length;}
function clearAllFilters(){S.session.hiddenSources=[];S.session.hiddenSeverities=[];S.session.hiddenCategories=[];S.session.hiddenRules=[];clearDimFilters();}

/* A chip is one value a filtering section is still showing, whichever state shape
   it lives in, so removing a chip always means the same thing: stop showing this. */
function filterChipRowHtml(includeSeverity){
  const groups=[];
  for(const section of FILTER_SECTIONS){
    if(!filterSectionActive(section))continue;
    if(section.id==='severity'&&!includeSeverity)continue;
    const kept=filterSectionKept(section);
    const chips=kept.length?kept.map(option=>`<span class="filter-chip"><span>${esc(option.label)}</span><button class="filter-chip-x" type="button" data-filter-chip="${esc(section.id)}" data-filter-chip-value="${esc(option.key)}" title="Stop showing ${esc(option.label)}" aria-label="Stop showing ${esc(option.label)}">${ic('x')}</button></span>`).join('')
      :'<span class="filter-chip is-none">Nothing shown</span>';
    groups.push(`<span class="filter-chip-group"><button type="button" class="filter-chip-group-label" data-filter-edit="${esc(section.id)}" title="Change the ${esc(section.label)} filter">${esc(section.label)}</button>${chips}</span>`);
  }
  if(!groups.length)return '<div class="filter-chip-row" id="filterChipRow" hidden></div>';
  return `<div class="filter-chip-row" id="filterChipRow"><button class="filter-chip-edit" type="button" data-filter-button title="Open the filters">${ic('filter')}Edit</button>${groups.join('')}<button class="filter-chip-clear" type="button" data-filter-clear-all title="Remove every filter and show everything">Reset filters</button></div>`;
}
function filterButtonHtml(id){
  const count=filterActiveCount();
  return `<button class="btn audit-filter-button ${count?'active':''}" id="${esc(id)}" type="button" data-filter-button aria-expanded="${S.session.panelOpen?'true':'false'}" aria-haspopup="dialog">${ic('filter')}Filters <b class="filter-badge">${count||''}</b></button>`;
}
function filterPanelSectionHtml(section,query){
  const options=filterSectionOptions(section);if(!options.length)return '';
  const matches=query?options.filter(option=>option.label.toUpperCase().includes(query)):options;
  const active=filterSectionActive(section),kept=filterSectionKept(section).length;
  const body=matches.length
    ?matches.map(option=>`<label class="audit-filter-option"><input type="checkbox" data-filter-option="${esc(section.id)}" value="${esc(option.key)}" ${filterSectionIsOn(section,option.key)?'checked':''}><span title="${esc(option.label)}">${esc(option.label)}</span><b>${option.count.toLocaleString()}</b></label>`).join('')
    :`<p class="filter-section-empty">Nothing in ${esc(section.label)} matches that search.</p>`;
  return `<section class="filter-section ${active?'is-active':''}" data-filter-section="${esc(section.id)}">
    <div class="filter-section-head"><div><b>${esc(section.label)}</b><span>${esc(section.hint)}</span></div>
      <span class="filter-section-state">${active?`${kept.toLocaleString()} of ${options.length.toLocaleString()}`:''}</span>
      <span class="filter-section-buttons"><button class="btn ghost sm" type="button" data-filter-all="${esc(section.id)}" ${active?'':'disabled'} title="Tick every ${esc(section.label.toLowerCase())}">All</button><button class="btn ghost sm" type="button" data-filter-none="${esc(section.id)}" title="Untick every ${esc(section.label.toLowerCase())}">None</button></span></div>
    <div class="filter-section-list">${body}</div></section>`;
}
function filterPanelBodyHtml(){
  const query=clean(S.session.panelSearch).toUpperCase();
  const sections=FILTER_SECTIONS.map(section=>filterPanelSectionHtml(section,query)).filter(Boolean).join('');
  return sections||'<p class="filter-section-empty">There is nothing to filter yet.</p>';
}
function filterPanelSummary(){const count=filterActiveCount();return count?`${count} ${count===1?'group':'groups'} filtering`:'Nothing filtered';}
function filterPanelHtml(){
  const open=!!S.session.panelOpen;
  return `<div class="filter-sheet-back ${open?'show':''}" id="filterSheetBack" ${open?'':'hidden'}>
    <aside class="filter-sheet" id="filterSheet" role="dialog" aria-modal="true" aria-label="Filters" tabindex="-1">
      <header class="filter-sheet-head"><div><b>Filters</b><span id="filterSheetSummary">${esc(filterPanelSummary())}</span></div><button class="xbtn icon-btn" type="button" id="filterSheetClose" aria-label="Close filters">${ic('x')}</button></header>
      ${filterViewsRowHtml()}
      <div class="filter-sheet-search"><div class="searchbox">${ic('search')}<input id="filterSheetSearch" aria-label="Search the filter options" placeholder="Search filter options" value="${esc(S.session.panelSearch)}"></div></div>
      <div class="filter-sheet-scroll" id="filterSheetBody">${filterPanelBodyHtml()}</div>
      <footer class="filter-sheet-foot"><button class="btn ghost" type="button" id="filterSheetSelectAll" title="Tick every box — show everything">${ic('check-check')}Select all</button><button class="btn ghost" type="button" id="filterSheetClear" title="Untick every box">${ic('x')}Clear all</button><button class="btn primary" type="button" id="filterSheetDone">Done</button></footer>
    </aside></div>`;
}
function focusFilterOption(sectionId,value){
  const match=$$('[data-filter-option]').find(box=>box.dataset.filterOption===sectionId&&box.value===value);
  (match||$(`[data-filter-all="${sectionId}"]`)||$('#filterSheetSearch'))?.focus();
}
function renderFilterPanelBody(){
  const body=$('#filterSheetBody');if(body){body.innerHTML=filterPanelBodyHtml();wireFilterPanelBody();}
  const summary=$('#filterSheetSummary');if(summary)summary.textContent=filterPanelSummary();
}
function updateFilterBadge(){
  const count=filterActiveCount();
  $$('[data-filter-button]').forEach(button=>{button.classList.toggle('active',!!count);const badge=button.querySelector('.filter-badge');if(badge)badge.textContent=count||'';});
}
/* Every change applies live. The panel body is rebuilt so per-section state stays
   truthful, which means focus has to be put back on the control that was used. */
function applyFilterChange(restoreFocus){
  S.session.cursor=-1;invalidateFindingCaches();
  if(S.session.panelOpen)renderFilterPanelBody();
  filterRerender?.();
  restoreFocus?.();
}
function wireFilterPanelBody(){
  $$('[data-filter-option]').forEach(input=>input.onchange=()=>{
    const section=filterSection(input.dataset.filterOption);if(!section)return;
    const value=input.value;filterSectionSet(section,value,input.checked);
    applyFilterChange(()=>focusFilterOption(section.id,value));
  });
  $$('[data-filter-all]').forEach(button=>button.onclick=()=>{
    const section=filterSection(button.dataset.filterAll);if(!section)return;
    filterSectionReset(section);applyFilterChange(()=>$(`[data-filter-all="${section.id}"]`)?.focus());
  });
  $$('[data-filter-none]').forEach(button=>button.onclick=()=>{
    const section=filterSection(button.dataset.filterNone);if(!section)return;
    filterSectionNone(section);applyFilterChange(()=>$(`[data-filter-none="${section.id}"]`)?.focus());
  });
}
function renderFilterChipRow(includeSeverity){
  const row=$('#filterChipRow');if(!row)return;
  const replacement=document.createElement('div');replacement.innerHTML=filterChipRowHtml(includeSeverity);
  row.replaceWith(replacement.firstElementChild);wireFilterChips();
}
function wireFilterChips(){
  $$('[data-filter-chip]').forEach(button=>button.onclick=()=>{
    const section=filterSection(button.dataset.filterChip);if(!section)return;
    filterSectionSet(section,button.dataset.filterChipValue,false);applyFilterChange();
  });
  /* The chip row carries one, and the dashboard's empty-scope card carries
     another — both are on screen at once when the filters exclude everything. */
  $$('[data-filter-clear-all]').forEach(button=>button.onclick=()=>{clearAllFilters();applyFilterChange();});
  $$('[data-filter-edit]').forEach(button=>button.onclick=()=>{
    const sectionId=button.dataset.filterEdit;
    setFilterPanelOpen(true);
    requestAnimationFrame(()=>requestAnimationFrame(()=>{const target=$(`[data-filter-section="${sectionId}"]`);target?.scrollIntoView({block:'start'});target?.querySelector('input')?.focus({preventScroll:true});}));
  });
  wireFilterButtons();
}
function setFilterPanelOpen(open){
  const back=$('#filterSheetBack');if(!back)return;
  S.session.panelOpen=!!open;
  if(open)animateOpen(back);else animateClose(back);
  $$('[data-filter-button]').forEach(button=>button.setAttribute('aria-expanded',open?'true':'false'));
  document.body.classList.toggle('filter-sheet-open',!!open);
  if(open){
    /* Whatever opened the panel gets focus back when it closes. A pointer click
       may leave the body focused, so the Filters button is the fallback. */
    const active=document.activeElement;
    filterPanelOpener=active&&active!==document.body?active:$('[data-filter-button]');
    renderFilterPanelBody();
    filterPanelTrap?.();filterPanelTrap=activateFocusTrap(back,()=>setFilterPanelOpen(false));
    requestAnimationFrame(()=>$('#filterSheetSearch')?.focus());
    return;
  }
  filterPanelTrap?.();filterPanelTrap=null;
  const opener=filterPanelOpener;filterPanelOpener=null;
  const target=opener&&document.contains(opener)&&typeof opener.focus==='function'?opener:$('[data-filter-button]');
  target?.focus();
}
/* The dashboard rebuilds its whole shell — Filters button included — after every
   live filter change, so the buttons are rewired separately from the panel. */
function wireFilterButtons(){
  $$('[data-filter-button]').forEach(button=>button.onclick=()=>setFilterPanelOpen(!S.session.panelOpen));
}
function wireFilterPanel(){
  const back=$('#filterSheetBack');if(!back)return;
  back.onpointerdown=event=>{if(event.target===back)setFilterPanelOpen(false);};
  $('#filterSheetClose').onclick=()=>setFilterPanelOpen(false);
  $('#filterSheetDone').onclick=()=>setFilterPanelOpen(false);
  $('#filterSheetClear').onclick=()=>{noneAllFilters();applyFilterChange(()=>$('#filterSheetClear')?.focus());};
  $('#filterSheetSelectAll').onclick=()=>{clearAllFilters();applyFilterChange(()=>$('#filterSheetSelectAll')?.focus());};
  wireFilterViewsRow();
  $('#filterSheetSearch').oninput=event=>{S.session.panelSearch=event.target.value;debounceSearch(renderFilterPanelBody);};
  wireFilterButtons();wireFilterPanelBody();
}
function sortOptions(){const options=[['severity-desc','Most serious first'],['severity-asc','Least serious first'],['equipment-asc','Equipment A to Z'],['equipment-desc','Equipment Z to A'],['rule-asc','Check A to Z'],['rule-desc','Check Z to A'],['row-asc','Row, low to high'],['row-desc','Row, high to low']];return options.map(([value,label])=>`<option value="${value}" ${S.session.sort===value?'selected':''}>${label}</option>`).join('');}
function groupOptions(){const options=[['none','No grouping'],['rule','Group by check'],['milestone','Group by L2 milestone']];return options.map(([value,label])=>`<option value="${value}" ${S.session.groupBy===value?'selected':''}>${label}</option>`).join('');}
function statusMarkup(summary){const blockers=summary.severity.blocker||0,label=summary.status==='blocked'?`${SEVERITY_LABELS.blocker} &middot; ${blockers.toLocaleString()} ${blockers===1?'row':'rows'}`:summary.status==='review'?'Review required':'Ready';return `<span class="audit-status ${summary.status}">${ic(summary.status==='ready'?'check-check':'triangle-alert')}${label}</span>`;}

export function renderAuditResult(navigate){
  const result=S.session&&S.session.result;if(!result){navigate('upload');return;}
  S.screen='audit';S.session.panelOpen=false;const summary=result.summary,fullscreen=!!S.session.fullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  $('#view').innerHTML=`<section id="auditShell" class="audit-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><div class="audit-title"><span class="audit-title-icon">${ic('check-check')}</span><div><h2>Audit Findings</h2><p>${esc(S.session.name)}</p></div></div><span class="spacer"></span>${statusMarkup(summary)}<button class="btn" id="auditDashboard">${ic('layout-dashboard')}Dashboard</button><button class="btn" id="auditBack">${ic('upload')}New registry</button><button class="btn" id="exportAudit">${ic('file-down')}Export report</button><button class="btn icon-btn" id="auditFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
    ${severityStrip(summary)}
    ${filterChipRowHtml(false)}
    <div class="audit-toolbar"><div class="searchbox">${ic('search')}<input id="auditSearch" aria-label="Search findings" placeholder="Search tags, checks, and explanations" value="${esc(S.session.search)}"></div>${filterButtonHtml('auditFilters')}<select id="auditGroup" aria-label="Group findings">${groupOptions()}</select><select id="auditSort" aria-label="Sort findings">${sortOptions()}</select>${actionedInResult()?`<button class="chip sm hide-actioned ${S.session.hideActioned?'on':''}" type="button" id="auditHideActioned" aria-pressed="${S.session.hideActioned?'true':'false'}" title="Hide findings marked actioned in the app">${ic('check')}Hide actioned<b>${actionedInResult().toLocaleString()}</b></button>`:''}<span class="audit-count" id="auditCount"></span></div>
    <div class="audit-table-wrap" id="auditTableWrap" tabindex="0" role="group" aria-label="Findings list. Use the arrow keys to move and Enter to open."><table class="audit-table"><thead><tr><th>Severity</th><th>What was found</th><th>Equipment ID</th><th>Source</th><th>Sheet &middot; row</th><th aria-label="Open details"></th></tr></thead><tbody id="auditRows"></tbody></table></div>
  </section>
  ${filterPanelHtml()}`;
  currentNavigate=navigate;renderSideNav(navigate);wireAuditResult(navigate);
}

function auditGroupRowHtml(item,index){
  const group=item.group;
  return `<tr class="audit-group-row" data-row-index="${index}" data-audit-group="${esc(group.key)}" tabindex="-1"><td colspan="6"><button class="audit-group-toggle ${item.collapsed?'':'open'}" type="button" data-audit-group-toggle="${esc(group.key)}" aria-expanded="${item.collapsed?'false':'true'}">${ic('chevron-right')}<b>${esc(group.label)}</b><small>${esc(group.note)}</small><span class="audit-group-count ${esc(group.worst||'')}">${group.items.length.toLocaleString()}</span></button></td></tr>`;
}
function auditRowHtml(item,index,query){
  const finding=item.finding;
  return `<tr data-row-index="${index}" data-audit-finding="${esc(finding.id)}" tabindex="-1" class="${isActioned(finding)?'is-actioned':''}" aria-label="Open finding for ${esc(finding.equipmentId||'the registry')}"><td><span class="audit-severity ${finding.severity}">${SEVERITY_LABELS[finding.severity]}</span>${isActioned(finding)?`<span class="actioned-flag" title="Marked actioned in the app">${ic('check')}</span>`:''}</td><td><b title="${esc(finding.why)}">${highlightHtml(finding.why,query)}</b><span>${esc(finding.rule.title)}</span><small class="audit-mobile-evidence">${esc(SOURCE_LABELS[finding.rule.source]||finding.rule.source)} &middot; ${esc(finding.sheet||'Registry')} &middot; row ${finding.row||'—'}</small></td><td>${copyTagHtml(finding.equipmentId,highlightHtml(finding.equipmentId,query))}</td><td>${esc(SOURCE_LABELS[finding.rule.source]||finding.rule.source)}</td><td class="audit-evidence-cell">${esc(finding.sheet||'Registry')} &middot; ${finding.row||'&mdash;'}</td><td>${ic('chevron-right')}</td></tr>`;
}
function renderRows(){
  const wrap=$('#auditTableWrap'),body=$('#auditRows');if(!wrap||!body)return;
  const rows=displayRows(),findingTotal=filteredFindings().length,count=$('#auditCount'),query=clean(S.session.search).toUpperCase();
  if(count)count.textContent=`${findingTotal.toLocaleString()} of ${S.session.result.findings.length.toLocaleString()} findings`;
  if(!rows.length){body.innerHTML=`<tr><td colspan="6"><div class="audit-empty-state">${ic('search')}<b>Nothing matches the filters you have set</b><span>Clear them to see the full list again.</span><button class="btn" type="button" id="auditClearFilters">Clear filters and search</button></div></td></tr>`;
    const clear=$('#auditClearFilters');if(clear)clear.onclick=()=>resetAuditFilters();return;}
  const visible=Math.min(AUDIT_MAX_ROWS,Math.max(28,Math.ceil(wrap.clientHeight/AUDIT_ROW_HEIGHT)+AUDIT_OVERSCAN*2));
  const start=Math.min(Math.max(0,rows.length-visible),Math.max(0,Math.floor(wrap.scrollTop/AUDIT_ROW_HEIGHT)-AUDIT_OVERSCAN)),end=Math.min(rows.length,start+visible);
  const spacer=height=>height?`<tr class="audit-spacer" aria-hidden="true"><td colspan="6" style="height:${height}px"></td></tr>`:'';
  body.innerHTML=spacer(start*AUDIT_ROW_HEIGHT)+rows.slice(start,end).map((item,offset)=>item.type==='group'?auditGroupRowHtml(item,start+offset):auditRowHtml(item,start+offset,query)).join('')+spacer((rows.length-end)*AUDIT_ROW_HEIGHT);
  wireCopyTags(body);
  $$('[data-audit-group-toggle]',body).forEach(button=>button.onclick=event=>{event.stopPropagation();toggleFindingGroup(button.dataset.auditGroupToggle);});
  $$('[data-audit-finding]',body).forEach(row=>{row.onclick=event=>{if(!event.target.closest('[data-copy-tag]')){S.session.cursor=Number(row.dataset.rowIndex);openFinding(row.dataset.auditFinding,row);}};});
  const cursorRow=body.querySelector(`[data-row-index="${S.session.cursor}"]`);if(cursorRow)cursorRow.classList.add('cursor');
}
function toggleFindingGroup(key){
  const collapsed=new Set(S.session.collapsedGroups||[]);if(collapsed.has(key))collapsed.delete(key);else collapsed.add(key);
  S.session.collapsedGroups=[...collapsed];S.session.displayCacheKey='';renderRows();
}
function focusCursorRow(){
  const wrap=$('#auditTableWrap'),body=$('#auditRows');if(!wrap||!body)return;
  const top=S.session.cursor*AUDIT_ROW_HEIGHT,bottom=top+AUDIT_ROW_HEIGHT;
  if(top<wrap.scrollTop)wrap.scrollTop=top;else if(bottom>wrap.scrollTop+wrap.clientHeight)wrap.scrollTop=bottom-wrap.clientHeight;
  S.session.scrollTop=wrap.scrollTop;renderRows();
  const row=body.querySelector(`[data-row-index="${S.session.cursor}"]`);if(row)row.focus({preventScroll:true});
}
function moveCursor(delta){
  const rows=displayRows();if(!rows.length)return;
  const next=S.session.cursor<0?(delta>0?0:rows.length-1):Math.min(rows.length-1,Math.max(0,S.session.cursor+delta));
  S.session.cursor=next;focusCursorRow();
}
function activateCursor(){
  const rows=displayRows(),item=rows[S.session.cursor];if(!item)return;
  if(item.type==='group'){toggleFindingGroup(item.group.key);focusCursorRow();return;}
  const row=$(`[data-row-index="${S.session.cursor}"]`);openFinding(item.finding.id,row);
}
function resetAuditFilters(){
  clearAllFilters();S.session.search='';
  const search=$('#auditSearch');if(search)search.value='';
  applyFilterChange();
}
function renderSeverityStrip(){
  const strip=$('.sev-strip'),result=S.session.result;if(!strip||!result)return;
  const replacement=document.createElement('div');replacement.innerHTML=severityStrip(result.summary);
  strip.replaceWith(replacement.firstElementChild);wireSeverityStrip();
}
function wireSeverityStrip(){
  $$('[data-audit-severity]').forEach(button=>button.onclick=()=>{
    const level=button.dataset.auditSeverity;
    /* From "everything on", the first pill narrows to just that level; further
       pills add or remove levels; removing the last one goes back to all. */
    if(level==='all')S.session.hiddenSeverities=[];
    else{const hidden=new Set(S.session.hiddenSeverities||[]);
      if(!hidden.size)S.session.hiddenSeverities=SSM_AUDIT_SEVERITIES.filter(item=>item!==level);
      else{if(hidden.has(level))hidden.delete(level);else hidden.add(level);if(hidden.size===SSM_AUDIT_SEVERITIES.length)hidden.clear();S.session.hiddenSeverities=[...hidden];}}
    applyFilterChange();
  });
}

function findingSection(title,body,className){return `<section class="finding-section ${className||''}"><h4>${esc(title)}</h4>${body}</section>`;}
function registryContextHtml(row){
  if(!row)return '';
  const field=(label,value,copy)=>{const text=clean(value);if(!text)return '';return `<div><dt>${esc(label)}</dt><dd>${copy?copyTagHtml(text):esc(text)}</dd></div>`;};
  const cells=[field('Description',row.equipmentDescription),field('Closest parent',row.closestParent,true),field('Discipline',row.discipline),field('UPN',row.upn),field('System Name',row.systemName),field('Building',row.building),field('L1 milestone',row.milestoneParent),field('L2 milestone',row.milestone),field('Item Master',row.itemMaster),field('Dependencies',row.dependencies)].filter(Boolean).join('');
  if(!cells)return '';
  return `<section class="finding-section"><h4>The registry row</h4><dl class="finding-context">${cells}</dl></section>`;
}
/* The engine attaches `finding.relationship` to relationship-shaped findings so the
   drawer can draw the link instead of only describing it. The Found/Expected block
   stays underneath — it remains the archival record of what was in the workbook. */
const RELATIONSHIP_LEGENDS={parent:'The equipment at the bottom is nested under the one at the top.',dependency:'The item on the left must be ready before the one on the right.',loop:'Following these links returns to where it started.'};
function relationshipDiffers(a,b){const left=auditNormId(a),right=auditNormId(b);return !!(left&&right&&left!==right);}
function relationshipMetaHtml(node,diff){
  const upn=clean(node&&node.upn),discipline=clean(node&&node.discipline);
  if(!upn&&!discipline)return '';
  const parts=[];
  if(upn)parts.push(`UPN <span class="rel-value ${diff&&diff.upn?'is-diff':''}">${esc(upn)}</span>`);
  if(discipline)parts.push(`<span class="rel-value ${diff&&diff.discipline?'is-diff':''}">${esc(discipline)}</span>`);
  const pills=`${diff&&diff.upn?'<span class="rel-pill">different UPN</span>':''}${diff&&diff.discipline?'<span class="rel-pill">different discipline</span>':''}`;
  return `<span class="rel-node-meta">${parts.join(' &middot; ')}${pills}</span>`;
}
function relationshipNodeHtml(node,roleLabel,roleClass,diff){
  const tag=clean(node&&node.tag);
  return `<div class="rel-node ${roleClass||''}">${roleLabel?`<span class="rel-node-role">${esc(roleLabel)}</span>`:''}<span class="rel-node-tag">${tag?copyTagHtml(tag):'<span class="muted">Not in this registry</span>'}</span>${relationshipMetaHtml(node,diff)}</div>`;
}
function relationshipParentHtml(self,parent){
  const diff={upn:relationshipDiffers(self&&self.upn,parent&&parent.upn),discipline:relationshipDiffers(self&&self.discipline,parent&&parent.discipline)};
  return `<div class="rel-diagram rel-parent">${relationshipNodeHtml(parent,'Closest Parent','role-parent',diff)}
    <div class="rel-link"><span class="rel-line"></span><span class="rel-link-label">nests under</span></div>
    ${relationshipNodeHtml(self,'This equipment','role-this')}</div>`;
}
function relationshipDependencyHtml(self,dependency){
  return `<div class="rel-diagram rel-dependency">${relationshipNodeHtml(dependency,'Dependency (starts first)','role-dependency')}
    <span class="rel-arrow" aria-hidden="true">${ic('arrow-right')}</span>
    ${relationshipNodeHtml(self,'This equipment','role-this')}</div>`;
}
/* The engine's cycle path closes on itself (A → B → C → A). The repeated tail is
   what the "back to 1" arrow already says, so it is dropped from the numbered list. */
function relationshipLoopNodes(nodes){
  const first=nodes[0],tail=nodes[nodes.length-1];
  return nodes.length>2&&auditNormId(first&&first.tag)===auditNormId(tail&&tail.tag)?nodes.slice(0,-1):nodes;
}
function relationshipLoopHtml(input){
  const nodes=relationshipLoopNodes(input),last=nodes.length-1;
  return `<ol class="rel-diagram rel-loop">${nodes.map((node,index)=>{
    const start=index===0,tag=clean(node&&node.tag);
    const connector=index<last?`<span class="rel-loop-arrow" aria-hidden="true">${ic('arrow-down')}</span>`:`<span class="rel-loop-arrow back">${ic('rotate-ccw')}back to 1</span>`;
    return `<li class="rel-loop-step ${start?'is-start':''}"><span class="rel-step-num">${index+1}</span>
      <div class="rel-node compact ${start?'role-this':'role-step'}">${start?'<span class="rel-node-role">This equipment (start of loop)</span>':''}<span class="rel-node-tag">${tag?copyTagHtml(tag):'<span class="muted">Not in this registry</span>'}</span>${relationshipMetaHtml(node)}</div>
      ${connector}</li>`;
  }).join('')}</ol>`;
}
function relationshipDiagramHtml(finding){
  const relationship=finding&&finding.relationship,nodes=relationship&&Array.isArray(relationship.nodes)?relationship.nodes:null;
  if(!nodes||!nodes.length)return '';
  const kind=relationship.kind,self=nodes.find(node=>node.role==='this')||nodes[0];
  let body='';
  if(kind==='parent'){const parent=nodes.find(node=>node.role==='parent')||nodes[1];if(!parent)return '';body=relationshipParentHtml(self,parent);}
  else if(kind==='dependency'){const dependency=nodes.find(node=>node.role==='dependency')||nodes[1];if(!dependency)return '';body=relationshipDependencyHtml(self,dependency);}
  else if(kind==='loop'){if(nodes.length<2)return '';body=relationshipLoopHtml(nodes);}
  else return '';
  return `<section class="finding-section rel-section"><h4>How these are linked</h4>${body}<p class="rel-legend">Reading this: ${esc(RELATIONSHIP_LEGENDS[kind]||'')}</p></section>`;
}
function openFinding(id,opener){
  /* Set-aside findings are not in the session result, but the Modifications
     screen still opens them -- the raw engine result is the fallback. */
  const finding=S.session.result.findings.find(item=>item.id===id)||(S.session.rawResult&&S.session.rawResult.findings.find(item=>item.id===id));if(!finding)return;
  const list=filteredFindings(),position=list.findIndex(item=>item.id===id);
  const sourceLabel=SOURCE_LABELS[finding.rule.source]||finding.rule.source,confidenceLabel=RULE_CONFIDENCE_LABELS[finding.rule.confidence]||finding.rule.confidence;
  const actual=typeof finding.actual==='string'?finding.actual:JSON.stringify(finding.actual),expected=typeof finding.expected==='string'?finding.expected:JSON.stringify(finding.expected);
  S.session.selectedFindingId=id;S.session.opener=opener;
  $('#drawerTitle').textContent='Finding';
  $('#drawerBody').innerHTML=`<div class="audit-drawer">
    <div class="audit-drawer-top"><span class="audit-severity ${finding.severity}">${SEVERITY_LABELS[finding.severity]}</span>${finding.equipmentId?`<span class="audit-drawer-tag">${copyTagHtml(finding.equipmentId)}</span>`:'<span class="audit-rule-source">Registry-wide</span>'}</div>
    ${findingSection('Why this was flagged',`<p class="finding-why">${esc(finding.why)}</p>`)}
    ${findingSection('What must be true',`<p class="finding-statement">${esc(finding.rule.statement)}</p><span class="finding-rule-name">${esc(finding.rule.title)} &middot; ${esc(sourceLabel)} &middot; ${esc(confidenceLabel)}</span>`,'muted-section')}
    ${relationshipDiagramHtml(finding)}
    <section class="finding-section"><h4>What we found and what we expected</h4><div class="finding-compare"><div class="found"><span>Found</span><p>${esc(actual)||'Blank'}</p></div><div class="expected"><span>Expected</span><p>${esc(expected)||'—'}</p></div></div></section>
    ${finding.recommendation?findingSection('What to do',`<p class="finding-action">${esc(finding.recommendation)}</p>`,'action-section'):''}
    ${registryContextHtml(registryRowFor(finding))}
    <div class="finding-evidence">${ic('file-spreadsheet')}${esc(finding.sheet||'Registry')} &middot; row ${finding.row||'—'}${finding.field?' &middot; '+esc(finding.field):''}</div>
    <div class="finding-action-row"><button class="btn ${isActioned(finding)?'done':''}" type="button" id="findingActioned" aria-pressed="${isActioned(finding)?'true':'false'}">${ic('check')}${isActioned(finding)?'Actioned — click to undo':'Mark actioned'}</button><button class="btn ${isExcludedId(finding.id)?'done':''}" type="button" id="findingExclude" title="${isExcludedId(finding.id)?'This finding is set aside — click to have it count again':'Disagree with this finding? Set it aside — it leaves every metric until restored on the Modifications screen'}">${ic(isExcludedId(finding.id)?'rotate-ccw':'circle-x')}${isExcludedId(finding.id)?'Set aside — click to restore':'Set aside'}</button>${finding.equipmentId?`<button class="btn" type="button" id="findingInHierarchy">${ic('list-tree')}Show in hierarchy</button>`:''}</div>
    <div class="finding-steps"><button class="btn ghost sm" type="button" id="findingPrev" ${position>0?'':'disabled'}>${ic('chevron-left')}Previous</button><span>${position>=0?`${(position+1).toLocaleString()} of ${list.length.toLocaleString()}`:''}</span><button class="btn ghost sm" type="button" id="findingNext" ${position>=0&&position<list.length-1?'':'disabled'}>Next${ic('chevron-right')}</button></div>
  </div>`;
  wireCopyTags($('#drawerBody'));
  const previous=$('#findingPrev'),next=$('#findingNext'),inTree=$('#findingInHierarchy');
  if(previous)previous.onclick=()=>{const target=list[position-1];if(target)openFinding(target.id,opener);};
  if(next)next.onclick=()=>{const target=list[position+1];if(target)openFinding(target.id,opener);};
  if(inTree)inTree.onclick=()=>{closeDrawer();focusHierarchyOnEquipment(finding.equipmentId);};
  const actionedButton=$('#findingActioned');
  if(actionedButton)actionedButton.onclick=()=>{setActioned(finding,!isActioned(finding));if(S.screen==='audit'&&currentNavigate)renderAuditResult(currentNavigate);else renderRows();openFinding(finding.id,opener);};
  const excludeButton=$('#findingExclude');
  if(excludeButton)excludeButton.onclick=()=>{
    const on=!isExcludedId(finding.id);
    setExcluded(finding.id,on);
    if(S.screen==='modify'){
      /* Update the row and counts in place; the drawer stays open on the finding. */
      const row=$(`[data-mod-finding="${finding.id.replace(/"/g,'\\"')}"]`);
      if(row){row.checked=!on;row.closest('.modify-row').classList.toggle('is-excluded',on);const pattern=row.closest('[data-mod-pattern]');if(pattern)syncModifyPatternBox(pattern.dataset.modPattern);}
      if(currentNavigate)updateModifyCounts(currentNavigate);
      openFinding(finding.id,opener);
      return;
    }
    closeDrawer();
    if(S.screen==='audit'&&currentNavigate)renderAuditResult(currentNavigate);else if(currentNavigate)renderSideNav(currentNavigate);
    toast(on?'Set aside — restore it on the Modifications screen':'This finding counts again');
  };
  const backdrop=$('#drawerBack');animateOpen(backdrop);backdrop.setAttribute('aria-hidden','false');
  drawerTrapCleanup?.();drawerTrapCleanup=activateFocusTrap(backdrop,closeDrawer);$('#drawer').focus();
}
export function closeDrawer(){
  drawerTrapCleanup?.();drawerTrapCleanup=null;$('#drawerBack').setAttribute('aria-hidden','true');animateClose($('#drawerBack'));const opener=S.session&&S.session.opener;if(opener&&document.contains(opener))opener.focus();
}
function rerenderRows(reset){const wrap=$('#auditTableWrap');if(reset&&wrap)wrap.scrollTop=0;renderRows();}
/* Every screen that can host the filter panel tears the previous one down first,
   so a panel never outlives the screen that opened it. */
function teardownAuditFilters(){
  filterPanelTrap?.();filterPanelTrap=null;filterPanelOpener=null;filterRerender=null;
  document.body.classList.remove('filter-sheet-open');
  if(S.session)S.session.panelOpen=false;
}
function wireAuditResult(navigate){
  teardownAuditFilters();
  filterRerender=()=>{renderSeverityStrip();renderFilterChipRow(false);updateFilterBadge();rerenderRows(true);};
  $('#auditBack').onclick=()=>{S.homeMode='audit';navigate('upload');};$('#exportAudit').onclick=()=>openExportOptions();
  $('#auditDashboard').onclick=()=>navigate('dashboard');
  $('#auditFullscreen').onclick=()=>{S.session.fullscreen=!S.session.fullscreen;renderAuditResult(navigate);};
  $('#auditSearch').oninput=event=>{S.session.search=event.target.value;S.session.cursor=-1;debounceSearch(()=>rerenderRows(true));};
  $('#auditSort').onchange=event=>{S.session.sort=event.target.value;S.session.cursor=-1;rerenderRows(true);};
  const hideActioned=$('#auditHideActioned');if(hideActioned)hideActioned.onclick=()=>{S.session.hideActioned=!S.session.hideActioned;S.session.cursor=-1;S.session.scrollTop=0;renderAuditResult(currentNavigate||(()=>{}));};
  $('#auditGroup').onchange=event=>{S.session.groupBy=event.target.value;S.session.collapsedGroups=[];S.session.cursor=-1;S.session.displayCacheKey='';rerenderRows(true);};
  wireSeverityStrip();wireFilterChips();wireFilterPanel();
  const wrap=$('#auditTableWrap');
  wrap.onkeydown=event=>{
    if(event.key==='ArrowDown'){event.preventDefault();moveCursor(1);}
    else if(event.key==='ArrowUp'){event.preventDefault();moveCursor(-1);}
    else if(event.key==='Home'){event.preventDefault();S.session.cursor=0;focusCursorRow();}
    else if(event.key==='End'){event.preventDefault();S.session.cursor=displayRows().length-1;focusCursorRow();}
    else if(event.key==='Enter'||event.key===' '){if(S.session.cursor<0)return;event.preventDefault();activateCursor();}
  };
  wrap.scrollTop=S.session.scrollTop||0;let frame=0;wrap.onscroll=()=>{S.session.scrollTop=wrap.scrollTop;if(frame)return;frame=requestAnimationFrame(()=>{frame=0;renderRows();});};requestAnimationFrame(renderRows);
}

/* --------------------------------------------------------- dashboard screen */

/* The dashboard is where an audit lands. It answers "what did this registry turn
   up" in one screen, every number on it is a way into the findings list, and it
   reads the same filters as the findings list so it can be scoped to one
   building, one discipline, or one phase. */
const DASH_RANK_LIMIT=8,DASH_NO_MILESTONE='No L2 milestone',DASH_DEEP_NEST=6,DASH_MAX_CHAIN=64;
const DASH_SEVERITY_MEANINGS={blocker:'Contradicts the registry or the approved lists',error:'Breaks an SSM SOP rule',warning:'A strong pattern says look',info:'Worth knowing'};

/* Findings are already filtered by everything. Rows are narrowed by the four
   registry dimensions only: a level, source, topic, or check filter says which
   findings to read, not which rows exist, so a row with nothing wrong with it
   still counts towards the structure numbers. */
function scopedResult(){
  const result=S.session&&S.session.result;if(!result)return {rows:[],findings:[],headerIds:[]};
  const findings=filteredFindings(),active=dimActiveSets();
  const key=[S.session.filteredCacheKey,DIM_KEYS.map(dimension=>(dimFilterMap()[dimension]||[]).join('~')).join('|')].join('#');
  if(S.session.scopedCacheKey===key&&S.session.scopedCache)return {...S.session.scopedCache,findings};
  const allRows=result.rows||[],allHeaders=result.headerIds||[];
  const rows=active.length?allRows.filter(row=>active.every(([field,values])=>values.has(auditNormId(row[field])))):[...allRows];
  let headerIds=[...allHeaders];
  if(active.length){const ids=new Set(rows.map(row=>auditNormId(row.equipmentId)));headerIds=allHeaders.filter(id=>ids.has(auditNormId(id)));}
  S.session.scopedCacheKey=key;S.session.scopedCache={rows,headerIds};
  return {rows,headerIds,findings};
}
function dashIsScoped(){return filterActiveCount()>0||!!clean(S.session.search);}
function dashSeverityCounts(findings){
  const counts=Object.fromEntries(SSM_AUDIT_SEVERITIES.map(level=>[level,0]));
  for(const finding of findings)if(counts[finding.severity]!=null)counts[finding.severity]++;
  return counts;
}
/* One pass that hands every block the same row-to-findings map, so two blocks can
   never disagree about whether a row is clean. */
function dashRowFindings(scoped){
  const map=new Map();
  for(const finding of scoped.findings){const row=registryRowFor(finding);if(!row)continue;const list=map.get(row);if(list)list.push(finding);else map.set(row,[finding]);}
  return map;
}
function dashRowsById(scoped){
  const rowsById=new Map();
  for(const row of scoped.rows){const id=auditNormId(row.equipmentId);if(id&&!rowsById.has(id))rowsById.set(id,row);}
  return rowsById;
}
function dashRuleCounts(scoped){
  const counts=new Map();for(const finding of scoped.findings)counts.set(finding.rule.id,(counts.get(finding.rule.id)||0)+1);return counts;
}
/* A UPN is labelled with the System Name most of its rows agree on — one row with
   a mistyped System Name should not become the name of the whole system. */
function dashRankBuckets(kind,findings){
  const field=DIM_FIELDS[kind],buckets=new Map();
  for(const finding of findings){
    const row=registryRowFor(finding);if(!row)continue;
    const value=clean(row[field]);if(!value&&kind!=='milestone')continue;
    const key=auditNormId(value);let bucket=buckets.get(key);
    if(!bucket){bucket={value,label:value||DASH_NO_MILESTONE,count:0,worst:'',names:new Map()};buckets.set(key,bucket);}
    bucket.count++;if(severityRank(finding.severity)>severityRank(bucket.worst))bucket.worst=finding.severity;
    const name=clean(row.systemName);if(name)bucket.names.set(name,(bucket.names.get(name)||0)+1);
  }
  const list=[...buckets.values()];
  if(kind==='upn')for(const bucket of list){
    const common=[...bucket.names.entries()].sort((a,b)=>b[1]-a[1]||natCmp(a[0],b[0]))[0];
    /* Approved System Names already lead with the UPN ("602  Medium Voltage"),
       so only prefix it when the name does not. */
    const name=common?clean(common[0]).replace(/\s+/g,' '):'';
    bucket.label=name&&auditNormId(name).startsWith(auditNormId(bucket.value)+' ')?name:clean(`${bucket.value}${name?' · '+name:''}`);
  }
  return list.sort((a,b)=>b.count-a.count||natCmp(a.label,b.label));
}

/* ---- milestone readiness ---- */

/* Can this phase be commissioned yet? One row per L2 milestone, counting the rows
   in it and how many of them are clean. */
function dashMilestoneReadiness(scoped){
  const byRow=dashRowFindings(scoped),buckets=new Map();
  for(const row of scoped.rows){
    const value=clean(row.milestone),key=auditNormId(value);
    let bucket=buckets.get(key);
    if(!bucket){bucket={key,label:value||DASH_NO_MILESTONE,rows:0,clean:0,findings:0,blocker:0,error:0,warning:0,info:0};buckets.set(key,bucket);}
    const list=byRow.get(row)||[];
    bucket.rows++;if(!list.length)bucket.clean++;
    for(const finding of list){bucket.findings++;if(bucket[finding.severity]!=null)bucket[finding.severity]++;}
  }
  return [...buckets.values()].sort((a,b)=>b.blocker-a.blocker||b.findings-a.findings||natCmp(a.label,b.label));
}
function dashSeverityNumeral(level,count){
  return `<span class="dash-numeral ${esc(level)} ${count?'':'is-zero'}" title="${esc(SEVERITY_LABELS[level])}">${count.toLocaleString()}</span>`;
}
function dashMilestoneTableHtml(scoped){
  const buckets=dashMilestoneReadiness(scoped);
  if(!buckets.length)return '<p class="dash-empty">No rows in scope.</p>';
  return `<div class="dash-table-scroll"><table class="dash-table dash-milestone-table">
    <thead><tr><th>L2 milestone</th><th>Rows</th><th>Findings</th><th class="dash-levels-head">By level</th><th>Clean</th></tr></thead>
    <tbody>${buckets.map(bucket=>{
      const percent=bucket.rows?Math.round(bucket.clean/bucket.rows*100):0;
      return `<tr data-dash-milestone="${esc(bucket.key)}" tabindex="0" title="Show the findings on ${esc(bucket.label)}">
        <td class="dash-table-name"><b>${esc(bucket.label)}</b></td>
        <td class="dash-table-num">${bucket.rows.toLocaleString()}</td>
        <td class="dash-table-num">${bucket.findings.toLocaleString()}</td>
        <td class="dash-levels"><span class="dash-levels-flex">${SSM_AUDIT_SEVERITIES.map(level=>dashSeverityNumeral(level,bucket[level])).join('')}</span></td>
        <td class="dash-clean"><span class="dash-clean-bar"><i style="width:${percent}%"></i></span><small>${percent}%</small></td></tr>`;
    }).join('')}</tbody></table></div>`;
}

/* ---- hierarchy health ---- */

/* Depth and "clean branch" both need the parent chain, so they are walked once.
   A chain that leaves the scope, or loops, is treated as a top: loops are already
   a finding of their own and must not turn into an endless walk here. */
function dashHierarchyHealth(scoped){
  const rowsById=dashRowsById(scoped),byRow=dashRowFindings(scoped),resolved=new Map();
  const resolveRow=start=>{
    const path=[],seen=new Set();let current=start;
    while(current&&!resolved.has(current)&&!seen.has(current)&&path.length<DASH_MAX_CHAIN){
      seen.add(current);path.push(current);
      const parentId=auditNormId(current.closestParent);
      const parent=parentId?rowsById.get(parentId)||null:null;
      current=parent===current?null:parent;
    }
    let base=current&&resolved.get(current)||{depth:-1,dirty:false};
    for(let index=path.length-1;index>=0;index--){
      const row=path[index];
      base={depth:base.depth+1,dirty:base.dirty||(byRow.get(row)||[]).length>0};
      resolved.set(row,base);
    }
    return resolved.get(start)||{depth:0,dirty:false};
  };
  let roots=0,maxDepth=0,deep=0,cleanBranch=0;
  for(const row of scoped.rows){
    const parent=auditNormId(row.closestParent);
    if(parent&&parent===auditNormId(row.systemName))roots++;
    const info=resolveRow(row);
    if(info.depth>maxDepth)maxDepth=info.depth;
    if(info.depth>=DASH_DEEP_NEST)deep++;
    if(!info.dirty)cleanBranch++;
  }
  const counts=dashRuleCounts(scoped);
  return {rows:scoped.rows.length,roots,maxDepth,deep,headers:scoped.headerIds.length,
    cleanBranchPercent:scoped.rows.length?Math.round(cleanBranch/scoped.rows.length*100):0,
    systemsNoRoot:counts.get(SSM_AUDIT_RULES.systemNoRoot.id)||0};
}
function dashHierarchyTiles(stats){
  return [
    {label:'Top of a system',value:stats.roots.toLocaleString(),note:'Parent is their own System Name'},
    {label:'Deepest nesting',value:stats.maxDepth.toLocaleString(),note:'Levels below the top of a system'},
    {label:`Nested ${DASH_DEEP_NEST}+ deep`,value:stats.deep.toLocaleString(),note:'Rows a long way down a branch',action:stats.deep?'hierarchy':'',title:'Browse these in the SSM hierarchy'},
    {label:'On a clean branch',value:`${stats.cleanBranchPercent}%`,note:'Nothing flagged on the row or above it'},
    {label:'Organizational headers',value:stats.headers.toLocaleString(),note:'Rows other equipment nests under',action:stats.headers?'hierarchy':'',title:'Browse these in the SSM hierarchy'},
    {label:'Systems with no top',value:stats.systemsNoRoot.toLocaleString(),note:stats.systemsNoRoot?'No row sits at the top of them':'Every system has a root row',
      action:stats.systemsNoRoot?`rule:${SSM_AUDIT_RULES.systemNoRoot.id}`:'',title:'Show the systems with no row at the top'},
  ];
}

/* ---- dependencies ---- */

const DASH_MOSTLY_ELECTRICAL=.7;
function dashDependencyStats(scoped){
  const rowsById=dashRowsById(scoped);
  let withDependencies=0,references=0,sameUpn=0,crossUpn=0,external=0;
  for(const row of scoped.rows){
    const list=auditSplitReferences(row.dependencies).filter(reference=>!/^N\/?A$/i.test(reference));
    if(!list.length)continue;
    withDependencies++;
    const hasProject=!!clean(row.dependencyProject);
    for(const reference of list){
      references++;
      const target=rowsById.get(auditNormId(reference));
      if(target)auditNormId(target.upn)===auditNormId(row.upn)?sameUpn++:crossUpn++;
      else if(hasProject)external++;
    }
  }
  let parentAlsoDependency=0,parentAlsoDependencyElectrical=0;
  for(const finding of scoped.findings){
    if(finding.rule.id!==SSM_AUDIT_RULES.parentAlsoDependency.id)continue;
    parentAlsoDependency++;
    const row=registryRowFor(finding);
    if(row&&auditPolarity(row.discipline)==='top-down')parentAlsoDependencyElectrical++;
  }
  const counts=dashRuleCounts(scoped);
  return {withDependencies,references,sameUpn,crossUpn,external,parentAlsoDependency,
    mostlyElectrical:!!parentAlsoDependency&&parentAlsoDependencyElectrical/parentAlsoDependency>=DASH_MOSTLY_ELECTRICAL,
    onHeaders:counts.get(SSM_AUDIT_RULES.dependencyOnHeader.id)||0};
}
function dashDependencyTiles(stats){
  return [
    {label:'Rows with dependencies',value:stats.withDependencies.toLocaleString(),note:'At least one dependency listed'},
    {label:'Dependency references',value:stats.references.toLocaleString(),note:'Every tag named across those rows'},
    {label:'Inside the same UPN',value:stats.sameUpn.toLocaleString(),note:'The hierarchy already carries the order'},
    {label:'Across UPNs',value:stats.crossUpn.toLocaleString(),note:'One system waiting on another'},
    {label:'In another project',value:stats.external.toLocaleString(),note:'Not in this registry, Dependency Project filled in'},
    {label:'Parent also a dependency',value:stats.parentAlsoDependency.toLocaleString(),note:stats.parentAlsoDependency?(stats.mostlyElectrical?'Mostly electrical':'Across disciplines'):'Nothing flagged',
      action:stats.parentAlsoDependency?`rule:${SSM_AUDIT_RULES.parentAlsoDependency.id}`:'',title:'Show the rows whose Closest Parent is also listed as a dependency'},
    {label:'Pointing at a header',value:stats.onHeaders.toLocaleString(),note:stats.onHeaders?'Depend on the equipment inside instead':'No dependency names a header',
      action:stats.onHeaders?`rule:${SSM_AUDIT_RULES.dependencyOnHeader.id}`:'',title:'Show the dependencies that name an organizational header'},
  ];
}

/* ---- checks overview ---- */

/* Every enabled check, including the ones that found nothing — silence is a
   result too, and an engineer reading this should be able to see which checks
   ran and came back clean. */
function dashCheckOverview(scoped){
  const counts=dashRuleCounts(scoped),worst=new Map();
  for(const finding of scoped.findings)if(severityRank(finding.severity)>severityRank(worst.get(finding.rule.id)||''))worst.set(finding.rule.id,finding.severity);
  const enabled=activeRules();
  let max=0;for(const rule of enabled)max=Math.max(max,counts.get(rule.id)||0);
  const groups=SSM_AUDIT_SOURCES.map(source=>{
    const entries=enabled.filter(rule=>rule.source===source.id)
      .map(rule=>({rule,count:counts.get(rule.id)||0,severity:worst.get(rule.id)||''}))
      .sort((a,b)=>b.count-a.count||severityRank(b.severity)-severityRank(a.severity)||natCmp(a.rule.title,b.rule.title));
    return {source,firing:entries.filter(entry=>entry.count>0),passing:entries.filter(entry=>!entry.count)};
  }).filter(group=>group.firing.length||group.passing.length);
  return {groups,max,firing:enabled.filter(rule=>(counts.get(rule.id)||0)>0).length,total:enabled.length};
}
function dashCheckRowHtml(entry,max){
  const width=max?Math.max(3,Math.round(entry.count/max*100)):0,firing=entry.count>0;
  const pill=firing?`<span class="audit-severity ${esc(entry.severity)}">${esc(SEVERITY_LABELS[entry.severity]||entry.severity)}</span>`:'<span class="audit-severity is-quiet">&mdash;</span>';
  return `<tr class="${firing?'':'is-passing'}" ${firing?`data-dash-rule="${esc(entry.rule.id)}" tabindex="0" title="Show only findings from this check"`:'title="This check ran and found nothing"'}>
    <td>${pill}</td>
    <td class="dash-check-name"><b>${esc(entry.rule.title)}</b><small>${esc(RULE_CATEGORY_LABELS[entry.rule.category]||entry.rule.category)}</small></td>
    <td class="dash-check-count">${firing?entry.count.toLocaleString():'0'}</td>
    <td><span class="dash-share"><span class="dash-share-track">${firing?`<i style="width:${width}%"></i>`:''}</span>${firing?'':'<small>Passing</small>'}</span></td></tr>`;
}
function dashCheckOverviewHtml(overview){
  if(!overview.groups.length)return '<p class="dash-empty">No checks to show.</p>';
  return overview.groups.map(group=>`<article class="dash-card dash-check-group">
    <header><div><h4>${esc(group.source.label)}</h4><span>${group.firing.length.toLocaleString()} of ${(group.firing.length+group.passing.length).toLocaleString()} checks found something</span></div></header>
    <div class="dash-table-scroll"><table class="dash-table dash-check-table">
      <thead><tr><th>Level</th><th>Check</th><th>Findings</th><th>Share</th></tr></thead>
      <tbody>${group.firing.map(entry=>dashCheckRowHtml(entry,overview.max)).join('')}${group.passing.map(entry=>dashCheckRowHtml(entry,overview.max)).join('')}</tbody>
    </table></div></article>`).join('');
}

/* ---- structure at a glance ---- */

const DASH_SITE_CLASS_RULE=SSM_AUDIT_RULES.siteClassification.id,DASH_ITEM_MASTER_RULE=SSM_AUDIT_RULES.itemMasterStandard.id;
function dashStructureStats(scoped){
  const upns=new Set(),disciplines=new Set(),siteClassifications=new Set();
  let fullyPhased=0,siteClassificationRows=0,itemMasterRows=0,itemMasterStandard=0;
  for(const row of scoped.rows){
    const classification=clean(row.equipmentClassification);
    if(classification&&!extoRev21Canonical('equipmentClassification',classification)){siteClassifications.add(auditNormId(classification));siteClassificationRows++;}
    const upn=auditNormId(row.upn),discipline=auditNormId(row.discipline);
    if(upn)upns.add(upn);if(discipline)disciplines.add(discipline);
    if(auditNormId(row.milestone)&&auditNormId(row.milestoneParent))fullyPhased++;
    /* Blank Item Masters are the header convention, not a migration state, so
       they are left out of the VF-versus-site-prefix split entirely. */
    const itemMaster=auditNormId(row.itemMaster);
    if(itemMaster&&!auditIsBlankItemMaster(row)){itemMasterRows++;if(/^VF[_\- ]/.test(itemMaster))itemMasterStandard++;}
  }
  const counts=dashRuleCounts(scoped);
  return {rows:scoped.rows.length,upns:upns.size,disciplines:disciplines.size,
    phased:scoped.rows.length?Math.round(fullyPhased/scoped.rows.length*100):0,
    siteClassifications:siteClassifications.size,siteClassificationRows,siteClassificationFindings:counts.get(DASH_SITE_CLASS_RULE)||0,
    itemMasterRows,itemMasterStandard,itemMasterPercent:itemMasterRows?Math.round(itemMasterStandard/itemMasterRows*100):0,
    itemMasterFindings:counts.get(DASH_ITEM_MASTER_RULE)||0};
}
function dashStatTiles(stats){
  return [
    {label:'Equipment rows',value:stats.rows.toLocaleString(),note:'Rows in scope'},
    {label:'Distinct UPNs',value:stats.upns.toLocaleString(),note:'Systems in this registry'},
    {label:'Disciplines',value:stats.disciplines.toLocaleString(),note:'Discipline values in use'},
    {label:'Rows with L1 and L2',value:`${stats.phased}%`,note:'Both milestone levels filled in'},
    {label:'Site-specific classifications',value:stats.siteClassifications.toLocaleString(),note:stats.siteClassifications?`${stats.siteClassificationRows.toLocaleString()} ${stats.siteClassificationRows===1?'row uses':'rows use'} them`:'Every classification is on the VF Exto Upload Template list',
      action:stats.siteClassificationFindings?`rule:${DASH_SITE_CLASS_RULE}`:'',title:'Show the rows whose Equipment Classification is not in the VF Exto Upload Template dropdown'},
    {label:'Item Master migration',value:`${stats.itemMasterPercent}%`,note:stats.itemMasterRows?`${stats.itemMasterStandard.toLocaleString()} of ${stats.itemMasterRows.toLocaleString()} on the VF standard`:'No Item Masters to migrate',
      bar:stats.itemMasterPercent,action:stats.itemMasterFindings?`rule:${DASH_ITEM_MASTER_RULE}`:'',title:'Show the rows still on a site-prefixed Item Master'},
  ];
}

/* ---- shared tile + block markup ---- */

/* The per-level breakdown: every check that fired, under its level, each row
   opening the findings narrowed to exactly that check at that level. */
function dashBreakdownHtml(scoped,severityCounts){
  const open=!!S.session.dashBreakdownOpen;
  const toggle=`<button class="btn ghost sm dash-breakdown-toggle" type="button" id="dashBreakdownToggle" aria-expanded="${open?'true':'false'}">${ic(open?'chevron-down':'chevron-right')}${open?'Hide the breakdown by check':'Breakdown by check'}</button>`;
  if(!open)return `<div class="dash-breakdown-bar">${toggle}</div>`;
  const byLevel=new Map(SSM_AUDIT_SEVERITIES.map(level=>[level,new Map()]));
  for(const finding of scoped.findings){const bucket=byLevel.get(finding.severity);if(!bucket)continue;const entry=bucket.get(finding.rule.id)||{rule:finding.rule,count:0};entry.count++;bucket.set(finding.rule.id,entry);}
  const columns=SSM_AUDIT_SEVERITIES.map(level=>{
    const entries=[...byLevel.get(level).values()].sort((left,right)=>right.count-left.count||natCmp(left.rule.title,right.rule.title));
    return `<div class="dash-breakdown-col ${level}"><header><span class="audit-severity ${level}">${esc(SEVERITY_LABELS[level])}</span><b>${(severityCounts[level]||0).toLocaleString()}</b></header>
      ${entries.length?entries.map(entry=>`<button type="button" class="dash-breakdown-item" data-dash-breakdown="${level}||${esc(entry.rule.id)}" title="Open just these findings"><span>${esc(entry.rule.title)}</span><b>${entry.count.toLocaleString()}</b></button>`).join(''):'<p class="dash-empty">Nothing at this level.</p>'}
    </div>`;
  }).join('');
  return `<div class="dash-breakdown-bar">${toggle}</div><div class="dash-card dash-breakdown">${columns}</div>`;
}
function dashSeverityTileHtml(level,count){
  const empty=!count,label=SEVERITY_LABELS[level];
  return `<button class="dash-tile ${level}${empty?' is-empty':''}" type="button" data-dash-severity="${level}" ${empty?'disabled':''} title="${empty?`Nothing at ${esc(label)} level`:`Show only ${esc(label)} findings`}">
    <span class="dash-tile-head"><span class="dash-tile-dot"></span>${esc(label)}</span>
    <b class="dash-tile-value">${count.toLocaleString()}</b>
    <span class="dash-tile-note">${esc(DASH_SEVERITY_MEANINGS[level]||'')}</span>
    <span class="dash-tile-go">${empty?'Nothing to show':`Open these findings${ic('arrow-right')}`}</span></button>`;
}
function dashRankRowHtml(bucket,kind,max){
  const width=max?Math.max(4,Math.round(bucket.count/max*100)):0;
  return `<button class="dash-rank-row" type="button" data-dash-rank="${esc(kind)}" data-dash-value="${esc(bucket.value)}" title="Show the ${bucket.count.toLocaleString()} findings on ${esc(bucket.label)}">
    <span class="dash-rank-name">${esc(bucket.label)}</span><b class="dash-rank-count">${bucket.count.toLocaleString()}</b>
    <span class="dash-rank-track"><i class="${esc(bucket.worst||'info')}" style="width:${width}%"></i></span></button>`;
}
function dashRankCardHtml(kind,title,note,findings){
  const buckets=dashRankBuckets(kind,findings),visible=buckets.slice(0,DASH_RANK_LIMIT),max=visible.length?visible[0].count:0,rest=buckets.length-visible.length;
  const body=visible.length?visible.map(bucket=>dashRankRowHtml(bucket,kind,max)).join(''):'<p class="dash-empty">Nothing flagged here.</p>';
  return `<article class="dash-card dash-rank-card"><header><h4>${esc(title)}</h4><span>${esc(note)}</span></header><div class="dash-rank-list">${body}</div>${rest>0?`<p class="dash-rank-more">+${rest.toLocaleString()} more</p>`:''}</article>`;
}
function dashStatHtml(stat){
  const bar=stat.bar==null?'':`<span class="dash-stat-bar"><i style="width:${Math.max(0,Math.min(100,stat.bar))}%"></i></span>`;
  const body=`<span>${esc(stat.label)}</span><b>${esc(stat.value)}</b>${bar}<small>${esc(stat.note)}</small>`;
  if(!stat.action)return `<div class="dash-stat">${body}</div>`;
  return `<button class="dash-stat is-action" type="button" data-dash-stat="${esc(stat.action)}" title="${esc(stat.title||'')}">${body}<span class="dash-stat-go">${ic('arrow-right')}</span></button>`;
}
/* The deep-dive sections share one tab strip so the dashboard stays one
   screen tall. The chosen tab survives filter changes and re-renders. */
const DASH_TABS=Object.freeze([
  {id:'milestones',label:'Milestone readiness',explainer:'One row per L2 phase: how much equipment is in it, what is flagged, and how much of it is already clean.'},
  {id:'hierarchy',label:'Hierarchy health',explainer:'How the registry is shaped, and how much of it sits on a branch with nothing flagged above it.'},
  {id:'dependencies',label:'Dependencies',explainer:'Every dependency named in scope, resolved against the rows in scope.'},
  {id:'checks',label:'Checks overview',explainer:''},
  {id:'structure',label:'Structure',explainer:'What this registry is made of, before any rule is applied.'},
]);
function dashTabBodyHtml(tab,scoped,overview){
  if(tab==='milestones')return `<div class="dash-card">${dashMilestoneTableHtml(scoped)}</div>`;
  if(tab==='hierarchy')return `<div class="dash-stats dash-stats-6">${dashHierarchyTiles(dashHierarchyHealth(scoped)).map(dashStatHtml).join('')}</div>`;
  if(tab==='dependencies')return `<div class="dash-stats dash-stats-4">${dashDependencyTiles(dashDependencyStats(scoped)).map(dashStatHtml).join('')}</div>`;
  if(tab==='checks')return `<div class="dash-check-grid">${dashCheckOverviewHtml(overview)}</div>`;
  return `<div class="dash-stats dash-stats-6">${dashStatTiles(dashStructureStats(scoped)).map(dashStatHtml).join('')}</div>`;
}
function dashTabsHtml(scoped,overview){
  const active=DASH_TABS.some(tab=>tab.id===S.session.dashTab)?S.session.dashTab:DASH_TABS[0].id;
  const current=DASH_TABS.find(tab=>tab.id===active);
  const explainer=active==='checks'?`${overview.firing.toLocaleString()} of ${overview.total.toLocaleString()} checks found something. The rest ran and came back clean.`:current.explainer;
  const extra=active==='checks'?`<button class="btn-link" type="button" id="dashSeeRules">See all in Rules${ic('arrow-right')}</button>`:'';
  return `<div class="dash-tabbar" role="tablist" aria-label="Dashboard sections">${DASH_TABS.map(tab=>`<button class="tabbtn ${tab.id===active?'on':''}" type="button" role="tab" aria-selected="${tab.id===active?'true':'false'}" data-dash-tab="${tab.id}">${esc(tab.label)}</button>`).join('')}</div>
    <div class="dash-tab-body">${dashBlockHtml(current.label,explainer,dashTabBodyHtml(active,scoped,overview),extra)}</div>`;
}
function dashBlockHtml(title,explainer,body,extra){
  return `<section class="dash-block"><div class="dash-block-head"><div><h3>${esc(title)}</h3><p>${esc(explainer)}</p></div>${extra||''}</div>${body}</section>`;
}

/* ---- the screen ---- */

function dashScopeLine(scoped){
  const summary=S.session.result.summary;
  if(!dashIsScoped())return `Audited ${summary.rows.toLocaleString()} rows &middot; ${summary.checks.toLocaleString()} checks &middot; ${esc(dashAuditedAt())}`;
  return `Scoped: ${scoped.rows.length.toLocaleString()} of ${summary.rows.toLocaleString()} rows &middot; ${scoped.findings.length.toLocaleString()} of ${summary.findings.toLocaleString()} findings`;
}
function dashAuditedAt(){return S.session.auditedAt?new Date(S.session.auditedAt).toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'}):'this session';}
function dashShellHtml(){
  const result=S.session.result,summary=result.summary,scoped=scopedResult();
  const head=`<header class="dash-head">
      <div class="dash-head-copy"><span class="eyebrow">Registry overview</span><h2>${esc(S.session.name||'This registry')}</h2><p class="${dashIsScoped()?'is-scoped':''}">${dashScopeLine(scoped)}</p></div>
      <div class="dash-head-actions">${statusMarkup(summary)}${filterButtonHtml('dashFilters')}<button class="btn primary" type="button" id="dashOpenFindings">${ic('check-check')}Open findings</button><button class="btn" type="button" id="dashExport">${ic('file-down')}Export report</button></div>
    </header>
    ${filterChipRowHtml(true)}`;
  if(!scoped.rows.length&&!scoped.findings.length)return `<section class="dash-shell" id="dashShell">${head}
    <div class="dash-card dash-nothing">${ic('search')}<b>Nothing in scope</b><span>The filters you have set leave no rows and no findings. Clear one to bring the registry back.</span><button class="btn" type="button" data-filter-clear-all>Clear all filters</button></div>
    <p class="dash-foot">${ic('lock')}Findings never leave this browser.</p></section>`;
  /* The level tiles are the level SUMMARY: they follow the dimension scope
     (discipline / milestone / UPN / building) but ignore the level filter itself,
     so a hidden level still shows how many findings it is hiding. */
  const levelScopeRows=new Set(scoped.rows.map(row=>auditNormId(row.equipmentId)));
  const levelScopeFindings=dimActiveSets().length?(S.session.result.findings||[]).filter(finding=>{const row=registryRowFor(finding);return row&&levelScopeRows.has(auditNormId(row.equipmentId));}):(S.session.result.findings||[]);
  const severityCounts=dashSeverityCounts(levelScopeFindings),overview=dashCheckOverview(scoped);
  return `<section class="dash-shell" id="dashShell">
    ${head}
    <div class="dash-tiles">${SSM_AUDIT_SEVERITIES.map(level=>dashSeverityTileHtml(level,severityCounts[level]||0)).join('')}</div>
    ${dashBreakdownHtml(scoped,severityCounts)}
    ${dashBlockHtml('Where the problems are','The same findings counted three ways. Pick a row to open the list narrowed to it.',
      `<div class="dash-rank-grid">${dashRankCardHtml('discipline','By discipline','Findings per discipline',scoped.findings)}${dashRankCardHtml('upn','By UPN and system','Findings per system',scoped.findings)}${dashRankCardHtml('milestone','By L2 milestone','Findings per L2 phase',scoped.findings)}</div>`)}
    ${dashTabsHtml(scoped,overview)}
    <p class="dash-foot">${ic('lock')}Findings never leave this browser.</p>
  </section>`;
}
export function renderDashboard(navigate){
  const result=S.session&&S.session.result;if(!result){S.homeMode='audit';navigate('upload');return;}
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='dashboard';S.homeMode='audit';
  $('#view').innerHTML=`<div id="dashHost">${dashShellHtml()}</div>${filterPanelHtml()}`;
  renderSideNav(navigate);
  /* The panel lives outside #dashHost, so a live filter change can rebuild every
     block underneath it without closing the panel or losing focus. */
  filterRerender=()=>redrawDashHost(navigate);
  wireDashboard(navigate);wireFilterPanel();
}
function dashOpenFindings(navigate,apply){
  S.session.hiddenSources=[];S.session.hiddenSeverities=[];S.session.hiddenCategories=[];S.session.hiddenRules=[];S.session.search='';
  /* The four registry dimensions are the dashboard's scope. A click narrows
     inside that scope rather than throwing it away. */
  S.session.groupBy='none';S.session.collapsedGroups=[];S.session.cursor=-1;S.session.scrollTop=0;
  if(apply)apply();
  invalidateFindingCaches();S.homeMode='audit';navigate('audit');
}
/* The scope and the four counting passes are pure given S.session, so they are
   exported for tests/ui-dashboard.test.mjs. The build strips the statement. */
export { dashCheckOverview, dashDependencyStats, dashHierarchyHealth, dashMilestoneReadiness, scopedResult };
/* Rebuild the dashboard without losing the page's scroll position. */
function redrawDashHost(navigate){
  const host=$('#dashHost'),view=$('#view');if(!host)return;
  const top=view?view.scrollTop:0;
  host.innerHTML=dashShellHtml();wireDashboard(navigate);
  if(view)view.scrollTop=top;
}
function wireDashboard(navigate){
  $('#dashOpenFindings')?.addEventListener('click',()=>dashOpenFindings(navigate));
  $('#dashExport')?.addEventListener('click',()=>openExportOptions());
  $('#dashSeeRules')?.addEventListener('click',()=>{S.homeMode='rules';navigate('rules');});
  $$('[data-dash-severity]').forEach(tile=>tile.onclick=()=>{const level=tile.dataset.dashSeverity;dashOpenFindings(navigate,()=>{S.session.hiddenSeverities=SSM_AUDIT_SEVERITIES.filter(item=>item!==level);});});
  $$('[data-dash-rank]').forEach(row=>row.onclick=()=>dashOpenFindings(navigate,()=>{dimFilterMap()[row.dataset.dashRank]=[auditNormId(row.dataset.dashValue)];}));
  $$('[data-dash-milestone]').forEach(row=>{
    const open=()=>dashOpenFindings(navigate,()=>{dimFilterMap().milestone=[row.dataset.dashMilestone];});
    row.onclick=open;row.onkeydown=event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();open();}};
  });
  $$('[data-dash-rule]').forEach(row=>{
    const open=()=>showOnlyRule(row.dataset.dashRule,navigate,true);
    row.onclick=open;row.onkeydown=event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();open();}};
  });
  $$('[data-dash-tab]').forEach(button=>button.onclick=()=>{S.session.dashTab=button.dataset.dashTab;redrawDashHost(navigate);});
  const breakdownToggle=$('#dashBreakdownToggle');if(breakdownToggle)breakdownToggle.onclick=()=>{S.session.dashBreakdownOpen=!S.session.dashBreakdownOpen;redrawDashHost(navigate);};
  $$('[data-dash-breakdown]').forEach(button=>button.onclick=()=>{
    const [level,ruleId]=button.dataset.dashBreakdown.split('||');
    dashOpenFindings(navigate,()=>{
      S.session.hiddenSeverities=SSM_AUDIT_SEVERITIES.filter(item=>item!==level);
      S.session.hiddenRules=[...new Set((S.session.result.findings||[]).map(finding=>finding.rule.id))].filter(id=>id!==ruleId);
    });
  });
  $$('[data-dash-stat]').forEach(tile=>tile.onclick=()=>{
    const action=tile.dataset.dashStat;
    if(action==='hierarchy'){if(S.session.snapshot)navigate('hierarchy');return;}
    if(action.startsWith('rule:'))showOnlyRule(action.slice(5),navigate,true);
  });
  wireFilterButtons();wireFilterChips();
}

/* --------------------------------------------------------- hierarchy screen */

function sessionHierarchy(){
  if(!S.session.hierarchy)S.session.hierarchy=buildSsmHierarchy(S.session.snapshot,S.session.result&&S.session.result.findings||[]);
  if(!S.session.hierarchyInitialized){S.session.hierarchyExpandedKeys=[...S.session.hierarchy.nodeByKey.values()].filter(node=>node.type==='building'||node.type==='discipline').map(node=>node.key);S.session.hierarchyInitialized=true;}
  return S.session.hierarchy;
}
function hierarchyHeaderIds(){
  if(S.session.headerIdSet)return S.session.headerIdSet;
  const result=S.session.result,ids=new Set((result&&result.headerIds||[]).map(auditNormId));
  S.session.headerIdSet=ids;return ids;
}
function nodeIsHeader(node){return !!node.isHeader||(node.type==='equipment'&&hierarchyHeaderIds().has(node.normalizedTag));}
function hierarchyDisciplineValue(node){return auditComparisonQuery(node&&node.label);}
function hierarchySystemValue(node){return `${auditComparisonQuery(node&&node.upn)}|${auditComparisonQuery(node&&node.label)}`;}
function hierarchyFilterSet(hierarchy){
  const building=S.session.hierarchyBuilding,discipline=S.session.hierarchyDiscipline,system=S.session.hierarchySystem;if(building==='all'&&discipline==='all'&&system==='all')return null;
  const visible=new Set(),addAncestors=node=>{const seen=new Set();while(node&&!seen.has(node.key)){seen.add(node.key);visible.add(node.key);node=hierarchy.nodeByKey.get(node.parentKey);}},addDescendants=node=>{const stack=[node],seen=new Set();while(stack.length){const current=stack.pop();if(!current||seen.has(current.key))continue;seen.add(current.key);visible.add(current.key);for(const child of current.children)stack.push(hierarchy.nodeByKey.get(child));}};
  for(const candidate of hierarchy.groups.systems){const disciplineNode=hierarchy.nodeByKey.get(candidate.parentKey);if(building!=='all'&&candidate.buildingKey!==building||discipline!=='all'&&hierarchyDisciplineValue(disciplineNode)!==discipline||system!=='all'&&hierarchySystemValue(candidate)!==system)continue;addAncestors(candidate);addDescendants(candidate);}
  return visible;
}
function hierarchySearchSets(hierarchy,query){
  if(!query)return {visible:null,direct:null};const visible=new Set(),direct=new Set(),addAncestors=node=>{const seen=new Set();while(node&&!seen.has(node.key)&&!visible.has(node.key)){seen.add(node.key);visible.add(node.key);node=hierarchy.nodeByKey.get(node.parentKey);}},addDescendants=node=>{const stack=[node],seen=new Set();while(stack.length){const current=stack.pop();if(!current||seen.has(current.key))continue;seen.add(current.key);visible.add(current.key);for(const child of current.children)stack.push(hierarchy.nodeByKey.get(child));}};
  for(const node of hierarchy.nodeByKey.values())if(node.searchKey.includes(query)){direct.add(node.key);visible.add(node.key);}
  for(const key of direct){const node=hierarchy.nodeByKey.get(key);addAncestors(hierarchy.nodeByKey.get(node&&node.parentKey));let parent=hierarchy.nodeByKey.get(node&&node.parentKey),covered=false;while(parent){if(direct.has(parent.key)){covered=true;break;}parent=hierarchy.nodeByKey.get(parent.parentKey);}if(!covered&&node&&node.children.length)addDescendants(node);}
  return {visible,direct};
}
function hierarchyRows(){
  const hierarchy=sessionHierarchy(),query=auditComparisonQuery(S.session.hierarchySearch),expanded=new Set(S.session.hierarchyExpandedKeys||[]),findingsOnly=!!S.session.hierarchyFindingsOnly;
  const key=[query,S.session.hierarchyBuilding,S.session.hierarchyDiscipline,S.session.hierarchySystem,findingsOnly?'F':'',...expanded].join('');
  if(S.session.hierarchyCacheKey===key&&S.session.hierarchyCacheRows)return S.session.hierarchyCacheRows;
  const filterVisible=hierarchyFilterSet(hierarchy),search=hierarchySearchSets(hierarchy,query),rows=[],stack=[...hierarchy.rootKeys].reverse().map(key=>({key,depth:0}));
  while(stack.length){
    const item=stack.pop(),node=hierarchy.nodeByKey.get(item.key);
    if(!node||filterVisible&&!filterVisible.has(node.key)||search.visible&&!search.visible.has(node.key)||findingsOnly&&!node.findingCount)continue;
    rows.push({node,depth:item.depth,isMatch:!!(search.direct&&search.direct.has(node.key)),isContext:!!(search.visible&&search.direct&&!search.direct.has(node.key))});
    const open=query||findingsOnly||expanded.has(node.key);if(open)for(let index=node.children.length-1;index>=0;index--)stack.push({key:node.children[index],depth:item.depth+1});
  }
  S.session.hierarchyCacheKey=key;S.session.hierarchyCacheRows=rows;return rows;
}
function hierarchyTypeLabel(node){return node.type==='building'?'Building':node.type==='discipline'?'Discipline':node.type==='system'?'System':node.isSyntheticHeader?'New header':nodeIsHeader(node)?'Header':'Equipment';}
function hierarchyNodeIcon(node){return node.type==='building'?'layers':node.type==='discipline'?'network':node.type==='system'?'hash':nodeIsHeader(node)?'folder-tree':'tag';}
function hierarchySeverity(node){return node.findings&&node.findings.reduce((winner,finding)=>severityRank(finding.severity)>severityRank(winner)?finding.severity:winner,'')||'';}
function hierarchyRowHtml(item,expanded,query){
  const node=item.node,hasChildren=!!node.children.length,open=!!query||!!S.session.hierarchyFindingsOnly||expanded.has(node.key);
  const toggle=hasChildren?`<button class="hierarchy-toggle ${open?'open':''}" type="button" data-hierarchy-toggle="${esc(node.key)}" aria-label="${open?'Collapse':'Expand'} ${esc(node.label)}" aria-expanded="${open}">${ic('chevron-right')}</button>`:'<span class="hierarchy-toggle-spacer"></span>';
  const equipment=node.type==='equipment',severity=hierarchySeverity(node),header=nodeIsHeader(node);
  const relationships=equipment?`${node.children.length.toLocaleString()} ${node.children.length===1?'child':'children'} &middot; ${node.dependencies.length.toLocaleString()} ${node.dependencies.length===1?'dependency':'dependencies'}`:`${node.equipmentCount.toLocaleString()} equipment`;
  const review=node.findingCount?`<span class="hierarchy-finding ${severity||'group'}">${node.findingCount.toLocaleString()}</span>`:`<span class="hierarchy-clear">${ic('check')}</span>`;
  const label=equipment?copyTagHtml(node.tag,highlightHtml(node.tag,query)):`<b>${highlightHtml(node.label,query)}</b>`;
  const subtitle=equipment?(node.description||'No description'):node.description;
  const warning=equipment&&(node.unresolvedParent||node.cycleBreak)?`<span class="hierarchy-link-warning" title="${node.cycleBreak?'This parent chain loops back on itself':'The closest parent is not in this registry'}">${ic('triangle-alert')}</span>`:'';
  return `<div class="hierarchy-row type-${node.type}${header?' is-header':''}${item.isContext?' search-context':''}${S.session.hierarchyFocusKey===node.key?' focused':''}" style="--hierarchy-depth:${Math.min(16,item.depth)}" data-hierarchy-node="${esc(node.key)}" ${equipment?'tabindex="0"':''}><div class="hierarchy-primary"><span class="hierarchy-guide" aria-hidden="true"></span>${toggle}<span class="hierarchy-node-icon">${ic(hierarchyNodeIcon(node))}</span><div class="hierarchy-node-copy">${label}<small>${highlightHtml(subtitle,query)}</small></div>${warning}</div><span class="hierarchy-type">${esc(hierarchyTypeLabel(node))}</span><code class="hierarchy-upn">${node.upn?esc(node.upn):'&mdash;'}</code><span class="hierarchy-relationships">${relationships}</span><span class="hierarchy-review">${review}${equipment?`<button class="icon-btn hierarchy-info" type="button" data-hierarchy-detail="${esc(node.key)}" aria-label="View ${esc(node.tag)} details">${ic('info')}</button>`:''}</span></div>`;
}
function renderHierarchyRows(){
  const wrap=$('#hierarchyTreeWrap'),body=$('#hierarchyTreeRows'),count=$('#hierarchyCount');if(!wrap||!body)return;
  const rows=hierarchyRows(),expanded=new Set(S.session.hierarchyExpandedKeys||[]),query=auditComparisonQuery(S.session.hierarchySearch);
  if(count)count.textContent=`${rows.length.toLocaleString()} rows shown`;
  if(!rows.length){body.innerHTML=`<div class="hierarchy-empty">${ic('search')}<b>Nothing here matches</b><span>${S.session.hierarchyFindingsOnly?'This part of the registry has no findings. Turn off "Findings only" to see everything.':'Clear the search or set the menus back to All.'}</span></div>`;return;}
  const visible=Math.min(HIERARCHY_MAX_ROWS,Math.max(26,Math.ceil(wrap.clientHeight/HIERARCHY_ROW_HEIGHT)+HIERARCHY_OVERSCAN*2)),treeScrollTop=Math.max(0,wrap.scrollTop-32),start=Math.min(Math.max(0,rows.length-visible),Math.max(0,Math.floor(treeScrollTop/HIERARCHY_ROW_HEIGHT)-HIERARCHY_OVERSCAN)),end=Math.min(rows.length,start+visible),spacer=height=>height?`<div class="hierarchy-spacer" style="height:${height}px" aria-hidden="true"></div>`:'';
  body.innerHTML=spacer(start*HIERARCHY_ROW_HEIGHT)+rows.slice(start,end).map(item=>hierarchyRowHtml(item,expanded,query)).join('')+spacer((rows.length-end)*HIERARCHY_ROW_HEIGHT);
  wireCopyTags(body);wireHierarchyRows();
}
function setHierarchyExpanded(values){S.session.hierarchyExpandedKeys=[...values];S.session.hierarchyCacheKey='';S.session.hierarchyCacheRows=null;}
function wireHierarchyRows(){
  $$('[data-hierarchy-toggle]').forEach(button=>button.onclick=event=>{event.stopPropagation();const expanded=new Set(S.session.hierarchyExpandedKeys||[]),key=button.dataset.hierarchyToggle;if(expanded.has(key))expanded.delete(key);else expanded.add(key);setHierarchyExpanded(expanded);renderHierarchyRows();$(`[data-hierarchy-toggle="${key}"]`)?.focus();});
  $$('[data-hierarchy-detail]').forEach(button=>button.onclick=event=>{event.stopPropagation();openHierarchyNode(button.dataset.hierarchyDetail,button);});
  $$('[data-hierarchy-node]').forEach(row=>{const node=sessionHierarchy().nodeByKey.get(row.dataset.hierarchyNode);row.onclick=event=>{if(event.target.closest('button'))return;if(node.type==='equipment')openHierarchyNode(node.key,row);else if(node.children.length)row.querySelector('[data-hierarchy-toggle]')?.click();};row.onkeydown=event=>{if(node.type==='equipment'&&(event.key==='Enter'||event.key===' ')){event.preventDefault();openHierarchyNode(node.key,row);}};});
}
function hierarchyDetailValue(label,value,copy=false){if(!value)return '';return `<dt>${esc(label)}</dt><dd>${copy?copyTagHtml(value):esc(value)}</dd>`;}
function openHierarchyNode(key,opener){
  const node=sessionHierarchy().nodeByKey.get(key);if(!node||node.type!=='equipment')return;S.session.selectedHierarchyNodeKey=key;S.session.opener=opener;
  const dependencies=node.dependencies.length?`<div class="hierarchy-dependency-list">${node.dependencies.map(tag=>copyTagHtml(tag)).join('')}</div>`:'<span class="hierarchy-none">No dependencies listed</span>';
  const findings=node.findings.length?`<div class="hierarchy-finding-list">${node.findings.slice(0,30).map(finding=>`<div><span class="audit-severity ${esc(finding.severity)}">${esc(SEVERITY_LABELS[finding.severity]||finding.severity)}</span><div><b>${esc(finding.why)}</b><small>${esc(finding.rule&&finding.rule.title||'Audit finding')}</small></div></div>`).join('')}</div>`:`<div class="hierarchy-no-findings">${ic('check-check')}Nothing flagged on this equipment</div>`;
  const source=node.isSyntheticHeader?'Created from a New closest-parent reference':`${clean(node.source.sheet)||'Registry'} / row ${node.source.row||'N/A'}`;
  $('#drawerTitle').textContent='Equipment';
  $('#drawerBody').innerHTML=`<div class="hierarchy-drawer"><div class="hierarchy-drawer-heading"><span class="hierarchy-node-icon">${ic(hierarchyNodeIcon(node))}</span><div><h3>${copyTagHtml(node.tag)}</h3><p>${esc(node.description||'No description')}</p></div></div>${node.isSyntheticHeader?'<div class="hierarchy-generated-note">This header was created from a closest-parent reference, not a registry row.</div>':''}${nodeIsHeader(node)&&!node.isSyntheticHeader?'<div class="hierarchy-generated-note">Other equipment nests under this row, so it acts as an organizational header.</div>':''}${node.unresolvedParent?'<div class="hierarchy-warning-note">'+ic('triangle-alert')+'The closest parent is not in this registry.</div>':''}${node.cycleBreak?'<div class="hierarchy-warning-note">'+ic('triangle-alert')+'This parent chain loops back on itself.</div>':''}<dl class="hierarchy-metadata">${hierarchyDetailValue('Building',node.building)}${hierarchyDetailValue('Discipline',node.discipline)}${hierarchyDetailValue('UPN',node.upn)}${hierarchyDetailValue('System Name',node.systemName)}${hierarchyDetailValue('Classification',node.classification)}${hierarchyDetailValue('Closest parent',node.closestParent,true)}${hierarchyDetailValue('Source',source)}</dl><section class="hierarchy-drawer-section"><h4>Dependencies <span>${node.dependencies.length.toLocaleString()}</span></h4>${dependencies}</section><section class="hierarchy-drawer-section"><h4>Findings <span>${node.findings.length.toLocaleString()}</span></h4>${findings}</section></div>`;
  wireCopyTags($('#drawerBody'));const backdrop=$('#drawerBack');animateOpen(backdrop);backdrop.setAttribute('aria-hidden','false');drawerTrapCleanup?.();drawerTrapCleanup=activateFocusTrap(backdrop,closeDrawer);$('#drawer').focus();
}
function focusHierarchyOnEquipment(equipmentId){
  const id=auditNormId(equipmentId);if(!id||!S.session.result||!S.session.snapshot)return;
  const hierarchy=sessionHierarchy(),node=[...hierarchy.nodeByKey.values()].find(candidate=>candidate.type==='equipment'&&candidate.normalizedTag===id);
  if(!node){toast('That equipment is not in the hierarchy');return;}
  S.session.hierarchyBuilding='all';S.session.hierarchyDiscipline='all';S.session.hierarchySystem='all';S.session.hierarchySearch='';S.session.hierarchyFindingsOnly=false;
  const expanded=new Set(S.session.hierarchyExpandedKeys||[]);let parent=hierarchy.nodeByKey.get(node.parentKey),guard=0;
  while(parent&&guard++<64){expanded.add(parent.key);parent=hierarchy.nodeByKey.get(parent.parentKey);}
  setHierarchyExpanded(expanded);S.session.hierarchyFocusKey=node.key;
  document.dispatchEvent(new CustomEvent('ssm-audit:navigate',{detail:{screen:'hierarchy'}}));
}
let hierarchyFocusTimer=0;
function applyHierarchyFocus(){
  const key=S.session.hierarchyFocusKey;if(!key)return;
  const rows=hierarchyRows(),index=rows.findIndex(item=>item.node.key===key),wrap=$('#hierarchyTreeWrap');
  /* Render BEFORE scrolling: on a freshly built screen the scroll container is
     still empty, so a scrollTop set first is clamped to 0 and the target lands
     off screen. The first render sizes the spacers; the scroll then sticks, and
     the second render draws the window around the target. */
  renderHierarchyRows();
  if(index>=0&&wrap){wrap.scrollTop=Math.max(0,index*HIERARCHY_ROW_HEIGHT-wrap.clientHeight/2+HIERARCHY_ROW_HEIGHT);S.session.hierarchyScrollTop=wrap.scrollTop;renderHierarchyRows();}
  $(`[data-hierarchy-node="${key}"]`)?.focus({preventScroll:true});
  /* A fresh jump gets its full three seconds even when an earlier timer is
     still pending. */
  clearTimeout(hierarchyFocusTimer);
  hierarchyFocusTimer=setTimeout(()=>{S.session.hierarchyFocusKey='';$('.hierarchy-row.focused')?.classList.remove('focused');},3200);
}
function hierarchyOptions(items,selected,label){return `<option value="all">${esc(label)}</option>`+items.map(item=>`<option value="${esc(item.value)}" ${selected===item.value?'selected':''}>${esc(item.label)}</option>`).join('');}
function uniqueHierarchyOptions(items,valueFor,labelFor){const options=new Map();for(const item of items){const value=valueFor(item);if(value&&!options.has(value))options.set(value,{value,label:labelFor(item)});}return [...options.values()].sort((left,right)=>left.label.localeCompare(right.label,undefined,{numeric:true}));}
function renderHierarchyResult(navigate){
  const result=S.session&&S.session.result;if(!result||!S.session.snapshot){navigate('upload');return;}
  teardownAuditFilters();S.screen='hierarchy';const hierarchy=sessionHierarchy(),summary=hierarchy.summary,fullscreen=!!S.session.hierarchyFullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  const buildingOptions=hierarchy.groups.buildings.map(node=>({value:node.key,label:node.label})),disciplineNodes=hierarchy.groups.disciplines.filter(node=>S.session.hierarchyBuilding==='all'||node.buildingKey===S.session.hierarchyBuilding),disciplineOptions=uniqueHierarchyOptions(disciplineNodes,hierarchyDisciplineValue,node=>node.label),systemNodes=hierarchy.groups.systems.filter(node=>(S.session.hierarchyBuilding==='all'||node.buildingKey===S.session.hierarchyBuilding)&&(S.session.hierarchyDiscipline==='all'||hierarchyDisciplineValue(hierarchy.nodeByKey.get(node.parentKey))===S.session.hierarchyDiscipline)),systemOptions=uniqueHierarchyOptions(systemNodes,hierarchySystemValue,node=>`${node.upn} · ${node.label}`);
  $('#view').innerHTML=`<section id="hierarchyShell" class="hierarchy-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><div class="audit-title"><span class="audit-title-icon">${ic('list-tree')}</span><div><h2>SSM Hierarchy</h2><p>${esc(S.session.name)}</p></div></div><span class="spacer"></span><span class="audit-status ready">${ic('list-tree')}${summary.equipment.toLocaleString()} equipment</span><span class="audit-status ${summary.findings?'review':'ready'}">${ic(summary.findings?'triangle-alert':'check-check')}${summary.findings.toLocaleString()} findings</span><button class="btn icon-btn" id="hierarchyFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
    <div class="hierarchy-toolbar"><div class="searchbox">${ic('search')}<input id="hierarchySearch" aria-label="Search the hierarchy" placeholder="Search tags, descriptions, parents, dependencies" value="${esc(S.session.hierarchySearch)}"></div><select id="hierarchyBuilding" aria-label="Filter by Building">${hierarchyOptions(buildingOptions,S.session.hierarchyBuilding,'All buildings')}</select><select id="hierarchyDiscipline" aria-label="Filter by Discipline">${hierarchyOptions(disciplineOptions,S.session.hierarchyDiscipline,'All disciplines')}</select><select id="hierarchySystem" aria-label="Filter by System">${hierarchyOptions(systemOptions,S.session.hierarchySystem,'All systems')}</select><button class="chip-toggle ${S.session.hierarchyFindingsOnly?'on':''}" type="button" id="hierarchyFindingsOnly" aria-pressed="${S.session.hierarchyFindingsOnly?'true':'false'}">${ic('triangle-alert')}Findings only</button><div class="hierarchy-actions"><button class="btn ghost sm" id="hierarchyCollapse" type="button">${ic('chevrons-up')}Collapse</button><button class="btn ghost sm" id="hierarchyExpand" type="button">${ic('chevrons-down')}Expand all</button></div><span id="hierarchyCount"></span></div>
    <div class="hierarchy-tree-panel"><div class="hierarchy-tree-wrap" id="hierarchyTreeWrap"><div class="hierarchy-tree-head"><b>Equipment</b><b>Type</b><b>UPN</b><b>Relationships</b><b>Findings</b></div><div id="hierarchyTreeRows"></div></div></div>
  </section>`;
  renderSideNav(navigate);wireHierarchyResult(navigate);
}
export { renderHierarchyResult };
function wireHierarchyResult(navigate){
  $('#hierarchyFullscreen').onclick=()=>{S.session.hierarchyFullscreen=!S.session.hierarchyFullscreen;renderHierarchyResult(navigate);};
  $('#hierarchySearch').oninput=event=>{S.session.hierarchySearch=event.target.value;debounceSearch(()=>{S.session.hierarchyCacheKey='';const wrap=$('#hierarchyTreeWrap');if(wrap)wrap.scrollTop=0;renderHierarchyRows();});};
  $('#hierarchyBuilding').onchange=event=>{S.session.hierarchyBuilding=event.target.value;S.session.hierarchyDiscipline='all';S.session.hierarchySystem='all';S.session.hierarchyScrollTop=0;S.session.hierarchyCacheKey='';renderHierarchyResult(navigate);};
  $('#hierarchyDiscipline').onchange=event=>{S.session.hierarchyDiscipline=event.target.value;S.session.hierarchySystem='all';S.session.hierarchyScrollTop=0;S.session.hierarchyCacheKey='';renderHierarchyResult(navigate);};
  $('#hierarchySystem').onchange=event=>{S.session.hierarchySystem=event.target.value;S.session.hierarchyScrollTop=0;S.session.hierarchyCacheKey='';renderHierarchyResult(navigate);};
  $('#hierarchyFindingsOnly').onclick=()=>{S.session.hierarchyFindingsOnly=!S.session.hierarchyFindingsOnly;S.session.hierarchyScrollTop=0;S.session.hierarchyCacheKey='';renderHierarchyResult(navigate);$('#hierarchyFindingsOnly')?.focus();};
  $('#hierarchyCollapse').onclick=()=>{setHierarchyExpanded(new Set());const wrap=$('#hierarchyTreeWrap');if(wrap)wrap.scrollTop=0;renderHierarchyRows();};
  $('#hierarchyExpand').onclick=()=>{const hierarchy=sessionHierarchy();setHierarchyExpanded(new Set([...hierarchy.nodeByKey.values()].filter(node=>node.children.length).map(node=>node.key)));const wrap=$('#hierarchyTreeWrap');if(wrap)wrap.scrollTop=0;renderHierarchyRows();};
  const wrap=$('#hierarchyTreeWrap');wrap.scrollTop=S.session.hierarchyScrollTop||0;let frame=0;wrap.onscroll=()=>{S.session.hierarchyScrollTop=wrap.scrollTop;if(frame)return;frame=requestAnimationFrame(()=>{frame=0;renderHierarchyRows();});};
  if(S.session.hierarchyFocusKey)requestAnimationFrame(applyHierarchyFocus);else requestAnimationFrame(renderHierarchyRows);
}

/* -------------------------------------------------------- comparison screen */

function comparisonSystems(){
  const query=auditComparisonQuery(S.comparison.systemSearch),filter=S.comparison.systemFilter||'different',sort=S.comparison.systemSort||'upn-asc';let systems=S.comparison.result.systems.filter(system=>(filter==='all'||filter==='different'&&system.status!=='aligned'||system.status===filter)&&(!query||auditComparisonQuery([system.upn,system.label,system.targetName,system.referenceName].join(' ')).includes(query)));
  systems=[...systems].sort((a,b)=>sort==='differences-desc'?(b.differenceCount-a.differenceCount||String(a.upn).localeCompare(String(b.upn),undefined,{numeric:true})):sort==='rows-desc'?(Math.max(b.targetRows,b.referenceRows)-Math.max(a.targetRows,a.referenceRows)||String(a.upn).localeCompare(String(b.upn),undefined,{numeric:true})):String(a.upn).localeCompare(String(b.upn),undefined,{numeric:true}));return systems;
}
function auditComparisonQuery(value){return clean(value).toUpperCase();}
function comparisonStatusLabel(status){return status==='aligned'?'Aligned':status==='target-only'?'Target only':status==='reference-only'?'Reference only':'Different';}
function reconcileComparisonSelection(){const systems=comparisonSystems();if(!systems.some(system=>system.upn===S.comparison.selectedUpn))S.comparison.selectedUpn=systems[0]&&systems[0].upn||'';return systems;}
function selectedComparisonSystem(){const systems=comparisonSystems();return systems.find(system=>system.upn===S.comparison.selectedUpn)||systems[0]||null;}
function systemListHtml(){const systems=comparisonSystems();if(!systems.length)return '<div class="compare-empty">No systems match these filters.</div>';return systems.map(system=>`<button class="compare-system ${system.upn===S.comparison.selectedUpn?'active':''}" type="button" data-compare-upn="${esc(system.upn)}"><span class="compare-upn">${esc(system.upn)}</span><span class="compare-system-copy"><b>${esc(stripSystemUpn(system.label))}</b><small>${system.targetRows.toLocaleString()} target &middot; ${system.referenceRows.toLocaleString()} reference</small></span><span class="compare-system-status ${system.status}">${system.status==='aligned'?ic('check'):system.differenceCount.toLocaleString()}</span></button>`).join('');}
function stripSystemUpn(value){return clean(value).replace(/^\s*[0-9]{3,4}\s*[-:]?\s*/,'')||clean(value);}
function countCard(label,target,reference){return `<div><span>${esc(label)}</span><b>${Number(target).toLocaleString()} <small>target</small></b><b>${Number(reference).toLocaleString()} <small>reference</small></b></div>`;}
function observationIcon(type){return type==='headers'?'folder-tree':type==='controls'?'network':type==='hierarchy'?'list-tree':type==='coverage'?'layers':'tag';}
function comparisonObservations(system){
  if(!system.observations.length)return '<div class="compare-aligned-note">'+ic('check-check')+'No pattern differences once Building is left out.</div>';
  const visible=system.observations.slice(0,COMPARE_MAX_OBSERVATIONS),remaining=system.observations.length-visible.length;
  return `<div class="compare-observation-list">${visible.map(item=>`<div class="compare-observation ${item.type}"><span>${ic(observationIcon(item.type))}</span><div><b>${esc(item.title)}</b><small>${esc(item.subject)}</small></div><code>${esc(item.target)} / ${esc(item.reference)}</code></div>`).join('')}${remaining?`<p class="compare-more">Showing ${visible.length.toLocaleString()} of ${system.observations.length.toLocaleString()}. Export the comparison for the full list.</p>`:''}</div>`;
}
function treeExpansion(system){
  const store=S.comparison.treeExpandedByUpn||(S.comparison.treeExpandedByUpn={});
  if(!Array.isArray(store[system.upn]))store[system.upn]=[];
  return new Set(store[system.upn]);
}
function setTreeExpansion(system,expanded){S.comparison.treeExpandedByUpn[system.upn]=[...expanded];}
function comparisonTreeRows(system){
  const byId=new Map(system.pairs.map(pair=>[pair.id,pair])),expanded=treeExpansion(system),rows=[],seen=new Set(),stack=[];
  const addRoots=ids=>{for(let index=ids.length-1;index>=0;index--)stack.push({id:ids[index],depth:0});};
  addRoots(system.treeRoots||[]);
  const collect=()=>{while(stack.length){const item=stack.pop(),pair=byId.get(item.id);if(!pair||seen.has(pair.id))continue;seen.add(pair.id);rows.push({pair,depth:item.depth});if(expanded.has(pair.id)){for(let index=pair.childrenIds.length-1;index>=0;index--)stack.push({id:pair.childrenIds[index],depth:item.depth+1});}}};
  collect();
  if(!rows.length&&system.pairs.length){addRoots([system.pairs[0].id]);collect();}
  return rows;
}
function comparisonTreeCell(pair,node,side,depth,expanded){
  const hasChildren=pair.childrenIds.length>0,isExpanded=expanded.has(pair.id),label=isExpanded?'Collapse matching branches':'Expand matching branches';
  const toggle=hasChildren?`<button class="compare-tree-toggle ${isExpanded?'open':''}" type="button" data-compare-tree-toggle="${esc(pair.id)}" aria-label="${label}" aria-expanded="${isExpanded}">${ic('chevron-right')}</button>`:'<span class="compare-tree-toggle-spacer"></span>';
  if(!node)return `<div class="compare-tree-cell missing" style="--tree-depth:${Math.min(10,depth)}">${toggle}<span>Not present</span></div>`;
  return `<div class="compare-tree-cell ${side}" style="--tree-depth:${Math.min(10,depth)}">${toggle}<span class="compare-tree-rail"></span><div class="compare-tree-copy"><b>${copyTagHtml(node.tag)}</b><small>${esc(node.role)}</small><span>Parent: ${esc(node.parentId||'System / external root')}</span></div></div>`;
}
function comparisonTreeRowHtml(item,expanded){
  const pair=item.pair,detail=pair.differences[0]||'Aligned';
  return `<div class="compare-tree-row ${pair.status} ${pair.placementMismatch?'placement-mismatch':''}" data-tree-pair="${esc(pair.id)}">${comparisonTreeCell(pair,pair.target,'target',item.depth,expanded)}<div class="compare-tree-state"><span class="compare-tree-state-mark"></span><b>${esc(comparisonStatusLabel(pair.status))}</b><small title="${esc(detail)}">${esc(detail)}</small></div>${comparisonTreeCell(pair,pair.reference,'reference',item.depth,expanded)}</div>`;
}
function renderComparisonTree(){
  const wrap=$('#compareTreeWrap'),body=$('#compareTreeRows'),count=$('#compareTreeCount'),system=selectedComparisonSystem();if(!wrap||!body||!system)return;const rows=comparisonTreeRows(system);if(count)count.textContent=`${rows.length.toLocaleString()} visible of ${system.pairs.length.toLocaleString()}`;
  const expanded=treeExpansion(system),visible=Math.min(COMPARE_TREE_MAX_ROWS,Math.max(16,Math.ceil(wrap.clientHeight/COMPARE_TREE_ROW_HEIGHT)+COMPARE_TREE_OVERSCAN*2)),start=Math.min(Math.max(0,rows.length-visible),Math.max(0,Math.floor(wrap.scrollTop/COMPARE_TREE_ROW_HEIGHT)-COMPARE_TREE_OVERSCAN)),end=Math.min(rows.length,start+visible),spacer=height=>height?`<div class="compare-tree-spacer" style="height:${height}px" aria-hidden="true"></div>`:'';
  body.innerHTML=spacer(start*COMPARE_TREE_ROW_HEIGHT)+rows.slice(start,end).map(item=>comparisonTreeRowHtml(item,expanded)).join('')+spacer((rows.length-end)*COMPARE_TREE_ROW_HEIGHT);wireCopyTags(body);wireComparisonTreeToggles();
}
function wireComparisonTreeToggles(){
  $$('[data-compare-tree-toggle]').forEach(button=>button.onclick=()=>{const system=selectedComparisonSystem(),expanded=treeExpansion(system),id=button.dataset.compareTreeToggle;if(expanded.has(id))expanded.delete(id);else expanded.add(id);setTreeExpansion(system,expanded);renderComparisonTree();});
}
function wireComparisonTree(){
  const system=selectedComparisonSystem(),wrap=$('#compareTreeWrap');
  $('#compareTreeExpand').onclick=()=>{setTreeExpansion(system,new Set(system.pairs.filter(pair=>pair.childrenIds.length).map(pair=>pair.id)));S.comparison.treeScrollTop=0;if(wrap)wrap.scrollTop=0;renderComparisonTree();};
  $('#compareTreeCollapse').onclick=()=>{setTreeExpansion(system,new Set());S.comparison.treeScrollTop=0;if(wrap)wrap.scrollTop=0;renderComparisonTree();};
  if(wrap){wrap.scrollTop=S.comparison.treeScrollTop||0;let frame=0;wrap.onscroll=()=>{S.comparison.treeScrollTop=wrap.scrollTop;if(frame)return;frame=requestAnimationFrame(()=>{frame=0;renderComparisonTree();});};requestAnimationFrame(renderComparisonTree);}
}
function pairVisible(pair){const filter=S.comparison.rowFilter||'different',query=auditComparisonQuery(S.comparison.rowSearch);if(filter!=='all'&&(filter==='different'?pair.status==='aligned':pair.status!==filter))return false;if(!query)return true;const nodes=[pair.target,pair.reference].filter(Boolean);return auditComparisonQuery(nodes.flatMap(node=>[node.tag,node.role,node.parentRole,node.classification,node.headerName,node.dependencyRoles.join(' ')]).join(' ')).includes(query);}
function comparisonNodeHtml(node,side){if(!node)return `<div class="compare-node missing"><span>${side==='target'?'Not in the registry being audited':'Not in the finished project'}</span></div>`;return `<div class="compare-node"><div class="compare-node-main" style="--node-depth:${Math.min(6,node.depth)}"><span class="compare-node-line"></span><div><b>${copyTagHtml(node.tag)}</b><small>${esc(node.role)}</small></div></div><dl><dt>Parent type</dt><dd>${esc(node.parentRole)}</dd>${node.underHeader?`<dt>Header</dt><dd>${copyTagHtml(node.headerName)}</dd>`:''}<dt>Discipline</dt><dd>${esc(node.disciplineKind)}</dd>${node.dependencyRoles.length?`<dt>Dependencies</dt><dd>${esc(node.dependencyRoles.join('; '))}</dd>`:''}</dl></div>`;}
function pairHtml(pair){return `<article class="compare-pair ${pair.status}" data-compare-pair><div class="compare-pair-status"><span class="comparison-pill ${pair.status}">${comparisonStatusLabel(pair.status)}</span><b>${esc(pair.differences.join(' · ')||'Same placement')}</b><small>${esc(pair.matchReason)}</small></div><div class="compare-side target"><div class="compare-side-label">Registry being audited</div>${comparisonNodeHtml(pair.target,'target')}</div><div class="compare-side reference"><div class="compare-side-label">Finished project</div>${comparisonNodeHtml(pair.reference,'reference')}</div></article>`;}
function renderComparisonPairs(){
  const wrap=$('#comparePairWrap'),body=$('#comparePairs'),count=$('#comparePairCount');if(!wrap||!body)return;const system=selectedComparisonSystem(),pairs=system?system.pairs.filter(pairVisible):[];if(count)count.textContent=`${pairs.length.toLocaleString()} of ${(system&&system.pairs.length||0).toLocaleString()} mappings`;
  const visible=Math.min(COMPARE_MAX_ROWS,Math.max(18,Math.ceil(wrap.clientHeight/COMPARE_ROW_HEIGHT)+COMPARE_OVERSCAN*2)),start=Math.min(Math.max(0,pairs.length-visible),Math.max(0,Math.floor(wrap.scrollTop/COMPARE_ROW_HEIGHT)-COMPARE_OVERSCAN)),end=Math.min(pairs.length,start+visible);
  if(!pairs.length){body.innerHTML='<div class="compare-detail-empty"><b>No equipment matches</b><span>Clear the search or choose a broader filter.</span></div>';return;}
  const spacer=height=>height?`<div class="compare-pair-spacer" style="height:${height}px" aria-hidden="true"></div>`:'';body.innerHTML=spacer(start*COMPARE_ROW_HEIGHT)+pairs.slice(start,end).map(pairHtml).join('')+spacer((pairs.length-end)*COMPARE_ROW_HEIGHT);wireCopyTags(body);
}
function comparisonDetailTabs(system){
  const active=S.comparison.detailTab||'hierarchy',tabs=[
    {id:'hierarchy',icon:'list-tree',label:'Hierarchy',count:system.pairs.length},
    {id:'differences',icon:'triangle-alert',label:'Differences',count:system.observations.length},
    {id:'mapping',icon:'table-2',label:'Equipment mapping',count:system.pairs.length},
  ];
  return `<div class="compare-detail-tabs" role="tablist" aria-label="Project comparison views">${tabs.map(tab=>`<button class="compare-detail-tab ${active===tab.id?'active':''}" type="button" data-compare-detail-tab="${tab.id}" role="tab" aria-selected="${active===tab.id}">${ic(tab.icon)}<span>${esc(tab.label)}</span><b>${tab.count.toLocaleString()}</b></button>`).join('')}</div>`;
}
function comparisonDetailPanel(system){
  const active=S.comparison.detailTab||'hierarchy';
  if(active==='differences')return `<section class="compare-subview compare-observations" role="tabpanel"><div class="compare-section-title"><div><b>Differences in this system</b><span>Counts compare meaning and nesting, not Building.</span></div><span>${system.observations.length.toLocaleString()} observations</span></div>${comparisonObservations(system)}</section>`;
  if(active==='mapping')return `<section class="compare-subview compare-mapping" role="tabpanel"><div class="compare-section-title"><div><b>Equipment side by side</b><span>Matched rows stay visible for context; filter to focus on changes.</span></div><span id="comparePairCount"></span></div><div class="compare-pair-toolbar"><div class="searchbox">${ic('search')}<input id="compareRowSearch" aria-label="Search mapped equipment" placeholder="Search tags, descriptions, parents" value="${esc(S.comparison.rowSearch)}"></div><select id="compareRowFilter" aria-label="Filter mappings"><option value="different" ${S.comparison.rowFilter==='different'?'selected':''}>Differences only</option><option value="all" ${S.comparison.rowFilter==='all'?'selected':''}>All mappings</option><option value="changed" ${S.comparison.rowFilter==='changed'?'selected':''}>Changed pairs</option><option value="target-only" ${S.comparison.rowFilter==='target-only'?'selected':''}>Target only</option><option value="reference-only" ${S.comparison.rowFilter==='reference-only'?'selected':''}>Reference only</option><option value="aligned" ${S.comparison.rowFilter==='aligned'?'selected':''}>Aligned pairs</option></select></div><div class="compare-pair-wrap" id="comparePairWrap"><div id="comparePairs"></div></div></section>`;
  return `<section class="compare-subview compare-tree-section" role="tabpanel"><div class="compare-section-title"><div><b>Both hierarchies, side by side</b><span>Open either side to reveal both matched branches.</span></div><span id="compareTreeCount"></span><div class="compare-tree-actions"><button class="btn ghost sm" id="compareTreeCollapse" type="button">${ic('chevrons-up')}Collapse</button><button class="btn ghost sm" id="compareTreeExpand" type="button">${ic('list-tree')}Expand all</button></div></div><div class="compare-tree-grid"><div class="compare-tree-head"><b>Registry being audited</b><span>Alignment</span><b>Finished project</b></div><div class="compare-tree-wrap" id="compareTreeWrap"><div id="compareTreeRows"></div></div></div></section>`;
}
function renderComparisonDetail(){
  const system=selectedComparisonSystem(),detail=$('#compareDetail');if(!detail)return;if(!system){detail.innerHTML=`<div class="compare-detail-empty">${ic('search')}<b>No systems match</b><span>Clear the search or broaden the filter.</span></div>`;return;}const types=comparisonSystemTypes(system);
  detail.innerHTML=`<div class="compare-detail-head"><div><span class="eyebrow">UPN ${esc(system.upn)}</span><h3>${esc(stripSystemUpn(system.label))}</h3><p>Building values are left out. Equipment is matched on tag identity, meaning, and where it sits.</p></div><span class="comparison-pill ${system.status}">${comparisonStatusLabel(system.status)}</span></div>
  <div class="compare-counts">${countCard('Equipment',system.targetRows,system.referenceRows)}${countCard('Headers',system.targetHeaders,system.referenceHeaders)}${countCard('I&C / Controls',system.targetControls,system.referenceControls)}<div><span>Pattern findings</span><b>${system.observations.length.toLocaleString()} <small>observed</small></b><b>${(types.hierarchy+types.headers+types.controls).toLocaleString()} <small>nesting</small></b></div></div>
  ${comparisonDetailTabs(system)}<div class="compare-detail-body">${comparisonDetailPanel(system)}</div>`;
  wireComparisonDetail();
}
function wireComparisonDetail(){
  $$('[data-compare-detail-tab]').forEach(button=>button.onclick=()=>{const id=button.dataset.compareDetailTab;S.comparison.detailTab=id;renderComparisonDetail();$(`[data-compare-detail-tab="${id}"]`)?.focus();});
  const search=$('#compareRowSearch'),filter=$('#compareRowFilter'),wrap=$('#comparePairWrap');if(search)search.oninput=event=>{S.comparison.rowSearch=event.target.value;debounceSearch(()=>{if(wrap)wrap.scrollTop=0;renderComparisonPairs();});};if(filter)filter.onchange=event=>{S.comparison.rowFilter=event.target.value;if(wrap)wrap.scrollTop=0;renderComparisonPairs();};
  if($('#compareTreeWrap'))wireComparisonTree();
  if(wrap){wrap.scrollTop=S.comparison.pairScrollTop||0;let frame=0;wrap.onscroll=()=>{S.comparison.pairScrollTop=wrap.scrollTop;if(frame)return;frame=requestAnimationFrame(()=>{frame=0;renderComparisonPairs();});};requestAnimationFrame(renderComparisonPairs);}
}

export function renderComparisonResult(navigate){
  const result=S.comparison&&S.comparison.result;if(!result){S.homeMode='compare';navigate('upload');return;}S.screen='compare';reconcileComparisonSelection();const summary=result.summary,fullscreen=!!S.comparison.fullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  $('#view').innerHTML=`<section id="compareShell" class="compare-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><div class="audit-title"><span class="audit-title-icon">${ic('square-stack')}</span><div><h2>Compare Projects</h2><p>${esc(S.comparison.targetName)} &middot; ${esc(S.comparison.referenceName)}</p></div></div><span class="spacer"></span><span class="compare-building-note">${ic('layers')}Building ignored</span><button class="btn" id="compareBack">${ic('upload')}Change files</button><button class="btn" id="exportComparison">${ic('file-down')}Export comparison</button><button class="btn icon-btn" id="compareFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
    <div class="compare-summary"><div><span>Systems compared</span><b>${summary.systems.toLocaleString()}</b></div><div class="difference"><span>With differences</span><b>${(summary.differentSystems+summary.targetOnlySystems+summary.referenceOnlySystems).toLocaleString()}</b></div><div class="aligned"><span>Aligned</span><b>${summary.alignedSystems.toLocaleString()}</b></div><div><span>Changed pairs</span><b>${summary.changedRows.toLocaleString()}</b></div><div><span>Target only</span><b>${summary.targetOnlyRows.toLocaleString()}</b></div><div><span>Reference only</span><b>${summary.referenceOnlyRows.toLocaleString()}</b></div></div>
    <div class="compare-workspace"><aside class="compare-system-panel"><div class="compare-system-tools"><div class="searchbox">${ic('search')}<input id="compareSystemSearch" aria-label="Search systems" placeholder="Find a UPN or system" value="${esc(S.comparison.systemSearch)}"></div><div><select id="compareSystemFilter" aria-label="Filter systems"><option value="different" ${S.comparison.systemFilter==='different'?'selected':''}>Differences</option><option value="all" ${S.comparison.systemFilter==='all'?'selected':''}>All systems</option><option value="aligned" ${S.comparison.systemFilter==='aligned'?'selected':''}>Aligned</option><option value="target-only" ${S.comparison.systemFilter==='target-only'?'selected':''}>Target only</option><option value="reference-only" ${S.comparison.systemFilter==='reference-only'?'selected':''}>Reference only</option></select><select id="compareSystemSort" aria-label="Sort systems"><option value="upn-asc" ${S.comparison.systemSort==='upn-asc'?'selected':''}>UPN</option><option value="differences-desc" ${S.comparison.systemSort==='differences-desc'?'selected':''}>Most differences</option><option value="rows-desc" ${S.comparison.systemSort==='rows-desc'?'selected':''}>Most equipment</option></select></div></div><div class="compare-system-list" id="compareSystemList">${systemListHtml()}</div></aside><section class="compare-detail" id="compareDetail" aria-label="Selected system comparison"></section></div>
  </section>`;
  renderSideNav(navigate);wireComparisonResult(navigate);renderComparisonDetail();
}

function renderSystemList(){const list=$('#compareSystemList');if(!list)return;reconcileComparisonSelection();list.innerHTML=systemListHtml();wireSystemButtons();}
function wireSystemButtons(){$$('[data-compare-upn]').forEach(button=>button.onclick=()=>{const upn=button.dataset.compareUpn;S.comparison.selectedUpn=upn;S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;renderSystemList();renderComparisonDetail();$(`[data-compare-upn="${upn}"]`)?.focus();});}
function wireComparisonResult(navigate){
  teardownAuditFilters();$('#compareBack').onclick=()=>{S.homeMode='compare';navigate('upload');};$('#exportComparison').onclick=exportSsmComparisonXlsx;$('#compareFullscreen').onclick=()=>{S.comparison.fullscreen=!S.comparison.fullscreen;renderComparisonResult(navigate);};
  $('#compareSystemSearch').oninput=event=>{S.comparison.systemSearch=event.target.value;debounceSearch(()=>{renderSystemList();renderComparisonDetail();});};$('#compareSystemFilter').onchange=event=>{S.comparison.systemFilter=event.target.value;renderSystemList();renderComparisonDetail();};$('#compareSystemSort').onchange=event=>{S.comparison.systemSort=event.target.value;renderSystemList();renderComparisonDetail();};wireSystemButtons();
}
