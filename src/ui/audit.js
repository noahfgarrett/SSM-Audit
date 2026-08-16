import { $, $$, clean, esc, natCmp } from '../core/text.js'
import { S, resetSession } from '../state.js'
import { readArrayBuffer } from '../io/workbook.js'
import { auditNormId, auditSnapshotFromWorkbook, auditSplitReferences } from '../audit/model.js'
import { runSsmAudit, SSM_AUDIT_CATEGORIES, SSM_AUDIT_RULES, SSM_AUDIT_SEVERITIES, SSM_AUDIT_SOURCES } from '../audit/engine.js'
import { compareSsmRegistries, comparisonSystemTypes } from '../audit/compare.js'
import { buildSsmHierarchy } from '../audit/hierarchy.js'
import { exportSsmAuditXlsx, exportSsmComparisonXlsx } from '../audit/export.js'
import { ic } from './icons.js'
import { activateFocusTrap, copyTagHtml, runWithProgress, toast, wireCopyTags } from './feedback.js'

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
let auditOutsideHandler=null,auditEscapeHandler=null,drawerTrapCleanup=null,searchDebounceTimer=0;

/* ---------------------------------------------------------------- shell nav */

const NAV_SECTIONS=[
  {id:'dashboard',icon:'layout-dashboard',label:'Dashboard',hint:'Everything this registry turned up, at a glance'},
  {id:'audit',icon:'check-check',label:'Audit findings',hint:'Every issue found in this registry'},
  {id:'hierarchy',icon:'list-tree',label:'SSM hierarchy',hint:'Browse the registry as a tree'},
  {id:'compare',icon:'square-stack',label:'Compare projects',hint:'Line this registry up beside a finished one'},
  {id:'rules',icon:'book-open',label:'Rules',hint:'What the audit checks, in plain language'},
];
function navActiveId(){
  if(S.screen==='dashboard')return 'dashboard';
  if(S.screen==='audit')return 'audit';
  if(S.screen==='hierarchy')return 'hierarchy';
  if(S.screen==='compare')return 'compare';
  if(S.screen==='rules')return 'rules';
  return S.homeMode==='compare'?'compare':S.homeMode==='rules'?'rules':'audit';
}
function navBadges(id){
  const result=S.session&&S.session.result,comparison=S.comparison&&S.comparison.result;
  if(id==='audit'&&result){
    const blockers=result.summary.severity.blocker;
    return `<b class="nav-count">${result.summary.findings.toLocaleString()}</b>${blockers?`<b class="nav-count blocker" title="${blockers.toLocaleString()} ${SEVERITY_LABELS.blocker.toLowerCase()}">${blockers.toLocaleString()}</b>`:''}`;
  }
  if(id==='hierarchy'&&S.session&&S.session.snapshot)return `<b class="nav-count">${S.session.snapshot.rows.length.toLocaleString()}</b>`;
  if(id==='compare'&&comparison)return `<b class="nav-count">${(comparison.summary.differentSystems+comparison.summary.targetOnlySystems+comparison.summary.referenceOnlySystems).toLocaleString()}</b>`;
  return '';
}
function navDisabled(id){
  if(id==='hierarchy')return !(S.session&&S.session.result&&S.session.snapshot);
  if(id==='dashboard')return !(S.session&&S.session.result);
  return false;
}
function navItemHtml(section,active,collapsed){
  const disabled=navDisabled(section.id),title=collapsed||disabled?`${section.label}${disabled?' — run an audit first':''}`:section.hint;
  return `<button class="sidenav-item ${active===section.id?'active':''}" type="button" data-nav="${section.id}" ${disabled?'disabled':''} title="${esc(title)}" aria-current="${active===section.id?'page':'false'}">${ic(section.icon)}<span>${esc(section.label)}</span>${navBadges(section.id)}</button>`;
}
export function renderSideNav(navigate){
  const nav=$('#sideNav');if(!nav)return;
  const active=navActiveId(),collapsed=!!S.ui.navCollapsed,canExport=S.screen==='compare'?!!(S.comparison&&S.comparison.result):!!(S.session&&S.session.result);
  nav.classList.toggle('collapsed',collapsed);
  nav.innerHTML=`<div class="sidenav-items">${NAV_SECTIONS.map(section=>navItemHtml(section,active,collapsed)).join('')}</div>
  <div class="sidenav-foot">
    <button class="sidenav-item" type="button" id="navGuide" title="Guide">${ic('book-open')}<span>Guide</span></button>
    <button class="sidenav-item" type="button" id="navExport" ${canExport?'':'disabled'} title="${canExport?'Download an Excel report':'Run an audit first'}">${ic('file-down')}<span>Export</span></button>
    <button class="sidenav-collapse" type="button" id="navCollapse" title="${collapsed?'Expand menu':'Collapse menu'}" aria-label="${collapsed?'Expand menu':'Collapse menu'}">${ic(collapsed?'panel-left-open':'panel-left-close')}<span>Collapse</span></button>
  </div>`;
  $$('[data-nav]',nav).forEach(button=>button.onclick=()=>navigateSection(button.dataset.nav,navigate));
  $('#navGuide').onclick=()=>document.dispatchEvent(new CustomEvent('ssm-audit:guide'));
  $('#navExport').onclick=()=>{if(S.screen==='compare')exportSsmComparisonXlsx();else exportSsmAuditXlsx();};
  $('#navCollapse').onclick=()=>{S.ui.navCollapsed=!S.ui.navCollapsed;renderSideNav(navigate);$('#navCollapse')?.focus();};
}
function navigateSection(id,navigate){
  if(id==='dashboard'){S.homeMode='audit';navigate(S.session&&S.session.result?'dashboard':'upload');return;}
  if(id==='rules'){S.homeMode='rules';navigate('rules');return;}
  if(id==='compare'){S.homeMode='compare';navigate(S.comparison&&S.comparison.result?'compare':'upload');return;}
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
function ruleFindingCounts(){
  const result=S.session&&S.session.result;if(!result)return null;
  const counts=new Map();for(const finding of result.findings)counts.set(finding.rule.id,(counts.get(finding.rule.id)||0)+1);return counts;
}
function ruleCatalogRows(){
  const query=clean(S.rules.search).toUpperCase(),source=S.rules.source,category=S.rules.category;
  return ruleCatalog().filter(rule=>(source==='all'||rule.source===source)&&(category==='all'||rule.category===category)&&(!query||[rule.title,rule.statement,SOURCE_LABELS[rule.source],RULE_CATEGORY_LABELS[rule.category]].join(' ').toUpperCase().includes(query)));
}
function ruleRowHtml(rule,counts,query){
  const confidence=RULE_CONFIDENCE_LABELS[rule.confidence]||'Guidance',found=counts?counts.get(rule.id)||0:null;
  const findings=found===null?'':found?`<button class="rule-finding-count" type="button" data-rule-findings="${esc(rule.id)}" title="Show only these findings">${found.toLocaleString()} found${ic('arrow-right')}</button>`:`<span class="rule-finding-count clear">${ic('check')}None found</span>`;
  return `<article class="rule-reference-row ${rule.enabled?'':'is-off'}">
    <span class="rule-reference-icon">${ic(rule.category==='dependencies'?'git-branch':rule.category==='metadata'?'database':rule.category==='headers'?'folder-tree':rule.category==='milestones'?'clipboard-list':rule.category==='item-masters'?'tag':'list-tree')}</span>
    <div><h4>${highlightHtml(rule.title,query)}</h4><p>${highlightHtml(rule.statement,query)}</p></div>
    <div class="rule-reference-tags"><span class="confidence-${esc(rule.confidence)}">${esc(confidence)}</span><span class="rule-state ${rule.enabled?'on':'off'}">${rule.enabled?'On':'Off'}</span>${findings}</div>
  </article>`;
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
}
function showOnlyRule(ruleId,navigate){
  const result=S.session&&S.session.result;if(!result)return;
  const others=[...new Set(result.findings.map(finding=>finding.rule.id))].filter(id=>id!==ruleId);
  S.session.hiddenRules=others;S.session.hiddenSources=[];S.session.hiddenSeverities=[];S.session.hiddenCategories=[];S.session.search='';S.session.dashFilter=null;S.session.scrollTop=0;S.session.cursor=-1;
  invalidateFindingCaches();S.homeMode='audit';navigate('audit');
}
export function renderRules(navigate){
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='rules';S.homeMode='rules';
  const rules=ruleCatalog(),sources=new Set(rules.map(rule=>rule.source)),topics=new Set(rules.map(rule=>rule.category)),hasResult=!!(S.session&&S.session.result);
  $('#view').innerHTML=`<section class="rules-shell">
    <div class="screen-heading"><div><span class="eyebrow">Plain-language reference</span><h2>What SSM Audit checks</h2><p>Every check the audit runs, grouped by where it comes from. ${hasResult?'Counts show how many times each check fired on the registry you loaded.':'Load a registry to see how many times each check fires.'}</p></div></div>
    <div class="rules-overview"><div><span>Checks in use</span><b>${rules.length}</b><p>Applied to every equipment row.</p></div><div><span>Sources</span><b>${sources.size}</b><p>Registry integrity, the SSM SOP, and commissioning logic.</p></div><div><span>Topics</span><b>${topics.size}</b><p>Hierarchy, dependencies, consistency, milestones, and headers.</p></div></div>
    <div class="rules-toolbar"><div class="searchbox">${ic('search')}<input id="ruleSearch" aria-label="Search audit rules" placeholder="Search checks and explanations" value="${esc(S.rules.search)}"></div><select id="ruleSource" aria-label="Filter by rule source"><option value="all">All sources</option>${SSM_AUDIT_SOURCES.map(source=>`<option value="${esc(source.id)}" ${S.rules.source===source.id?'selected':''}>${esc(source.label)}</option>`).join('')}</select><select id="ruleCategory" aria-label="Filter by topic"><option value="all">All topics</option>${Object.entries(RULE_CATEGORY_LABELS).filter(([key])=>rules.some(rule=>rule.category===key)).map(([key,label])=>`<option value="${esc(key)}" ${S.rules.category===key?'selected':''}>${esc(label)}</option>`).join('')}</select><span id="ruleResultCount"></span></div>
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
      const workbook=XLSX.read(bytes,{type:'array',dense:true});report(.08,'Registry opened');await checkpoint();
      const snapshot=await auditSnapshotFromWorkbook(workbook,file.name,checkpoint,(fraction,label)=>report(.08+fraction*.42,label));
      report(.55,`${snapshot.rows.length.toLocaleString()} rows parsed`);await checkpoint();
      const result=runSsmAudit(snapshot);report(1,`${result.findings.length.toLocaleString()} findings`);
      S.session={...S.session,snapshot,result,error:'',auditedAt:Date.now()};
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
    if(side==='target'){resetSession();S.session={...S.session,name:file.name,snapshot,result:auditResult,error:'',auditedAt:Date.now()};S.comparison.targetName=file.name;S.comparison.targetSnapshot=snapshot;S.comparison.targetError='';}
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

function invalidateFindingCaches(){S.session.filteredCacheKey='';S.session.filteredCacheRows=null;S.session.displayCacheKey='';S.session.displayCacheRows=null;}
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
function filteredFindings(){
  const query=clean(S.session.search).toUpperCase(),hiddenSources=new Set(S.session.hiddenSources||[]),hiddenSeverities=new Set(S.session.hiddenSeverities||[]),hiddenCategories=new Set(S.session.hiddenCategories||[]),hiddenRules=new Set(S.session.hiddenRules||[]),slice=S.session.dashFilter;
  const key=[...hiddenSources,...hiddenSeverities,...hiddenCategories,...hiddenRules,query,S.session.sort,slice?`${slice.kind}|${slice.value}`:''].join('');
  if(S.session.filteredCacheKey===key&&S.session.filteredCacheRows)return S.session.filteredCacheRows;
  let rows=S.session.result.findings.filter(finding=>!hiddenSources.has(finding.rule.source)&&!hiddenSeverities.has(finding.severity)&&!hiddenCategories.has(finding.category)&&!hiddenRules.has(finding.rule.id)&&(!query||finding.searchKey.includes(query))&&(!slice||dashSliceMatches(finding,slice)));
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
  return `<div class="sev-strip"><button class="sev-chip all ${active.length===SSM_AUDIT_SEVERITIES.length?'on':''}" type="button" data-audit-severity="all">All findings<b>${summary.findings.toLocaleString()}</b></button>${SSM_AUDIT_SEVERITIES.map(chip).join('')}${dashSliceChipHtml()}<span class="sev-meta">${summary.rows.toLocaleString()} rows checked &middot; ${summary.checks.toLocaleString()} checks run</span></div>`;
}
function filterCount(){return ['hiddenSources','hiddenSeverities','hiddenCategories','hiddenRules'].reduce((total,key)=>total+(S.session[key]||[]).length,0);}
function filterCheck(key,value,label,count){const checked=!(S.session[key]||[]).includes(value);return `<label class="audit-filter-option"><input type="checkbox" data-audit-filter-key="${key}" value="${esc(value)}" ${checked?'checked':''}><span>${esc(label)}</span><b>${Number(count||0).toLocaleString()}</b></label>`;}
function filterMenu(result){
  const counts=new Map();for(const finding of result.findings)counts.set(finding.rule.id,(counts.get(finding.rule.id)||0)+1);
  const rules=[...new Map(result.findings.map(finding=>[finding.rule.id,finding.rule])).values()].sort((a,b)=>a.title.localeCompare(b.title)||a.id.localeCompare(b.id));
  const categories=SSM_AUDIT_CATEGORIES.filter(category=>(result.summary.category[category]||0)>0);
  return `<div class="audit-filter-menu ${S.session.filterOpen?'open':''}" id="auditFilterMenu" role="dialog" aria-label="Filter findings" ${S.session.filterOpen?'':'hidden'}>
    <div class="audit-filter-head"><div><b>Show findings</b><span>Ticked items stay visible</span></div><button class="btn ghost sm" id="auditResetFilters" type="button">${ic('rotate-ccw')}Clear all</button></div>
    <div class="audit-filter-scroll"><fieldset><legend>Where the check comes from</legend>${SSM_AUDIT_SOURCES.map(source=>filterCheck('hiddenSources',source.id,source.label,result.summary.source[source.id])).join('')}</fieldset>
    ${categories.length?`<fieldset><legend>Topic</legend>${categories.map(category=>filterCheck('hiddenCategories',category,CATEGORY_LABELS[category]||category,result.summary.category[category])).join('')}</fieldset>`:''}
    <fieldset><legend>Individual checks</legend>${rules.map(rule=>filterCheck('hiddenRules',rule.id,rule.title,counts.get(rule.id))).join('')}</fieldset></div>
  </div>`;
}
function sortOptions(){const options=[['severity-desc','Most serious first'],['severity-asc','Least serious first'],['equipment-asc','Equipment A to Z'],['equipment-desc','Equipment Z to A'],['rule-asc','Check A to Z'],['rule-desc','Check Z to A'],['row-asc','Row, low to high'],['row-desc','Row, high to low']];return options.map(([value,label])=>`<option value="${value}" ${S.session.sort===value?'selected':''}>${label}</option>`).join('');}
function groupOptions(){const options=[['none','No grouping'],['rule','Group by check'],['milestone','Group by L2 milestone']];return options.map(([value,label])=>`<option value="${value}" ${S.session.groupBy===value?'selected':''}>${label}</option>`).join('');}
function statusMarkup(summary){const blockers=summary.severity.blocker||0,label=summary.status==='blocked'?`${SEVERITY_LABELS.blocker} &middot; ${blockers.toLocaleString()} ${blockers===1?'row':'rows'}`:summary.status==='review'?'Review required':'Ready';return `<span class="audit-status ${summary.status}">${ic(summary.status==='ready'?'check-check':'triangle-alert')}${label}</span>`;}

export function renderAuditResult(navigate){
  const result=S.session&&S.session.result;if(!result){navigate('upload');return;}
  S.screen='audit';const summary=result.summary,fullscreen=!!S.session.fullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  $('#view').innerHTML=`<section id="auditShell" class="audit-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><div class="audit-title"><span class="audit-title-icon">${ic('check-check')}</span><div><h2>Audit findings</h2><p>${esc(S.session.name)}</p></div></div><span class="spacer"></span>${statusMarkup(summary)}<button class="btn" id="auditDashboard">${ic('layout-dashboard')}Dashboard</button><button class="btn" id="auditBack">${ic('upload')}New registry</button><button class="btn" id="exportAudit">${ic('file-down')}Export report</button><button class="btn icon-btn" id="auditFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
    ${severityStrip(summary)}
    <div class="audit-toolbar"><div class="searchbox">${ic('search')}<input id="auditSearch" aria-label="Search findings" placeholder="Search tags, checks, and explanations" value="${esc(S.session.search)}"></div><div class="audit-filter-wrap"><button class="btn audit-filter-button ${filterCount()?'active':''}" id="auditFilters" type="button" aria-expanded="${S.session.filterOpen?'true':'false'}" aria-haspopup="dialog">${ic('filter')}Filters <b id="auditFilterCount">${filterCount()||''}</b></button>${filterMenu(result)}</div><select id="auditGroup" aria-label="Group findings">${groupOptions()}</select><select id="auditSort" aria-label="Sort findings">${sortOptions()}</select><span class="audit-count" id="auditCount"></span></div>
    <div class="audit-table-wrap" id="auditTableWrap" tabindex="0" role="group" aria-label="Findings list. Use the arrow keys to move and Enter to open."><table class="audit-table"><thead><tr><th>Severity</th><th>What was found</th><th>Equipment ID</th><th>Source</th><th>Sheet &middot; row</th><th aria-label="Open details"></th></tr></thead><tbody id="auditRows"></tbody></table></div>
  </section>`;
  renderSideNav(navigate);wireAuditResult(navigate);
}

function auditGroupRowHtml(item,index){
  const group=item.group;
  return `<tr class="audit-group-row" data-row-index="${index}" data-audit-group="${esc(group.key)}" tabindex="-1"><td colspan="6"><button class="audit-group-toggle ${item.collapsed?'':'open'}" type="button" data-audit-group-toggle="${esc(group.key)}" aria-expanded="${item.collapsed?'false':'true'}">${ic('chevron-right')}<b>${esc(group.label)}</b><small>${esc(group.note)}</small><span class="audit-group-count ${esc(group.worst||'')}">${group.items.length.toLocaleString()}</span></button></td></tr>`;
}
function auditRowHtml(item,index,query){
  const finding=item.finding;
  return `<tr data-row-index="${index}" data-audit-finding="${esc(finding.id)}" tabindex="-1" aria-label="Open finding for ${esc(finding.equipmentId||'the registry')}"><td><span class="audit-severity ${finding.severity}">${SEVERITY_LABELS[finding.severity]}</span></td><td><b title="${esc(finding.why)}">${highlightHtml(finding.why,query)}</b><span>${esc(finding.rule.title)}</span><small class="audit-mobile-evidence">${esc(SOURCE_LABELS[finding.rule.source]||finding.rule.source)} &middot; ${esc(finding.sheet||'Registry')} &middot; row ${finding.row||'—'}</small></td><td>${copyTagHtml(finding.equipmentId,highlightHtml(finding.equipmentId,query))}</td><td>${esc(SOURCE_LABELS[finding.rule.source]||finding.rule.source)}</td><td class="audit-evidence-cell">${esc(finding.sheet||'Registry')} &middot; ${finding.row||'&mdash;'}</td><td>${ic('chevron-right')}</td></tr>`;
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
  S.session.hiddenSources=[];S.session.hiddenSeverities=[];S.session.hiddenCategories=[];S.session.hiddenRules=[];S.session.search='';S.session.dashFilter=null;S.session.cursor=-1;
  invalidateFindingCaches();const search=$('#auditSearch');if(search)search.value='';
  $$('[data-audit-filter-key]').forEach(input=>input.checked=true);updateFilterButton();renderSeverityStrip();rerenderRows(true);
}
function renderSeverityStrip(){
  const strip=$('.sev-strip'),result=S.session.result;if(!strip||!result)return;
  const replacement=document.createElement('div');replacement.innerHTML=severityStrip(result.summary);
  strip.replaceWith(replacement.firstElementChild);wireSeverityStrip();
}
function wireSeverityStrip(){
  const clearSlice=$('#dashClearSlice');
  if(clearSlice)clearSlice.onclick=()=>{S.session.dashFilter=null;S.session.cursor=-1;invalidateFindingCaches();renderSeverityStrip();rerenderRows(true);};
  $$('[data-audit-severity]').forEach(button=>button.onclick=()=>{
    const level=button.dataset.auditSeverity;
    if(level==='all')S.session.hiddenSeverities=[];
    else{const hidden=new Set(S.session.hiddenSeverities||[]);if(hidden.has(level))hidden.delete(level);else hidden.add(level);
      if(hidden.size===SSM_AUDIT_SEVERITIES.length)hidden.clear();S.session.hiddenSeverities=[...hidden];}
    S.session.cursor=-1;invalidateFindingCaches();updateFilterButton();
    $$('[data-audit-filter-key]').forEach(input=>{if(input.dataset.auditFilterKey==='hiddenSeverities')input.checked=!(S.session.hiddenSeverities||[]).includes(input.value);});
    renderSeverityStrip();rerenderRows(true);
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
  const finding=S.session.result.findings.find(item=>item.id===id);if(!finding)return;
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
    ${finding.equipmentId?`<button class="btn" type="button" id="findingInHierarchy">${ic('list-tree')}Show in hierarchy</button>`:''}
    <div class="finding-steps"><button class="btn ghost sm" type="button" id="findingPrev" ${position>0?'':'disabled'}>${ic('chevron-left')}Previous</button><span>${position>=0?`${(position+1).toLocaleString()} of ${list.length.toLocaleString()}`:''}</span><button class="btn ghost sm" type="button" id="findingNext" ${position>=0&&position<list.length-1?'':'disabled'}>Next${ic('chevron-right')}</button></div>
  </div>`;
  wireCopyTags($('#drawerBody'));
  const previous=$('#findingPrev'),next=$('#findingNext'),inTree=$('#findingInHierarchy');
  if(previous)previous.onclick=()=>{const target=list[position-1];if(target)openFinding(target.id,opener);};
  if(next)next.onclick=()=>{const target=list[position+1];if(target)openFinding(target.id,opener);};
  if(inTree)inTree.onclick=()=>{closeDrawer();focusHierarchyOnEquipment(finding.equipmentId);};
  const backdrop=$('#drawerBack');backdrop.classList.add('show');backdrop.setAttribute('aria-hidden','false');
  drawerTrapCleanup?.();drawerTrapCleanup=activateFocusTrap(backdrop,closeDrawer);$('#drawer').focus();
}
export function closeDrawer(){
  drawerTrapCleanup?.();drawerTrapCleanup=null;$('#drawerBack').classList.remove('show');$('#drawerBack').setAttribute('aria-hidden','true');const opener=S.session&&S.session.opener;if(opener&&document.contains(opener))opener.focus();
}
function rerenderRows(reset){const wrap=$('#auditTableWrap');if(reset&&wrap)wrap.scrollTop=0;renderRows();}
function updateFilterButton(){const count=filterCount(),button=$('#auditFilters'),badge=$('#auditFilterCount');if(button)button.classList.toggle('active',!!count);if(badge)badge.textContent=count||'';}
function setFilterOpen(open){S.session.filterOpen=!!open;const menu=$('#auditFilterMenu'),button=$('#auditFilters');if(menu){menu.hidden=!open;menu.classList.toggle('open',!!open);}if(button)button.setAttribute('aria-expanded',open?'true':'false');}
function teardownAuditFilters(){if(auditOutsideHandler)document.removeEventListener('pointerdown',auditOutsideHandler);if(auditEscapeHandler)document.removeEventListener('keydown',auditEscapeHandler);auditOutsideHandler=null;auditEscapeHandler=null;}
function wireAuditResult(navigate){
  teardownAuditFilters();
  $('#auditBack').onclick=()=>{S.homeMode='audit';navigate('upload');};$('#exportAudit').onclick=exportSsmAuditXlsx;
  $('#auditDashboard').onclick=()=>navigate('dashboard');
  $('#auditFullscreen').onclick=()=>{S.session.fullscreen=!S.session.fullscreen;renderAuditResult(navigate);};
  $('#auditSearch').oninput=event=>{S.session.search=event.target.value;S.session.cursor=-1;debounceSearch(()=>rerenderRows(true));};
  $('#auditSort').onchange=event=>{S.session.sort=event.target.value;S.session.cursor=-1;rerenderRows(true);};
  $('#auditGroup').onchange=event=>{S.session.groupBy=event.target.value;S.session.collapsedGroups=[];S.session.cursor=-1;S.session.displayCacheKey='';rerenderRows(true);};
  $('#auditFilters').onclick=event=>{event.stopPropagation();setFilterOpen(!S.session.filterOpen);};
  $('#auditResetFilters').onclick=()=>resetAuditFilters();
  $$('[data-audit-filter-key]').forEach(input=>input.onchange=()=>{const key=input.dataset.auditFilterKey,values=new Set(S.session[key]||[]);if(input.checked)values.delete(input.value);else values.add(input.value);S.session[key]=[...values];S.session.cursor=-1;invalidateFindingCaches();updateFilterButton();renderSeverityStrip();rerenderRows(true);});
  wireSeverityStrip();
  auditOutsideHandler=event=>{if(S.session.filterOpen&&!event.target.closest('.audit-filter-wrap'))setFilterOpen(false);};document.addEventListener('pointerdown',auditOutsideHandler);
  auditEscapeHandler=event=>{if(event.key==='Escape'&&S.session.filterOpen){setFilterOpen(false);$('#auditFilters')?.focus();}};document.addEventListener('keydown',auditEscapeHandler);
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
   up" in one screen, and every number on it is a way into the findings list. */
const DASH_RANK_LIMIT=8,DASH_RULE_LIMIT=10,DASH_NO_MILESTONE='No L2 milestone';
const DASH_SEVERITY_MEANINGS={blocker:'Contradicts the registry or the approved lists',error:'Breaks an SSM SOP rule',warning:'A strong pattern says look',info:'Worth knowing'};
const DASH_SLICE_LABELS={discipline:'Discipline',upn:'System',milestone:'L2 milestone'};
const DASH_SLICE_FIELDS={discipline:'discipline',upn:'upn',milestone:'milestone'};

/* A dashboard row click narrows the findings list by a registry column the
   findings themselves do not carry, so the match runs through the registry row. */
function dashSliceMatches(finding,slice){
  const field=DASH_SLICE_FIELDS[slice&&slice.kind];if(!field)return true;
  const row=registryRowFor(finding);if(!row)return false;
  return auditNormId(row[field])===auditNormId(slice.value);
}
function dashSliceChipHtml(){
  const slice=S.session&&S.session.dashFilter;if(!slice)return '';
  const label=clean(slice.label)||clean(slice.value)||DASH_NO_MILESTONE;
  return `<span class="dash-slice-chip">${ic('filter')}<b>${esc(DASH_SLICE_LABELS[slice.kind]||'Filter')}</b><span>${esc(label)}</span><button class="dash-slice-clear" type="button" id="dashClearSlice" title="Clear this filter" aria-label="Clear the ${esc(label)} filter">${ic('x')}</button></span>`;
}

/* A UPN is labelled with the System Name most of its rows agree on — one row with
   a mistyped System Name should not become the name of the whole system. */
function dashRankBuckets(kind){
  const field=DASH_SLICE_FIELDS[kind],buckets=new Map();
  for(const finding of S.session.result.findings){
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
    bucket.label=name&&auditNormId(name).startsWith(auditNormId(bucket.value)+' ')?name:clean(`${bucket.value}${name?' \u00b7 '+name:''}`);
  }
  return list.sort((a,b)=>b.count-a.count||natCmp(a.label,b.label));
}
function dashRuleRanking(){
  const entries=new Map();
  for(const finding of S.session.result.findings){
    let entry=entries.get(finding.rule.id);
    if(!entry){entry={rule:finding.rule,count:0,worst:'',rows:new Set()};entries.set(finding.rule.id,entry);}
    entry.count++;if(severityRank(finding.severity)>severityRank(entry.worst))entry.worst=finding.severity;
    entry.rows.add(auditNormId(finding.equipmentId)||`${auditNormId(finding.sheet)}|${finding.row||0}`);
  }
  return [...entries.values()].sort((a,b)=>b.count-a.count||severityRank(b.worst)-severityRank(a.worst)||natCmp(a.rule.title,b.rule.title));
}
function dashStructureStats(){
  const result=S.session.result,rows=result.rows||[],upns=new Set(),disciplines=new Set(),milestones=new Set();
  let roots=0,withDependencies=0,fullyPhased=0;
  for(const row of rows){
    const upn=auditNormId(row.upn),discipline=auditNormId(row.discipline),milestone=auditNormId(row.milestone),parent=auditNormId(row.closestParent);
    if(upn)upns.add(upn);if(discipline)disciplines.add(discipline);if(milestone)milestones.add(milestone);
    if(parent&&parent===auditNormId(row.systemName))roots++;
    if(auditSplitReferences(row.dependencies).length)withDependencies++;
    if(milestone&&auditNormId(row.milestoneParent))fullyPhased++;
  }
  return {rows:rows.length,roots,headers:(result.headerIds||[]).length,withDependencies,upns:upns.size,disciplines:disciplines.size,milestones:milestones.size,
    phased:rows.length?Math.round(fullyPhased/rows.length*100):0};
}
function dashStatTiles(stats){
  return [
    {label:'Equipment rows',value:stats.rows.toLocaleString(),note:'Rows the audit read'},
    {label:'Top of a system',value:stats.roots.toLocaleString(),note:'Parent is their own System Name'},
    {label:'Organizational headers',value:stats.headers.toLocaleString(),note:'Rows other equipment nests under',action:stats.headers?'headers':'',title:'Browse these in the SSM hierarchy'},
    {label:'Rows with dependencies',value:stats.withDependencies.toLocaleString(),note:'At least one dependency listed'},
    {label:'Distinct UPNs',value:stats.upns.toLocaleString(),note:'Systems in this registry'},
    {label:'Disciplines',value:stats.disciplines.toLocaleString(),note:'Discipline values in use'},
    {label:'L2 milestones in use',value:stats.milestones.toLocaleString(),note:'Distinct L2 phases',action:stats.milestones?'milestones':'',title:'Group the findings by L2 milestone'},
    {label:'Rows with L1 and L2',value:`${stats.phased}%`,note:'Both milestone levels filled in'},
  ];
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
  return `<button class="dash-rank-row" type="button" data-dash-rank="${esc(kind)}" data-dash-value="${esc(bucket.value)}" data-dash-label="${esc(bucket.label)}" title="Show the ${bucket.count.toLocaleString()} findings on ${esc(bucket.label)}">
    <span class="dash-rank-name">${esc(bucket.label)}</span><b class="dash-rank-count">${bucket.count.toLocaleString()}</b>
    <span class="dash-rank-track"><i class="${esc(bucket.worst||'info')}" style="width:${width}%"></i></span></button>`;
}
function dashRankCardHtml(kind,title,note){
  const buckets=dashRankBuckets(kind),visible=buckets.slice(0,DASH_RANK_LIMIT),max=visible.length?visible[0].count:0,rest=buckets.length-visible.length;
  const body=visible.length?visible.map(bucket=>dashRankRowHtml(bucket,kind,max)).join(''):'<p class="dash-empty">Nothing flagged here.</p>';
  return `<article class="dash-card dash-rank-card"><header><h4>${esc(title)}</h4><span>${esc(note)}</span></header><div class="dash-rank-list">${body}</div>${rest>0?`<p class="dash-rank-more">+${rest.toLocaleString()} more</p>`:''}</article>`;
}
function dashCheckTableHtml(){
  const ranking=dashRuleRanking(),visible=ranking.slice(0,DASH_RULE_LIMIT),total=Math.max(1,S.session.result.summary.rows);
  if(!visible.length)return '<p class="dash-empty">No check fired on this registry.</p>';
  return `<table class="dash-check-table"><thead><tr><th>Level</th><th>Check</th><th>Findings</th><th>Rows touched</th></tr></thead><tbody>${visible.map(entry=>{
    const share=entry.rows.size/total*100,text=share>=1?`${Math.round(share)}%`:'&lt;1%';
    return `<tr data-dash-rule="${esc(entry.rule.id)}" tabindex="0" title="Show only findings from this check"><td><span class="audit-severity ${esc(entry.worst)}">${esc(SEVERITY_LABELS[entry.worst]||entry.worst)}</span></td><td class="dash-check-name"><b>${esc(entry.rule.title)}</b><small>${esc(SOURCE_LABELS[entry.rule.source]||entry.rule.source)}</small></td><td class="dash-check-count">${entry.count.toLocaleString()}</td><td><span class="dash-share"><span class="dash-share-track"><i style="width:${Math.max(3,Math.min(100,Math.round(share)))}%"></i></span><small>${text} of rows</small></span></td></tr>`;
  }).join('')}</tbody></table>`;
}
function dashStatHtml(stat){
  const body=`<span>${esc(stat.label)}</span><b>${esc(stat.value)}</b><small>${esc(stat.note)}</small>`;
  if(!stat.action)return `<div class="dash-stat">${body}</div>`;
  return `<button class="dash-stat is-action" type="button" data-dash-stat="${esc(stat.action)}" title="${esc(stat.title||'')}">${body}<span class="dash-stat-go">${ic('arrow-right')}</span></button>`;
}

export function renderDashboard(navigate){
  const result=S.session&&S.session.result;if(!result){S.homeMode='audit';navigate('upload');return;}
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='dashboard';S.homeMode='audit';
  const summary=result.summary,audited=S.session.auditedAt?new Date(S.session.auditedAt).toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'}):'this session';
  $('#view').innerHTML=`<section class="dash-shell">
    <header class="dash-head">
      <div class="dash-head-copy"><span class="eyebrow">Registry overview</span><h2>${esc(S.session.name||'This registry')}</h2><p>Audited ${summary.rows.toLocaleString()} rows &middot; ${summary.checks.toLocaleString()} checks &middot; ${esc(audited)}</p></div>
      <div class="dash-head-actions">${statusMarkup(summary)}<button class="btn primary" type="button" id="dashOpenFindings">${ic('check-check')}Open findings</button><button class="btn" type="button" id="dashExport">${ic('file-down')}Export report</button></div>
    </header>
    <div class="dash-tiles">${SSM_AUDIT_SEVERITIES.map(level=>dashSeverityTileHtml(level,summary.severity[level]||0)).join('')}</div>
    <section class="dash-block">
      <div class="dash-block-head"><div><h3>Where the problems are</h3><p>The same findings counted three ways. Pick a row to open the list narrowed to it.</p></div></div>
      <div class="dash-rank-grid">${dashRankCardHtml('discipline','By discipline','Findings per discipline')}${dashRankCardHtml('upn','By UPN and system','Findings per system')}${dashRankCardHtml('milestone','By L2 milestone','Findings per L2 phase')}</div>
    </section>
    <section class="dash-block">
      <div class="dash-block-head"><div><h3>Top checks firing</h3><p>The checks producing the most findings. Open one to see only its findings.</p></div><button class="btn-link" type="button" id="dashSeeRules">See all in Rules${ic('arrow-right')}</button></div>
      <div class="dash-card dash-check-card">${dashCheckTableHtml()}</div>
    </section>
    <section class="dash-block">
      <div class="dash-block-head"><div><h3>Structure at a glance</h3><p>What this registry is made of, before any rule is applied.</p></div></div>
      <div class="dash-stats">${dashStatTiles(dashStructureStats()).map(dashStatHtml).join('')}</div>
    </section>
    <p class="dash-foot">${ic('lock')}Findings never leave this browser.</p>
  </section>`;
  renderSideNav(navigate);wireDashboard(navigate);
}
function dashOpenFindings(navigate,apply){
  S.session.hiddenSources=[];S.session.hiddenSeverities=[];S.session.hiddenCategories=[];S.session.hiddenRules=[];S.session.search='';S.session.dashFilter=null;
  S.session.groupBy='none';S.session.collapsedGroups=[];S.session.cursor=-1;S.session.scrollTop=0;
  if(apply)apply();
  invalidateFindingCaches();S.homeMode='audit';navigate('audit');
}
function wireDashboard(navigate){
  $('#dashOpenFindings').onclick=()=>dashOpenFindings(navigate);
  $('#dashExport').onclick=exportSsmAuditXlsx;
  $('#dashSeeRules').onclick=()=>{S.homeMode='rules';navigate('rules');};
  $$('[data-dash-severity]').forEach(tile=>tile.onclick=()=>{const level=tile.dataset.dashSeverity;dashOpenFindings(navigate,()=>{S.session.hiddenSeverities=SSM_AUDIT_SEVERITIES.filter(item=>item!==level);});});
  $$('[data-dash-rank]').forEach(row=>row.onclick=()=>dashOpenFindings(navigate,()=>{S.session.dashFilter={kind:row.dataset.dashRank,value:row.dataset.dashValue,label:row.dataset.dashLabel};}));
  $$('[data-dash-rule]').forEach(row=>{
    const open=()=>showOnlyRule(row.dataset.dashRule,navigate);
    row.onclick=open;row.onkeydown=event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();open();}};
  });
  $$('[data-dash-stat]').forEach(tile=>tile.onclick=()=>{
    const action=tile.dataset.dashStat;
    if(action==='headers'){if(S.session.snapshot)navigate('hierarchy');return;}
    if(action==='milestones')dashOpenFindings(navigate,()=>{S.session.groupBy='milestone';});
  });
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
  wireCopyTags($('#drawerBody'));const backdrop=$('#drawerBack');backdrop.classList.add('show');backdrop.setAttribute('aria-hidden','false');drawerTrapCleanup?.();drawerTrapCleanup=activateFocusTrap(backdrop,closeDrawer);$('#drawer').focus();
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
function applyHierarchyFocus(){
  const key=S.session.hierarchyFocusKey;if(!key)return;
  const rows=hierarchyRows(),index=rows.findIndex(item=>item.node.key===key),wrap=$('#hierarchyTreeWrap');
  if(index>=0&&wrap){wrap.scrollTop=Math.max(0,index*HIERARCHY_ROW_HEIGHT-wrap.clientHeight/2+HIERARCHY_ROW_HEIGHT);S.session.hierarchyScrollTop=wrap.scrollTop;}
  renderHierarchyRows();$(`[data-hierarchy-node="${key}"]`)?.focus({preventScroll:true});
  setTimeout(()=>{S.session.hierarchyFocusKey='';$('.hierarchy-row.focused')?.classList.remove('focused');},2600);
}
function hierarchyOptions(items,selected,label){return `<option value="all">${esc(label)}</option>`+items.map(item=>`<option value="${esc(item.value)}" ${selected===item.value?'selected':''}>${esc(item.label)}</option>`).join('');}
function uniqueHierarchyOptions(items,valueFor,labelFor){const options=new Map();for(const item of items){const value=valueFor(item);if(value&&!options.has(value))options.set(value,{value,label:labelFor(item)});}return [...options.values()].sort((left,right)=>left.label.localeCompare(right.label,undefined,{numeric:true}));}
function renderHierarchyResult(navigate){
  const result=S.session&&S.session.result;if(!result||!S.session.snapshot){navigate('upload');return;}
  teardownAuditFilters();S.screen='hierarchy';const hierarchy=sessionHierarchy(),summary=hierarchy.summary,fullscreen=!!S.session.hierarchyFullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  const buildingOptions=hierarchy.groups.buildings.map(node=>({value:node.key,label:node.label})),disciplineNodes=hierarchy.groups.disciplines.filter(node=>S.session.hierarchyBuilding==='all'||node.buildingKey===S.session.hierarchyBuilding),disciplineOptions=uniqueHierarchyOptions(disciplineNodes,hierarchyDisciplineValue,node=>node.label),systemNodes=hierarchy.groups.systems.filter(node=>(S.session.hierarchyBuilding==='all'||node.buildingKey===S.session.hierarchyBuilding)&&(S.session.hierarchyDiscipline==='all'||hierarchyDisciplineValue(hierarchy.nodeByKey.get(node.parentKey))===S.session.hierarchyDiscipline)),systemOptions=uniqueHierarchyOptions(systemNodes,hierarchySystemValue,node=>`${node.upn} · ${node.label}`);
  $('#view').innerHTML=`<section id="hierarchyShell" class="hierarchy-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><div class="audit-title"><span class="audit-title-icon">${ic('list-tree')}</span><div><h2>SSM hierarchy</h2><p>${esc(S.session.name)}</p></div></div><span class="spacer"></span><span class="audit-status ready">${ic('list-tree')}${summary.equipment.toLocaleString()} equipment</span><span class="audit-status ${summary.findings?'review':'ready'}">${ic(summary.findings?'triangle-alert':'check-check')}${summary.findings.toLocaleString()} findings</span><button class="btn icon-btn" id="hierarchyFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
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
    <div class="audit-head"><div class="audit-title"><span class="audit-title-icon">${ic('square-stack')}</span><div><h2>Compare projects</h2><p>${esc(S.comparison.targetName)} &middot; ${esc(S.comparison.referenceName)}</p></div></div><span class="spacer"></span><span class="compare-building-note">${ic('layers')}Building ignored</span><button class="btn" id="compareBack">${ic('upload')}Change files</button><button class="btn" id="exportComparison">${ic('file-down')}Export comparison</button><button class="btn icon-btn" id="compareFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
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
