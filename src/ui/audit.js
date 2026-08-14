import { $, $$, clean, esc } from '../core/text.js'
import { S, resetSession } from '../state.js'
import { readArrayBuffer } from '../io/workbook.js'
import { auditSnapshotFromWorkbook } from '../audit/model.js'
import { runSsmAudit, SSM_AUDIT_CATEGORIES, SSM_AUDIT_SEVERITIES, SSM_AUDIT_SOURCES } from '../audit/engine.js'
import { compareSsmRegistries, comparisonSystemTypes } from '../audit/compare.js'
import { exportSsmAuditXlsx, exportSsmComparisonXlsx } from '../audit/export.js'
import { ic } from './icons.js'
import { copyTagHtml, runWithProgress, toast, wireCopyTags } from './feedback.js'

const AUDIT_ROW_HEIGHT=48,AUDIT_OVERSCAN=18,AUDIT_MAX_ROWS=160;
const COMPARE_ROW_HEIGHT=76,COMPARE_OVERSCAN=14,COMPARE_MAX_ROWS=120;
const COMPARE_TREE_ROW_HEIGHT=52,COMPARE_TREE_OVERSCAN=10,COMPARE_TREE_MAX_ROWS=120;
const CATEGORY_LABELS={structure:'Structure',dependencies:'Dependencies',metadata:'Metadata',milestones:'Milestones','item-masters':'Item Masters',headers:'Headers / Rollups'};
const RULE_CATEGORY_LABELS={structure:'Hierarchy',dependencies:'Dependencies',metadata:'Registry consistency',headers:'Headers / Rollups'};
const RULE_CONFIDENCE_LABELS={required:'Required',strong:'Strong pattern','description-rated':'Description based'};
const RULE_SOURCE_DESCRIPTIONS={registry:'Identity, references, and metadata consistency within the registry.',sop:'Required parent-child, dependency, sequencing, and header checks.',logic:'Confidence-rated control, electrical, and process-enabling relationships.'};
const SEVERITY_LABELS={blocker:'Blocker',error:'Error',warning:'Warning',info:'Advisory'};
const SOURCE_LABELS=Object.fromEntries(SSM_AUDIT_SOURCES.map(source=>[source.id,source.label]));
let auditOutsideHandler=null,auditEscapeHandler=null;

function workspaceTabs(active){return `<div class="workspace-tabs" role="tablist" aria-label="SSM Audit tools"><button class="workspace-tab ${active==='audit'?'active':''}" type="button" data-workspace-mode="audit" role="tab" aria-selected="${active==='audit'}">${ic('check-check')}Audit Registry</button><button class="workspace-tab ${active==='compare'?'active':''}" type="button" data-workspace-mode="compare" role="tab" aria-selected="${active==='compare'}">${ic('square-stack')}Compare Projects</button><button class="workspace-tab ${active==='rules'?'active':''}" type="button" data-workspace-mode="rules" role="tab" aria-selected="${active==='rules'}">${ic('book-open')}Rules</button></div>`;}
function wireWorkspaceTabs(navigate){$$('[data-workspace-mode]').forEach(button=>button.onclick=()=>{S.homeMode=button.dataset.workspaceMode;if(S.homeMode==='rules')navigate('rules');else if(S.homeMode==='compare'&&S.comparison.result)navigate('compare');else if(S.homeMode==='audit'&&S.session.result)navigate('audit');else renderUpload(navigate);});}

function activeRuleCatalog(){return Object.values(SSM_AUDIT_RULES).filter(rule=>rule.enabled);}
function ruleCatalogRows(){
  const query=clean(S.rules.search).toUpperCase(),source=S.rules.source,category=S.rules.category;
  return activeRuleCatalog().filter(rule=>(source==='all'||rule.source===source)&&(category==='all'||rule.category===category)&&(!query||[rule.title,rule.statement,SOURCE_LABELS[rule.source],RULE_CATEGORY_LABELS[rule.category]].join(' ').toUpperCase().includes(query)));
}
function ruleRow(rule){
  const topic=RULE_CATEGORY_LABELS[rule.category]||'General',confidence=RULE_CONFIDENCE_LABELS[rule.confidence]||'Guidance';
  return `<article class="rule-reference-row"><span class="rule-reference-icon">${ic(rule.category==='dependencies'?'git-branch':rule.category==='metadata'?'database':rule.category==='headers'?'folder-tree':'list-tree')}</span><div><h4>${esc(rule.title)}</h4><p>${esc(rule.statement)}</p></div><div class="rule-reference-tags"><span>${esc(topic)}</span><span class="confidence-${esc(rule.confidence)}">${esc(confidence)}</span></div></article>`;
}
function renderRuleCatalog(){
  const rows=ruleCatalogRows(),container=$('#ruleCatalog'),count=$('#ruleResultCount');if(!container)return;
  if(count)count.textContent=`${rows.length} of ${activeRuleCatalog().length} checks`;
  if(!rows.length){container.innerHTML=`<div class="rule-reference-empty">${ic('search')}<b>No checks found</b><span>Try a broader search or choose All topics.</span></div>`;return;}
  container.innerHTML=SSM_AUDIT_SOURCES.map(source=>{const matches=rows.filter(rule=>rule.source===source.id);if(!matches.length)return '';return `<section class="rule-source-section"><header><span>${ic(source.id==='sop'?'book-open':source.id==='logic'?'network':'check-check')}</span><div><h3>${esc(source.label)}</h3><p>${esc(RULE_SOURCE_DESCRIPTIONS[source.id]||source.description)}</p></div><b>${matches.length}</b></header><div>${matches.map(ruleRow).join('')}</div></section>`;}).join('');
}
export function renderRules(navigate){
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='rules';S.homeMode='rules';
  const rules=activeRuleCatalog(),sources=new Set(rules.map(rule=>rule.source)),topics=new Set(rules.map(rule=>rule.category));
  $('#view').innerHTML=`<section class="rules-shell">
    ${workspaceTabs('rules')}
    <div class="screen-heading rules-heading"><div><span class="eyebrow">Plain-language audit reference</span><h2>Rules used by SSM Audit</h2><p>See exactly what the app checks and why each relationship matters.</p></div><button class="btn ghost" id="openGuide">${ic('book-open')}Guide</button></div>
    <div class="rules-overview"><div><span>Checks in use</span><b>${rules.length}</b><p>Every check currently applied during an audit.</p></div><div><span>Governing sources</span><b>${sources.size}</b><p>Registry integrity, SSM requirements, and commissioning logic.</p></div><div><span>Review topics</span><b>${topics.size}</b><p>Hierarchy, dependencies, consistency, and headers.</p></div></div>
    <div class="rules-note">${ic('info')}<span><b>Transparent by design.</b> Project comparisons provide context, but never create or change these rules.</span></div>
    <div class="rules-toolbar"><div class="searchbox">${ic('search')}<input id="ruleSearch" aria-label="Search audit rules" placeholder="Search checks and explanations" value="${esc(S.rules.search)}"></div><select id="ruleSource" aria-label="Filter by rule source"><option value="all">All sources</option>${SSM_AUDIT_SOURCES.map(source=>`<option value="${esc(source.id)}" ${S.rules.source===source.id?'selected':''}>${esc(source.label)}</option>`).join('')}</select><select id="ruleCategory" aria-label="Filter by topic"><option value="all">All topics</option>${Object.entries(RULE_CATEGORY_LABELS).filter(([key])=>rules.some(rule=>rule.category===key)).map(([key,label])=>`<option value="${esc(key)}" ${S.rules.category===key?'selected':''}>${esc(label)}</option>`).join('')}</select><span id="ruleResultCount"></span></div>
    <div class="rule-catalog" id="ruleCatalog"></div>
  </section>`;
  wireWorkspaceTabs(navigate);$('#openGuide').onclick=()=>document.dispatchEvent(new CustomEvent('ssm-audit:guide'));
  $('#ruleSearch').oninput=event=>{S.rules.search=event.target.value;renderRuleCatalog();};
  $('#ruleSource').onchange=event=>{S.rules.source=event.target.value;renderRuleCatalog();};
  $('#ruleCategory').onchange=event=>{S.rules.category=event.target.value;renderRuleCatalog();};
  renderRuleCatalog();
}

function importStatus(){
  const result=S.session.result,summary=result&&result.summary;
  if(result)return `<div class="audit-import ready"><span class="file-icon">${ic(summary.status==='ready'?'check-check':'triangle-alert')}</span><div class="file-meta"><b>${esc(S.session.name)}</b><span>${summary.rows.toLocaleString()} rows audited &middot; ${summary.findings.toLocaleString()} findings &middot; ${esc(summary.status.toUpperCase())}</span></div><button class="btn" id="openAuditResult">Open audit</button><button class="xbtn icon-btn" id="removeAuditTarget" type="button" aria-label="Remove audit target">${ic('x')}</button></div>`;
  if(S.session.error)return `<div class="audit-import error"><span class="file-icon">${ic('triangle-alert')}</span><div class="file-meta"><b>Could not run SSM Audit</b><span>${esc(S.session.error)}</span></div><button class="btn" id="chooseAuditAgain">Choose another</button></div>`;
  return '';
}

export function renderUpload(navigate){
  if(S.homeMode==='rules'){renderRules(navigate);return;}
  if(S.homeMode==='compare'){renderComparisonUpload(navigate);return;}
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='upload';
  $('#view').innerHTML=`<section class="upload-shell">
    ${workspaceTabs('audit')}
    <div class="screen-heading"><div><span class="eyebrow">Registry Integrity + SSM Rules</span><h2>Audit an existing SSM</h2><p>Check a completed Cx Registry and receive an evidence-backed correction report.</p></div><button class="btn ghost" id="openGuide">${ic('book-open')}Guide</button></div>
    <div class="upload-grid">
      <div class="dropzone" id="dropzone" tabindex="0" role="button" aria-label="Choose a Cx Registry workbook">
        <span class="drop-icon">${ic('file-spreadsheet')}</span>
        <h3>Choose a completed Cx Registry</h3>
        <p>Drop an Excel workbook here or browse this device.</p>
        <button class="btn primary" id="auditBrowse">${ic('upload')}Choose workbook</button>
        <span class="file-types">.xlsx or .xls</span>
      </div>
      <aside class="privacy-panel">
        <span class="privacy-icon">${ic('lock')}</span><div><h3>Private by design</h3><p>The workbook is analyzed only in this browser session. It is never uploaded, stored, or used to teach future rules.</p></div>
      </aside>
    </div>
    <div id="importStatus">${importStatus()}</div>
    <input id="auditFile" type="file" accept=".xlsx,.xls" hidden>
    <div class="coverage-band">
      <div><span>01</span><b>Hierarchy</b><p>Parents, roots, cycles, and discipline or UPN crossings.</p></div>
      <div><span>02</span><b>Dependencies</b><p>Unresolved, repeated, self-referencing, and circular predecessors.</p></div>
      <div><span>03</span><b>Commissioning logic</b><p>Control sources, power paths, equipment roles, milestones, and headers.</p></div>
    </div>
  </section>`;
  wireWorkspaceTabs(navigate);wireUpload(navigate);
}

function chooseFile(){const input=$('#auditFile');if(input)input.click();}
function wireUpload(navigate){
  const input=$('#auditFile'),drop=$('#dropzone'),browse=$('#auditBrowse'),again=$('#chooseAuditAgain');
  if(browse)browse.onclick=event=>{event.stopPropagation();chooseFile();};if(again)again.onclick=chooseFile;
  if(input)input.onchange=()=>{const file=input.files[0];input.value='';if(file)addAuditTarget(file,navigate);};
  if(drop){
    drop.onclick=event=>{if(!event.target.closest('button'))chooseFile();};
    drop.onkeydown=event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();chooseFile();}};
    drop.ondragover=event=>{event.preventDefault();drop.classList.add('dragging');};
    drop.ondragleave=()=>drop.classList.remove('dragging');
    drop.ondrop=event=>{event.preventDefault();drop.classList.remove('dragging');const file=event.dataTransfer.files[0];if(file)addAuditTarget(file,navigate);};
  }
  const open=$('#openAuditResult');if(open)open.onclick=()=>navigate('audit');
  const remove=$('#removeAuditTarget');if(remove)remove.onclick=()=>{resetSession();clearComparisonTarget();renderUpload(navigate);};
  $('#openGuide').onclick=()=>document.dispatchEvent(new CustomEvent('ssm-audit:guide'));
}

export async function addAuditTarget(file,navigate){
  if(!/\.(xlsx|xls)$/i.test(file.name)){toast('SSM Audit requires an Excel workbook');return;}
  resetSession();clearComparisonTarget();S.session.name=file.name;
  try{
    await runWithProgress('Running SSM Audit',file.name,async(checkpoint,report)=>{
      const bytes=new Uint8Array(await readArrayBuffer(file));await checkpoint();
      const workbook=XLSX.read(bytes,{type:'array',dense:true});report(.08,'Registry opened');await checkpoint();
      const snapshot=await auditSnapshotFromWorkbook(workbook,file.name,checkpoint,(fraction,label)=>report(.08+fraction*.42,label));
      report(.55,`${snapshot.rows.length.toLocaleString()} rows parsed`);await checkpoint();
      const result=runSsmAudit(snapshot);report(1,`${result.findings.length.toLocaleString()} findings`);
      S.session={...S.session,snapshot,result,error:''};
      S.comparison.targetName=file.name;S.comparison.targetSnapshot=snapshot;S.comparison.targetError='';S.comparison.result=null;
    });
    navigate('audit');
  }catch(error){
    console.error('SSM Audit failed',error);S.session.error=error&&error.message||'Could not read this registry';S.session.snapshot=null;S.session.result=null;renderUpload(navigate);
  }
}

function clearComparisonTarget(){S.comparison.targetName='';S.comparison.targetSnapshot=null;S.comparison.targetError='';S.comparison.result=null;S.comparison.selectedUpn='';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};}
function clearComparisonReference(){S.comparison.referenceName='';S.comparison.referenceSnapshot=null;S.comparison.referenceError='';S.comparison.result=null;S.comparison.selectedUpn='';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};}

function comparisonSlot(side){
  const target=side==='target',name=target?S.comparison.targetName:S.comparison.referenceName,snapshot=target?S.comparison.targetSnapshot:S.comparison.referenceSnapshot,error=target?S.comparison.targetError:S.comparison.referenceError;
  const title=target?'Registry to audit':'Completed project reference',detail=target?'Receives the full SSM Audit and comparison findings.':'Used only for this side-by-side session.';
  if(snapshot)return `<section class="compare-upload-slot ready" data-compare-side="${side}"><div class="compare-slot-label"><span>${target?'01':'02'}</span><div><b>${title}</b><small>${detail}</small></div></div><div class="compare-ready-icon">${ic(target?'check-check':'square-stack')}</div><h3>${esc(name)}</h3><p>${snapshot.rows.length.toLocaleString()} registry rows ready</p><div class="compare-slot-actions"><button class="btn sm" type="button" data-compare-replace="${side}">${ic('upload')}Replace</button><button class="xbtn icon-btn" type="button" data-compare-remove="${side}" aria-label="Remove ${esc(title)}">${ic('x')}</button></div></section>`;
  return `<section class="compare-upload-slot ${error?'error':''}" data-compare-side="${side}" tabindex="0" role="button" aria-label="Choose ${esc(title)}"><div class="compare-slot-label"><span>${target?'01':'02'}</span><div><b>${title}</b><small>${detail}</small></div></div><div class="compare-ready-icon">${ic(error?'triangle-alert':'file-spreadsheet')}</div><h3>${error?'Could not read workbook':'Choose workbook'}</h3><p>${esc(error||'Drop an .xlsx or .xls file here, or browse this device.')}</p><button class="btn ${target?'primary':''} sm" type="button" data-compare-browse="${side}">${ic('upload')}Browse</button></section>`;
}

function renderComparisonUpload(navigate){
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='upload';
  if(!S.comparison.targetSnapshot&&S.session.snapshot){S.comparison.targetName=S.session.name;S.comparison.targetSnapshot=S.session.snapshot;}
  const ready=!!(S.comparison.targetSnapshot&&S.comparison.referenceSnapshot);
  $('#view').innerHTML=`<section class="upload-shell compare-upload-shell">
    ${workspaceTabs('compare')}
    <div class="screen-heading"><div><span class="eyebrow">Building-neutral registry comparison</span><h2>Compare project hierarchies</h2><p>Audit the target registry, align both projects by UPN, and inspect equipment nesting, header use, and I&amp;C placement side by side.</p></div><button class="btn ghost" id="openGuide">${ic('book-open')}Guide</button></div>
    <div class="compare-upload-grid">${comparisonSlot('target')}${comparisonSlot('reference')}</div>
    <div class="compare-runbar"><div class="compare-method">${ic('lock')}<span><b>Local and session-only</b> Building is excluded from matching. The reference never changes audit rules or future results.</span></div><button class="btn primary" id="runComparison" type="button" ${ready?'':'disabled'}>${ic('square-stack')}Compare registries</button></div>
    <input id="compareTargetFile" type="file" accept=".xlsx,.xls" hidden><input id="compareReferenceFile" type="file" accept=".xlsx,.xls" hidden>
    <div class="coverage-band compare-coverage"><div><span>01</span><b>UPN alignment</b><p>Maps corresponding systems without comparing project Building values.</p></div><div><span>02</span><b>Semantic nesting</b><p>Pairs equipment by tag core, description, classification, and parent role.</p></div><div><span>03</span><b>Project differences</b><p>Highlights hierarchy patterns, headers, dependencies, and I&amp;C placement.</p></div></div>
  </section>`;
  wireWorkspaceTabs(navigate);wireComparisonUpload(navigate);
}

function wireComparisonUpload(navigate){
  const choose=side=>$('#compare'+(side==='target'?'Target':'Reference')+'File')?.click();
  for(const side of ['target','reference']){
    const input=$('#compare'+(side==='target'?'Target':'Reference')+'File'),slot=$(`[data-compare-side="${side}"]`),browse=$(`[data-compare-browse="${side}"]`),replace=$(`[data-compare-replace="${side}"]`),remove=$(`[data-compare-remove="${side}"]`);
    if(input)input.onchange=()=>{const file=input.files[0];input.value='';if(file)addComparisonFile(file,side,navigate);};
    if(browse)browse.onclick=event=>{event.stopPropagation();choose(side);};if(replace)replace.onclick=()=>choose(side);
    if(remove)remove.onclick=()=>{if(side==='target'){resetSession();clearComparisonTarget();}else clearComparisonReference();renderUpload(navigate);};
    if(slot&&!slot.classList.contains('ready')){
      slot.onclick=event=>{if(!event.target.closest('button'))choose(side);};slot.onkeydown=event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();choose(side);}};
      slot.ondragover=event=>{event.preventDefault();slot.classList.add('dragging');};slot.ondragleave=()=>slot.classList.remove('dragging');slot.ondrop=event=>{event.preventDefault();slot.classList.remove('dragging');const file=event.dataTransfer.files[0];if(file)addComparisonFile(file,side,navigate);};
    }
  }
  $('#runComparison').onclick=()=>runComparison(navigate);$('#openGuide').onclick=()=>document.dispatchEvent(new CustomEvent('ssm-audit:guide'));
}

async function addComparisonFile(file,side,navigate){
  if(!/\.(xlsx|xls)$/i.test(file.name)){toast('Choose an Excel workbook');return;}
  const errorKey=side==='target'?'targetError':'referenceError';S.comparison[errorKey]='';
  try{
    let snapshot,auditResult;
    await runWithProgress(side==='target'?'Loading registry to audit':'Loading completed project',file.name,async(checkpoint,report)=>{
      const bytes=new Uint8Array(await readArrayBuffer(file));await checkpoint();const workbook=XLSX.read(bytes,{type:'array',dense:true});report(.1,'Registry opened');await checkpoint();
      snapshot=await auditSnapshotFromWorkbook(workbook,file.name,checkpoint,(fraction,label)=>report(.1+fraction*.68,label));
      report(.82,`${snapshot.rows.length.toLocaleString()} rows parsed`);await checkpoint();if(side==='target'){auditResult=runSsmAudit(snapshot);report(1,`${auditResult.findings.length.toLocaleString()} audit findings`);}else report(1,'Reference ready');
    });
    if(side==='target'){resetSession();S.session={...S.session,name:file.name,snapshot,result:auditResult,error:''};S.comparison.targetName=file.name;S.comparison.targetSnapshot=snapshot;S.comparison.targetError='';}
    else{S.comparison.referenceName=file.name;S.comparison.referenceSnapshot=snapshot;S.comparison.referenceError='';}
    S.comparison.result=null;S.comparison.selectedUpn='';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};renderUpload(navigate);
  }catch(error){console.error('Registry comparison import failed',error);S.comparison[errorKey]=error&&error.message||'Could not read this registry';renderUpload(navigate);}
}

async function runComparison(navigate){
  if(!S.comparison.targetSnapshot||!S.comparison.referenceSnapshot){toast('Choose both registries first');return;}
  await runWithProgress('Comparing project hierarchies','Building values are excluded',async(checkpoint,report)=>{report(.15,'Indexing target systems');await checkpoint();report(.38,'Aligning equipment by UPN and meaning');await checkpoint();const result=compareSsmRegistries(S.comparison.targetSnapshot,S.comparison.referenceSnapshot);S.comparison.result=result;const selected=result.systems.find(system=>system.status!=='aligned')||result.systems[0];S.comparison.selectedUpn=selected&&selected.upn||'';S.comparison.pairScrollTop=0;S.comparison.treeScrollTop=0;S.comparison.treeExpandedByUpn={};report(1,`${result.summary.systems.toLocaleString()} systems compared`);});
  navigate('compare');
}

function filteredFindings(){
  const query=clean(S.session.search).toUpperCase(),hiddenSources=new Set(S.session.hiddenSources||[]),hiddenSeverities=new Set(S.session.hiddenSeverities||[]),hiddenCategories=new Set(S.session.hiddenCategories||[]),hiddenRules=new Set(S.session.hiddenRules||[]);
  const key=[...hiddenSources,...hiddenSeverities,...hiddenCategories,...hiddenRules,query,S.session.sort].join('\u0001');
  if(S.session.filteredCacheKey===key&&S.session.filteredCacheRows)return S.session.filteredCacheRows;
  let rows=S.session.result.findings.filter(finding=>!hiddenSources.has(finding.rule.source)&&!hiddenSeverities.has(finding.severity)&&!hiddenCategories.has(finding.category)&&!hiddenRules.has(finding.rule.id)&&(!query||finding.searchKey.includes(query)));
  const natural=(a,b)=>String(a||'').localeCompare(String(b||''),undefined,{numeric:true,sensitivity:'base'}),sort=S.session.sort;
  if(sort==='severity-asc')rows=[...rows].reverse();
  else if(sort==='equipment-asc'||sort==='equipment-desc')rows=[...rows].sort((a,b)=>(sort.endsWith('desc')?-1:1)*(natural(a.equipmentId,b.equipmentId)||a.row-b.row));
  else if(sort==='rule-asc'||sort==='rule-desc')rows=[...rows].sort((a,b)=>(sort.endsWith('desc')?-1:1)*(natural(a.rule.title,b.rule.title)||natural(a.rule.id,b.rule.id)||a.row-b.row));
  else if(sort==='row-asc'||sort==='row-desc')rows=[...rows].sort((a,b)=>(sort.endsWith('desc')?-1:1)*((a.row-b.row)||natural(a.sheet,b.sheet)));
  S.session.filteredCacheKey=key;S.session.filteredCacheRows=rows;return rows;
}
function statusMarkup(summary){const label=summary.status==='blocked'?'Blocked':summary.status==='review'?'Review required':'Ready';return `<span class="audit-status ${summary.status}">${ic(summary.status==='ready'?'check-check':'triangle-alert')}${label}</span>`;}
function categoryTabs(summary){const hidden=new Set(S.session.hiddenCategories||[]);return `<button class="audit-cat ${hidden.size?'':'active'}" data-audit-category="all">All <b>${summary.findings.toLocaleString()}</b></button>`+SSM_AUDIT_CATEGORIES.map(category=>`<button class="audit-cat ${hidden.has(category)?'':'active'}" data-audit-category="${category}" aria-pressed="${hidden.has(category)?'false':'true'}">${CATEGORY_LABELS[category]} <b>${(summary.category[category]||0).toLocaleString()}</b></button>`).join('');}
function filterCount(){return ['hiddenSources','hiddenSeverities','hiddenCategories','hiddenRules'].reduce((total,key)=>total+(S.session[key]||[]).length,0);}
function filterCheck(key,value,label,count){const checked=!(S.session[key]||[]).includes(value);return `<label class="audit-filter-option"><input type="checkbox" data-audit-filter-key="${key}" value="${esc(value)}" ${checked?'checked':''}><span>${esc(label)}</span><b>${Number(count||0).toLocaleString()}</b></label>`;}
function filterMenu(result){
  const counts=new Map();for(const finding of result.findings)counts.set(finding.rule.id,(counts.get(finding.rule.id)||0)+1);
  const rules=[...new Map(result.findings.map(finding=>[finding.rule.id,finding.rule])).values()].sort((a,b)=>a.title.localeCompare(b.title)||a.id.localeCompare(b.id));
  return `<div class="audit-filter-menu ${S.session.filterOpen?'open':''}" id="auditFilterMenu" ${S.session.filterOpen?'':'hidden'}>
    <div class="audit-filter-head"><div><b>Show findings</b><span>Checked items stay visible</span></div><button class="btn ghost sm" id="auditResetFilters" type="button">${ic('rotate-ccw')}Reset</button></div>
    <div class="audit-filter-scroll"><fieldset><legend>Rule source</legend>${SSM_AUDIT_SOURCES.map(source=>filterCheck('hiddenSources',source.id,source.label,result.summary.source[source.id])).join('')}</fieldset>
    <fieldset><legend>Severity</legend>${SSM_AUDIT_SEVERITIES.map(level=>filterCheck('hiddenSeverities',level,SEVERITY_LABELS[level],result.summary.severity[level])).join('')}</fieldset>
    <fieldset><legend>Individual rules</legend>${rules.map(rule=>filterCheck('hiddenRules',rule.id,rule.title,counts.get(rule.id))).join('')}</fieldset></div>
  </div>`;
}
function sortOptions(){const options=[['severity-desc','Severity: high to low'],['severity-asc','Severity: low to high'],['equipment-asc','Equipment: A to Z'],['equipment-desc','Equipment: Z to A'],['rule-asc','Rule: A to Z'],['rule-desc','Rule: Z to A'],['row-asc','Row: low to high'],['row-desc','Row: high to low']];return options.map(([value,label])=>`<option value="${value}" ${S.session.sort===value?'selected':''}>${label}</option>`).join('');}

export function renderAuditResult(navigate){
  const result=S.session&&S.session.result;if(!result){navigate('upload');return;}
  S.screen='audit';const summary=result.summary,fullscreen=!!S.session.fullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  $('#view').innerHTML=`<section id="auditShell" class="audit-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><button class="btn ghost" id="auditBack">${ic('arrow-left')}Files</button><div class="audit-title"><span class="audit-title-icon">${ic('check-check')}</span><div><span class="eyebrow">${esc(result.standard)}</span><h2>SSM Audit</h2><p>${esc(S.session.name)}</p></div></div><span class="spacer"></span>${statusMarkup(summary)}<button class="btn" id="exportAudit">${ic('file-down')}Export report</button><button class="btn icon-btn" id="auditFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
    <div class="result-tabs" role="tablist" aria-label="Registry results"><button class="result-tab active" type="button" role="tab" aria-selected="true">${ic('check-check')}Audit findings <b>${summary.findings.toLocaleString()}</b></button><button class="result-tab" id="openComparisonResult" type="button" role="tab" aria-selected="false">${ic('square-stack')}Project comparison${S.comparison.result?` <b>${(S.comparison.result.summary.differentSystems+S.comparison.result.summary.targetOnlySystems+S.comparison.result.summary.referenceOnlySystems).toLocaleString()}</b>`:''}</button></div>
    <div class="audit-summary"><div><span>Rows audited</span><b>${summary.rows.toLocaleString()}</b></div><div class="critical"><span>Blockers</span><b>${summary.severity.blocker.toLocaleString()}</b></div><div class="error"><span>Errors</span><b>${summary.severity.error.toLocaleString()}</b></div><div class="warning"><span>Warnings</span><b>${summary.severity.warning.toLocaleString()}</b></div><div><span>Advisories</span><b>${summary.severity.info.toLocaleString()}</b></div><div><span>Checks completed</span><b>${summary.checks.toLocaleString()}</b></div></div>
    <div class="audit-categories">${categoryTabs(summary)}</div>
    <div class="audit-toolbar"><div class="searchbox">${ic('search')}<input id="auditSearch" aria-label="Search findings" placeholder="Search tags, rules, and explanations" value="${esc(S.session.search)}"></div><div class="audit-filter-wrap"><button class="btn audit-filter-button ${filterCount()?'active':''}" id="auditFilters" type="button" aria-expanded="${S.session.filterOpen?'true':'false'}">${ic('filter')}Filters <b id="auditFilterCount">${filterCount()||''}</b></button>${filterMenu(result)}</div><select id="auditSort" aria-label="Sort findings">${sortOptions()}</select><span class="audit-count" id="auditCount"></span></div>
    <div class="audit-table-wrap" id="auditTableWrap"><table class="audit-table"><thead><tr><th>Severity</th><th>Finding</th><th>Equipment ID</th><th>Rule source</th><th>Evidence</th><th aria-label="Open details"></th></tr></thead><tbody id="auditRows"></tbody></table></div>
  </section>`;
  wireAuditResult(navigate);
}

function auditRowHtml(finding){return `<tr data-audit-finding="${esc(finding.id)}" tabindex="0" aria-label="Open ${esc(finding.rule.title)} finding for ${esc(finding.equipmentId||'registry')}"><td><span class="audit-severity ${finding.severity}">${SEVERITY_LABELS[finding.severity]}</span></td><td><b>${esc(finding.why)}</b><span>${esc(finding.rule.title)} &middot; ${esc(CATEGORY_LABELS[finding.category]||finding.category)}</span></td><td>${copyTagHtml(finding.equipmentId)}</td><td>${esc(SOURCE_LABELS[finding.rule.source]||finding.rule.source)}</td><td>${esc(finding.sheet||'Registry')} &middot; ${finding.row||'&mdash;'}</td><td>${ic('chevron-right')}</td></tr>`;}
function renderRows(){
  const wrap=$('#auditTableWrap'),body=$('#auditRows');if(!wrap||!body)return;const findings=filteredFindings(),count=$('#auditCount');if(count)count.textContent=`${findings.length.toLocaleString()} of ${S.session.result.findings.length.toLocaleString()} findings`;
  const visible=Math.min(AUDIT_MAX_ROWS,Math.max(28,Math.ceil(wrap.clientHeight/AUDIT_ROW_HEIGHT)+AUDIT_OVERSCAN*2));
  const start=Math.min(Math.max(0,findings.length-visible),Math.max(0,Math.floor(wrap.scrollTop/AUDIT_ROW_HEIGHT)-AUDIT_OVERSCAN)),end=Math.min(findings.length,start+visible);
  const spacer=height=>height?`<tr class="audit-spacer" aria-hidden="true"><td colspan="6" style="height:${height}px"></td></tr>`:'';
  body.innerHTML=spacer(start*AUDIT_ROW_HEIGHT)+findings.slice(start,end).map(auditRowHtml).join('')+spacer((findings.length-end)*AUDIT_ROW_HEIGHT);wireCopyTags(body);
  $$('[data-audit-finding]',body).forEach(row=>{row.onclick=event=>{if(!event.target.closest('[data-copy-tag]'))openFinding(row.dataset.auditFinding,row);};row.onkeydown=event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();openFinding(row.dataset.auditFinding,row);}};});
}
function openFinding(id,opener){
  const finding=S.session.result.findings.find(item=>item.id===id);if(!finding)return;
  S.session.selectedFindingId=id;S.session.opener=opener;$('#drawerTitle').textContent='Audit finding';$('#drawerBody').innerHTML=`<div class="audit-drawer"><div class="audit-drawer-top"><span class="audit-severity ${finding.severity}">${SEVERITY_LABELS[finding.severity]}</span><code>${esc(finding.rule.id)}</code></div><h3>${esc(finding.why)}</h3>${finding.equipmentId?`<div class="audit-subject">${copyTagHtml(finding.equipmentId)}</div>`:''}<div class="audit-detail"><span>Found</span><p>${esc(typeof finding.actual==='string'?finding.actual:JSON.stringify(finding.actual))||'Blank'}</p></div><div class="audit-detail expected"><span>Expected</span><p>${esc(typeof finding.expected==='string'?finding.expected:JSON.stringify(finding.expected))}</p></div><div class="audit-detail action"><span>Recommended correction</span><p>${esc(finding.recommendation)}</p></div><dl class="audit-evidence"><dt>Rule source</dt><dd>${esc(SOURCE_LABELS[finding.rule.source]||finding.rule.source)}</dd><dt>Confidence</dt><dd>${esc(finding.rule.confidence)}</dd><dt>Rule</dt><dd><b>${esc(finding.rule.title)}</b><br>${esc(finding.rule.statement)}</dd><dt>Evidence</dt><dd>${esc(finding.sheet||'Registry')} &middot; row ${finding.row||'&mdash;'}${finding.field?' &middot; '+esc(finding.field):''}</dd><dt>Fingerprint</dt><dd><code>${esc(finding.fingerprint)}</code></dd></dl></div>`;
  wireCopyTags($('#drawerBody'));$('#drawerBack').classList.add('show');$('#drawer').focus();
}
export function closeDrawer(){
  $('#drawerBack').classList.remove('show');const opener=S.session&&S.session.opener;if(opener&&document.contains(opener))opener.focus();
}
function rerenderRows(reset){const wrap=$('#auditTableWrap');if(reset&&wrap)wrap.scrollTop=0;renderRows();}
function updateFilterButton(){const count=filterCount(),button=$('#auditFilters'),badge=$('#auditFilterCount');if(button)button.classList.toggle('active',!!count);if(badge)badge.textContent=count||'';}
function setFilterOpen(open){S.session.filterOpen=!!open;const menu=$('#auditFilterMenu'),button=$('#auditFilters');if(menu){menu.hidden=!open;menu.classList.toggle('open',!!open);}if(button)button.setAttribute('aria-expanded',open?'true':'false');}
function teardownAuditFilters(){if(auditOutsideHandler)document.removeEventListener('pointerdown',auditOutsideHandler);if(auditEscapeHandler)document.removeEventListener('keydown',auditEscapeHandler);auditOutsideHandler=null;auditEscapeHandler=null;}
function wireAuditResult(navigate){
  teardownAuditFilters();
  $('#auditBack').onclick=()=>{S.homeMode='audit';navigate('upload');};$('#exportAudit').onclick=exportSsmAuditXlsx;
  $('#openComparisonResult').onclick=()=>{S.homeMode='compare';if(S.comparison.result)navigate('compare');else navigate('upload');};
  $('#auditFullscreen').onclick=()=>{S.session.fullscreen=!S.session.fullscreen;renderAuditResult(navigate);};
  $('#auditSearch').oninput=event=>{S.session.search=event.target.value;rerenderRows(true);};
  $('#auditSort').onchange=event=>{S.session.sort=event.target.value;rerenderRows(true);};
  $('#auditFilters').onclick=event=>{event.stopPropagation();setFilterOpen(!S.session.filterOpen);};
  $('#auditResetFilters').onclick=()=>{S.session.hiddenSources=[];S.session.hiddenSeverities=[];S.session.hiddenCategories=[];S.session.hiddenRules=[];S.session.filteredCacheKey='';renderAuditResult(navigate);};
  $$('[data-audit-filter-key]').forEach(input=>input.onchange=()=>{const key=input.dataset.auditFilterKey,values=new Set(S.session[key]||[]);if(input.checked)values.delete(input.value);else values.add(input.value);S.session[key]=[...values];S.session.filteredCacheKey='';updateFilterButton();rerenderRows(true);});
  $$('[data-audit-category]').forEach(button=>button.onclick=()=>{const category=button.dataset.auditCategory;if(category==='all')S.session.hiddenCategories=[];else{const hidden=new Set(S.session.hiddenCategories||[]);if(hidden.has(category))hidden.delete(category);else hidden.add(category);S.session.hiddenCategories=[...hidden];}S.session.filteredCacheKey='';renderAuditResult(navigate);});
  auditOutsideHandler=event=>{if(S.session.filterOpen&&!event.target.closest('.audit-filter-wrap'))setFilterOpen(false);};document.addEventListener('pointerdown',auditOutsideHandler);
  auditEscapeHandler=event=>{if(event.key==='Escape'&&S.session.filterOpen){setFilterOpen(false);$('#auditFilters')?.focus();}};document.addEventListener('keydown',auditEscapeHandler);
  const wrap=$('#auditTableWrap');wrap.scrollTop=S.session.scrollTop||0;let frame=0;wrap.onscroll=()=>{S.session.scrollTop=wrap.scrollTop;if(frame)return;frame=requestAnimationFrame(()=>{frame=0;renderRows();});};requestAnimationFrame(renderRows);
}

function comparisonSystems(){
  const query=auditComparisonQuery(S.comparison.systemSearch),filter=S.comparison.systemFilter||'different',sort=S.comparison.systemSort||'upn-asc';let systems=S.comparison.result.systems.filter(system=>(filter==='all'||filter==='different'&&system.status!=='aligned'||system.status===filter)&&(!query||auditComparisonQuery([system.upn,system.label,system.targetName,system.referenceName].join(' ')).includes(query)));
  systems=[...systems].sort((a,b)=>sort==='differences-desc'?(b.differenceCount-a.differenceCount||String(a.upn).localeCompare(String(b.upn),undefined,{numeric:true})):sort==='rows-desc'?(Math.max(b.targetRows,b.referenceRows)-Math.max(a.targetRows,a.referenceRows)||String(a.upn).localeCompare(String(b.upn),undefined,{numeric:true})):String(a.upn).localeCompare(String(b.upn),undefined,{numeric:true}));return systems;
}
function auditComparisonQuery(value){return clean(value).toUpperCase();}
function comparisonStatusLabel(status){return status==='aligned'?'Aligned':status==='target-only'?'Target only':status==='reference-only'?'Reference only':'Different';}
function selectedComparisonSystem(){return S.comparison.result.systems.find(system=>system.upn===S.comparison.selectedUpn)||S.comparison.result.systems[0];}
function systemListHtml(){const systems=comparisonSystems();if(!systems.length)return '<div class="compare-empty">No systems match these filters.</div>';return systems.map(system=>`<button class="compare-system ${system.upn===S.comparison.selectedUpn?'active':''}" type="button" data-compare-upn="${esc(system.upn)}"><span class="compare-upn">${esc(system.upn)}</span><span class="compare-system-copy"><b>${esc(stripSystemUpn(system.label))}</b><small>${system.targetRows.toLocaleString()} target &middot; ${system.referenceRows.toLocaleString()} reference</small></span><span class="compare-system-status ${system.status}">${system.status==='aligned'?ic('check'):system.differenceCount.toLocaleString()}</span></button>`).join('');}
function stripSystemUpn(value){return clean(value).replace(/^\s*[0-9]{3,4}\s*[-:]?\s*/,'')||clean(value);}
function countCard(label,target,reference){return `<div><span>${esc(label)}</span><b>${Number(target).toLocaleString()} <small>target</small></b><b>${Number(reference).toLocaleString()} <small>reference</small></b></div>`;}
function observationIcon(type){return type==='headers'?'folder-tree':type==='controls'?'network':type==='hierarchy'?'list-tree':type==='coverage'?'layers':'tag';}
function comparisonObservations(system){
  if(!system.observations.length)return '<div class="compare-aligned-note">'+ic('check-check')+'No system-level pattern differences were found after Building was excluded.</div>';
  return `<div class="compare-observation-list">${system.observations.slice(0,80).map(item=>`<div class="compare-observation ${item.type}"><span>${ic(observationIcon(item.type))}</span><div><b>${esc(item.title)}</b><small>${esc(item.subject)}</small></div><code>${esc(item.target)} / ${esc(item.reference)}</code></div>`).join('')}${system.observations.length>80?`<p class="compare-more">${(system.observations.length-80).toLocaleString()} additional observations are included in the Excel export.</p>`:''}</div>`;
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
function comparisonNodeHtml(node,side){if(!node)return `<div class="compare-node missing"><span>${side==='target'?'Not present in target':'Not present in completed project'}</span></div>`;return `<div class="compare-node"><div class="compare-node-main" style="--node-depth:${Math.min(6,node.depth)}"><span class="compare-node-line"></span><div><b>${copyTagHtml(node.tag)}</b><small>${esc(node.role)}</small></div></div><dl><dt>Parent type</dt><dd>${esc(node.parentRole)}</dd>${node.underHeader?`<dt>Header</dt><dd>${copyTagHtml(node.headerName)}</dd>`:''}<dt>Discipline</dt><dd>${esc(node.disciplineKind)}</dd>${node.dependencyRoles.length?`<dt>Dependencies</dt><dd>${esc(node.dependencyRoles.join('; '))}</dd>`:''}</dl></div>`;}
function pairHtml(pair){return `<article class="compare-pair ${pair.status}" data-compare-pair><div class="compare-pair-status"><span class="comparison-pill ${pair.status}">${comparisonStatusLabel(pair.status)}</span><b>${esc(pair.differences.join(' · ')||'Same semantic placement')}</b><small>${esc(pair.matchReason)}</small></div><div class="compare-side target"><div class="compare-side-label">Target registry</div>${comparisonNodeHtml(pair.target,'target')}</div><div class="compare-side reference"><div class="compare-side-label">Completed project</div>${comparisonNodeHtml(pair.reference,'reference')}</div></article>`;}
function renderComparisonPairs(){
  const wrap=$('#comparePairWrap'),body=$('#comparePairs'),count=$('#comparePairCount');if(!wrap||!body)return;const system=selectedComparisonSystem(),pairs=system?system.pairs.filter(pairVisible):[];if(count)count.textContent=`${pairs.length.toLocaleString()} of ${(system&&system.pairs.length||0).toLocaleString()} mappings`;
  const visible=Math.min(COMPARE_MAX_ROWS,Math.max(18,Math.ceil(wrap.clientHeight/COMPARE_ROW_HEIGHT)+COMPARE_OVERSCAN*2)),start=Math.min(Math.max(0,pairs.length-visible),Math.max(0,Math.floor(wrap.scrollTop/COMPARE_ROW_HEIGHT)-COMPARE_OVERSCAN)),end=Math.min(pairs.length,start+visible);
  const spacer=height=>height?`<div class="compare-pair-spacer" style="height:${height}px" aria-hidden="true"></div>`:'';body.innerHTML=spacer(start*COMPARE_ROW_HEIGHT)+pairs.slice(start,end).map(pairHtml).join('')+spacer((pairs.length-end)*COMPARE_ROW_HEIGHT);wireCopyTags(body);
}
function renderComparisonDetail(){
  const system=selectedComparisonSystem(),detail=$('#compareDetail');if(!detail||!system)return;const types=comparisonSystemTypes(system);
  detail.innerHTML=`<div class="compare-detail-head"><div><span class="eyebrow">UPN ${esc(system.upn)}</span><h3>${esc(stripSystemUpn(system.label))}</h3><p>Building values are excluded. Equipment is paired by tag core and semantic role.</p></div><span class="comparison-pill ${system.status}">${comparisonStatusLabel(system.status)}</span></div>
  <div class="compare-counts">${countCard('Equipment',system.targetRows,system.referenceRows)}${countCard('Headers',system.targetHeaders,system.referenceHeaders)}${countCard('I&C / Controls',system.targetControls,system.referenceControls)}<div><span>Pattern findings</span><b>${system.observations.length.toLocaleString()} <small>observed</small></b><b>${(types.hierarchy+types.headers+types.controls).toLocaleString()} <small>nesting</small></b></div></div>
  <section class="compare-tree-section"><div class="compare-section-title"><div><b>Synchronized hierarchy</b><span>Open either side to reveal both matched branches. Rows and scrolling stay aligned.</span></div><span id="compareTreeCount"></span><div class="compare-tree-actions"><button class="btn ghost sm" id="compareTreeCollapse" type="button">${ic('chevrons-up')}Collapse</button><button class="btn ghost sm" id="compareTreeExpand" type="button">${ic('list-tree')}Expand all</button></div></div><div class="compare-tree-grid"><div class="compare-tree-head"><b>Registry to audit</b><span>Alignment</span><b>Completed project</b></div><div class="compare-tree-wrap" id="compareTreeWrap"><div id="compareTreeRows"></div></div></div></section>
  <section class="compare-observations"><div class="compare-section-title"><div><b>System differences</b><span>Counts compare equipment meaning and nesting patterns, not Building.</span></div></div>${comparisonObservations(system)}</section>
  <section class="compare-mapping"><div class="compare-section-title"><div><b>Side-by-side equipment mapping</b><span>Matched rows remain visible for context; filter to focus on changes.</span></div><span id="comparePairCount"></span></div><div class="compare-pair-toolbar"><div class="searchbox">${ic('search')}<input id="compareRowSearch" aria-label="Search mapped equipment" placeholder="Search tags, descriptions, parents" value="${esc(S.comparison.rowSearch)}"></div><select id="compareRowFilter" aria-label="Filter mappings"><option value="different" ${S.comparison.rowFilter==='different'?'selected':''}>Differences only</option><option value="all" ${S.comparison.rowFilter==='all'?'selected':''}>All mappings</option><option value="changed" ${S.comparison.rowFilter==='changed'?'selected':''}>Changed pairs</option><option value="target-only" ${S.comparison.rowFilter==='target-only'?'selected':''}>Target only</option><option value="reference-only" ${S.comparison.rowFilter==='reference-only'?'selected':''}>Reference only</option><option value="aligned" ${S.comparison.rowFilter==='aligned'?'selected':''}>Aligned pairs</option></select></div><div class="compare-pair-wrap" id="comparePairWrap"><div id="comparePairs"></div></div></section>`;
  wireComparisonDetail();
}
function wireComparisonDetail(){
  const search=$('#compareRowSearch'),filter=$('#compareRowFilter'),wrap=$('#comparePairWrap');if(search)search.oninput=event=>{S.comparison.rowSearch=event.target.value;if(wrap)wrap.scrollTop=0;renderComparisonPairs();};if(filter)filter.onchange=event=>{S.comparison.rowFilter=event.target.value;if(wrap)wrap.scrollTop=0;renderComparisonPairs();};
  wireComparisonTree();
  if(wrap){wrap.scrollTop=S.comparison.pairScrollTop||0;let frame=0;wrap.onscroll=()=>{S.comparison.pairScrollTop=wrap.scrollTop;if(frame)return;frame=requestAnimationFrame(()=>{frame=0;renderComparisonPairs();});};requestAnimationFrame(renderComparisonPairs);}
}

export function renderComparisonResult(navigate){
  const result=S.comparison&&S.comparison.result;if(!result){S.homeMode='compare';navigate('upload');return;}S.screen='compare';const summary=result.summary,fullscreen=!!S.comparison.fullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  $('#view').innerHTML=`<section id="compareShell" class="compare-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><button class="btn ghost" id="compareBack">${ic('arrow-left')}Files</button><div class="audit-title"><span class="audit-title-icon">${ic('square-stack')}</span><div><span class="eyebrow">${esc(result.standard)}</span><h2>Project Comparison</h2><p>${esc(S.comparison.targetName)} &middot; ${esc(S.comparison.referenceName)}</p></div></div><span class="spacer"></span><span class="compare-building-note">${ic('layers')}Building ignored</span><button class="btn" id="exportComparison">${ic('file-down')}Export comparison</button><button class="btn icon-btn" id="compareFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
    <div class="result-tabs" role="tablist" aria-label="Registry results"><button class="result-tab" id="openAuditFromCompare" type="button" role="tab" aria-selected="false">${ic('check-check')}Audit findings <b>${S.session.result.summary.findings.toLocaleString()}</b></button><button class="result-tab active" type="button" role="tab" aria-selected="true">${ic('square-stack')}Project comparison <b>${(summary.differentSystems+summary.targetOnlySystems+summary.referenceOnlySystems).toLocaleString()}</b></button></div>
    <div class="compare-summary"><div><span>Systems compared</span><b>${summary.systems.toLocaleString()}</b></div><div class="difference"><span>Systems with differences</span><b>${(summary.differentSystems+summary.targetOnlySystems+summary.referenceOnlySystems).toLocaleString()}</b></div><div class="aligned"><span>Aligned systems</span><b>${summary.alignedSystems.toLocaleString()}</b></div><div><span>Changed pairs</span><b>${summary.changedRows.toLocaleString()}</b></div><div><span>Target only</span><b>${summary.targetOnlyRows.toLocaleString()}</b></div><div><span>Reference only</span><b>${summary.referenceOnlyRows.toLocaleString()}</b></div></div>
    <div class="compare-workspace"><aside class="compare-system-panel"><div class="compare-system-tools"><div class="searchbox">${ic('search')}<input id="compareSystemSearch" aria-label="Search systems" placeholder="Find UPN or system" value="${esc(S.comparison.systemSearch)}"></div><div><select id="compareSystemFilter" aria-label="Filter systems"><option value="different" ${S.comparison.systemFilter==='different'?'selected':''}>Differences</option><option value="all" ${S.comparison.systemFilter==='all'?'selected':''}>All systems</option><option value="aligned" ${S.comparison.systemFilter==='aligned'?'selected':''}>Aligned</option><option value="target-only" ${S.comparison.systemFilter==='target-only'?'selected':''}>Target only</option><option value="reference-only" ${S.comparison.systemFilter==='reference-only'?'selected':''}>Reference only</option></select><select id="compareSystemSort" aria-label="Sort systems"><option value="upn-asc" ${S.comparison.systemSort==='upn-asc'?'selected':''}>UPN</option><option value="differences-desc" ${S.comparison.systemSort==='differences-desc'?'selected':''}>Most differences</option><option value="rows-desc" ${S.comparison.systemSort==='rows-desc'?'selected':''}>Most equipment</option></select></div></div><div class="compare-system-list" id="compareSystemList">${systemListHtml()}</div></aside><main class="compare-detail" id="compareDetail"></main></div>
  </section>`;
  wireComparisonResult(navigate);renderComparisonDetail();
}

function renderSystemList(){const list=$('#compareSystemList');if(!list)return;list.innerHTML=systemListHtml();wireSystemButtons();}
function wireSystemButtons(){$$('[data-compare-upn]').forEach(button=>button.onclick=()=>{S.comparison.selectedUpn=button.dataset.compareUpn;S.comparison.pairScrollTop=0;renderSystemList();renderComparisonDetail();});}
function wireComparisonResult(navigate){
  teardownAuditFilters();$('#compareBack').onclick=()=>{S.homeMode='compare';navigate('upload');};$('#exportComparison').onclick=exportSsmComparisonXlsx;$('#openAuditFromCompare').onclick=()=>navigate('audit');$('#compareFullscreen').onclick=()=>{S.comparison.fullscreen=!S.comparison.fullscreen;renderComparisonResult(navigate);};
  $('#compareSystemSearch').oninput=event=>{S.comparison.systemSearch=event.target.value;renderSystemList();};$('#compareSystemFilter').onchange=event=>{S.comparison.systemFilter=event.target.value;renderSystemList();};$('#compareSystemSort').onchange=event=>{S.comparison.systemSort=event.target.value;renderSystemList();};wireSystemButtons();
}
