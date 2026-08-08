import { $, $$, clean, esc } from '../core/text.js'
import { S, resetSession } from '../state.js'
import { readArrayBuffer } from '../io/workbook.js'
import { auditSnapshotFromWorkbook } from '../audit/model.js'
import { runSsmAudit, SSM_AUDIT_CATEGORIES, SSM_AUDIT_SEVERITIES } from '../audit/engine.js'
import { exportSsmAuditXlsx } from '../audit/export.js'
import { ic } from './icons.js'
import { copyTagHtml, runWithProgress, toast, wireCopyTags } from './feedback.js'

const AUDIT_ROW_HEIGHT=48,AUDIT_OVERSCAN=18,AUDIT_MAX_ROWS=160;
const CATEGORY_LABELS={upload:'Upload Readiness',structure:'Structure',dependencies:'Dependencies',metadata:'Metadata',milestones:'Milestones','item-masters':'Item Masters',headers:'Headers / Rollups'};
const SEVERITY_LABELS={blocker:'Blocker',error:'Error',warning:'Warning',info:'Advisory'};

function importStatus(){
  const result=S.session.result,summary=result&&result.summary;
  if(result)return `<div class="audit-import ready"><span class="file-icon">${ic(summary.status==='ready'?'check-check':'triangle-alert')}</span><div class="file-meta"><b>${esc(S.session.name)}</b><span>${summary.rows.toLocaleString()} rows audited &middot; ${summary.findings.toLocaleString()} findings &middot; ${esc(summary.status.toUpperCase())}</span></div><button class="btn" id="openAuditResult">Open audit</button><button class="xbtn icon-btn" id="removeAuditTarget" type="button" aria-label="Remove audit target">${ic('x')}</button></div>`;
  if(S.session.error)return `<div class="audit-import error"><span class="file-icon">${ic('triangle-alert')}</span><div class="file-meta"><b>Could not run SSM Audit</b><span>${esc(S.session.error)}</span></div><button class="btn" id="chooseAuditAgain">Choose another</button></div>`;
  return '';
}

export function renderUpload(navigate){
  document.body.classList.remove('audit-fullscreen');S.screen='upload';
  $('#view').innerHTML=`<section class="upload-shell">
    <div class="screen-heading"><div><span class="eyebrow">SSM SOP + Exto Rev21</span><h2>Audit an existing SSM</h2><p>Check a completed Exto Cx Registry and receive an evidence-backed correction report.</p></div><button class="btn ghost" id="openGuide">${ic('book-open')}Guide</button></div>
    <div class="upload-grid">
      <div class="dropzone" id="dropzone" tabindex="0" role="button" aria-label="Choose an Exto Cx Registry workbook">
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
      <div><span>03</span><b>Exto readiness</b><p>Approved values, gating fields, milestones, Item Masters, and headers.</p></div>
    </div>
  </section>`;
  wireUpload(navigate);
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
  const remove=$('#removeAuditTarget');if(remove)remove.onclick=()=>{resetSession();renderUpload(navigate);};
  $('#openGuide').onclick=()=>document.dispatchEvent(new CustomEvent('ssm-audit:guide'));
}

export async function addAuditTarget(file,navigate){
  if(!/\.(xlsx|xls)$/i.test(file.name)){toast('SSM Audit requires an Excel workbook');return;}
  resetSession();S.session.name=file.name;
  try{
    await runWithProgress('Running SSM Audit',file.name,async(checkpoint,report)=>{
      const bytes=new Uint8Array(await readArrayBuffer(file));await checkpoint();
      const workbook=XLSX.read(bytes,{type:'array',dense:true});report(.08,'Registry opened');await checkpoint();
      const snapshot=await auditSnapshotFromWorkbook(workbook,file.name,checkpoint,(fraction,label)=>report(.08+fraction*.42,label));
      report(.55,`${snapshot.rows.length.toLocaleString()} rows parsed`);await checkpoint();
      const result=runSsmAudit(snapshot);report(1,`${result.findings.length.toLocaleString()} findings`);
      S.session={...S.session,snapshot,result,error:''};
    });
    navigate('audit');
  }catch(error){
    console.error('SSM Audit failed',error);S.session.error=error&&error.message||'Could not read this registry';S.session.snapshot=null;S.session.result=null;renderUpload(navigate);
  }
}

function filteredFindings(){
  const query=clean(S.session.search).toUpperCase();let rows=S.session.result.findings.filter(finding=>(S.session.severity==='all'||finding.severity===S.session.severity)&&(S.session.category==='all'||finding.category===S.session.category)&&(!query||finding.searchKey.includes(query)));
  if(S.session.sort==='equipment')rows=[...rows].sort((a,b)=>a.equipmentId.localeCompare(b.equipmentId,undefined,{numeric:true,sensitivity:'base'})||a.row-b.row);
  else if(S.session.sort==='rule')rows=[...rows].sort((a,b)=>a.rule.id.localeCompare(b.rule.id)||a.row-b.row);
  return rows;
}
function statusMarkup(summary){const label=summary.status==='blocked'?'Blocked':summary.status==='review'?'Review required':'Ready';return `<span class="audit-status ${summary.status}">${ic(summary.status==='ready'?'check-check':'triangle-alert')}${label}</span>`;}
function severityOptions(){return `<option value="all">All severities</option>`+SSM_AUDIT_SEVERITIES.map(value=>`<option value="${value}" ${S.session.severity===value?'selected':''}>${SEVERITY_LABELS[value]}</option>`).join('');}
function categoryOptions(){return `<option value="all">All categories</option>`+SSM_AUDIT_CATEGORIES.map(value=>`<option value="${value}" ${S.session.category===value?'selected':''}>${CATEGORY_LABELS[value]}</option>`).join('');}
function categoryTabs(summary){return `<button class="audit-cat ${S.session.category==='all'?'active':''}" data-audit-category="all">All <b>${summary.findings.toLocaleString()}</b></button>`+SSM_AUDIT_CATEGORIES.map(category=>`<button class="audit-cat ${S.session.category===category?'active':''}" data-audit-category="${category}">${CATEGORY_LABELS[category]} <b>${(summary.category[category]||0).toLocaleString()}</b></button>`).join('');}

export function renderAuditResult(navigate){
  const result=S.session&&S.session.result;if(!result){navigate('upload');return;}
  S.screen='audit';const summary=result.summary,fullscreen=!!S.session.fullscreen;document.body.classList.toggle('audit-fullscreen',fullscreen);
  $('#view').innerHTML=`<section id="auditShell" class="audit-shell ${fullscreen?'fullscreen':''}">
    <div class="audit-head"><button class="btn ghost" id="auditBack">${ic('arrow-left')}Files</button><div class="audit-title"><span class="audit-title-icon">${ic('check-check')}</span><div><span class="eyebrow">${esc(result.standard)}</span><h2>SSM Audit</h2><p>${esc(S.session.name)}</p></div></div><span class="spacer"></span>${statusMarkup(summary)}<button class="btn" id="exportAudit">${ic('file-down')}Export report</button><button class="btn icon-btn" id="auditFullscreen" title="${fullscreen?'Exit full screen':'Full screen'}" aria-label="${fullscreen?'Exit full screen':'Full screen'}">${ic(fullscreen?'minimize-2':'maximize-2')}</button></div>
    <div class="audit-summary"><div><span>Rows audited</span><b>${summary.rows.toLocaleString()}</b></div><div class="critical"><span>Blockers</span><b>${summary.severity.blocker.toLocaleString()}</b></div><div class="error"><span>Errors</span><b>${summary.severity.error.toLocaleString()}</b></div><div class="warning"><span>Warnings</span><b>${summary.severity.warning.toLocaleString()}</b></div><div><span>Advisories</span><b>${summary.severity.info.toLocaleString()}</b></div><div><span>Checks completed</span><b>${summary.checks.toLocaleString()}</b></div></div>
    <div class="audit-categories">${categoryTabs(summary)}</div>
    <div class="audit-toolbar"><div class="searchbox">${ic('search')}<input id="auditSearch" aria-label="Search findings" placeholder="Search tags, rules, and explanations" value="${esc(S.session.search)}"></div><select id="auditSeverity" aria-label="Severity">${severityOptions()}</select><select id="auditCategory" aria-label="Category">${categoryOptions()}</select><select id="auditSort" aria-label="Sort findings"><option value="severity" ${S.session.sort==='severity'?'selected':''}>Severity order</option><option value="equipment" ${S.session.sort==='equipment'?'selected':''}>Equipment ID</option><option value="rule" ${S.session.sort==='rule'?'selected':''}>Rule</option></select><span class="audit-count" id="auditCount"></span></div>
    <div class="audit-table-wrap" id="auditTableWrap"><table class="audit-table"><thead><tr><th>Severity</th><th>Finding</th><th>Equipment ID</th><th>Category</th><th>Source</th><th aria-label="Open details"></th></tr></thead><tbody id="auditRows"></tbody></table></div>
  </section>`;
  wireAuditResult(navigate);
}

function auditRowHtml(finding){return `<tr data-audit-finding="${esc(finding.id)}" tabindex="0" aria-label="Open ${esc(finding.rule.id)} finding for ${esc(finding.equipmentId||'registry')}"><td><span class="audit-severity ${finding.severity}">${SEVERITY_LABELS[finding.severity]}</span></td><td><b>${esc(finding.why)}</b><span>${esc(finding.rule.id)}</span></td><td>${copyTagHtml(finding.equipmentId)}</td><td>${esc(CATEGORY_LABELS[finding.category]||finding.category)}</td><td>${esc(finding.sheet||'Registry')} &middot; ${finding.row||'&mdash;'}</td><td>${ic('chevron-right')}</td></tr>`;}
function renderRows(){
  const wrap=$('#auditTableWrap'),body=$('#auditRows');if(!wrap||!body)return;const findings=filteredFindings(),count=$('#auditCount');if(count)count.textContent=`${findings.length.toLocaleString()} finding${findings.length===1?'':'s'}`;
  const visible=Math.min(AUDIT_MAX_ROWS,Math.max(28,Math.ceil(wrap.clientHeight/AUDIT_ROW_HEIGHT)+AUDIT_OVERSCAN*2));
  const start=Math.min(Math.max(0,findings.length-visible),Math.max(0,Math.floor(wrap.scrollTop/AUDIT_ROW_HEIGHT)-AUDIT_OVERSCAN)),end=Math.min(findings.length,start+visible);
  const spacer=height=>height?`<tr class="audit-spacer" aria-hidden="true"><td colspan="6" style="height:${height}px"></td></tr>`:'';
  body.innerHTML=spacer(start*AUDIT_ROW_HEIGHT)+findings.slice(start,end).map(auditRowHtml).join('')+spacer((findings.length-end)*AUDIT_ROW_HEIGHT);wireCopyTags(body);
  $$('[data-audit-finding]',body).forEach(row=>{row.onclick=event=>{if(!event.target.closest('[data-copy-tag]'))openFinding(row.dataset.auditFinding,row);};row.onkeydown=event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();openFinding(row.dataset.auditFinding,row);}};});
}
function openFinding(id,opener){
  const finding=S.session.result.findings.find(item=>item.id===id);if(!finding)return;
  S.session.selectedFindingId=id;S.session.opener=opener;$('#drawerTitle').textContent='Audit finding';$('#drawerBody').innerHTML=`<div class="audit-drawer"><div class="audit-drawer-top"><span class="audit-severity ${finding.severity}">${SEVERITY_LABELS[finding.severity]}</span><code>${esc(finding.rule.id)}</code></div><h3>${esc(finding.why)}</h3>${finding.equipmentId?`<div class="audit-subject">${copyTagHtml(finding.equipmentId)}</div>`:''}<div class="audit-detail"><span>Found</span><p>${esc(typeof finding.actual==='string'?finding.actual:JSON.stringify(finding.actual))||'Blank'}</p></div><div class="audit-detail expected"><span>Expected</span><p>${esc(typeof finding.expected==='string'?finding.expected:JSON.stringify(finding.expected))}</p></div><div class="audit-detail action"><span>Recommended correction</span><p>${esc(finding.recommendation)}</p></div><dl class="audit-evidence"><dt>Standard</dt><dd>${esc(finding.rule.standardRef)}</dd><dt>Evidence</dt><dd>${esc(finding.sheet||'Registry')} &middot; row ${finding.row||'&mdash;'}${finding.field?' &middot; '+esc(finding.field):''}</dd><dt>Fingerprint</dt><dd><code>${esc(finding.fingerprint)}</code></dd></dl></div>`;
  wireCopyTags($('#drawerBody'));$('#drawerBack').classList.add('show');$('#drawer').focus();
}
export function closeDrawer(){
  $('#drawerBack').classList.remove('show');const opener=S.session&&S.session.opener;if(opener&&document.contains(opener))opener.focus();
}
function rerenderRows(reset){const wrap=$('#auditTableWrap');if(reset&&wrap)wrap.scrollTop=0;renderRows();}
function wireAuditResult(navigate){
  $('#auditBack').onclick=()=>navigate('upload');$('#exportAudit').onclick=exportSsmAuditXlsx;
  $('#auditFullscreen').onclick=()=>{S.session.fullscreen=!S.session.fullscreen;renderAuditResult(navigate);};
  $('#auditSearch').oninput=event=>{S.session.search=event.target.value;rerenderRows(true);};
  $('#auditSeverity').onchange=event=>{S.session.severity=event.target.value;rerenderRows(true);};
  $('#auditCategory').onchange=event=>{S.session.category=event.target.value;renderAuditResult(navigate);};
  $('#auditSort').onchange=event=>{S.session.sort=event.target.value;rerenderRows(true);};
  $$('[data-audit-category]').forEach(button=>button.onclick=()=>{S.session.category=button.dataset.auditCategory;renderAuditResult(navigate);});
  const wrap=$('#auditTableWrap');wrap.scrollTop=S.session.scrollTop||0;let frame=0;wrap.onscroll=()=>{S.session.scrollTop=wrap.scrollTop;if(frame)return;frame=requestAnimationFrame(()=>{frame=0;renderRows();});};requestAnimationFrame(renderRows);
}
