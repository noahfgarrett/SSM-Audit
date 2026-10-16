import { $, $$, clean, esc, natCmp } from '../core/text.js'
import { S } from '../state.js'
import { ic } from './icons.js'
import { copyTagHtml, toast, wireCopyTags } from './feedback.js'
import { downloadBlob, workbookBlobCompact } from '../core/download.js'
import { buildDependencyCheckWorkbook } from '../audit/dependency-export.js'
import { engineeringUploadsHtml, renderSideNav, teardownAuditFilters, wireEngineeringUploads } from './audit.js'

/* Dependencies tab: what the power and control documents say each piece of
   equipment should depend on, checked against this registry. Read-only view
   over result.dependencyCheck plus a one-click workbook download. */

const DEPENDENCY_DOCUMENT_KINDS=Object.freeze(['easyPower','cable','pmd']);
const DEPENDENCY_CHUNK=200;
const DEPENDENCY_STATUS=Object.freeze({
  match:{label:'Match',icon:'circle-check'},
  missing:{label:'Missing',icon:'circle-x'},
  absent:{label:'Not in registry',icon:'triangle-alert'},
});

export function dependencyCheckOf(session){return session?.rawResult?.dependencyCheck||null;}

function dependencyLocation(source){return `${source.label} ${source.sheet}, row ${source.row}`;}
function dependencySourceHtml(sources){
  const unique=[...new Set(sources.map(dependencyLocation))];
  if(!unique.length)return '<span class="muted">—</span>';
  return `<span title="${esc(unique.join('\n'))}">${esc(unique[0])}${unique.length>1?` <span class="muted">+${unique.length-1}</span>`:''}</span>`;
}
function dependencySentence(text){const value=clean(text);return value.charAt(0).toUpperCase()+value.slice(1);}
function dependencyTagsHtml(tags,emptyText='—'){return tags.length?tags.map(tag=>`<code>${esc(tag)}</code>`).join(' '):`<span class="muted">${esc(emptyText)}</span>`;}

/* One list of display items: compared registry rows, then tags missing from the registry. */
export function dependencyItems(check){
  return [
    ...check.rows.map(entry=>({key:`row:${entry.equipmentId}`,kind:entry.status,tag:entry.equipmentId,entry,discipline:clean(entry.discipline),building:clean(entry.building)})),
    ...check.notInRegistry.map(entry=>({key:`absent:${entry.tag}`,kind:'absent',tag:entry.tag,entry,discipline:'',building:''})),
  ];
}
export function dependencyFiltered(items,session){
  const query=clean(session.depSearch||'').toUpperCase(),status=session.depStatus||'all',discipline=session.depDiscipline||'all',building=session.depBuilding||'all';
  return items.filter(item=>{
    if(status==='compared'?item.kind==='absent':status!=='all'&&item.kind!==status)return false;
    if(discipline!=='all'&&item.discipline!==discipline)return false;
    if(building!=='all'&&item.building!==building)return false;
    if(query){
      const haystack=[item.tag,...(item.kind==='absent'?item.entry.affects:[...item.entry.missing,...item.entry.expected.map(expected=>expected.tag)])].join(' ').toUpperCase();
      if(!haystack.includes(query))return false;
    }
    return true;
  });
}
function dependencyDetailHtml(item){
  if(item.kind==='absent'){
    const entry=item.entry;
    return `<div class="dep-detail"><p>This tag appears in the documents but not in the registry. Rows fed through it were compared against the next tag upstream.</p>
      <dl><dt>Found at</dt><dd>${entry.sources.map(source=>esc(dependencyLocation(source))).join('<br>')}</dd>
      <dt>Registry rows affected</dt><dd>${dependencyTagsHtml(entry.affects,'None')}</dd></dl></div>`;
  }
  const entry=item.entry;
  return `<div class="dep-detail">${entry.expected.map(expected=>{
    const chain=[expected.tag,...[...expected.skipped].reverse(),entry.equipmentId];
    const present=!entry.missing.includes(expected.tag);
    return `<section class="dep-expect ${present?'ok':'bad'}"><div class="dep-chain">${ic(present?'circle-check':'circle-x')}${chain.map((tag,index)=>`<code class="${expected.skipped.includes(tag)?'skipped':''}">${esc(tag)}</code>${index<chain.length-1?ic('arrow-right'):''}`).join('')}</div>
      <p>${esc(dependencySentence(expected.reason))} — ${present?'listed in Dependencies.':'not listed in Dependencies.'}</p>
      <small>${expected.sources.map(source=>esc(dependencyLocation(source))).join(' · ')}</small></section>`;
  }).join('')}${entry.notes.length?`<ul class="dep-notes">${entry.notes.map(note=>`<li>${ic('info')}${esc(note)}</li>`).join('')}</ul>`:''}
  <dl><dt>Closest Parent</dt><dd>${entry.closestParent?`<code>${esc(entry.closestParent)}</code>`:'<span class="muted">—</span>'}</dd><dt>Registry dependencies</dt><dd>${dependencyTagsHtml(entry.registryDependencies,'None')}</dd></dl></div>`;
}
export function dependencyRowHtml(item,expanded){
  const status=DEPENDENCY_STATUS[item.kind],entry=item.entry,absent=item.kind==='absent';
  const expected=absent?'<span class="muted">—</span>':dependencyTagsHtml(entry.expected.map(expected=>expected.tag));
  const registry=absent?'<span class="muted">Not in registry</span>':dependencyTagsHtml(entry.registryDependencies,'None');
  const missing=absent?'<span class="muted">—</span>':dependencyTagsHtml(entry.missing,'None');
  const note=absent?(entry.affects.length?`Affects ${entry.affects.length.toLocaleString()} registry row${entry.affects.length===1?'':'s'}`:''):entry.notes[0]||'';
  const sources=absent?entry.sources:entry.expected.flatMap(expected=>expected.sources);
  return `<div class="dep-item ${expanded?'open':''}" data-dep-key="${esc(item.key)}">
    <div class="dep-row" role="button" tabindex="0" aria-expanded="${expanded}">
      <span class="dep-caret">${ic(expanded?'chevron-down':'chevron-right')}</span>
      <span class="dep-tag">${copyTagHtml(item.tag)}</span>
      <span class="dep-status ${item.kind}">${ic(status.icon)}${esc(status.label)}</span>
      <span class="dep-cell" data-label="Expected">${expected}</span>
      <span class="dep-cell" data-label="Registry">${registry}</span>
      <span class="dep-cell" data-label="Missing">${missing}</span>
      <span class="dep-cell dep-note" data-label="Note" title="${esc(note)}">${esc(note)||'<span class="muted">—</span>'}</span>
      <span class="dep-cell dep-source" data-label="Source">${dependencySourceHtml(sources)}</span>
    </div>${expanded?dependencyDetailHtml(item):''}</div>`;
}
function dependencyEmptyHtml(){
  return `<div class="rule-reference-empty">${ic('git-branch')}<b>Add EasyPower, a Cable Schedule or a PMD export</b><span>The check reads which panel, MCC, RIO or PLC feeds each tag. A feed that crosses a Building, Discipline or System Name boundary should be listed as a dependency. Excel or CSV both work.</span></div>`;
}

export async function exportDependencyCheckXlsx(){
  const session=S.session,check=dependencyCheckOf(session);if(!check){toast('Add a power or control document first');return;}
  const base=clean(session.name).replace(/\.[^.]+$/,'').replace(/[^a-z0-9_-]+/gi,'-').replace(/^-+|-+$/g,'')||'SSM';
  try{
    const blob=await workbookBlobCompact(buildDependencyCheckWorkbook(check,session.name),{});
    downloadBlob(`${base}-Dependency-Check.xlsx`,blob);toast('Dependency check exported');
  }catch(_){toast('The dependency check could not be built');}
}

export function renderDependencyCheck(navigate){
  const session=S.session;
  if(!(session&&session.rawResult)){navigate('upload');return;}
  teardownAuditFilters();document.body.classList.remove('audit-fullscreen');S.screen='dependencies';S.homeMode='audit';
  const rerender=()=>{renderDependencyCheck(navigate);renderSideNav(navigate);};
  const check=dependencyCheckOf(session);
  session.depExpanded||=new Set();
  const items=check?dependencyItems(check):[],list=dependencyFiltered(items,session);
  const disciplines=[...new Set(items.map(item=>item.discipline).filter(Boolean))].sort(natCmp),buildings=[...new Set(items.map(item=>item.building).filter(Boolean))].sort(natCmp);
  const status=session.depStatus||'all',counts=check?.counts;
  const countButton=(id,label,value)=>`<button class="dep-count ${id} ${status===id?'active':''}" type="button" data-dep-status="${id}" aria-pressed="${status===id}" title="${status===id?'Show every row':`Show only ${esc(label.toLowerCase())}`}"><b>${(value||0).toLocaleString()}</b><span>${esc(label)}</span></button>`;
  const hasDocuments=DEPENDENCY_DOCUMENT_KINDS.some(kind=>session.references?.[kind]?.entries?.length);
  const filtered=!!clean(session.depSearch)||status!=='all'||(session.depDiscipline||'all')!=='all'||(session.depBuilding||'all')!=='all';
  $('#view').innerHTML=`<section class="completed-shell dep-shell">
    <div class="screen-heading"><div><span class="eyebrow">Power and control documents</span><h2>Dependencies</h2><p>${esc(session.name)} — dependencies that EasyPower, the Cable Schedule and the PMD call for, checked against this registry. A feed that crosses a Building, Discipline or System Name boundary should be a dependency; electrical equipment fed inside one building nests instead.</p></div>
      <button class="btn primary" type="button" id="depExport" ${check?'':'disabled'} title="${check?'Download the Summary, Dependencies and Not in registry sheets':'Add a power or control document first'}">${ic('download')}Download dependency check</button></div>
    ${check&&hasDocuments?`<div class="dep-counts" role="group" aria-label="Filter by status">
      ${countButton('compared','Compared',counts.compared)}${countButton('match','Match',counts.match)}${countButton('missing','Missing',counts.missing)}${countButton('absent','Not in registry',counts.notInRegistry)}</div>
    <div class="modify-toolbar"><div class="searchbox">${ic('search')}<input id="depSearch" aria-label="Search dependency check" placeholder="Search tags" value="${esc(session.depSearch||'')}"></div>
      <select id="depDiscipline" class="modify-dim" aria-label="Filter by discipline"><option value="all">All disciplines</option>${disciplines.map(name=>`<option value="${esc(name)}" ${session.depDiscipline===name?'selected':''}>${esc(name)}</option>`).join('')}</select>
      <select id="depBuilding" class="modify-dim" aria-label="Filter by building"><option value="all">All buildings</option>${buildings.map(name=>`<option value="${esc(name)}" ${session.depBuilding===name?'selected':''}>${esc(name)}</option>`).join('')}</select>
      <span class="spacer"></span><span class="completed-count">${filtered?`${list.length.toLocaleString()} of ${items.length.toLocaleString()} shown`:`${items.length.toLocaleString()} rows`}</span></div>
    <div class="completed-body" id="depBody">${list.length?`<div class="dep-list"><div class="dep-row dep-head" aria-hidden="true"><span></span><b>Equipment ID</b><span>Status</span><span>Expected</span><span>Registry has</span><span>Missing</span><span>Note</span><span>Source</span></div>
      ${list.slice(0,DEPENDENCY_CHUNK).map(item=>dependencyRowHtml(item,session.depExpanded.has(item.key))).join('')}
      ${list.length>DEPENDENCY_CHUNK?`<button class="btn ghost sm modify-more" type="button" id="depMore" data-offset="${DEPENDENCY_CHUNK}">${ic('chevrons-down')}Show ${Math.min(list.length-DEPENDENCY_CHUNK,DEPENDENCY_CHUNK).toLocaleString()} more of ${(list.length-DEPENDENCY_CHUNK).toLocaleString()}</button>`:''}</div>`
      :`<div class="rule-reference-empty">${ic(filtered?'search':'check-check')}<b>${filtered?'Nothing matches those filters':'No feeds cross a group boundary'}</b><span>${filtered?'Try a different search, status, discipline or building.':'Every feed in the selected documents stays inside one Building, Discipline and System Name, so no dependencies are expected.'}</span></div>`}</div>
    <details class="dep-documents"><summary>${ic('file-spreadsheet')}Documents in this check</summary>${engineeringUploadsHtml(DEPENDENCY_DOCUMENT_KINDS,'Power and control documents','Excel or CSV')}</details>`
    :`${dependencyEmptyHtml()}${engineeringUploadsHtml(DEPENDENCY_DOCUMENT_KINDS,'Power and control documents','Excel or CSV')}`}
  </section>`;
  wireEngineeringUploads(navigate,rerender);
  const exportButton=$('#depExport');
  if(exportButton)exportButton.onclick=async()=>{exportButton.disabled=true;try{await exportDependencyCheckXlsx();}finally{exportButton.disabled=false;}};
  if(!check||!hasDocuments)return;
  for(const button of $$('[data-dep-status]'))button.onclick=()=>{session.depStatus=status===button.dataset.depStatus?'all':button.dataset.depStatus;rerender();};
  const search=$('#depSearch');let timer=0;
  search.oninput=()=>{session.depSearch=search.value;clearTimeout(timer);timer=setTimeout(()=>{renderDependencyCheck(navigate);const input=$('#depSearch');if(input){input.focus();input.setSelectionRange(input.value.length,input.value.length);}},150);};
  $('#depDiscipline').onchange=event=>{session.depDiscipline=event.target.value;rerender();};
  $('#depBuilding').onchange=event=>{session.depBuilding=event.target.value;rerender();};
  const body=$('#depBody'),byKey=new Map(list.map(item=>[item.key,item]));
  const toggle=element=>{
    const holder=element.closest('[data-dep-key]');if(!holder)return;
    const item=byKey.get(holder.dataset.depKey);if(!item)return;
    if(session.depExpanded.has(item.key))session.depExpanded.delete(item.key);else session.depExpanded.add(item.key);
    const next=document.createElement('div');next.innerHTML=dependencyRowHtml(item,session.depExpanded.has(item.key));
    const replacement=next.firstElementChild;holder.replaceWith(replacement);wireCopyTags(replacement);replacement.querySelector('.dep-row')?.focus();
  };
  body.onclick=event=>{if(event.target.closest('.copy-tag')||event.target.closest('#depMore'))return;if(event.target.closest('.dep-row:not(.dep-head)'))toggle(event.target);};
  body.onkeydown=event=>{if((event.key==='Enter'||event.key===' ')&&event.target.classList.contains('dep-row')){event.preventDefault();toggle(event.target);}};
  const more=$('#depMore');
  if(more)more.onclick=()=>{
    const offset=Number(more.dataset.offset)||0,slice=list.slice(offset,offset+DEPENDENCY_CHUNK),fragment=document.createElement('div');
    fragment.innerHTML=slice.map(item=>dependencyRowHtml(item,session.depExpanded.has(item.key))).join('');
    wireCopyTags(fragment);more.before(...fragment.children);
    const remaining=list.length-offset-slice.length;
    if(remaining>0){more.dataset.offset=String(offset+slice.length);more.innerHTML=`${ic('chevrons-down')}Show ${Math.min(remaining,DEPENDENCY_CHUNK).toLocaleString()} more of ${remaining.toLocaleString()}`;}
    else more.remove();
  };
  wireCopyTags(body);
}
