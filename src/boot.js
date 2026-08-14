import { $, $$, esc } from './core/text.js'
import { APP_VERSION } from './version.js'
import { S } from './state.js'
import { ic } from './ui/icons.js'
import { closeDrawer, renderAuditResult, renderComparisonResult, renderHierarchyResult, renderRules, renderUpload } from './ui/audit.js'
import { closeUpdateModal, initUpdate } from './update/update.js'

const ruleGuide=()=>'<p>Open the Rules workspace for a searchable, plain-language list of every check currently used by SSM Audit.</p><div class="guide-note">'+ic('info')+'Rules are grouped by governing source and review topic so they are easy to scan.</div><button class="btn primary guide-rules-link" id="guideOpenRules" type="button">'+ic('book-open')+'View audit rules</button>';
const GUIDE_SECTIONS=[
  {id:'start',label:'Getting started',icon:'upload',title:'Run an audit',body:'<p>Select a completed Cx Registry workbook. SSM Audit finds the registry tab automatically, reads it in local browser memory, and evaluates every populated equipment row.</p><ol><li>Choose or drop an .xlsx or .xls workbook.</li><li>Wait for the local audit to complete.</li><li>Start with Blockers, then work through Errors and Warnings.</li></ol>'},
  {id:'checks',label:'What is checked',icon:'check-check',title:'Standards coverage',body:'<p>The auditor checks an existing registry for internal integrity, SSM SOP requirements, and confidence-rated commissioning relationships.</p><ol><li>Parent-child structure, roots, missing parents, self-parenting, and cycles.</li><li>UPN and Discipline crossings, System Name consistency, and metadata alignment.</li><li>Dependencies, electrical and control paths, equipment-role patterns, and organizational headers.</li></ol><div class="guide-note">'+ic('info')+'Current upload-template vocabulary is checked before export in SSManagement, not against an already-uploaded registry.</div>'},
  {id:'hierarchy',label:'SSM hierarchy',icon:'list-tree',title:'Explore one registry as a tree',body:'<p>Open the SSM Hierarchy tab after an audit to review the registry from Building through Discipline, System Name, and Equipment.</p><ol><li>Search tags, descriptions, parents, or dependencies.</li><li>Filter to a Building, Discipline, or System and expand only the branches you need.</li><li>Open an equipment row to inspect its assigned metadata, closest parent, dependencies, source row, and audit findings.</li></ol><div class="guide-note">'+ic('info')+'The tree reflects the registry as uploaded. It does not silently repair relationships that the audit has flagged.</div>'},
  {id:'rules',label:'Rule catalog',icon:'check-check',title:'Rules used by the audit',body:ruleGuide()},
  {id:'findings',label:'Understanding findings',icon:'triangle-alert',title:'Read each finding as evidence',body:'<p>Every finding identifies the rule, source row, observed value, expected result, recommended correction, and a stable fingerprint. Use filters to isolate one severity or category, then open a row for the complete explanation.</p>'},
  {id:'report',label:'Exporting results',icon:'file-down',title:'Create a resolution report',body:'<p>Export report creates an Excel workbook containing the audit summary, every finding, and a rule-level count. The report is generated locally and can be used to assign and track corrections.</p>'},
  {id:'compare',label:'Project comparison',icon:'square-stack',title:'Compare with a completed project',body:'<p>Project Comparison audits the target registry normally, then aligns both registries by UPN for a separate side-by-side review. Building values are excluded so project locations do not create false differences.</p><ol><li>Choose the registry to audit and a completed-project reference.</li><li>Select a UPN to compare equipment nesting, organizational headers, dependencies, and I&amp;C placement.</li><li>Treat differences as review prompts. The completed project is precedent, not an automatic rule.</li></ol><div class="guide-note">'+ic('lock')+'Both workbooks stay in browser memory for this session. The reference is never saved or used to change audit rules.</div>'},
  {id:'privacy',label:'Privacy',icon:'lock',title:'The registry stays on this device',body:'<p>The selected workbook is never uploaded, persisted, or used to learn rules. Refreshing or closing the HTML clears the audit session.</p><div class="guide-note">'+ic('info')+'The only automatic network request is an anonymous check for a newer SSM Audit release when the HTML opens or refreshes.</div>'},
];
let activeGuide='start';

function go(screen){S.screen=screen;if(screen==='audit')renderAuditResult(go);else if(screen==='hierarchy')renderHierarchyResult(go);else if(screen==='compare')renderComparisonResult(go);else if(screen==='rules')renderRules(go);else renderUpload(go);}
function renderGuide(){
  $('#guideNav').innerHTML=GUIDE_SECTIONS.map(section=>`<button type="button" class="${section.id===activeGuide?'active':''}" data-guide="${section.id}">${ic(section.icon)}${esc(section.label)}</button>`).join('');
  const section=GUIDE_SECTIONS.find(item=>item.id===activeGuide)||GUIDE_SECTIONS[0];$('#guideBody').innerHTML=`<div class="guide-copy"><span class="eyebrow">SSM Audit Guide</span><h3>${esc(section.title)}</h3>${section.body}</div>`;
  $$('[data-guide]').forEach(button=>button.onclick=()=>{activeGuide=button.dataset.guide;renderGuide();});
  const openRules=$('#guideOpenRules');if(openRules)openRules.onclick=()=>{closeGuide();S.homeMode='rules';go('rules');};
}
function openGuide(){renderGuide();const modal=$('#guideModal');modal.hidden=false;modal.classList.add('show');modal.setAttribute('aria-hidden','false');}
function closeGuide(){const modal=$('#guideModal');modal.classList.remove('show');modal.hidden=true;modal.setAttribute('aria-hidden','true');}

function init(){
  $('#brandmark').innerHTML=ic('zap');$('#headerRules').innerHTML=ic('check-check')+'<span>Rules</span>';$('#headerRules').onclick=()=>{S.homeMode='rules';go('rules');};$('#headerGuide').innerHTML=ic('book-open')+'<span>Guide</span>';$('#headerGuide').onclick=openGuide;
  $('#versionLink').textContent=`v${APP_VERSION}`;$('#guideClose').classList.add('xbtn');$('#guideClose').innerHTML=ic('x');$('#guideClose').onclick=closeGuide;
  $('#guideModal').onclick=event=>{if(event.target===$('#guideModal'))closeGuide();};document.addEventListener('ssm-audit:guide',openGuide);
  $('#drawerClose').innerHTML=ic('x');$('#drawerClose').onclick=closeDrawer;$('#drawerBack').onclick=event=>{if(event.target===$('#drawerBack'))closeDrawer();};
  document.addEventListener('keydown',event=>{if(event.key!=='Escape')return;if($('#drawerBack').classList.contains('show'))closeDrawer();else if($('#guideModal').classList.contains('show'))closeGuide();else if($('#updateModal').classList.contains('show'))closeUpdateModal();});
  initUpdate();go('upload');
}

init();
