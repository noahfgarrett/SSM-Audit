import { $, $$, esc } from './core/text.js'
import { APP_VERSION } from './version.js'
import { S } from './state.js'
import { ic } from './ui/icons.js'
import { activateFocusTrap } from './ui/feedback.js'
import { closeDrawer, renderAuditResult, renderComparisonResult, renderHierarchyResult, renderRules, renderSideNav, renderUpload } from './ui/audit.js'
import { closeUpdateModal, initUpdate } from './update/update.js'

const ruleGuide=()=>'<p>Open Rules in the menu on the left for a searchable, plain-language list of every check the audit runs.</p><div class="guide-note">'+ic('info')+'Checks are grouped by where they come from and what they cover, so they are easy to scan. Once a registry is loaded, each check also shows how many times it fired.</div><button class="btn primary guide-rules-link" id="guideOpenRules" type="button">'+ic('book-open')+'View the rules</button>';
const GUIDE_SECTIONS=[
  {id:'start',label:'Getting started',icon:'upload',title:'Run an audit',body:'<p>Select a completed Cx Registry workbook. SSM Audit finds the registry tab automatically, reads it in local browser memory, and evaluates every populated equipment row.</p><ol><li>Choose or drop an .xlsx or .xls workbook.</li><li>Wait for the local audit to complete.</li><li>Start with Blockers, then work through Errors and Warnings.</li></ol>'},
  {id:'checks',label:'What is checked',icon:'check-check',title:'Standards coverage',body:'<p>The auditor checks an existing registry for internal integrity, SSM SOP requirements, and confidence-rated commissioning relationships.</p><ol><li>Parent-child structure, roots, missing parents, self-parenting, and cycles.</li><li>UPN and Discipline crossings, System Name consistency, and metadata alignment.</li><li>Dependencies, electrical and control paths, equipment-role patterns, and organizational headers.</li></ol><div class="guide-note">'+ic('info')+'Current upload-template vocabulary is checked before export in SSManagement, not against an already-uploaded registry.</div>'},
  {id:'hierarchy',label:'SSM hierarchy',icon:'list-tree',title:'Explore one registry as a tree',body:'<p>Choose SSM hierarchy in the menu after an audit to walk the registry from Building down through Discipline, System Name, and Equipment.</p><ol><li>Search tags, descriptions, parents, or dependencies — matches are highlighted and opened for you.</li><li>Filter to a Building, Discipline, or System, or switch on Findings only to hide the branches that are already clean.</li><li>Open an equipment row to see its metadata, closest parent, dependencies, source row, and findings.</li></ol><div class="guide-note">'+ic('info')+'The tree shows the registry exactly as uploaded. It never quietly repairs a relationship the audit has flagged.</div>'},
  {id:'rules',label:'Rule catalog',icon:'check-check',title:'Rules used by the audit',body:ruleGuide()},
  {id:'findings',label:'Understanding findings',icon:'triangle-alert',title:'Read each finding as evidence',body:'<p>The headline of every finding is what was actually seen. Open a row and you also get the rule behind it, what was found against what was expected, what to do about it, and the registry row it came from.</p><ol><li>Click a severity chip above the table to show or hide that level.</li><li>Use Filters for source, topic, or a single check, and Group by to gather findings under one check or one L2 milestone.</li><li>Arrow keys move down the list, Enter opens a finding, and Escape closes it.</li></ol>'},
  {id:'report',label:'Exporting results',icon:'file-down',title:'Create a resolution report',body:'<p>Export report creates an Excel workbook containing the audit summary, every finding, and a rule-level count. The report is generated locally and can be used to assign and track corrections.</p>'},
  {id:'compare',label:'Project comparison',icon:'square-stack',title:'Compare with a completed project',body:'<p>Project Comparison audits the target registry normally, then aligns both registries by UPN for a separate side-by-side review. Building values are excluded so project locations do not create false differences.</p><ol><li>Choose the registry to audit and a completed-project reference.</li><li>Select a UPN to compare equipment nesting, organizational headers, dependencies, and I&amp;C placement.</li><li>Treat differences as review prompts. The completed project is precedent, not an automatic rule.</li></ol><div class="guide-note">'+ic('lock')+'Both workbooks stay in browser memory for this session. The reference is never saved or used to change audit rules.</div>'},
  {id:'privacy',label:'Privacy',icon:'lock',title:'The registry stays on this device',body:'<p>The selected workbook is never uploaded, persisted, or used to learn rules. Refreshing or closing the HTML clears the audit session.</p><div class="guide-note">'+ic('info')+'The only automatic network request is an anonymous check for a newer SSM Audit release when the HTML opens or refreshes.</div>'},
];
let activeGuide='start';
let guideOpener=null;
let guideTrapCleanup=null;

function go(screen){S.screen=screen;if(screen==='audit')renderAuditResult(go);else if(screen==='hierarchy')renderHierarchyResult(go);else if(screen==='compare')renderComparisonResult(go);else if(screen==='rules')renderRules(go);else renderUpload(go);renderSideNav(go);}
function renderGuide(){
  $('#guideNav').innerHTML=GUIDE_SECTIONS.map(section=>`<button type="button" class="${section.id===activeGuide?'active':''}" data-guide="${section.id}">${ic(section.icon)}${esc(section.label)}</button>`).join('');
  const section=GUIDE_SECTIONS.find(item=>item.id===activeGuide)||GUIDE_SECTIONS[0];$('#guideBody').innerHTML=`<div class="guide-copy"><span class="eyebrow">SSM Audit Guide</span><h3>${esc(section.title)}</h3>${section.body}</div>`;
  $$('[data-guide]').forEach(button=>button.onclick=()=>{activeGuide=button.dataset.guide;renderGuide();});
  const openRules=$('#guideOpenRules');if(openRules)openRules.onclick=()=>{closeGuide();S.homeMode='rules';go('rules');};
}
function openGuide(){guideOpener=document.activeElement;renderGuide();const modal=$('#guideModal');modal.hidden=false;modal.classList.add('show');modal.setAttribute('aria-hidden','false');guideTrapCleanup?.();guideTrapCleanup=activateFocusTrap(modal,closeGuide);requestAnimationFrame(()=>$('#guideClose').focus());}
function closeGuide(){const modal=$('#guideModal');if(!modal.classList.contains('show'))return;guideTrapCleanup?.();guideTrapCleanup=null;modal.classList.remove('show');modal.hidden=true;modal.setAttribute('aria-hidden','true');const opener=guideOpener;guideOpener=null;if(opener&&document.contains(opener)&&typeof opener.focus==='function')opener.focus();}

function init(){
  $('#brandmark').innerHTML=ic('zap');$('#headerGuide').innerHTML=ic('book-open')+'<span>Guide</span>';$('#headerGuide').onclick=openGuide;
  $('#versionLink').textContent=`v${APP_VERSION}`;$('#guideClose').classList.add('xbtn');$('#guideClose').innerHTML=ic('x');$('#guideClose').onclick=closeGuide;
  $('#guideModal').onclick=event=>{if(event.target===$('#guideModal'))closeGuide();};document.addEventListener('ssm-audit:guide',openGuide);
  document.addEventListener('ssm-audit:navigate',event=>go(event.detail&&event.detail.screen||'upload'));
  $('#drawerClose').innerHTML=ic('x');$('#drawerClose').onclick=closeDrawer;$('#drawerBack').onclick=event=>{if(event.target===$('#drawerBack'))closeDrawer();};
  document.addEventListener('keydown',event=>{if(event.key!=='Escape')return;if($('#drawerBack').classList.contains('show'))closeDrawer();else if($('#guideModal').classList.contains('show'))closeGuide();else if($('#updateModal').classList.contains('show'))closeUpdateModal();});
  initUpdate();go('upload');
}

init();
