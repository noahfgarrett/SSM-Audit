import { $, $$, esc } from './core/text.js'
import { APP_VERSION } from './version.js'
import { S } from './state.js'
import { SSM_AUDIT_RULES, SSM_AUDIT_SOURCES } from './audit/engine.js'
import { ic } from './ui/icons.js'
import { closeDrawer, renderAuditResult, renderUpload } from './ui/audit.js'
import { closeUpdateModal, initUpdate } from './update/update.js'

const ruleGuide=()=>SSM_AUDIT_SOURCES.map(source=>`<h4>${esc(source.label)}</h4><ul>${Object.values(SSM_AUDIT_RULES).filter(rule=>rule.source===source.id).map(rule=>`<li class="${rule.enabled?'':'rule-disabled'}"><b>${esc(rule.title)}.</b> ${esc(rule.statement)} <span class="rule-confidence">${rule.enabled?esc(rule.confidence):'Disabled'}</span>${rule.disabledReason?`<small>${esc(rule.disabledReason)}</small>`:''}</li>`).join('')}</ul>`).join('');
const GUIDE_SECTIONS=[
  {id:'start',label:'Getting started',icon:'upload',title:'Run an audit',body:'<p>Select a completed Cx Registry workbook. SSM Audit finds the registry tab automatically, reads it in local browser memory, and evaluates every populated equipment row.</p><ol><li>Choose or drop an .xlsx or .xls workbook.</li><li>Wait for the local audit to complete.</li><li>Start with Blockers, then work through Errors and Warnings.</li></ol>'},
  {id:'checks',label:'What is checked',icon:'check-check',title:'Standards coverage',body:'<p>The auditor checks an existing registry for internal integrity, SSM SOP requirements, and confidence-rated commissioning relationships.</p><ol><li>Parent-child structure, roots, missing parents, self-parenting, and cycles.</li><li>UPN and Discipline crossings, System Name consistency, and milestone completeness.</li><li>Dependencies, electrical and control paths, equipment-role patterns, and organizational headers.</li></ol><div class="guide-note">'+ic('info')+'Current upload-template vocabulary is checked before export in SSManagement, not against an already-uploaded registry.</div>'},
  {id:'rules',label:'Rule catalog',icon:'check-check',title:'Rules used by the audit',body:ruleGuide()},
  {id:'findings',label:'Understanding findings',icon:'triangle-alert',title:'Read each finding as evidence',body:'<p>Every finding identifies the rule, source row, observed value, expected result, recommended correction, and a stable fingerprint. Use filters to isolate one severity or category, then open a row for the complete explanation.</p>'},
  {id:'report',label:'Exporting results',icon:'file-down',title:'Create a resolution report',body:'<p>Export report creates an Excel workbook containing the audit summary, every finding, and a rule-level count. The report is generated locally and can be used to assign and track corrections.</p>'},
  {id:'privacy',label:'Privacy',icon:'lock',title:'The registry stays on this device',body:'<p>The selected workbook is never uploaded, persisted, or used to learn rules. Refreshing or closing the HTML clears the audit session.</p><div class="guide-note">'+ic('info')+'The only automatic network request is an anonymous check for a newer SSM Audit release when the HTML opens or refreshes.</div>'},
];
let activeGuide='start';

function go(screen){S.screen=screen;if(screen==='audit')renderAuditResult(go);else renderUpload(go);}
function renderGuide(){
  $('#guideNav').innerHTML=GUIDE_SECTIONS.map(section=>`<button type="button" class="${section.id===activeGuide?'active':''}" data-guide="${section.id}">${ic(section.icon)}${esc(section.label)}</button>`).join('');
  const section=GUIDE_SECTIONS.find(item=>item.id===activeGuide)||GUIDE_SECTIONS[0];$('#guideBody').innerHTML=`<div class="guide-copy"><span class="eyebrow">SSM Audit Guide</span><h3>${esc(section.title)}</h3>${section.body}</div>`;
  $$('[data-guide]').forEach(button=>button.onclick=()=>{activeGuide=button.dataset.guide;renderGuide();});
}
function openGuide(){renderGuide();const modal=$('#guideModal');modal.hidden=false;modal.classList.add('show');modal.setAttribute('aria-hidden','false');}
function closeGuide(){const modal=$('#guideModal');modal.classList.remove('show');modal.hidden=true;modal.setAttribute('aria-hidden','true');}

function init(){
  $('#brandmark').innerHTML=ic('zap');$('#headerGuide').innerHTML=ic('book-open')+'<span>Guide</span>';$('#headerGuide').onclick=openGuide;
  $('#versionLink').textContent=`v${APP_VERSION}`;$('#guideClose').classList.add('xbtn');$('#guideClose').innerHTML=ic('x');$('#guideClose').onclick=closeGuide;
  $('#guideModal').onclick=event=>{if(event.target===$('#guideModal'))closeGuide();};document.addEventListener('ssm-audit:guide',openGuide);
  $('#drawerClose').innerHTML=ic('x');$('#drawerClose').onclick=closeDrawer;$('#drawerBack').onclick=event=>{if(event.target===$('#drawerBack'))closeDrawer();};
  document.addEventListener('keydown',event=>{if(event.key!=='Escape')return;if($('#drawerBack').classList.contains('show'))closeDrawer();else if($('#guideModal').classList.contains('show'))closeGuide();else if($('#updateModal').classList.contains('show'))closeUpdateModal();});
  initUpdate();go('upload');
}

init();
