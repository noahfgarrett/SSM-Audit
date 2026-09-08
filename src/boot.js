import { $, $$, esc } from './core/text.js'
import { APP_VERSION } from './version.js'
import { S } from './state.js'
import { ic } from './ui/icons.js'
import { activateFocusTrap, animateClose, animateOpen, motionReduced } from './ui/feedback.js'
import { closeDrawer, renderAuditResult, renderComparisonResult, renderDashboard, renderCompletedEquipment, renderHierarchyResult, renderModifications, renderRules, renderSideNav, renderUpload, closeRuleExample, closeExportOptions, loadRulePreferences } from './ui/audit.js'
import { closeUpdateModal, initUpdate } from './update/update.js'

import { GUIDE_SECTIONS } from './ui/guide-content.js'
let activeGuide='start';
let guideOpener=null;
let guideTrapCleanup=null;

function go(screen){const changed=S.screen!==screen;S.screen=screen;if(screen==='dashboard')renderDashboard(go);else if(screen==='audit')renderAuditResult(go);else if(screen==='modify')renderModifications(go);else if(screen==='completed')renderCompletedEquipment(go);else if(screen==='hierarchy')renderHierarchyResult(go);else if(screen==='compare')renderComparisonResult(go);else if(screen==='rules')renderRules(go);else renderUpload(go);renderSideNav(go);if(changed)markViewEnter();}
/* A screen change slides the new content in; re-renders of the same screen (filters, sorting) stay still. */
function markViewEnter(){const view=$('#view');if(!view||motionReduced())return;view.classList.remove('view-enter');void view.offsetWidth;view.classList.add('view-enter');view.addEventListener('animationend',()=>view.classList.remove('view-enter'),{once:true});}
function renderGuide(){
  $('#guideNav').innerHTML=GUIDE_SECTIONS.map(section=>`<button type="button" class="${section.id===activeGuide?'active':''}" data-guide="${section.id}">${ic(section.icon)}${esc(section.label)}</button>`).join('');
  const section=GUIDE_SECTIONS.find(item=>item.id===activeGuide)||GUIDE_SECTIONS[0];$('#guideBody').innerHTML=`<div class="guide-copy"><span class="eyebrow">SSM Audit Guide</span><h3>${esc(section.title)}</h3>${section.body}</div>`;
  const main=$('#guideBody').parentElement;if(main)main.scrollTop=0;
  $$('[data-guide]').forEach(button=>button.onclick=()=>{activeGuide=button.dataset.guide;renderGuide();});
  const openRules=$('#guideOpenRules');if(openRules)openRules.onclick=()=>{closeGuide();S.homeMode='rules';go('rules');};
}
function openGuide(){guideOpener=document.activeElement;renderGuide();const modal=$('#guideModal');animateOpen(modal);modal.setAttribute('aria-hidden','false');guideTrapCleanup?.();guideTrapCleanup=activateFocusTrap(modal,closeGuide);requestAnimationFrame(()=>$('#guideClose').focus());}
function closeGuide(){const modal=$('#guideModal');if(!modal.classList.contains('show'))return;guideTrapCleanup?.();guideTrapCleanup=null;modal.setAttribute('aria-hidden','true');animateClose(modal);const opener=guideOpener;guideOpener=null;if(opener&&document.contains(opener)&&typeof opener.focus==='function')opener.focus();}

function init(){
  $('#brandmark').innerHTML=ic('zap');$('#headerGuide').innerHTML=ic('book-open')+'<span>Guide</span>';$('#headerGuide').onclick=openGuide;
  $('#versionLink').textContent=`v${APP_VERSION}`;$('#guideClose').classList.add('xbtn');$('#guideClose').innerHTML=ic('x');$('#guideClose').onclick=closeGuide;
  $('#guideModal').onclick=event=>{if(event.target===$('#guideModal'))closeGuide();};$('#exampleClose').innerHTML=ic('x');$('#exportModalClose').innerHTML=ic('x');document.addEventListener('ssm-audit:guide',openGuide);
  document.addEventListener('ssm-audit:navigate',event=>go(event.detail&&event.detail.screen||'upload'));
  $('#drawerClose').innerHTML=ic('x');$('#drawerClose').onclick=closeDrawer;$('#drawerBack').onclick=event=>{if(event.target===$('#drawerBack'))closeDrawer();};
  document.addEventListener('keydown',event=>{if(event.key!=='Escape')return;if($('#drawerBack').classList.contains('show'))closeDrawer();else if($('#exportModal').classList.contains('show'))closeExportOptions();else if($('#exampleModal').classList.contains('show'))closeRuleExample();else if($('#guideModal').classList.contains('show'))closeGuide();else if($('#updateModal').classList.contains('show'))closeUpdateModal();});
  /* Narrow screens: the fixed version badge floats over content, so it moves
     into the topbar actions and back again as the window crosses 900px. */
  const badgeHome=$('#versionLink').parentElement,mq=window.matchMedia('(max-width:900px)');
  const placeBadge=()=>{const badge=$('#versionLink'),actions=$('.top-actions');if(!badge||!actions)return;if(mq.matches){if(badge.parentElement!==actions)actions.insertBefore(badge,actions.firstChild);}else if(badge.parentElement!==badgeHome)badgeHome.appendChild(badge);};
  placeBadge();if(typeof mq.addEventListener==='function')mq.addEventListener('change',placeBadge);
  loadRulePreferences();initUpdate();go('upload');
}

init();
