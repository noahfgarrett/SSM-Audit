import { $ } from '../core/text.js'

/* ------------------------------------------------------------------ motion */
/* One shared open/close cycle for every overlay in the app — drawer, filter
   sheet, guide, changelog, and the progress overlay. The element is displayed
   by removing `hidden`, painted once in its closed state, then given `.show`
   so the CSS transition has two frames to run between. Closing reverses it via
   `.is-closing` and only hides the element once that transition ends, with a
   timer fallback so a dropped `transitionend` can never leave an overlay
   stuck on screen. */
export const MOTION_EXIT_MS=140;
const motionPending=new WeakMap();

export function motionReduced(){
  return typeof window!=='undefined'&&typeof window.matchMedia==='function'&&window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}
function motionSettle(element){
  const pending=motionPending.get(element);if(!pending)return;
  element.removeEventListener('transitionend',pending.onEnd);clearTimeout(pending.timer);motionPending.delete(element);
}
export function animateOpen(element){
  if(!element)return;
  motionSettle(element);element.classList.remove('is-closing');element.hidden=false;
  /* Forces the closed frame to paint; without it the element goes straight from
     display:none to its open state and no transition runs. */
  void element.offsetWidth;
  element.classList.add('show');
}
export function animateClose(element,done){
  if(!element)return;
  motionSettle(element);
  const finish=()=>{motionSettle(element);element.classList.remove('is-closing');element.hidden=true;done?.();};
  if(!element.classList.contains('show')){finish();return;}
  element.classList.remove('show');
  if(motionReduced()){finish();return;}
  element.classList.add('is-closing');
  const onEnd=event=>{if(event.target===element)finish();};
  element.addEventListener('transitionend',onEnd);
  motionPending.set(element,{onEnd,timer:setTimeout(finish,MOTION_EXIT_MS+140)});
}

let toastTimer=0;
export function toast(message){
  const element=$('#toast');if(!element)return;
  element.textContent=message;element.classList.add('show');clearTimeout(toastTimer);
  toastTimer=setTimeout(()=>element.classList.remove('show'),2800);
}

export function showProgress(title,detail=''){
  $('#lmsg').textContent=title;$('#lsub').textContent=detail;$('#lpct').textContent='0';
  $('#lring').style.strokeDashoffset=String(239);animateOpen($('#overlay'));
}
export function setProgress(fraction,label){
  const value=Math.max(0,Math.min(1,Number(fraction)||0)),percent=Math.round(value*100);
  $('#lpct').textContent=String(percent);$('#lring').style.strokeDashoffset=String(239*(1-value));
  $('.ring')?.setAttribute('aria-valuenow',String(percent));
  if(label)$('#lsub').textContent=label;
}
export function hideProgress(){setProgress(1);setTimeout(()=>animateClose($('#overlay')),140);}

export async function runWithProgress(title,detail,task){
  showProgress(title,detail);
  const checkpoint=()=>new Promise(resolve=>setTimeout(resolve,0));
  try{return await task(checkpoint,setProgress);}finally{hideProgress();}
}

/* `labelHtml` lets callers pass already-escaped markup (search highlighting) for the
   visible label. The copied value always comes from the raw tag. */
export function copyTagHtml(tag,labelHtml){
  if(!tag)return '<span class="muted">Registry</span>';
  const safe=String(tag).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  return `<button class="copy-tag" type="button" data-copy-tag="${safe}" title="Copy ${safe}">${labelHtml||safe}</button>`;
}
export async function copyText(value){
  const text=String(value||'');if(!text)return false;
  try{
    if(navigator.clipboard&&typeof navigator.clipboard.writeText==='function'){await navigator.clipboard.writeText(text);return true;}
  }catch(_){/* Direct-from-disk browsers may block the async Clipboard API. */}
  const input=document.createElement('textarea'),previous=document.activeElement;
  input.value=text;input.setAttribute('readonly','');input.setAttribute('aria-hidden','true');input.style.cssText='position:fixed;left:-9999px;top:0;opacity:0';document.body.appendChild(input);
  input.focus();input.select();input.setSelectionRange(0,text.length);
  let copied=false;try{copied=!!document.execCommand&&document.execCommand('copy');}catch(_){}input.remove();
  if(previous&&previous!==document.body&&typeof previous.focus==='function')try{previous.focus({preventScroll:true});}catch(_){previous.focus();}
  return copied;
}
export function activateFocusTrap(container,onEscape){
  if(!container)return ()=>{};
  const focusable=()=>[...container.querySelectorAll('a[href],button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex="-1"])')].filter(element=>!element.hidden&&element.getClientRects().length);
  const handler=event=>{if(event.key==='Escape'){event.preventDefault();onEscape?.();return;}if(event.key!=='Tab')return;const items=focusable();if(!items.length){event.preventDefault();container.focus?.();return;}const first=items[0],last=items[items.length-1];if(event.shiftKey&&document.activeElement===first){event.preventDefault();last.focus();}else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first.focus();}};
  container.addEventListener('keydown',handler);return ()=>container.removeEventListener('keydown',handler);
}
export function wireCopyTags(root=document){
  root.querySelectorAll('[data-copy-tag]').forEach(button=>button.onclick=async event=>{
    event.stopPropagation();
    const copied=await copyText(button.dataset.copyTag);toast(copied?'Tag copied':'Could not copy tag');
  });
}
