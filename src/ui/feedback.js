import { $ } from '../core/text.js'

let toastTimer=0;
export function toast(message){
  const element=$('#toast');if(!element)return;
  element.textContent=message;element.classList.add('show');clearTimeout(toastTimer);
  toastTimer=setTimeout(()=>element.classList.remove('show'),2800);
}

export function showProgress(title,detail=''){
  $('#lmsg').textContent=title;$('#lsub').textContent=detail;$('#lpct').textContent='0';
  $('#lring').style.strokeDashoffset=String(239);$('#overlay').classList.add('show');
}
export function setProgress(fraction,label){
  const value=Math.max(0,Math.min(1,Number(fraction)||0)),percent=Math.round(value*100);
  $('#lpct').textContent=String(percent);$('#lring').style.strokeDashoffset=String(239*(1-value));
  if(label)$('#lsub').textContent=label;
}
export function hideProgress(){setProgress(1);setTimeout(()=>$('#overlay').classList.remove('show'),140);}

export async function runWithProgress(title,detail,task){
  showProgress(title,detail);
  const checkpoint=()=>new Promise(resolve=>setTimeout(resolve,0));
  try{return await task(checkpoint,setProgress);}finally{hideProgress();}
}

export function copyTagHtml(tag){
  if(!tag)return '<span class="muted">Registry</span>';
  const safe=String(tag).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  return `<button class="copy-tag" type="button" data-copy-tag="${safe}" title="Copy ${safe}">${safe}</button>`;
}
export function wireCopyTags(root=document){
  root.querySelectorAll('[data-copy-tag]').forEach(button=>button.onclick=async event=>{
    event.stopPropagation();
    try{await navigator.clipboard.writeText(button.dataset.copyTag);toast('Tag copied');}
    catch(_){toast('Could not copy tag');}
  });
}
