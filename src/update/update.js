import { $, $$, esc } from '../core/text.js'
import { downloadBlob } from '../core/download.js'
import { APP_VERSION, UPDATE_REPOSITORY } from '../version.js'
import { S } from '../state.js'
import { ic } from '../ui/icons.js'
import { activateFocusTrap, animateOpen, animateClose } from '../ui/feedback.js'

let updateOpener=null;
let updateTrapCleanup=null;

function semver(value){return String(value||'').replace(/^v/i,'').split('.').map(part=>Number(part)||0);}
export function compareVersions(left,right){
  const a=semver(left),b=semver(right),length=Math.max(a.length,b.length);
  for(let index=0;index<length;index++){const diff=(a[index]||0)-(b[index]||0);if(diff)return diff;}
  return 0;
}
export function versionedFilename(version){return `SSM-Audit-v${String(version).replace(/^v/i,'')}.html`;}
export function selectReleaseAsset(release){
  const version=String(release&&release.tag_name||'').replace(/^v/i,''),assets=release&&release.assets||[];
  const exact=`SSM-Audit-v${version}.html`.toLowerCase();
  const asset=assets.find(item=>String(item.name).toLowerCase()===exact);
  return asset?{version,assetName:asset.name,downloadUrl:asset.browser_download_url,releaseNotes:release.body||'',publishedAt:release.published_at||''}:null;
}

/* The Releases repo also carries latest.json (written at release time). It is served
   from raw.githubusercontent.com, which is not subject to the 60-requests-an-hour
   limit that the GitHub API applies per public IP -- the reason a whole site behind
   one NAT address could see the API check fail silently. */
export function latestManifestUrl(repository){return `https://raw.githubusercontent.com/${repository}/main/latest.json`;}
export function selectLatestManifest(manifest){
  const version=String(manifest&&manifest.version||'').replace(/^v/i,'');if(!/^\d+\.\d+\.\d+$/.test(version))return null;
  const assetName=versionedFilename(version);
  const downloadUrl=String(manifest.downloadUrl||'');if(!downloadUrl.toLowerCase().endsWith('/'+assetName.toLowerCase()))return null;
  return {version,assetName,downloadUrl,releaseNotes:String(manifest.releaseNotes||''),publishedAt:String(manifest.publishedAt||'')};
}
export function classifyCheckFailure(status,headers){
  if(status===403||status===429){const remaining=headers&&typeof headers.get==='function'?headers.get('x-ratelimit-remaining'):null;if(remaining===null||remaining==='0'||status===429)return 'rate-limited';}
  return 'unreachable';
}
export function describeUpdateCheck(check,localVersion){
  const state=check&&check.state;
  if(state==='checking')return 'Checking for a newer version\u2026';
  if(state==='update')return `Version ${check.version} is available.`;
  if(state==='current')return `You are on the latest version (v${localVersion}).`;
  if(state==='rate-limited')return 'Could not check just now: GitHub is limiting update checks from this network. Try again in a few minutes.';
  if(state==='unreachable')return 'Could not reach GitHub to check for updates. Check the connection and try again.';
  return 'Not checked yet.';
}

function notesHtml(markdown){
  const lines=String(markdown||'').split(/\r?\n/).map(line=>line.trim()).filter(Boolean);let html='',open=false;
  const close=()=>{if(open){html+='</ul>';open=false;}};
  for(const line of lines){const heading=line.match(/^###\s+(.+)/),bullet=line.match(/^[-*]\s+(.+)/);if(heading){close();html+=`<h3>${esc(heading[1])}</h3>`;}else if(bullet){if(!open){html+='<ul>';open=true;}html+=`<li>${esc(bullet[1])}</li>`;}else{close();html+=`<p>${esc(line)}</p>`;}}
  close();return html;
}
function releaseType(version){const parts=semver(version);return parts[0]>1&&parts[1]===0&&parts[2]===0?'major':parts[1]>0&&parts[2]===0?'feature':'fix';}
function typeLabel(type){return type==='major'?'Major':type==='feature'?'Feature':'Fix';}
function formattedDate(value){const source=String(value||''),date=new Date(/^\d{4}-\d{2}-\d{2}$/.test(source)?source+'T12:00:00':source);return Number.isNaN(date.getTime())?'':date.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});}
function changelogEntries(){
  const entries=[...CHANGELOG];const info=S.updateInfo;
  if(info&&info.version&&!entries.some(entry=>entry.version===info.version))entries.unshift({version:info.version,date:info.publishedAt,type:releaseType(info.version),notes:info.releaseNotes});
  return entries;
}
function checkStatusHtml(){
  const check=S.updateCheck,state=check&&check.state,busy=state==='checking';
  const action=state==='update'?`<button class="btn ghost sm" type="button" id="updateViewUpdate">${ic('download')}View update</button>`:'';
  return `<div class="update-check ${esc(state||'idle')}"><span class="update-check-text">${esc(describeUpdateCheck(check,APP_VERSION))}</span>${action}<button class="btn ghost sm" type="button" id="updateRecheck" ${busy?'disabled':''}>${ic('rotate-ccw')}${busy?'Checking\u2026':'Check again'}</button></div>`;
}
function renderChangelog(){
  return `${checkStatusHtml()}<p class="changelog-intro">Recent SSM Audit changes, newest first.</p><div class="changelog-list">${changelogEntries().map((entry,index)=>`<div class="change-entry ${esc(entry.type)}${index?' closed':''}"><button class="change-head" type="button"><span class="change-chev">${ic('chevron-down')}</span><span class="change-ver">v${esc(entry.version)}</span><span class="change-type">${typeLabel(entry.type)}</span>${index===0?'<span class="change-latest">latest</span>':''}<span class="change-date">${esc(formattedDate(entry.date))}</span></button><div class="change-body">${notesHtml(entry.notes)}</div></div>`).join('')}</div>`;
}
function setTab(tab){
  const hasUpdate=!!S.updateInfo;$('#updateTab').hidden=!hasUpdate;$('#updatePanel').hidden=tab!=='update'||!hasUpdate;$('#changelogPanel').hidden=tab!=='changelog';
  $('#updateTab').classList.toggle('on',tab==='update');$('#changelogTab').classList.toggle('on',tab==='changelog');$('#updateTitle').textContent=tab==='update'&&hasUpdate?'Update Available':'Changelog';
  for(const [id,name] of [['updateTab','update'],['changelogTab','changelog']]){const button=$('#'+id),active=name===tab;if(button){button.setAttribute('aria-selected',String(active));button.tabIndex=active?0:-1;}}
}
export function closeUpdateModal(){const modal=$('#updateModal');if(!modal.classList.contains('show'))return;updateTrapCleanup?.();updateTrapCleanup=null;modal.setAttribute('aria-hidden','true');animateClose(modal);const opener=updateOpener;updateOpener=null;if(opener&&opener!==document.body&&document.contains(opener)&&typeof opener.focus==='function')opener.focus();}
export function openUpdateModal(tab='changelog'){
  $('#changelogPanel').innerHTML=renderChangelog();$$('#changelogPanel .change-head').forEach(button=>button.onclick=()=>button.closest('.change-entry').classList.toggle('closed'));
  const recheck=$('#updateRecheck');if(recheck)recheck.onclick=()=>checkForUpdate({manual:true});
  const view=$('#updateViewUpdate');if(view)view.onclick=()=>{if(S.updateInfo)showUpdate(S.updateInfo);};
  setTab(tab);const modal=$('#updateModal');if(!modal.classList.contains('show'))updateOpener=document.activeElement;animateOpen(modal);modal.setAttribute('aria-hidden','false');updateTrapCleanup?.();updateTrapCleanup=activateFocusTrap(modal,closeUpdateModal);requestAnimationFrame(()=>$('#updateCloseX').focus());
}
function showUpdate(info){
  S.updateInfo=info;$('#updateLocal').textContent=`v${APP_VERSION}`;$('#updateRemote').textContent=`v${info.version}`;
  $('#updateMsg').textContent='A newer offline SSM Audit file is ready. Download it, then use the new version for future audits.';
  const notes=$('#updateNotes'),html=notesHtml(info.releaseNotes);notes.innerHTML=html;notes.hidden=!html;$('#updateStatus').innerHTML='';
  const button=$('#updateDownload');button.disabled=false;button.innerHTML=ic('download')+`Download v${esc(info.version)}`;openUpdateModal('update');
}
async function downloadUpdate(){
  const info=S.updateInfo;if(!info)return;const button=$('#updateDownload'),status=$('#updateStatus');button.disabled=true;button.textContent='Preparing...';status.innerHTML='';
  try{
    const response=await fetch(info.downloadUrl,{cache:'no-store'});if(!response.ok)throw new Error(`HTTP ${response.status}`);
    const blob=await response.blob(),text=await blob.text();
    if(text.length<100000||!text.includes('<title>SSM Audit</title>')||!text.includes(`APP_VERSION='${info.version}'`))throw new Error('The release asset did not pass validation');
    downloadBlob(versionedFilename(info.version),new Blob([text],{type:'text/html'}));status.innerHTML=`<div class="update-status">Download ready: ${esc(versionedFilename(info.version))}</div>`;button.innerHTML=ic('check')+'Downloaded';
  }catch(error){console.error('Update download failed',error);status.innerHTML='<div class="update-error">Download failed. Check the connection and try again.</div>';button.disabled=false;button.innerHTML=ic('download')+`Download v${esc(info.version)}`;}
}
async function fetchLatestRelease(){
  let reason='unreachable';
  try{
    const response=await fetch(`https://api.github.com/repos/${UPDATE_REPOSITORY}/releases/latest`,{headers:{Accept:'application/vnd.github+json'},cache:'no-store'});
    if(response.ok){const info=selectReleaseAsset(await response.json());if(info)return {info};}
    else reason=classifyCheckFailure(response.status,response.headers);
  }catch(error){reason='unreachable';}
  try{
    const response=await fetch(latestManifestUrl(UPDATE_REPOSITORY),{cache:'no-store'});
    if(response.ok){const info=selectLatestManifest(await response.json());if(info)return {info};}
  }catch(error){/* the API reason stands */}
  return {reason};
}
function setUpdateCheck(check){S.updateCheck=check;if($('#changelogPanel')&&!$('#changelogPanel').hidden&&$('#updateModal').classList.contains('show'))openUpdateModal('changelog');}
let updateRetryTimer=0;
export async function checkForUpdate({manual=false}={}){
  if(S.updateCheck&&S.updateCheck.state==='checking')return;
  clearTimeout(updateRetryTimer);setUpdateCheck({state:'checking'});
  const {info,reason}=await fetchLatestRelease();
  if(info){
    const newer=compareVersions(info.version,APP_VERSION)>0;
    S.updateInfo=newer?info:null;setUpdateCheck({state:newer?'update':'current',version:info.version,checkedAt:Date.now()});
    if(newer&&(manual||!S.updateModalDismissed))showUpdate(info);
    return;
  }
  setUpdateCheck({state:reason,checkedAt:Date.now()});
  /* One quiet retry for the automatic check: a laptop that has just woken up often
     has no network for the first second or two. */
  if(!manual&&!S.updateRetried){S.updateRetried=true;updateRetryTimer=setTimeout(()=>checkForUpdate(),6000);}
}
export function initUpdate(){
  $('#updateCloseX').innerHTML=ic('x');$('#updateCloseX').onclick=()=>{S.updateModalDismissed=true;closeUpdateModal();};
  $('#updateSkip').onclick=()=>{S.updateModalDismissed=true;closeUpdateModal();};$('#updateDownload').onclick=downloadUpdate;
  $('#updateTab').onclick=()=>setTab('update');$('#changelogTab').onclick=()=>setTab('changelog');
  $('#updateTabs').onkeydown=event=>{if(!['ArrowLeft','ArrowRight'].includes(event.key))return;event.preventDefault();const next=document.activeElement===$('#updateTab')?$('#changelogTab'):$('#updateTab');if(!next.hidden){setTab(next===$('#updateTab')?'update':'changelog');next.focus();}};
  $('#versionLink').onclick=()=>{S.updateModalDismissed=false;openUpdateModal('changelog');};
  checkForUpdate();
}
