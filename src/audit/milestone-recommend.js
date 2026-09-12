import { clean } from '../core/text.js'
import { auditNormId } from './model.js'

/* ---- Milestone recommendation ----
   Recommends L2 / L1 milestones for rows that have none, from evidence in a
   fixed order. Measured on real registries: milestones live on the hierarchy
   branch, so the closest-parent chain (95%) and the rest of the branch (99%
   when it agrees) come first. Milestone NUMBERS carry across sites while
   wording and prefixes do not, so completed registries from other sites are
   consulted by number: a counterpart with the same tag core on the same UPN,
   then the usual number for that UPN / building type / discipline. Rows
   without evidence are surfaced with candidates, never guessed.
   Shared by SSM Audit and SSManagement; keep both copies identical. */

export const MILESTONE_LAYER_CONFIDENCE=Object.freeze({
  'parent-chain':.95,'branch-strong':.99,'branch':.9,'counterpart-electrical':.95,'counterpart':.7,
  'schedule-upn':.9,'pattern-high':.83,'pattern-good':.78,'pattern':.72,
});
const BUILDING_TYPES=[['FAB',/\bFAB\b|^F\d/i],['CUB',/\bCUB|\bBC\d|\bBCP/i],['PSSS',/PSSS/i],['WT',/\bWT\d|\bPW\d/i],['PB',/\bPB\d/i],['GIS',/\bGIS\b/i],['EYARD',/EYARD|EGEN|\bSES\b|POWERSTATION/i],['BGY',/\bBGY|\bBG\d|BGC/i],['ASU',/\bASU/i],['SS',/\bSS\d|\bSR\d/i],['SITE',/\bSITE\b/i]];
const SITE_PREFIX=/^(?:(?:LC\d+[A-Z]?|OC\d+|F\d{2}(?:\.\d)?|EA|SP|AZ|WT\d|BC\d|PB\d|BG\d|SS\d|PSSS|FAB|CUB\d?)(?:\s*\([^)]*\))?[-_ ]+)+/i;
const L2_NUMBER=/L2[-_ ]?M\d[-_ ]?(\d+(?:[._]\d+)?)/i,L1_NUMBER=/L1[-_ ]?M\d[-_ ]?(\d+(?:[._]\d+)?)/i;

export function milestoneBuildingType(building){
  const value=clean(building);if(!value)return '';
  for(const [type,pattern] of BUILDING_TYPES)if(pattern.test(value))return type;
  return 'OTHER';
}
export function milestoneTagCore(tag){return auditNormId(tag).replace(SITE_PREFIX,'').replace(/[\s_]+/g,'-');}
export function milestoneL2Number(text){const match=clean(text).match(L2_NUMBER);return match?match[1].replace('_','.'):'';}
export function milestoneL1Number(text){const match=clean(text).match(L1_NUMBER);return match?match[1].replace('_','.'):'';}
function isElectrical(row){return /ELEC/i.test(row&&row.discipline||'');}
function patternKey(row){return [auditNormId(row.upn),milestoneBuildingType(row.building),auditNormId(row.discipline)].join('|');}
function count(map,key){map.set(key,(map.get(key)||0)+1);}
function best(map){let top=null,total=0;for(const [key,n] of map){total+=n;if(!top||n>top.n)top={key,n};}return top?{...top,share:top.n/total,total}:null;}

/* A completed registry (any site) distilled to what recommendations need. */
export function buildMilestoneReference(rows,site=''){
  const byNumber=new Map(),coreIndex=new Map(),patternIndex=new Map();let labelled=0;
  for(const row of rows||[]){
    const number=milestoneL2Number(row&&row.milestone);if(!number)continue;labelled++;
    const entry=byNumber.get(number)||{label:clean(row.milestone),l1Label:clean(row.milestoneParent),count:0};entry.count++;byNumber.set(number,entry);
    const core=milestoneTagCore(row.equipmentId);
    if(core){const key=core+'|'+auditNormId(row.upn);const votes=coreIndex.get(key)||new Map();count(votes,number);coreIndex.set(key,votes);}
    const votes=patternIndex.get(patternKey(row))||new Map();count(votes,number);patternIndex.set(patternKey(row),votes);
  }
  return {site:clean(site),count:(rows||[]).length,labelled,byNumber,coreIndex,patternIndex};
}

/* Schedule definitions from a milestone register (id/title/parentId/parentTitle
   entries) -- the vocabulary a recommendation must resolve to. */
export function milestoneScheduleFromRegister(entries){
  const schedule=[];
  for(const entry of entries||[]){
    const code=clean(entry&&entry.id),number=milestoneL2Number(code)||milestoneL2Number(entry&&entry.title);if(!code||!number)continue;
    const label=clean(entry.milestone)||[code,clean(entry.title)].filter(Boolean).join(' '),l1Label=clean(entry.milestoneParent)||[clean(entry.parentId),clean(entry.parentTitle)].filter(Boolean).join(' ');
    const upns=[...`${code} ${clean(entry.title)}`.matchAll(/UPN\s*([\d ,&\/]+)/gi)].flatMap(match=>match[1].split(/[ ,&\/]+/)).filter(value=>/^\d{3}$/.test(value));
    schedule.push({code,number,label,l1Label,upns});
  }
  return schedule;
}

function localVocabulary(rows,schedule){
  const byNumber=new Map();
  for(const definition of schedule||[])if(!byNumber.has(definition.number))byNumber.set(definition.number,{code:definition.code,label:definition.label,l1Label:definition.l1Label});
  const l1Votes=new Map();
  for(const row of rows||[]){
    const number=milestoneL2Number(row&&row.milestone);if(!number)continue;
    if(!byNumber.has(number))byNumber.set(number,{code:clean(row.milestone),label:clean(row.milestone),l1Label:''});
    const l1=clean(row.milestoneParent);if(l1){const votes=l1Votes.get(number)||new Map();count(votes,l1);l1Votes.set(number,votes);}
  }
  for(const [number,votes] of l1Votes){const entry=byNumber.get(number);if(entry&&!entry.l1Label)entry.l1Label=best(votes).key;}
  const scheduleUpns=new Map();
  for(const definition of schedule||[])for(const upn of definition.upns){const list=scheduleUpns.get(upn)||[];if(!list.includes(definition.number))list.push(definition.number);scheduleUpns.set(upn,list);}
  return {byNumber,scheduleUpns};
}

function labelFor(vocabulary,number,references){
  const local=vocabulary.byNumber.get(number);if(local)return {code:local.code,label:local.label,l1Label:local.l1Label,local:true};
  for(const reference of references||[]){const entry=reference.byNumber.get(number);if(entry)return {code:'',label:`(no local milestone carries number ${number}; ${reference.site||'reference'} calls it "${entry.label}")`,l1Label:'',local:false};}
  return {code:'',label:`(number ${number} is unknown here)`,l1Label:'',local:false};
}

/* rows: the registry being worked (model row shape). schedule: definitions from
   milestoneScheduleFromRegister. references: buildMilestoneReference results.
   Returns Map(normalized tag -> recommendation) for rows without an L2. */
export function recommendMilestones(rows,{schedule=[],references=[],confidence=MILESTONE_LAYER_CONFIDENCE,holdOut=false}={}){
  const list=rows||[],byTag=new Map();
  for(const row of list){const id=auditNormId(row&&row.equipmentId);if(id&&!byTag.has(id))byTag.set(id,row);}
  const rootOf=row=>{let current=row,guard=0;while(guard++<40){const parent=byTag.get(auditNormId(current.closestParent));if(!parent||parent===current)return current;current=parent;}return current;};
  const rootKey=row=>auditNormId(rootOf(row).equipmentId)||('SYS:'+auditNormId(row.systemName));
  const branchVotes=new Map();
  for(const row of list){const number=milestoneL2Number(row.milestone);if(!number)continue;const key=rootKey(row);const votes=branchVotes.get(key)||new Map();count(votes,number);branchVotes.set(key,votes);}
  const vocabulary=localVocabulary(list,schedule),out=new Map();
  const decide=row=>{
    const truth=holdOut?milestoneL2Number(row.milestone):'';
    let parent=byTag.get(auditNormId(row.closestParent)),guard=0;
    while(parent&&guard++<12){const number=milestoneL2Number(parent.milestone);
      if(number)return {layer:'parent-chain',number,l2Label:clean(parent.milestone),l1Label:clean(parent.milestoneParent),confidence:confidence['parent-chain'],evidence:`Closest-parent chain reaches ${parent.equipmentId}, which carries this milestone`};
      parent=byTag.get(auditNormId(parent.closestParent));}
    const votes=new Map(branchVotes.get(rootKey(row))||[]);
    if(truth){votes.set(truth,(votes.get(truth)||0)-1);if(votes.get(truth)<=0)votes.delete(truth);}
    const branch=best(votes);
    if(branch){const strong=branch.share>=.8&&branch.n>=3,named=labelFor(vocabulary,branch.key,references);
      return {layer:strong?'branch-strong':'branch',number:branch.key,l2Label:named.label,l1Label:named.l1Label,confidence:strong?confidence['branch-strong']:Math.round(confidence.branch*branch.share*100)/100,evidence:`${branch.n} of ${branch.total} rows under the same branch root carry this milestone`};}
    const core=milestoneTagCore(row.equipmentId),coreKey=core+'|'+auditNormId(row.upn),counterVotes=new Map();let counterSite='';
    if(core)for(const reference of references){const hits=reference.coreIndex.get(coreKey);if(!hits)continue;for(const [number,n] of hits){count(counterVotes,number);if(!counterSite)counterSite=reference.site;}}
    const counterpart=best(counterVotes);
    if(counterpart){const named=labelFor(vocabulary,counterpart.key,references),electrical=isElectrical(row);
      return {layer:electrical?'counterpart-electrical':'counterpart',number:counterpart.key,l2Label:named.label,l1Label:named.l1Label,confidence:confidence[electrical?'counterpart-electrical':'counterpart'],evidence:`The same equipment tag on UPN ${clean(row.upn)} at ${counterSite||'a reference site'} carries milestone number ${counterpart.key}`,local:named.local};}
    const declared=vocabulary.scheduleUpns.get(auditNormId(row.upn))||[];
    const patternVotes=new Map();for(const reference of references){const hits=reference.patternIndex.get(patternKey(row));if(hits)for(const [number,n] of hits)patternVotes.set(number,(patternVotes.get(number)||0)+n);}
    const pattern=best(patternVotes);
    if(declared.length===1){const named=labelFor(vocabulary,declared[0],references);
      return {layer:'schedule-upn',number:declared[0],l2Label:named.label,l1Label:named.l1Label,confidence:confidence['schedule-upn'],evidence:`The schedule names UPN ${clean(row.upn)} in only this milestone`,candidates:pattern?candidatesFrom(patternVotes,vocabulary,references):[]};}
    if(pattern&&pattern.share>=.6){const layer=pattern.share>=.9&&references.length>1?'pattern-high':pattern.share>=.9?'pattern-good':'pattern',named=labelFor(vocabulary,pattern.key,references);
      return {layer,number:pattern.key,l2Label:named.label,l1Label:named.l1Label,confidence:confidence[layer],evidence:`${Math.round(pattern.share*100)}% of ${pattern.total} reference rows on UPN ${clean(row.upn)} in ${milestoneBuildingType(row.building)||'this building type'} ${clean(row.discipline)} carry milestone number ${pattern.key}`,candidates:candidatesFrom(patternVotes,vocabulary,references),local:named.local};}
    const candidates=[...new Set([...declared,...patternVotes.keys()])].map(number=>({number,...labelFor(vocabulary,number,references),support:patternVotes.get(number)||0,declared:declared.includes(number)}));
    return candidates.length?{layer:'candidates',number:'',l2Label:'',l1Label:'',confidence:0,evidence:`${candidates.length} candidate milestone${candidates.length===1?'':'s'} — an engineer decides`,candidates}:{layer:'none',number:'',l2Label:'',l1Label:'',confidence:0,evidence:'No parent, branch, counterpart, schedule, or reference evidence',candidates:[]};
  };
  let assigned=0,candidates=0,none=0;
  for(const row of list){
    const id=auditNormId(row&&row.equipmentId);if(!id||out.has(id))continue;
    if(!holdOut&&milestoneL2Number(row.milestone))continue;
    if(holdOut&&!milestoneL2Number(row.milestone))continue;
    const recommendation=decide(row);out.set(id,{equipmentId:row.equipmentId,...recommendation});
    if(recommendation.number)assigned++;else if(recommendation.layer==='candidates')candidates++;else none++;
  }
  return {byId:out,summary:{rows:out.size,assigned,candidates,none}};
}

/* Replays each layer against rows that already carry a milestone, so the
   confidence a user sees is a measured hit rate on their own data. */
export function calibrateMilestoneLayers(rows,options={}){
  const {byId}=recommendMilestones(rows,{...options,holdOut:true}),truthById=new Map();
  for(const row of rows||[]){const id=auditNormId(row&&row.equipmentId),number=milestoneL2Number(row&&row.milestone);if(id&&number&&!truthById.has(id))truthById.set(id,number);}
  const layers={};
  for(const [id,recommendation] of byId){const stat=layers[recommendation.layer]||(layers[recommendation.layer]={tested:0,right:0});stat.tested++;if(recommendation.number&&recommendation.number===truthById.get(id))stat.right++;}
  for(const stat of Object.values(layers))stat.accuracy=stat.tested?Math.round(stat.right/stat.tested*100)/100:0;
  return layers;
}
function candidatesFrom(votes,vocabulary,references){return [...votes.entries()].sort((left,right)=>right[1]-left[1]).slice(0,6).map(([number,support])=>({number,...labelFor(vocabulary,number,references),support,declared:false}));}
